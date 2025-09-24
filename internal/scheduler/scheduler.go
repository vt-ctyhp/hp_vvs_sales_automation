package scheduler

import (
	"context"
	"time"

	"github.com/example/vvsapp/internal/jobs"
	"github.com/example/vvsapp/internal/logging"
)

// Scheduler triggers background jobs on an interval.
type Scheduler struct {
	interval time.Duration
	jobs     *jobs.Service
	logger   *logging.Logger
}

// New constructs a scheduler instance.
func New(interval time.Duration, jobsSvc *jobs.Service, logger *logging.Logger) *Scheduler {
	if interval <= 0 {
		interval = 5 * time.Minute
	}
	return &Scheduler{interval: interval, jobs: jobsSvc, logger: logger}
}

// Start launches the scheduler loop until the context is done.
func (s *Scheduler) Start(ctx context.Context) {
	if s == nil || s.jobs == nil {
		return
	}
	go s.loop(ctx)
}

func (s *Scheduler) loop(ctx context.Context) {
	ticker := time.NewTicker(s.interval)
	defer ticker.Stop()

	s.runOnce(ctx)

	for {
		select {
		case <-ctx.Done():
			return
		case <-ticker.C:
			s.runOnce(ctx)
		}
	}
}

func (s *Scheduler) runOnce(ctx context.Context) {
	if s.jobs == nil {
		return
	}
	res, err := s.jobs.Run(ctx)
	fields := map[string]any{
		"start3d_due_updated": res.Start3DDueUpdated,
		"start3d_flagged":     res.Start3DFlagged,
	}
	if err != nil {
		fields["error"] = err.Error()
		s.logger.Error("scheduler_tick_failed", fields)
		return
	}
	s.logger.Info("scheduler_tick_complete", fields)
}
