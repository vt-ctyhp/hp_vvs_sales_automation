package jobs

import (
	"context"
	"database/sql"
	"encoding/json"
	"fmt"
	"time"
)

// Service coordinates background job execution.
type Service struct {
	db *sql.DB
}

// Result captures the effects of a scheduler run.
type Result struct {
	Start3DDueUpdated int `json:"start3d_due_updated"`
	Start3DFlagged    int `json:"start3d_flagged"`
}

// NewService constructs a job service.
func NewService(db *sql.DB) *Service {
	return &Service{db: db}
}

// Run executes all configured jobs once.
func (s *Service) Run(ctx context.Context) (Result, error) {
	now := time.Now().UTC()
	res, err := s.runStart3DCheck(ctx, now)
	status := "completed"
	var lastErr string
	if err != nil {
		status = "failed"
		lastErr = err.Error()
	}
	payload, _ := json.Marshal(res)
	if logErr := s.recordRun(ctx, "start3d_check", status, payload, now, lastErr); logErr != nil {
		if err != nil {
			return res, fmt.Errorf("record run: %w (original: %v)", logErr, err)
		}
		return res, fmt.Errorf("record run: %w", logErr)
	}
	return res, err
}

func (s *Service) runStart3DCheck(ctx context.Context, now time.Time) (Result, error) {
	tx, err := s.db.BeginTx(ctx, nil)
	if err != nil {
		return Result{}, fmt.Errorf("begin tx: %w", err)
	}
	defer func() {
		_ = tx.Rollback()
	}()

	rows, err := tx.QueryContext(ctx, `SELECT id, started_at, start3d_due_at, start3d_flagged_at FROM sales_orders WHERE started_at IS NOT NULL`)
	if err != nil {
		return Result{}, fmt.Errorf("select start3d candidates: %w", err)
	}
	defer rows.Close()

	var dueUpdated, flaggedCount int
	for rows.Next() {
		var (
			id      int64
			started time.Time
			due     sql.NullTime
			flagged sql.NullTime
		)
		if err := rows.Scan(&id, &started, &due, &flagged); err != nil {
			return Result{}, fmt.Errorf("scan start3d candidate: %w", err)
		}
		dueTarget := started.Add(72 * time.Hour)
		if !due.Valid || !due.Time.Equal(dueTarget) {
			if _, err := tx.ExecContext(ctx, `UPDATE sales_orders SET start3d_due_at = ? WHERE id = ?`, dueTarget, id); err != nil {
				return Result{}, fmt.Errorf("update due_at: %w", err)
			}
			dueUpdated++
		}
		if flagged.Valid {
			continue
		}
		if !now.Before(dueTarget) {
			if _, err := tx.ExecContext(ctx, `UPDATE sales_orders SET start3d_flagged_at = ? WHERE id = ?`, now, id); err != nil {
				return Result{}, fmt.Errorf("update flagged_at: %w", err)
			}
			flaggedCount++
		}
	}
	if err := rows.Err(); err != nil {
		return Result{}, fmt.Errorf("iterate start3d candidates: %w", err)
	}

	if err := tx.Commit(); err != nil {
		return Result{}, fmt.Errorf("commit start3d check: %w", err)
	}

	return Result{Start3DDueUpdated: dueUpdated, Start3DFlagged: flaggedCount}, nil
}

func (s *Service) recordRun(ctx context.Context, jobType, status string, payload []byte, runAt time.Time, lastErr string) error {
	var lastErrVal any
	if lastErr != "" {
		lastErrVal = lastErr
	}
	_, err := s.db.ExecContext(ctx, `INSERT INTO jobs (type, payload_json, status, attempts, run_at, last_error, updated_at)
VALUES (?, ?, ?, 1, ?, ?, CURRENT_TIMESTAMP)`, jobType, string(payload), status, runAt, lastErrVal)
	return err
}
