package main

import (
	"context"
	"errors"
	"net/http"
	"os"
	"os/signal"
	"syscall"
	"time"

	"github.com/google/uuid"

	"github.com/example/vvsapp/internal/auth"
	"github.com/example/vvsapp/internal/config"
	"github.com/example/vvsapp/internal/db"
	"github.com/example/vvsapp/internal/logging"
	"github.com/example/vvsapp/internal/server"
)

func main() {
	ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt, syscall.SIGINT, syscall.SIGTERM)
	defer stop()

	cfg, err := config.Load("config/app.yaml")
	if err != nil {
		panic(err)
	}

	baseLogger := logging.New(cfg.Logging.Level)
	logger := baseLogger.With(map[string]any{"request_id": uuid.NewString()})
	logger.Info("config_loaded", map[string]any{"config": cfg.Summary()})

	database, err := db.Open(cfg.Database)
	if err != nil {
		logger.Error("database_open_failed", map[string]any{"error": err.Error()})
		os.Exit(1)
	}
	defer database.Close()

	if applied, err := db.RunMigrations(ctx, database); err != nil {
		logger.Error("migrations_failed", map[string]any{"error": err.Error()})
		os.Exit(1)
	} else if len(applied) > 0 {
		logger.Info("migrations_applied", map[string]any{"count": len(applied)})
	}

	if err := auth.SeedAdmin(ctx, database, cfg.Seed, logger); err != nil {
		logger.Error("seed_admin_failed", map[string]any{"error": err.Error()})
		os.Exit(1)
	}

	authSvc, err := auth.NewService(database, cfg.Auth, logger)
	if err != nil {
		logger.Error("auth_service_init_failed", map[string]any{"error": err.Error()})
		os.Exit(1)
	}

	srv := server.New(cfg, logger, database, authSvc)
	httpServer := &http.Server{
		Addr:    cfg.Server.Address,
		Handler: srv,
	}

	serverErr := make(chan error, 1)
	go func() {
		logger.Info("server_listening", map[string]any{"address": cfg.Server.Address})
		if err := httpServer.ListenAndServe(); err != nil && !errors.Is(err, http.ErrServerClosed) {
			serverErr <- err
		}
		close(serverErr)
	}()

	select {
	case <-ctx.Done():
		shutdownCtx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
		defer cancel()
		if err := httpServer.Shutdown(shutdownCtx); err != nil {
			logger.Error("server_shutdown_failed", map[string]any{"error": err.Error()})
		}
	case err := <-serverErr:
		if err != nil {
			logger.Error("server_listen_failed", map[string]any{"error": err.Error()})
		}
	}

	logger.Info("server_stopped", map[string]any{})
}
