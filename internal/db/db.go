package db

import (
	"context"
	"database/sql"
	"fmt"
	"time"

	"github.com/example/vvsapp/internal/config"

	_ "modernc.org/sqlite"
)

// Open establishes a SQLite connection using the provided configuration.
func Open(cfg config.DatabaseConfig) (*sql.DB, error) {
	if cfg.Path == "" {
		return nil, fmt.Errorf("database path is required")
	}

	dsn := cfg.Path
	conn, err := sql.Open("sqlite", dsn)
	if err != nil {
		return nil, fmt.Errorf("open sqlite: %w", err)
	}

	conn.SetConnMaxIdleTime(0)
	conn.SetMaxOpenConns(1)
	conn.SetMaxIdleConns(1)

	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := conn.PingContext(ctx); err != nil {
		conn.Close()
		return nil, fmt.Errorf("ping sqlite: %w", err)
	}

	if _, err := conn.ExecContext(ctx, "PRAGMA foreign_keys = ON"); err != nil {
		conn.Close()
		return nil, fmt.Errorf("enable foreign keys: %w", err)
	}

	return conn, nil
}

// Ping verifies the database connection is alive.
type pinger interface {
	PingContext(ctx context.Context) error
}

func Ping(ctx context.Context, db pinger) error {
	ctx, cancel := context.WithTimeout(ctx, 2*time.Second)
	defer cancel()
	return db.PingContext(ctx)
}
