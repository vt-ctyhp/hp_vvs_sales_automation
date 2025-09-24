package db

import (
	"context"
	"database/sql"
	"fmt"
	"sort"
)

type Migration struct {
	Version int
	Name    string
	Up      string
}

var migrations = []Migration{
	{
		Version: 1,
		Name:    "create_users",
		Up: `CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
        );`,
	},
	{
		Version: 2,
		Name:    "create_customers",
		Up: `CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            business_name TEXT NOT NULL,
            contact_name TEXT,
            phone TEXT,
            email TEXT,
            city TEXT,
            state TEXT,
            zip TEXT,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
        );`,
	},
}

// RunMigrations applies pending migrations and returns the versions that were newly applied.
func RunMigrations(ctx context.Context, conn *sql.DB) ([]Migration, error) {
	if err := ensureSchemaTable(ctx, conn); err != nil {
		return nil, err
	}

	applied, err := fetchApplied(ctx, conn)
	if err != nil {
		return nil, err
	}

	sort.Slice(migrations, func(i, j int) bool {
		return migrations[i].Version < migrations[j].Version
	})

	var newlyApplied []Migration
	for _, m := range migrations {
		if applied[m.Version] {
			continue
		}
		if err := applyMigration(ctx, conn, m); err != nil {
			return newlyApplied, err
		}
		newlyApplied = append(newlyApplied, m)
	}

	return newlyApplied, nil
}

func ensureSchemaTable(ctx context.Context, conn *sql.DB) error {
	_, err := conn.ExecContext(ctx, `CREATE TABLE IF NOT EXISTS schema_migrations (
        version INTEGER PRIMARY KEY
    );`)
	if err != nil {
		return fmt.Errorf("ensure schema_migrations: %w", err)
	}
	return nil
}

func fetchApplied(ctx context.Context, conn *sql.DB) (map[int]bool, error) {
	rows, err := conn.QueryContext(ctx, `SELECT version FROM schema_migrations`)
	if err != nil {
		return nil, fmt.Errorf("select schema_migrations: %w", err)
	}
	defer rows.Close()

	applied := make(map[int]bool)
	for rows.Next() {
		var version int
		if err := rows.Scan(&version); err != nil {
			return nil, fmt.Errorf("scan schema_migrations: %w", err)
		}
		applied[version] = true
	}
	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("iterate schema_migrations: %w", err)
	}
	return applied, nil
}

func applyMigration(ctx context.Context, conn *sql.DB, m Migration) error {
	tx, err := conn.BeginTx(ctx, nil)
	if err != nil {
		return fmt.Errorf("begin migration %d: %w", m.Version, err)
	}
	defer func() {
		_ = tx.Rollback()
	}()

	if _, err := tx.ExecContext(ctx, m.Up); err != nil {
		return fmt.Errorf("exec migration %d: %w", m.Version, err)
	}

	if _, err := tx.ExecContext(ctx, `INSERT INTO schema_migrations(version) VALUES(?)`, m.Version); err != nil {
		return fmt.Errorf("record migration %d: %w", m.Version, err)
	}

	if err := tx.Commit(); err != nil {
		return fmt.Errorf("commit migration %d: %w", m.Version, err)
	}

	return nil
}
