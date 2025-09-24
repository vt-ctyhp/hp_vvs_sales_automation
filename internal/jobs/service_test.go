package jobs

import (
	"context"
	"database/sql"
	"testing"
	"time"

	_ "modernc.org/sqlite"
)

func TestRunStart3DCheck(t *testing.T) {
	db := openTestDB(t)
	defer db.Close()

	now := time.Now().UTC()
	started := now.AddDate(0, 0, -4)

	if _, err := db.Exec(`INSERT INTO sales_orders (customer_id, so_code, status, priority, lead_time_days, started_at, created_at)
        VALUES (1, 'SO-1', 'In Progress', 'P1', 14, ?, ?)`, started, started); err != nil {
		t.Fatalf("insert sales order: %v", err)
	}

	svc := NewService(db)
	res, err := svc.Run(context.Background())
	if err != nil {
		t.Fatalf("run scheduler: %v", err)
	}
	if res.Start3DDueUpdated == 0 {
		t.Fatalf("expected due update count > 0, got %d", res.Start3DDueUpdated)
	}
	if res.Start3DFlagged != 1 {
		t.Fatalf("expected one flagged order, got %d", res.Start3DFlagged)
	}

	var flaggedAt sql.NullTime
	var dueAt sql.NullTime
	row := db.QueryRow(`SELECT start3d_due_at, start3d_flagged_at FROM sales_orders WHERE so_code = 'SO-1'`)
	if err := row.Scan(&dueAt, &flaggedAt); err != nil {
		t.Fatalf("scan order: %v", err)
	}
	if !dueAt.Valid {
		t.Fatalf("expected due_at to be set, got %v", dueAt)
	}
	expectedDue := started.AddDate(0, 0, 3)
	if dueAt.Time.Before(expectedDue.Add(-time.Second)) || dueAt.Time.After(expectedDue.Add(time.Second)) {
		t.Fatalf("expected due_at near %v, got %v", expectedDue, dueAt.Time)
	}
	if !flaggedAt.Valid {
		t.Fatal("expected flagged_at to be set")
	}

	// Idempotent run should not change flagged count
	res2, err := svc.Run(context.Background())
	if err != nil {
		t.Fatalf("second run failed: %v", err)
	}
	if res2.Start3DFlagged != 0 {
		t.Fatalf("expected no additional flagging, got %d", res2.Start3DFlagged)
	}

	var jobCount int
	if err := db.QueryRow(`SELECT COUNT(*) FROM jobs`).Scan(&jobCount); err != nil {
		t.Fatalf("count jobs: %v", err)
	}
	if jobCount != 2 {
		t.Fatalf("expected 2 job records, got %d", jobCount)
	}
}

func openTestDB(t *testing.T) *sql.DB {
	t.Helper()
	db, err := sql.Open("sqlite", ":memory:")
	if err != nil {
		t.Fatalf("open sqlite: %v", err)
	}
	if _, err := db.Exec(`PRAGMA foreign_keys = ON`); err != nil {
		t.Fatalf("enable foreign keys: %v", err)
	}
	schema := []string{
		`CREATE TABLE sales_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            so_code TEXT NOT NULL,
            status TEXT NOT NULL,
            priority TEXT NOT NULL,
            lead_time_days INTEGER NOT NULL,
            started_at DATETIME,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
            start3d_due_at DATETIME,
            start3d_flagged_at DATETIME
        );`,
		`CREATE TABLE jobs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            payload_json TEXT,
            status TEXT NOT NULL,
            attempts INTEGER NOT NULL DEFAULT 0,
            run_at DATETIME NOT NULL,
            last_error TEXT,
            updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
        );`,
	}
	for _, stmt := range schema {
		if _, err := db.Exec(stmt); err != nil {
			t.Fatalf("exec schema: %v", err)
		}
	}
	return db
}
