package revisions

import (
	"context"
	"database/sql"
	"testing"
	"time"

	_ "modernc.org/sqlite"
)

func TestNormalizeInputDefaults(t *testing.T) {
	normalized, err := normalizeInput(CreateInput{SalesOrderID: 1, FilePath: "files/2024/05/test.pdf"})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if normalized.Status != "pending" {
		t.Fatalf("expected default status pending, got %s", normalized.Status)
	}
	if normalized.Note != "" {
		t.Fatalf("expected empty note, got %s", normalized.Note)
	}
}

func TestServiceCreateAndList(t *testing.T) {
	db := openTestDB(t)
	defer db.Close()

	svc := NewService(db)
	ctx := context.Background()

	res, err := db.ExecContext(ctx, `INSERT INTO sales_orders (customer_id, so_code, status, priority, lead_time_days, created_at)
                VALUES (1, 'SO-1', 'open', 'P2', 28, ?)`, time.Now().UTC())
	if err != nil {
		t.Fatalf("insert sales order: %v", err)
	}
	orderID, err := res.LastInsertId()
	if err != nil {
		t.Fatalf("last insert id: %v", err)
	}

	created, err := svc.Create(ctx, CreateInput{SalesOrderID: orderID, FilePath: "files/2024/05/test.pdf", Note: "rev", Status: "approved"})
	if err != nil {
		t.Fatalf("create revision: %v", err)
	}
	if created.ID == 0 {
		t.Fatal("expected revision id to be set")
	}
	if created.Status != "approved" {
		t.Fatalf("expected status to remain 'approved', got %s", created.Status)
	}

	revisions, err := svc.ListBySalesOrder(ctx, orderID)
	if err != nil {
		t.Fatalf("list revisions: %v", err)
	}
	if len(revisions) != 1 {
		t.Fatalf("expected 1 revision, got %d", len(revisions))
	}
	if revisions[0].ID != created.ID {
		t.Fatalf("expected revision id %d, got %d", created.ID, revisions[0].ID)
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
                        created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
                );`,
		`CREATE TABLE revisions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        sales_order_id INTEGER NOT NULL,
                        note TEXT,
                        file_path TEXT NOT NULL,
                        status TEXT,
                        created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(sales_order_id) REFERENCES sales_orders(id)
                );`,
	}
	for _, stmt := range schema {
		if _, err := db.Exec(stmt); err != nil {
			t.Fatalf("exec schema: %v", err)
		}
	}
	return db
}
