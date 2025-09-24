package reports

import (
	"context"
	"database/sql"
	"testing"
	"time"

	_ "modernc.org/sqlite"
)

func TestKPIsAndExports(t *testing.T) {
	db := openTestDB(t)
	defer db.Close()

	svc := NewService(db)
	start := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
	end := time.Date(2024, 12, 31, 23, 59, 0, 0, time.UTC)

	kpi, err := svc.KPIs(context.Background(), &start, &end)
	if err != nil {
		t.Fatalf("kpi query failed: %v", err)
	}
	if kpi.TotalCustomers != 2 {
		t.Fatalf("expected 2 customers, got %d", kpi.TotalCustomers)
	}
	if kpi.Start3DDue != 1 {
		t.Fatalf("expected 1 order due for 3d check, got %d", kpi.Start3DDue)
	}
	if kpi.Start3DFlagged != 1 {
		t.Fatalf("expected 1 flagged order, got %d", kpi.Start3DFlagged)
	}
	if kpi.PaymentsTotal <= 0 {
		t.Fatalf("expected positive payment total, got %f", kpi.PaymentsTotal)
	}

	if data, err := svc.CustomersCSV(context.Background()); err != nil {
		t.Fatalf("customers csv: %v", err)
	} else if len(data) == 0 {
		t.Fatal("expected customers csv to have content")
	}
	if data, err := svc.OrdersCSV(context.Background()); err != nil {
		t.Fatalf("orders csv: %v", err)
	} else if len(data) == 0 {
		t.Fatal("expected orders csv to have content")
	}
	if data, err := svc.PaymentsCSV(context.Background()); err != nil {
		t.Fatalf("payments csv: %v", err)
	} else if len(data) == 0 {
		t.Fatal("expected payments csv to have content")
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
		`CREATE TABLE customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            business_name TEXT,
            contact_name TEXT,
            phone TEXT,
            email TEXT,
            city TEXT,
            state TEXT,
            zip TEXT,
            created_at DATETIME NOT NULL
        );`,
		`CREATE TABLE sales_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            so_code TEXT NOT NULL,
            status TEXT NOT NULL,
            priority TEXT NOT NULL,
            lead_time_days INTEGER NOT NULL,
            started_at DATETIME,
            created_at DATETIME NOT NULL,
            start3d_due_at DATETIME,
            start3d_flagged_at DATETIME
        );`,
		`CREATE TABLE payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sales_order_id INTEGER,
            date DATETIME NOT NULL,
            method TEXT NOT NULL,
            amount REAL NOT NULL,
            reference TEXT,
            created_at DATETIME NOT NULL
        );`,
	}
	for _, stmt := range schema {
		if _, err := db.Exec(stmt); err != nil {
			t.Fatalf("exec schema: %v", err)
		}
	}

	seedData(t, db)
	return db
}

func seedData(t *testing.T, db *sql.DB) {
	t.Helper()
	now := time.Date(2024, 5, 1, 10, 0, 0, 0, time.UTC)
	if _, err := db.Exec(`INSERT INTO customers (business_name, created_at) VALUES ('A', ?), ('B', ?)`, now, now.AddDate(0, 0, -10)); err != nil {
		t.Fatalf("seed customers: %v", err)
	}
	dueStart := now.AddDate(0, 0, -4)
	flaggedStart := now.AddDate(0, 0, -3)
	if _, err := db.Exec(`INSERT INTO sales_orders (customer_id, so_code, status, priority, lead_time_days, started_at, created_at, start3d_due_at, start3d_flagged_at)
        VALUES (1, 'SO-1', 'Started', 'P1', 10, ?, ?, ?, ?),
               (2, 'SO-2', 'Started', 'P2', 14, ?, ?, ?, NULL)`, flaggedStart, flaggedStart, flaggedStart.Add(72*time.Hour), flaggedStart, dueStart, dueStart, dueStart.Add(72*time.Hour)); err != nil {
		t.Fatalf("seed orders: %v", err)
	}
	if _, err := db.Exec(`INSERT INTO payments (sales_order_id, date, method, amount, reference, created_at)
        VALUES (1, ?, 'Card', 500, 'REF-1', ?)`, now, now); err != nil {
		t.Fatalf("seed payments: %v", err)
	}
}
