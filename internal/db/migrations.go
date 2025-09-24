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
	{
		Version: 3,
		Name:    "create_inquiries_sales_orders",
		Up: `CREATE TABLE IF NOT EXISTS inquiries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            product_description TEXT,
            status TEXT,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(customer_id) REFERENCES customers(id)
        );

        CREATE TABLE IF NOT EXISTS sales_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            so_code TEXT NOT NULL UNIQUE,
            status TEXT NOT NULL,
            priority TEXT NOT NULL,
            lead_time_days INTEGER NOT NULL,
            started_at DATETIME,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(customer_id) REFERENCES customers(id)
        );

        CREATE INDEX IF NOT EXISTS idx_customers_email ON customers(email);
        CREATE INDEX IF NOT EXISTS idx_customers_phone ON customers(phone);
        CREATE INDEX IF NOT EXISTS idx_sales_orders_so_code ON sales_orders(so_code);
        CREATE INDEX IF NOT EXISTS idx_sales_orders_customer_status ON sales_orders(customer_id, status);
        `,
	},
	{
		Version: 4,
		Name:    "create_revisions",
		Up: `CREATE TABLE IF NOT EXISTS revisions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sales_order_id INTEGER NOT NULL,
            note TEXT,
            file_path TEXT NOT NULL,
            status TEXT,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(sales_order_id) REFERENCES sales_orders(id)
        );

        CREATE INDEX IF NOT EXISTS idx_revisions_sales_order ON revisions(sales_order_id, created_at DESC);
        `,
	},
	{
		Version: 5,
		Name:    "create_payments_allocations_documents",
		Up: `CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sales_order_id INTEGER,
            date DATETIME NOT NULL,
            method TEXT NOT NULL,
            amount REAL NOT NULL,
            reference TEXT,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(sales_order_id) REFERENCES sales_orders(id)
        );

        CREATE TABLE IF NOT EXISTS allocations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            payment_id INTEGER NOT NULL,
            sales_order_id INTEGER NOT NULL,
            amount REAL NOT NULL,
            FOREIGN KEY(payment_id) REFERENCES payments(id) ON DELETE CASCADE,
            FOREIGN KEY(sales_order_id) REFERENCES sales_orders(id)
        );

        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sales_order_id INTEGER NOT NULL,
            doc_type TEXT NOT NULL,
            amount REAL NOT NULL,
            file_path TEXT NOT NULL,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(sales_order_id) REFERENCES sales_orders(id)
        );

        CREATE INDEX IF NOT EXISTS idx_payments_order_date ON payments(sales_order_id, date);
        CREATE INDEX IF NOT EXISTS idx_allocations_payment ON allocations(payment_id);
        CREATE INDEX IF NOT EXISTS idx_allocations_sales_order ON allocations(sales_order_id);
        CREATE INDEX IF NOT EXISTS idx_documents_sales_order ON documents(sales_order_id, created_at DESC);
        `,
	},

	{
		Version: 6,
		Name:    "create_jobs_and_start3d_columns",
		Up: `ALTER TABLE sales_orders ADD COLUMN start3d_due_at DATETIME;
        ALTER TABLE sales_orders ADD COLUMN start3d_flagged_at DATETIME;

        CREATE TABLE IF NOT EXISTS jobs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            payload_json TEXT,
            status TEXT NOT NULL,
            attempts INTEGER NOT NULL DEFAULT 0,
            run_at DATETIME NOT NULL,
            last_error TEXT,
            updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE INDEX IF NOT EXISTS idx_sales_orders_start3d_flagged ON sales_orders(start3d_flagged_at);
        CREATE INDEX IF NOT EXISTS idx_jobs_status_run_at ON jobs(status, run_at);
        `,
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
