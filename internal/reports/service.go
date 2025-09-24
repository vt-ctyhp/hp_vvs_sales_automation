package reports

import (
	"context"
	"database/sql"
	"encoding/csv"
	"fmt"
	"strings"
	"time"
)

// Service exposes reporting queries and exports.
type Service struct {
	db *sql.DB
}

// KPIResult aggregates summary metrics.
type KPIResult struct {
	TotalCustomers   int     `json:"total_customers"`
	NewCustomers     int     `json:"new_customers"`
	TotalSalesOrders int     `json:"total_sales_orders"`
	OrdersStarted    int     `json:"orders_started"`
	PaymentsTotal    float64 `json:"payments_total"`
	Start3DFlagged   int     `json:"start3d_flagged"`
	Start3DDue       int     `json:"start3d_due"`
}

// NewService constructs a report service.
func NewService(db *sql.DB) *Service {
	return &Service{db: db}
}

// KPIs calculates summary metrics between optional start/end bounds.
func (s *Service) KPIs(ctx context.Context, start, end *time.Time) (KPIResult, error) {
	var res KPIResult

	if err := s.db.QueryRowContext(ctx, `SELECT COUNT(*) FROM customers`).Scan(&res.TotalCustomers); err != nil {
		return KPIResult{}, fmt.Errorf("count customers: %w", err)
	}
	if err := s.db.QueryRowContext(ctx, `SELECT COUNT(*) FROM sales_orders`).Scan(&res.TotalSalesOrders); err != nil {
		return KPIResult{}, fmt.Errorf("count sales orders: %w", err)
	}

	now := time.Now().UTC()
	if start != nil {
		endTime := now
		if end != nil {
			endTime = end.UTC()
		}
		query := `SELECT COUNT(*) FROM customers WHERE created_at >= ? AND created_at <= ?`
		if err := s.db.QueryRowContext(ctx, query, start.UTC(), endTime).Scan(&res.NewCustomers); err != nil {
			return KPIResult{}, fmt.Errorf("count new customers: %w", err)
		}

		query = `SELECT COUNT(*) FROM sales_orders WHERE started_at IS NOT NULL AND started_at >= ? AND started_at <= ?`
		if err := s.db.QueryRowContext(ctx, query, start.UTC(), endTime).Scan(&res.OrdersStarted); err != nil {
			return KPIResult{}, fmt.Errorf("count orders started: %w", err)
		}

		query = `SELECT COALESCE(SUM(amount), 0) FROM payments WHERE date >= ? AND date <= ?`
		if err := s.db.QueryRowContext(ctx, query, start.UTC(), endTime).Scan(&res.PaymentsTotal); err != nil {
			return KPIResult{}, fmt.Errorf("sum payments: %w", err)
		}
	} else {
		query := `SELECT COUNT(*) FROM sales_orders WHERE started_at IS NOT NULL`
		if err := s.db.QueryRowContext(ctx, query).Scan(&res.OrdersStarted); err != nil {
			return KPIResult{}, fmt.Errorf("count orders started total: %w", err)
		}
		query = `SELECT COALESCE(SUM(amount), 0) FROM payments`
		if err := s.db.QueryRowContext(ctx, query).Scan(&res.PaymentsTotal); err != nil {
			return KPIResult{}, fmt.Errorf("sum payments total: %w", err)
		}
	}

	if err := s.db.QueryRowContext(ctx, `SELECT COUNT(*) FROM sales_orders WHERE start3d_flagged_at IS NOT NULL`).Scan(&res.Start3DFlagged); err != nil {
		return KPIResult{}, fmt.Errorf("count flagged: %w", err)
	}
	if err := s.db.QueryRowContext(ctx, `SELECT COUNT(*) FROM sales_orders WHERE start3d_due_at IS NOT NULL AND start3d_flagged_at IS NULL AND start3d_due_at <= ?`, now).Scan(&res.Start3DDue); err != nil {
		return KPIResult{}, fmt.Errorf("count start3d due: %w", err)
	}

	return res, nil
}

// CustomersCSV exports customer data.
func (s *Service) CustomersCSV(ctx context.Context) ([]byte, error) {
	rows, err := s.db.QueryContext(ctx, `SELECT id, business_name, contact_name, phone, email, city, state, zip, created_at FROM customers ORDER BY created_at DESC`)
	if err != nil {
		return nil, fmt.Errorf("query customers: %w", err)
	}
	defer rows.Close()
	headers := []string{"ID", "Business Name", "Contact Name", "Phone", "Email", "City", "State", "ZIP", "Created At"}
	return writeCSV(rows, headers)
}

// OrdersCSV exports sales orders with Start3D fields.
func (s *Service) OrdersCSV(ctx context.Context) ([]byte, error) {
	rows, err := s.db.QueryContext(ctx, `SELECT id, customer_id, so_code, status, priority, lead_time_days, started_at, start3d_due_at, start3d_flagged_at, created_at FROM sales_orders ORDER BY created_at DESC`)
	if err != nil {
		return nil, fmt.Errorf("query orders: %w", err)
	}
	defer rows.Close()
	headers := []string{"ID", "Customer ID", "SO Code", "Status", "Priority", "Lead Time Days", "Started At", "Start3D Due At", "Start3D Flagged At", "Created At"}
	return writeCSV(rows, headers)
}

// PaymentsCSV exports payment history.
func (s *Service) PaymentsCSV(ctx context.Context) ([]byte, error) {
	rows, err := s.db.QueryContext(ctx, `SELECT id, sales_order_id, date, method, amount, reference, created_at FROM payments ORDER BY date DESC, id DESC`)
	if err != nil {
		return nil, fmt.Errorf("query payments: %w", err)
	}
	defer rows.Close()
	headers := []string{"ID", "Sales Order ID", "Date", "Method", "Amount", "Reference", "Created At"}
	return writeCSV(rows, headers)
}

func writeCSV(rows *sql.Rows, headers []string) ([]byte, error) {
	var buf strings.Builder
	writer := csv.NewWriter(&buf)
	if err := writer.Write(headers); err != nil {
		return nil, err
	}
	cols, err := rows.Columns()
	if err != nil {
		return nil, err
	}
	values := make([]any, len(cols))
	scanArgs := make([]any, len(cols))
	for i := range values {
		scanArgs[i] = &values[i]
	}
	for rows.Next() {
		if err := rows.Scan(scanArgs...); err != nil {
			return nil, err
		}
		record := make([]string, len(values))
		for i, v := range values {
			if v == nil {
				record[i] = ""
				continue
			}
			switch val := v.(type) {
			case time.Time:
				record[i] = val.UTC().Format(time.RFC3339)
			case []byte:
				record[i] = string(val)
			default:
				record[i] = fmt.Sprint(val)
			}
		}
		if err := writer.Write(record); err != nil {
			return nil, err
		}
	}
	if err := rows.Err(); err != nil {
		return nil, err
	}
	writer.Flush()
	if err := writer.Error(); err != nil {
		return nil, err
	}
	return []byte(buf.String()), nil
}
