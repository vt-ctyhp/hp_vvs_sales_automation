package salesorders

import (
	"context"
	"database/sql"
	"errors"
	"fmt"
	"strings"
	"time"
)

// ErrNotFound indicates a requested sales order does not exist.
var (
	ErrNotFound   = errors.New("sales order not found")
	ErrValidation = errors.New("sales order validation error")
)

// Service coordinates CRUD operations for sales orders.
type Service struct {
	db *sql.DB
}

// NewService constructs the service.
func NewService(db *sql.DB) *Service {
	return &Service{db: db}
}

// SalesOrder represents a persisted order.
type SalesOrder struct {
	ID           int64      `json:"id"`
	CustomerID   int64      `json:"customer_id"`
	SOCode       string     `json:"so_code"`
	Status       string     `json:"status"`
	Priority     string     `json:"priority"`
	LeadTimeDays int        `json:"lead_time_days"`
	StartedAt    *time.Time `json:"started_at,omitempty"`
	CreatedAt    time.Time  `json:"created_at"`
}

// Input represents user-provided fields.
type Input struct {
	CustomerID   int64      `json:"customer_id"`
	SOCode       string     `json:"so_code"`
	Status       string     `json:"status"`
	Priority     string     `json:"priority"`
	LeadTimeDays int        `json:"lead_time_days"`
	StartedAt    *time.Time `json:"started_at"`
}

// UpdateInput captures mutable fields for updates.
type UpdateInput struct {
	Status       *string     `json:"status"`
	Priority     *string     `json:"priority"`
	LeadTimeDays *int        `json:"lead_time_days"`
	StartedAt    **time.Time `json:"started_at"`
}

// List returns sales orders filtered by customer or status.
func (s *Service) List(ctx context.Context, customerID int64, status string) ([]SalesOrder, error) {
	base := `SELECT id, customer_id, so_code, status, priority, lead_time_days, started_at, created_at FROM sales_orders`
	var clauses []string
	var args []any
	if customerID > 0 {
		clauses = append(clauses, "customer_id = ?")
		args = append(args, customerID)
	}
	if status = strings.TrimSpace(status); status != "" {
		clauses = append(clauses, "status = ?")
		args = append(args, status)
	}
	if len(clauses) > 0 {
		base += " WHERE " + strings.Join(clauses, " AND ")
	}
	base += " ORDER BY created_at DESC, id DESC LIMIT 200"

	rows, err := s.db.QueryContext(ctx, base, args...)
	if err != nil {
		return nil, fmt.Errorf("query sales orders: %w", err)
	}
	defer rows.Close()

	var results []SalesOrder
	for rows.Next() {
		var so SalesOrder
		var started sql.NullTime
		if err := rows.Scan(&so.ID, &so.CustomerID, &so.SOCode, &so.Status, &so.Priority, &so.LeadTimeDays, &started, &so.CreatedAt); err != nil {
			return nil, fmt.Errorf("scan sales order: %w", err)
		}
		if started.Valid {
			so.StartedAt = &started.Time
		}
		results = append(results, so)
	}
	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("iterate sales orders: %w", err)
	}
	return results, nil
}

// Create inserts a new sales order.
func (s *Service) Create(ctx context.Context, input Input) (*SalesOrder, error) {
	cleaned, err := normalizeInput(input)
	if err != nil {
		return nil, err
	}

	res, err := s.db.ExecContext(ctx, `INSERT INTO sales_orders (customer_id, so_code, status, priority, lead_time_days, started_at)
                VALUES (?, ?, ?, ?, ?, ?)`, cleaned.CustomerID, cleaned.SOCode, cleaned.Status, cleaned.Priority, cleaned.LeadTimeDays, cleaned.StartedAt)
	if err != nil {
		return nil, fmt.Errorf("insert sales order: %w", err)
	}
	id, err := res.LastInsertId()
	if err != nil {
		return nil, fmt.Errorf("last insert id: %w", err)
	}
	return s.Get(ctx, id)
}

// Update mutates selected fields.
func (s *Service) Update(ctx context.Context, id int64, input UpdateInput) (*SalesOrder, error) {
	cleaned, err := normalizeInputForUpdate(input)
	if err != nil {
		return nil, err
	}

	setClauses := []string{
		"status = COALESCE(?, status)",
		"priority = COALESCE(?, priority)",
		"lead_time_days = COALESCE(?, lead_time_days)",
	}
	args := []any{cleaned.Status, cleaned.Priority, cleaned.LeadTimeDays}
	if cleaned.HasStartedAt {
		setClauses = append(setClauses, "started_at = ?")
		args = append(args, cleaned.StartedAt)
	}
	args = append(args, id)

	query := "UPDATE sales_orders SET " + strings.Join(setClauses, ", ") + " WHERE id = ?"
	_, err = s.db.ExecContext(ctx, query, args...)
	if err != nil {
		return nil, fmt.Errorf("update sales order: %w", err)
	}

	so, err := s.Get(ctx, id)
	if err != nil {
		return nil, err
	}
	return so, nil
}

// Get fetches a single sales order.
func (s *Service) Get(ctx context.Context, id int64) (*SalesOrder, error) {
	var so SalesOrder
	var started sql.NullTime
	row := s.db.QueryRowContext(ctx, `SELECT id, customer_id, so_code, status, priority, lead_time_days, started_at, created_at FROM sales_orders WHERE id = ?`, id)
	if err := row.Scan(&so.ID, &so.CustomerID, &so.SOCode, &so.Status, &so.Priority, &so.LeadTimeDays, &started, &so.CreatedAt); err != nil {
		if errors.Is(err, sql.ErrNoRows) {
			return nil, ErrNotFound
		}
		return nil, fmt.Errorf("scan sales order: %w", err)
	}
	if started.Valid {
		so.StartedAt = &started.Time
	}
	return &so, nil
}

func normalizeInput(input Input) (Input, error) {
	if input.CustomerID <= 0 {
		return Input{}, fmt.Errorf("%w: customer_id is required", ErrValidation)
	}
	input.SOCode = strings.TrimSpace(input.SOCode)
	if input.SOCode == "" {
		return Input{}, fmt.Errorf("%w: so_code is required", ErrValidation)
	}
	input.Status = strings.TrimSpace(input.Status)
	if input.Status == "" {
		return Input{}, fmt.Errorf("%w: status is required", ErrValidation)
	}
	input.Priority = strings.TrimSpace(input.Priority)
	if input.Priority == "" {
		input.Priority = "P2"
	}
	if input.LeadTimeDays == 0 {
		input.LeadTimeDays = 28
	}
	return input, nil
}

type updateValues struct {
	Status       any
	Priority     any
	LeadTimeDays any
	StartedAt    any
	HasStartedAt bool
}

func normalizeInputForUpdate(input UpdateInput) (updateValues, error) {
	var status any
	if input.Status != nil {
		trimmed := strings.TrimSpace(*input.Status)
		if trimmed == "" {
			return updateValues{}, fmt.Errorf("%w: status cannot be empty", ErrValidation)
		}
		status = trimmed
	}

	var priority any
	if input.Priority != nil {
		trimmed := strings.TrimSpace(*input.Priority)
		if trimmed == "" {
			return updateValues{}, fmt.Errorf("%w: priority cannot be empty", ErrValidation)
		}
		priority = trimmed
	}

	var lead any
	if input.LeadTimeDays != nil {
		if *input.LeadTimeDays < 0 {
			return updateValues{}, fmt.Errorf("%w: lead_time_days cannot be negative", ErrValidation)
		}
		lead = *input.LeadTimeDays
	}

	var started any
	hasStarted := false
	if input.StartedAt != nil {
		started = *input.StartedAt
		hasStarted = true
	}

	return updateValues{Status: status, Priority: priority, LeadTimeDays: lead, StartedAt: started, HasStartedAt: hasStarted}, nil
}
