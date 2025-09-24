package revisions

import (
	"context"
	"database/sql"
	"errors"
	"fmt"
	"strings"
	"time"
)

var (
	// ErrNotFound indicates a revision could not be located.
	ErrNotFound = errors.New("revision not found")
	// ErrValidation indicates inputs failed validation rules.
	ErrValidation = errors.New("revision validation error")
)

// Service persists revision metadata.
type Service struct {
	db *sql.DB
}

// NewService constructs a Service instance.
func NewService(db *sql.DB) *Service {
	return &Service{db: db}
}

// Revision represents an uploaded file tied to a sales order.
type Revision struct {
	ID           int64     `json:"id"`
	SalesOrderID int64     `json:"sales_order_id"`
	Note         string    `json:"note"`
	FilePath     string    `json:"file_path"`
	Status       string    `json:"status"`
	CreatedAt    time.Time `json:"created_at"`
}

// CreateInput defines fields required when persisting a revision.
type CreateInput struct {
	SalesOrderID int64
	Note         string
	Status       string
	FilePath     string
}

// Create stores metadata for an uploaded revision file.
func (s *Service) Create(ctx context.Context, input CreateInput) (*Revision, error) {
	normalized, err := normalizeInput(input)
	if err != nil {
		return nil, err
	}

	res, err := s.db.ExecContext(ctx, `INSERT INTO revisions (sales_order_id, note, file_path, status)
                VALUES (?, ?, ?, ?)`, normalized.SalesOrderID, normalized.Note, normalized.FilePath, normalized.Status)
	if err != nil {
		return nil, fmt.Errorf("insert revision: %w", err)
	}
	id, err := res.LastInsertId()
	if err != nil {
		return nil, fmt.Errorf("last insert id: %w", err)
	}
	return s.Get(ctx, id)
}

// ListBySalesOrder returns revisions linked to a specific sales order.
func (s *Service) ListBySalesOrder(ctx context.Context, salesOrderID int64) ([]Revision, error) {
	if salesOrderID <= 0 {
		return nil, fmt.Errorf("list revisions: %w", ErrValidation)
	}
	rows, err := s.db.QueryContext(ctx, `SELECT id, sales_order_id, note, file_path, status, created_at
                FROM revisions WHERE sales_order_id = ? ORDER BY created_at DESC, id DESC`, salesOrderID)
	if err != nil {
		return nil, fmt.Errorf("query revisions: %w", err)
	}
	defer rows.Close()

	var results []Revision
	for rows.Next() {
		var rev Revision
		if err := rows.Scan(&rev.ID, &rev.SalesOrderID, &rev.Note, &rev.FilePath, &rev.Status, &rev.CreatedAt); err != nil {
			return nil, fmt.Errorf("scan revision: %w", err)
		}
		results = append(results, rev)
	}
	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("iterate revisions: %w", err)
	}
	return results, nil
}

// Get fetches a single revision by ID.
func (s *Service) Get(ctx context.Context, id int64) (*Revision, error) {
	var rev Revision
	row := s.db.QueryRowContext(ctx, `SELECT id, sales_order_id, note, file_path, status, created_at
                FROM revisions WHERE id = ?`, id)
	if err := row.Scan(&rev.ID, &rev.SalesOrderID, &rev.Note, &rev.FilePath, &rev.Status, &rev.CreatedAt); err != nil {
		if errors.Is(err, sql.ErrNoRows) {
			return nil, ErrNotFound
		}
		return nil, fmt.Errorf("scan revision: %w", err)
	}
	return &rev, nil
}

func normalizeInput(input CreateInput) (CreateInput, error) {
	if input.SalesOrderID <= 0 {
		return CreateInput{}, fmt.Errorf("sales_order_id: %w", ErrValidation)
	}
	filePath := strings.TrimSpace(input.FilePath)
	if filePath == "" {
		return CreateInput{}, fmt.Errorf("file_path: %w", ErrValidation)
	}
	status := strings.TrimSpace(input.Status)
	if status == "" {
		status = "pending"
	}
	note := strings.TrimSpace(input.Note)
	return CreateInput{
		SalesOrderID: input.SalesOrderID,
		Note:         note,
		Status:       status,
		FilePath:     filePath,
	}, nil
}
