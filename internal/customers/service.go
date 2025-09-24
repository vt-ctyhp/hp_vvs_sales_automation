package customers

import (
	"context"
	"database/sql"
	"errors"
	"fmt"
	"strings"
	"time"
)

// ErrNotFound indicates a requested customer could not be located.
var (
	ErrNotFound   = errors.New("customer not found")
	ErrValidation = errors.New("customer validation error")
)

// Service provides customer persistence helpers.
type Service struct {
	db *sql.DB
}

// NewService constructs a Service.
func NewService(db *sql.DB) *Service {
	return &Service{db: db}
}

// Customer represents a CRM entry persisted to the database.
type Customer struct {
	ID           int64     `json:"id"`
	BusinessName string    `json:"business_name"`
	ContactName  string    `json:"contact_name"`
	Phone        string    `json:"phone"`
	Email        string    `json:"email"`
	City         string    `json:"city"`
	State        string    `json:"state"`
	ZIP          string    `json:"zip"`
	CreatedAt    time.Time `json:"created_at"`
}

// Input captures the writable fields for creates/updates.
type Input struct {
	BusinessName string `json:"business_name"`
	ContactName  string `json:"contact_name"`
	Phone        string `json:"phone"`
	Email        string `json:"email"`
	City         string `json:"city"`
	State        string `json:"state"`
	ZIP          string `json:"zip"`
}

// List returns customers filtered by a loose query.
func (s *Service) List(ctx context.Context, query string) ([]Customer, error) {
	likeQuery := strings.TrimSpace(query)
	var rows *sql.Rows
	var err error
	if likeQuery == "" {
		rows, err = s.db.QueryContext(ctx, `SELECT id, business_name, contact_name, phone, email, city, state, zip, created_at
                FROM customers ORDER BY created_at DESC, id DESC LIMIT 200`)
	} else {
		likeQuery = "%" + strings.ToLower(likeQuery) + "%"
		rows, err = s.db.QueryContext(ctx, `SELECT id, business_name, contact_name, phone, email, city, state, zip, created_at
                FROM customers
                WHERE lower(business_name) LIKE ? OR lower(contact_name) LIKE ? OR lower(phone) LIKE ? OR lower(email) LIKE ?
                ORDER BY created_at DESC, id DESC LIMIT 200`, likeQuery, likeQuery, likeQuery, likeQuery)
	}
	if err != nil {
		return nil, fmt.Errorf("query customers: %w", err)
	}
	defer rows.Close()

	var results []Customer
	for rows.Next() {
		var c Customer
		if err := rows.Scan(&c.ID, &c.BusinessName, &c.ContactName, &c.Phone, &c.Email, &c.City, &c.State, &c.ZIP, &c.CreatedAt); err != nil {
			return nil, fmt.Errorf("scan customer: %w", err)
		}
		results = append(results, c)
	}
	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("iterate customers: %w", err)
	}
	return results, nil
}

// Create inserts a customer and returns warning fields for soft duplicates.
func (s *Service) Create(ctx context.Context, input Input) (*Customer, []string, error) {
	cleaned, err := normalizeInput(input)
	if err != nil {
		return nil, nil, err
	}

	warn, err := s.findDuplicates(ctx, 0, cleaned)
	if err != nil {
		return nil, nil, err
	}

	res, err := s.db.ExecContext(ctx, `INSERT INTO customers (business_name, contact_name, phone, email, city, state, zip)
                VALUES (?, ?, ?, ?, ?, ?, ?)`, cleaned.BusinessName, cleaned.ContactName, cleaned.Phone, cleaned.Email, cleaned.City, cleaned.State, cleaned.ZIP)
	if err != nil {
		return nil, nil, fmt.Errorf("insert customer: %w", err)
	}
	id, err := res.LastInsertId()
	if err != nil {
		return nil, nil, fmt.Errorf("last insert id: %w", err)
	}

	created, err := s.Get(ctx, id)
	if err != nil {
		return nil, nil, err
	}
	return created, warn, nil
}

// Update writes new values to an existing record and reports warnings.
func (s *Service) Update(ctx context.Context, id int64, input Input) (*Customer, []string, error) {
	cleaned, err := normalizeInput(input)
	if err != nil {
		return nil, nil, err
	}

	if _, err := s.Get(ctx, id); err != nil {
		return nil, nil, err
	}

	warn, err := s.findDuplicates(ctx, id, cleaned)
	if err != nil {
		return nil, nil, err
	}

	_, err = s.db.ExecContext(ctx, `UPDATE customers SET business_name = ?, contact_name = ?, phone = ?, email = ?, city = ?, state = ?, zip = ? WHERE id = ?`,
		cleaned.BusinessName, cleaned.ContactName, cleaned.Phone, cleaned.Email, cleaned.City, cleaned.State, cleaned.ZIP, id)
	if err != nil {
		return nil, nil, fmt.Errorf("update customer: %w", err)
	}

	updated, err := s.Get(ctx, id)
	if err != nil {
		return nil, nil, err
	}
	return updated, warn, nil
}

// Get fetches a customer by ID.
func (s *Service) Get(ctx context.Context, id int64) (*Customer, error) {
	var c Customer
	row := s.db.QueryRowContext(ctx, `SELECT id, business_name, contact_name, phone, email, city, state, zip, created_at FROM customers WHERE id = ?`, id)
	if err := row.Scan(&c.ID, &c.BusinessName, &c.ContactName, &c.Phone, &c.Email, &c.City, &c.State, &c.ZIP, &c.CreatedAt); err != nil {
		if errors.Is(err, sql.ErrNoRows) {
			return nil, ErrNotFound
		}
		return nil, fmt.Errorf("scan customer: %w", err)
	}
	return &c, nil
}

func (s *Service) findDuplicates(ctx context.Context, selfID int64, input Input) ([]string, error) {
	var warns []string
	type check struct {
		field string
		value string
	}
	checks := []check{{"business_name", input.BusinessName}}
	if input.Phone != "" {
		checks = append(checks, check{"phone", input.Phone})
	}
	if input.Email != "" {
		checks = append(checks, check{"email", strings.ToLower(input.Email)})
	}

	for _, chk := range checks {
		query := fmt.Sprintf("SELECT 1 FROM customers WHERE %s = ?", chk.field)
		args := []any{chk.value}
		if chk.field == "email" {
			query = fmt.Sprintf("SELECT 1 FROM customers WHERE lower(%s) = ?", chk.field)
		}
		if selfID > 0 {
			query += " AND id != ?"
			args = append(args, selfID)
		}
		row := s.db.QueryRowContext(ctx, query+" LIMIT 1", args...)
		var exists int
		if err := row.Scan(&exists); err != nil {
			if errors.Is(err, sql.ErrNoRows) {
				continue
			}
			return nil, fmt.Errorf("duplicate check %s: %w", chk.field, err)
		}
		warns = append(warns, chk.field)
	}
	return warns, nil
}

func normalizeInput(in Input) (Input, error) {
	in.BusinessName = strings.TrimSpace(in.BusinessName)
	in.ContactName = strings.TrimSpace(in.ContactName)
	in.Phone = normalizePhone(in.Phone)
	in.Email = strings.TrimSpace(strings.ToLower(in.Email))
	in.City = strings.TrimSpace(in.City)
	in.State = strings.ToUpper(strings.TrimSpace(in.State))
	in.ZIP = strings.TrimSpace(in.ZIP)

	if in.BusinessName == "" {
		return Input{}, fmt.Errorf("%w: business_name is required", ErrValidation)
	}

	return in, nil
}

func normalizePhone(raw string) string {
	raw = strings.TrimSpace(raw)
	if raw == "" {
		return ""
	}
	var b strings.Builder
	for _, r := range raw {
		if r >= '0' && r <= '9' {
			b.WriteRune(r)
		}
	}
	return b.String()
}
