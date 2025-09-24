package payments

import (
	"context"
	"database/sql"
	"errors"
	"fmt"
	"math"
	"sort"
	"strings"
	"time"
)

var (
	// ErrNotFound indicates the payment does not exist.
	ErrNotFound = errors.New("payment not found")
	// ErrValidation indicates invalid input data.
	ErrValidation = errors.New("payment validation error")
)

// Service coordinates payment persistence and allocation logic.
type Service struct {
	db *sql.DB
}

// NewService constructs a Service instance.
func NewService(db *sql.DB) *Service {
	return &Service{db: db}
}

// Payment represents a stored payment with its allocations.
type Payment struct {
	ID           int64        `json:"id"`
	SalesOrderID *int64       `json:"sales_order_id,omitempty"`
	Date         time.Time    `json:"date"`
	Method       string       `json:"method"`
	Amount       float64      `json:"amount"`
	Reference    string       `json:"reference,omitempty"`
	CreatedAt    time.Time    `json:"created_at"`
	Allocations  []Allocation `json:"allocations"`
}

// Allocation represents a payment allocation row.
type Allocation struct {
	ID           int64   `json:"id"`
	PaymentID    int64   `json:"payment_id"`
	SalesOrderID int64   `json:"sales_order_id"`
	Amount       float64 `json:"amount"`
}

// AllocationInput captures user supplied allocations for creation.
type AllocationInput struct {
	SalesOrderID int64   `json:"sales_order_id"`
	Amount       float64 `json:"amount"`
}

// CreateInput captures payload fields for creating a payment.
type CreateInput struct {
	SalesOrderID *int64            `json:"sales_order_id"`
	Date         string            `json:"date"`
	Method       string            `json:"method"`
	Amount       float64           `json:"amount"`
	Reference    string            `json:"reference"`
	Allocations  []AllocationInput `json:"allocations"`
}

// Filter captures list filters for payments.
type Filter struct {
	SalesOrderID int64
	From         *time.Time
	To           *time.Time
}

const epsilon = 0.00001

// Create records a payment and deterministically allocates its amount.
func (s *Service) Create(ctx context.Context, input CreateInput) (*Payment, error) {
	cleaned, err := normalizeCreateInput(input)
	if err != nil {
		return nil, err
	}

	tx, err := s.db.BeginTx(ctx, nil)
	if err != nil {
		return nil, fmt.Errorf("begin tx: %w", err)
	}
	defer tx.Rollback()

	outstandingMap, err := s.fetchOutstanding(ctx, tx)
	if err != nil {
		return nil, err
	}

	explicit, explicitTotal, err := s.applyExplicitAllocations(cleaned.Allocations, outstandingMap)
	if err != nil {
		return nil, err
	}

	remaining := cleaned.Amount - explicitTotal
	if remaining < -epsilon {
		return nil, fmt.Errorf("%w: allocations exceed payment amount", ErrValidation)
	}

	autoAlloc := autoAllocate(remaining, outstandingMap)
	finalAllocations := append(explicit, autoAlloc...)

	var salesOrderID any
	if cleaned.SalesOrderID != nil && *cleaned.SalesOrderID > 0 {
		salesOrderID = *cleaned.SalesOrderID
	}

	res, err := tx.ExecContext(ctx, `INSERT INTO payments (sales_order_id, date, method, amount, reference) VALUES (?, ?, ?, ?, ?)`,
		salesOrderID, cleaned.Date, cleaned.Method, cleaned.Amount, cleaned.Reference)
	if err != nil {
		return nil, fmt.Errorf("insert payment: %w", err)
	}
	paymentID, err := res.LastInsertId()
	if err != nil {
		return nil, fmt.Errorf("last insert id: %w", err)
	}

	for _, alloc := range finalAllocations {
		if alloc.Amount <= epsilon {
			continue
		}
		if _, err := tx.ExecContext(ctx, `INSERT INTO allocations (payment_id, sales_order_id, amount) VALUES (?, ?, ?)`,
			paymentID, alloc.SalesOrderID, alloc.Amount); err != nil {
			return nil, fmt.Errorf("insert allocation: %w", err)
		}
	}

	if err := tx.Commit(); err != nil {
		return nil, fmt.Errorf("commit payment: %w", err)
	}

	return s.Get(ctx, paymentID)
}

// List returns payments matching the provided filter.
func (s *Service) List(ctx context.Context, filter Filter) ([]Payment, error) {
	base := `SELECT id, sales_order_id, date, method, amount, reference, created_at FROM payments`
	var clauses []string
	var args []any
	if filter.SalesOrderID > 0 {
		clauses = append(clauses, `(sales_order_id = ? OR id IN (SELECT payment_id FROM allocations WHERE sales_order_id = ?))`)
		args = append(args, filter.SalesOrderID, filter.SalesOrderID)
	}
	if filter.From != nil {
		clauses = append(clauses, "date >= ?")
		args = append(args, filter.From.Format(time.RFC3339))
	}
	if filter.To != nil {
		clauses = append(clauses, "date <= ?")
		args = append(args, filter.To.Format(time.RFC3339))
	}
	if len(clauses) > 0 {
		base += " WHERE " + strings.Join(clauses, " AND ")
	}
	base += " ORDER BY date DESC, id DESC LIMIT 200"

	rows, err := s.db.QueryContext(ctx, base, args...)
	if err != nil {
		return nil, fmt.Errorf("query payments: %w", err)
	}
	defer rows.Close()

	var payments []Payment
	var ids []int64
	for rows.Next() {
		var p Payment
		var so sql.NullInt64
		var date time.Time
		if err := rows.Scan(&p.ID, &so, &date, &p.Method, &p.Amount, &p.Reference, &p.CreatedAt); err != nil {
			return nil, fmt.Errorf("scan payment: %w", err)
		}
		p.Date = date
		if so.Valid {
			p.SalesOrderID = &so.Int64
		}
		payments = append(payments, p)
		ids = append(ids, p.ID)
	}
	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("iterate payments: %w", err)
	}
	if len(ids) == 0 {
		return payments, nil
	}

	allocations, err := s.fetchAllocations(ctx, ids)
	if err != nil {
		return nil, err
	}
	for i := range payments {
		payments[i].Allocations = allocations[payments[i].ID]
	}
	return payments, nil
}

// Get fetches a payment by id.
func (s *Service) Get(ctx context.Context, id int64) (*Payment, error) {
	row := s.db.QueryRowContext(ctx, `SELECT id, sales_order_id, date, method, amount, reference, created_at FROM payments WHERE id = ?`, id)
	var p Payment
	var so sql.NullInt64
	var date time.Time
	if err := row.Scan(&p.ID, &so, &date, &p.Method, &p.Amount, &p.Reference, &p.CreatedAt); err != nil {
		if errors.Is(err, sql.ErrNoRows) {
			return nil, ErrNotFound
		}
		return nil, fmt.Errorf("scan payment: %w", err)
	}
	p.Date = date
	if so.Valid {
		p.SalesOrderID = &so.Int64
	}
	allocations, err := s.fetchAllocations(ctx, []int64{id})
	if err != nil {
		return nil, err
	}
	p.Allocations = allocations[id]
	return &p, nil
}

func (s *Service) fetchAllocations(ctx context.Context, paymentIDs []int64) (map[int64][]Allocation, error) {
	if len(paymentIDs) == 0 {
		return make(map[int64][]Allocation), nil
	}
	placeholders := make([]string, len(paymentIDs))
	args := make([]any, len(paymentIDs))
	for i, id := range paymentIDs {
		placeholders[i] = "?"
		args[i] = id
	}
	query := fmt.Sprintf(`SELECT id, payment_id, sales_order_id, amount FROM allocations WHERE payment_id IN (%s) ORDER BY id`, strings.Join(placeholders, ","))
	rows, err := s.db.QueryContext(ctx, query, args...)
	if err != nil {
		return nil, fmt.Errorf("query allocations: %w", err)
	}
	defer rows.Close()

	result := make(map[int64][]Allocation)
	for rows.Next() {
		var a Allocation
		if err := rows.Scan(&a.ID, &a.PaymentID, &a.SalesOrderID, &a.Amount); err != nil {
			return nil, fmt.Errorf("scan allocation: %w", err)
		}
		result[a.PaymentID] = append(result[a.PaymentID], a)
	}
	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("iterate allocations: %w", err)
	}
	return result, nil
}

func normalizeCreateInput(input CreateInput) (CreateInputNormalized, error) {
	input.Method = strings.TrimSpace(input.Method)
	if input.Method == "" {
		return CreateInputNormalized{}, fmt.Errorf("%w: method is required", ErrValidation)
	}
	if input.Amount <= 0 {
		return CreateInputNormalized{}, fmt.Errorf("%w: amount must be positive", ErrValidation)
	}
	date := time.Now().UTC()
	if strings.TrimSpace(input.Date) != "" {
		parsed, err := parseDate(input.Date)
		if err != nil {
			return CreateInputNormalized{}, err
		}
		date = parsed
	}
	ref := strings.TrimSpace(input.Reference)
	allocations := make([]AllocationInput, 0, len(input.Allocations))
	for _, alloc := range input.Allocations {
		if alloc.SalesOrderID <= 0 {
			return CreateInputNormalized{}, fmt.Errorf("%w: allocation sales_order_id required", ErrValidation)
		}
		if alloc.Amount <= 0 {
			return CreateInputNormalized{}, fmt.Errorf("%w: allocation amount must be positive", ErrValidation)
		}
		allocations = append(allocations, AllocationInput{SalesOrderID: alloc.SalesOrderID, Amount: roundCurrency(alloc.Amount)})
	}
	var soID *int64
	if input.SalesOrderID != nil && *input.SalesOrderID > 0 {
		soID = input.SalesOrderID
	}
	return CreateInputNormalized{
		SalesOrderID: soID,
		Date:         date,
		Method:       input.Method,
		Amount:       roundCurrency(input.Amount),
		Reference:    ref,
		Allocations:  allocations,
	}, nil
}

// CreateInputNormalized represents sanitized fields for persistence.
type CreateInputNormalized struct {
	SalesOrderID *int64
	Date         time.Time
	Method       string
	Amount       float64
	Reference    string
	Allocations  []AllocationInput
}

func parseDate(value string) (time.Time, error) {
	trimmed := strings.TrimSpace(value)
	layouts := []string{
		time.RFC3339,
		"2006-01-02",
	}
	var err error
	for _, layout := range layouts {
		if t, parseErr := time.Parse(layout, trimmed); parseErr == nil {
			return t.UTC(), nil
		} else {
			err = parseErr
		}
	}
	if err != nil {
		return time.Time{}, fmt.Errorf("%w: invalid date", ErrValidation)
	}
	return time.Time{}, fmt.Errorf("%w: invalid date", ErrValidation)
}

func (s *Service) fetchOutstanding(ctx context.Context, tx *sql.Tx) (map[int64]OutstandingBalance, error) {
	rows, err := tx.QueryContext(ctx, `SELECT so.id, so.created_at,
        COALESCE(SUM(CASE WHEN d.doc_type IN ('Deposit Invoice','Sales Invoice') THEN d.amount
                          WHEN d.doc_type = 'Credit' THEN -d.amount ELSE 0 END), 0) -
        COALESCE((SELECT SUM(a.amount) FROM allocations a WHERE a.sales_order_id = so.id), 0)
        AS outstanding
        FROM sales_orders so
        LEFT JOIN documents d ON d.sales_order_id = so.id
        GROUP BY so.id`)
	if err != nil {
		return nil, fmt.Errorf("query outstanding: %w", err)
	}
	defer rows.Close()

	result := make(map[int64]OutstandingBalance)
	for rows.Next() {
		var ob OutstandingBalance
		if err := rows.Scan(&ob.SalesOrderID, &ob.CreatedAt, &ob.Outstanding); err != nil {
			return nil, fmt.Errorf("scan outstanding: %w", err)
		}
		result[ob.SalesOrderID] = ob
	}
	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("iterate outstanding: %w", err)
	}
	return result, nil
}

// OutstandingBalance tracks remaining balance for an order.
type OutstandingBalance struct {
	SalesOrderID int64
	Outstanding  float64
	CreatedAt    time.Time
}

func (s *Service) applyExplicitAllocations(inputs []AllocationInput, outstanding map[int64]OutstandingBalance) ([]AllocationInput, float64, error) {
	var total float64
	result := make([]AllocationInput, 0, len(inputs))
	for _, alloc := range inputs {
		bal, ok := outstanding[alloc.SalesOrderID]
		if !ok {
			return nil, 0, fmt.Errorf("%w: sales order %d has no outstanding balance", ErrValidation, alloc.SalesOrderID)
		}
		if bal.Outstanding > epsilon && bal.Outstanding+epsilon < alloc.Amount {
			return nil, 0, fmt.Errorf("%w: allocation exceeds outstanding for order %d", ErrValidation, alloc.SalesOrderID)
		}
		total += alloc.Amount
		bal.Outstanding = roundCurrency(bal.Outstanding - alloc.Amount)
		outstanding[alloc.SalesOrderID] = bal
		result = append(result, AllocationInput{SalesOrderID: alloc.SalesOrderID, Amount: alloc.Amount})
	}
	return result, total, nil
}

func autoAllocate(amount float64, outstanding map[int64]OutstandingBalance) []AllocationInput {
	remaining := roundCurrency(amount)
	if remaining <= epsilon {
		return nil
	}
	candidates := make([]OutstandingBalance, 0, len(outstanding))
	for _, bal := range outstanding {
		if bal.Outstanding > epsilon {
			candidates = append(candidates, bal)
		}
	}
	sort.SliceStable(candidates, func(i, j int) bool {
		if candidates[i].CreatedAt.Equal(candidates[j].CreatedAt) {
			return candidates[i].SalesOrderID < candidates[j].SalesOrderID
		}
		return candidates[i].CreatedAt.Before(candidates[j].CreatedAt)
	})

	var allocations []AllocationInput
	for _, candidate := range candidates {
		if remaining <= epsilon {
			break
		}
		toAlloc := math.Min(candidate.Outstanding, remaining)
		toAlloc = roundCurrency(toAlloc)
		if toAlloc <= epsilon {
			continue
		}
		allocations = append(allocations, AllocationInput{SalesOrderID: candidate.SalesOrderID, Amount: toAlloc})
		remaining = roundCurrency(remaining - toAlloc)
	}
	return allocations
}

func roundCurrency(value float64) float64 {
	return math.Round(value*100) / 100
}
