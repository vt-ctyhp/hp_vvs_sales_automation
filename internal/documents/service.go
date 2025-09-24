package documents

import (
	"bytes"
	"context"
	"database/sql"
	"errors"
	"fmt"
	"html/template"
	"strings"
	"time"

	"github.com/example/vvsapp/internal/rules"
	"github.com/example/vvsapp/internal/storage"
)

var (
	// ErrValidation indicates invalid document input.
	ErrValidation = errors.New("document validation error")
)

var allowedDocTypes = map[string]struct{}{
	"Deposit Invoice": {},
	"Deposit Receipt": {},
	"Sales Invoice":   {},
	"Sales Receipt":   {},
	"Credit":          {},
}

// Service coordinates document persistence and rendering.
type Service struct {
	db      *sql.DB
	storage storage.Adapter
}

// NewService constructs a Service instance.
func NewService(db *sql.DB, storage storage.Adapter) *Service {
	return &Service{db: db, storage: storage}
}

// Document represents a stored document record.
type Document struct {
	ID           int64     `json:"id"`
	SalesOrderID int64     `json:"sales_order_id"`
	DocType      string    `json:"doc_type"`
	Amount       float64   `json:"amount"`
	FilePath     string    `json:"file_path"`
	CreatedAt    time.Time `json:"created_at"`
	URL          string    `json:"url"`
}

// CreateInput captures payload fields for drafting a document.
type CreateInput struct {
	SalesOrderID int64   `json:"sales_order_id"`
	DocType      string  `json:"doc_type"`
	Amount       float64 `json:"amount"`
}

// Create drafts a document, saves the HTML, and records the row.
func (s *Service) Create(ctx context.Context, input CreateInput) (*Document, error) {
	cleaned, err := normalizeInput(input)
	if err != nil {
		return nil, err
	}

	meta, err := s.fetchSalesOrderMeta(ctx, cleaned.SalesOrderID)
	if err != nil {
		return nil, err
	}

	html, err := renderHTML(cleaned, meta)
	if err != nil {
		return nil, err
	}

	fileName := fmt.Sprintf("%s_%d.html", strings.ReplaceAll(strings.ToLower(strings.ReplaceAll(cleaned.DocType, " ", "_")), "__", "_"), cleaned.SalesOrderID)
	path, err := s.storage.Save(ctx, fileName, bytes.NewReader(html))
	if err != nil {
		return nil, fmt.Errorf("store document: %w", err)
	}

	res, err := s.db.ExecContext(ctx, `INSERT INTO documents (sales_order_id, doc_type, amount, file_path) VALUES (?, ?, ?, ?)`,
		cleaned.SalesOrderID, cleaned.DocType, cleaned.Amount, path)
	if err != nil {
		return nil, fmt.Errorf("insert document: %w", err)
	}
	id, err := res.LastInsertId()
	if err != nil {
		return nil, fmt.Errorf("last insert id: %w", err)
	}

	row := s.db.QueryRowContext(ctx, `SELECT id, sales_order_id, doc_type, amount, file_path, created_at FROM documents WHERE id = ?`, id)
	var doc Document
	if err := row.Scan(&doc.ID, &doc.SalesOrderID, &doc.DocType, &doc.Amount, &doc.FilePath, &doc.CreatedAt); err != nil {
		return nil, fmt.Errorf("scan document: %w", err)
	}
	if url, err := s.storage.URL(doc.FilePath); err == nil {
		doc.URL = url
	}
	return &doc, nil
}

func normalizeInput(input CreateInput) (CreateInput, error) {
	if input.SalesOrderID <= 0 {
		return CreateInput{}, fmt.Errorf("%w: sales_order_id is required", ErrValidation)
	}
	input.DocType = strings.TrimSpace(input.DocType)
	if input.DocType == "" {
		return CreateInput{}, fmt.Errorf("%w: doc_type is required", ErrValidation)
	}
	if _, ok := allowedDocTypes[input.DocType]; !ok {
		return CreateInput{}, fmt.Errorf("%w: unsupported doc_type", ErrValidation)
	}
	if input.Amount < 0 {
		return CreateInput{}, fmt.Errorf("%w: amount must be zero or positive", ErrValidation)
	}
	return input, nil
}

type salesOrderMeta struct {
	SOCode       string
	CustomerName string
	CreatedAt    time.Time
}

func (s *Service) fetchSalesOrderMeta(ctx context.Context, id int64) (salesOrderMeta, error) {
	var meta salesOrderMeta
	row := s.db.QueryRowContext(ctx, `SELECT so.so_code, c.business_name, so.created_at
        FROM sales_orders so
        JOIN customers c ON c.id = so.customer_id
        WHERE so.id = ?`, id)
	if err := row.Scan(&meta.SOCode, &meta.CustomerName, &meta.CreatedAt); err != nil {
		if errors.Is(err, sql.ErrNoRows) {
			return salesOrderMeta{}, fmt.Errorf("%w: sales order not found", ErrValidation)
		}
		return salesOrderMeta{}, fmt.Errorf("fetch sales order: %w", err)
	}
	return meta, nil
}

func renderHTML(input CreateInput, meta salesOrderMeta) ([]byte, error) {
	data := struct {
		DocType   string
		SOCode    string
		Customer  string
		Amount    float64
		Shipping  float64
		Total     float64
		Generated string
	}{
		DocType:   input.DocType,
		SOCode:    meta.SOCode,
		Customer:  meta.CustomerName,
		Amount:    input.Amount,
		Shipping:  shippingForDocument(input.DocType, input.Amount),
		Generated: time.Now().UTC().Format(time.RFC1123),
	}
	data.Total = data.Amount + data.Shipping

	buf := &bytes.Buffer{}
	if err := docTemplate.Execute(buf, data); err != nil {
		return nil, fmt.Errorf("render document: %w", err)
	}
	return buf.Bytes(), nil
}

var docTemplate = template.Must(template.New("document").Parse(`<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{{.DocType}} · {{.SOCode}}</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 2rem; }
      header { margin-bottom: 2rem; }
      table { border-collapse: collapse; width: 100%; }
      th, td { border: 1px solid #ccc; padding: 0.75rem; text-align: left; }
      tfoot td { font-weight: bold; }
    </style>
  </head>
  <body>
    <header>
      <h1>{{.DocType}}</h1>
      <p><strong>Sales Order:</strong> {{.SOCode}}</p>
      <p><strong>Customer:</strong> {{.Customer}}</p>
      <p><strong>Generated:</strong> {{.Generated}}</p>
    </header>
    <table>
      <thead>
        <tr>
          <th>Description</th>
          <th>Amount</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>Document Amount</td>
          <td>$ {{printf "%.2f" .Amount}}</td>
        </tr>
        <tr>
          <td>Shipping</td>
          <td>$ {{printf "%.2f" .Shipping}}</td>
        </tr>
      </tbody>
      <tfoot>
        <tr>
          <td>Total</td>
          <td>$ {{printf "%.2f" .Total}}</td>
        </tr>
      </tfoot>
    </table>
  </body>
</html>`))

func shippingForDocument(docType string, subtotal float64) float64 {
	switch docType {
	case "Deposit Invoice", "Sales Invoice":
		return rules.ShippingForSubtotal(subtotal)
	default:
		return 0
	}
}
