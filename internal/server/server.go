package server

import (
	"context"
	"database/sql"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/example/vvsapp/internal/auth"
	"github.com/example/vvsapp/internal/config"
	"github.com/example/vvsapp/internal/customers"
	"github.com/example/vvsapp/internal/db"
	"github.com/example/vvsapp/internal/documents"
	"github.com/example/vvsapp/internal/jobs"
	"github.com/example/vvsapp/internal/logging"
	"github.com/example/vvsapp/internal/payments"
	"github.com/example/vvsapp/internal/reports"
	"github.com/example/vvsapp/internal/revisions"
	"github.com/example/vvsapp/internal/salesorders"
	"github.com/example/vvsapp/internal/storage"
)

// contextKey helps avoid collisions when storing values in request contexts.
type contextKey string

const contextKeyClaims contextKey = "authClaims"

// Server wires HTTP handlers, authentication, and diagnostics.
type Server struct {
	cfg          *config.Config
	logger       *logging.Logger
	authSvc      *auth.Service
	db           *sql.DB
	router       http.Handler
	customerSvc  *customers.Service
	orderSvc     *salesorders.Service
	revisionSvc  *revisions.Service
	paymentSvc   *payments.Service
	documentSvc  *documents.Service
	jobsSvc      *jobs.Service
	reportsSvc   *reports.Service
	storage      storage.Adapter
	staticServer http.Handler
	filesServer  http.Handler
	filesPrefix  string
}

// New constructs a server with routes and middleware applied.
func New(cfg *config.Config, logger *logging.Logger, database *sql.DB, authSvc *auth.Service, storageAdapter storage.Adapter, jobsSvc *jobs.Service, reportsSvc *reports.Service) *Server {
	srv := &Server{
		cfg:          cfg,
		logger:       logger,
		authSvc:      authSvc,
		db:           database,
		customerSvc:  customers.NewService(database),
		orderSvc:     salesorders.NewService(database),
		revisionSvc:  revisions.NewService(database),
		paymentSvc:   payments.NewService(database),
		documentSvc:  documents.NewService(database, storageAdapter),
		jobsSvc:      jobsSvc,
		reportsSvc:   reportsSvc,
		storage:      storageAdapter,
		staticServer: http.StripPrefix("/web/", http.FileServer(http.Dir("web"))),
	}
	srv.configureFilesServer()
	srv.router = srv.routes()
	return srv
}

func (s *Server) routes() http.Handler {
	mux := http.NewServeMux()
	mux.HandleFunc("/api/health", s.handleHealth)
	mux.HandleFunc("/api/auth/login", s.handleLogin)
	mux.HandleFunc("/api/customers", s.handleCustomers)
	mux.HandleFunc("/api/customers/", s.handleCustomerByID)
	mux.HandleFunc("/api/sales-orders", s.handleSalesOrders)
	mux.HandleFunc("/api/sales-orders/", s.handleSalesOrderByID)
	mux.HandleFunc("/api/revisions", s.handleRevisions)
	mux.HandleFunc("/api/revisions/upload", s.handleRevisionsUpload)
	mux.HandleFunc("/api/payments", s.handlePayments)
	mux.HandleFunc("/api/documents", s.handleDocuments)
	mux.HandleFunc("/api/jobs/run", s.handleJobsRun)
	mux.HandleFunc("/api/reports/kpis", s.handleReportsKPIs)
	mux.HandleFunc("/api/reports/export/customers", s.handleReportsExportCustomers)
	mux.HandleFunc("/api/reports/export/orders", s.handleReportsExportOrders)
	mux.HandleFunc("/api/reports/export/payments", s.handleReportsExportPayments)
	mux.Handle("/web/", s.staticServer)
	if s.filesPrefix != "" && s.filesServer != nil {
		mux.Handle(s.filesPrefix, s.filesServer)
	}
	mux.HandleFunc("/", s.handleIndex)

	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if strings.HasPrefix(r.URL.Path, "/api/") && !isPublicAPI(r.URL.Path) {
			claims, err := s.authorizeRequest(r)
			if err != nil {
				s.writeError(w, http.StatusUnauthorized, err)
				return
			}
			ctx := context.WithValue(r.Context(), contextKeyClaims, claims)
			r = r.WithContext(ctx)
		}

		mux.ServeHTTP(w, r)
	})
}

// ServeHTTP makes Server implement http.Handler.
func (s *Server) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	s.router.ServeHTTP(w, r)
}

func (s *Server) configureFilesServer() {
	prefix := s.cfg.Storage.URLPrefix
	if prefix == "" {
		prefix = "/files/"
	}
	if !strings.HasPrefix(prefix, "/") {
		prefix = "/" + prefix
	}
	if !strings.HasSuffix(prefix, "/") {
		prefix += "/"
	}

	dir := s.cfg.Storage.LocalPath
	if dir == "" {
		dir = "./files"
	}
	if err := os.MkdirAll(dir, 0o755); err != nil {
		s.logger.Error("files_server_init_failed", map[string]any{"error": err.Error(), "dir": dir})
		s.filesPrefix = prefix
		s.filesServer = http.NotFoundHandler()
		return
	}
	s.filesPrefix = prefix
	s.filesServer = http.StripPrefix(strings.TrimSuffix(prefix, "/"), http.FileServer(http.Dir(dir)))
}

func isPublicAPI(path string) bool {
	switch path {
	case "/api/health", "/api/auth/login":
		return true
	default:
		return false
	}
}

func (s *Server) handleHealth(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
		return
	}

	ctx := r.Context()
	dbStatus := "ok"
	if err := db.Ping(ctx, s.db); err != nil {
		dbStatus = "error"
	}

	s.writeJSON(w, http.StatusOK, map[string]any{
		"status": "ok",
		"db":     dbStatus,
		"nowIso": time.Now().UTC().Format(time.RFC3339),
	})
}

func (s *Server) handleIndex(w http.ResponseWriter, r *http.Request) {
	if r.URL.Path != "/" {
		http.NotFound(w, r)
		return
	}
	if r.Method != http.MethodGet {
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
		return
	}
	http.ServeFile(w, r, filepath.Join("web", "index.html"))
}

func (s *Server) handleLogin(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
		return
	}

	var payload struct {
		Email    string `json:"email"`
		Password string `json:"password"`
	}
	if err := json.NewDecoder(r.Body).Decode(&payload); err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid JSON payload"))
		return
	}
	if payload.Email == "" || payload.Password == "" {
		s.writeError(w, http.StatusBadRequest, errors.New("email and password are required"))
		return
	}

	token, claims, err := s.authSvc.Authenticate(r.Context(), strings.ToLower(strings.TrimSpace(payload.Email)), payload.Password)
	if err != nil {
		s.writeError(w, http.StatusUnauthorized, errors.New("invalid credentials"))
		return
	}

	s.writeJSON(w, http.StatusOK, map[string]any{
		"token": token,
		"role":  claims.Role,
		"email": claims.Email,
	})
}

func (s *Server) handleCustomers(w http.ResponseWriter, r *http.Request) {
	switch r.Method {
	case http.MethodGet:
		s.listCustomers(w, r)
	case http.MethodPost:
		s.createCustomer(w, r)
	default:
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
	}
}

func (s *Server) handleCustomerByID(w http.ResponseWriter, r *http.Request) {
	idStr := strings.TrimPrefix(r.URL.Path, "/api/customers/")
	id, err := strconv.ParseInt(idStr, 10, 64)
	if err != nil || id <= 0 {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid customer id"))
		return
	}

	switch r.Method {
	case http.MethodPut:
		s.updateCustomer(w, r, id)
	default:
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
	}
}

func (s *Server) listCustomers(w http.ResponseWriter, r *http.Request) {
	query := r.URL.Query().Get("query")
	customersList, err := s.customerSvc.List(r.Context(), query)
	if err != nil {
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusOK, map[string]any{"customers": customersList})
}

func (s *Server) createCustomer(w http.ResponseWriter, r *http.Request) {
	var input customers.Input
	if err := json.NewDecoder(r.Body).Decode(&input); err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid JSON payload"))
		return
	}
	created, warnings, err := s.customerSvc.Create(r.Context(), input)
	if err != nil {
		if errors.Is(err, customers.ErrNotFound) {
			s.writeError(w, http.StatusNotFound, err)
			return
		}
		if isValidationError(err) {
			s.writeError(w, http.StatusBadRequest, err)
			return
		}
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusCreated, map[string]any{
		"customer":       created,
		"warnDuplicates": warnings,
	})
}

func (s *Server) updateCustomer(w http.ResponseWriter, r *http.Request, id int64) {
	var input customers.Input
	if err := json.NewDecoder(r.Body).Decode(&input); err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid JSON payload"))
		return
	}
	updated, warnings, err := s.customerSvc.Update(r.Context(), id, input)
	if err != nil {
		if errors.Is(err, customers.ErrNotFound) {
			s.writeError(w, http.StatusNotFound, err)
			return
		}
		if isValidationError(err) {
			s.writeError(w, http.StatusBadRequest, err)
			return
		}
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusOK, map[string]any{
		"customer":       updated,
		"warnDuplicates": warnings,
	})
}

func (s *Server) handleSalesOrders(w http.ResponseWriter, r *http.Request) {
	switch r.Method {
	case http.MethodGet:
		s.listSalesOrders(w, r)
	case http.MethodPost:
		s.createSalesOrder(w, r)
	default:
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
	}
}

func (s *Server) handleSalesOrderByID(w http.ResponseWriter, r *http.Request) {
	idStr := strings.TrimPrefix(r.URL.Path, "/api/sales-orders/")
	id, err := strconv.ParseInt(idStr, 10, 64)
	if err != nil || id <= 0 {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid sales order id"))
		return
	}
	switch r.Method {
	case http.MethodPut:
		s.updateSalesOrder(w, r, id)
	default:
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
	}
}

func (s *Server) handleRevisions(w http.ResponseWriter, r *http.Request) {
	switch r.Method {
	case http.MethodGet:
		s.listRevisions(w, r)
	default:
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
	}
}

func (s *Server) handleRevisionsUpload(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
		return
	}
	if err := r.ParseMultipartForm(32 << 20); err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid multipart payload"))
		return
	}
	salesOrderStr := strings.TrimSpace(r.FormValue("sales_order_id"))
	if salesOrderStr == "" {
		s.writeError(w, http.StatusBadRequest, errors.New("sales_order_id is required"))
		return
	}
	salesOrderID, err := strconv.ParseInt(salesOrderStr, 10, 64)
	if err != nil || salesOrderID <= 0 {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid sales_order_id"))
		return
	}

	file, header, err := r.FormFile("file")
	if err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("file is required"))
		return
	}
	defer file.Close()

	storedPath, err := s.storage.Save(r.Context(), header.Filename, file)
	if err != nil {
		s.writeServerError(w, err)
		return
	}

	revision, err := s.revisionSvc.Create(r.Context(), revisions.CreateInput{
		SalesOrderID: salesOrderID,
		Note:         r.FormValue("note"),
		Status:       r.FormValue("status"),
		FilePath:     storedPath,
	})
	if err != nil {
		if isValidationError(err) {
			s.writeError(w, http.StatusBadRequest, err)
			return
		}
		s.writeServerError(w, err)
		return
	}

	dto, err := s.revisionDTO(*revision)
	if err != nil {
		s.writeServerError(w, err)
		return
	}

	s.writeJSON(w, http.StatusCreated, map[string]any{"revision": dto})
}

func (s *Server) listRevisions(w http.ResponseWriter, r *http.Request) {
	salesOrderStr := strings.TrimSpace(r.URL.Query().Get("sales_order_id"))
	if salesOrderStr == "" {
		s.writeError(w, http.StatusBadRequest, errors.New("sales_order_id is required"))
		return
	}
	salesOrderID, err := strconv.ParseInt(salesOrderStr, 10, 64)
	if err != nil || salesOrderID <= 0 {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid sales_order_id"))
		return
	}

	revisionsList, err := s.revisionSvc.ListBySalesOrder(r.Context(), salesOrderID)
	if err != nil {
		if isValidationError(err) {
			s.writeError(w, http.StatusBadRequest, err)
			return
		}
		s.writeServerError(w, err)
		return
	}

	responses := make([]revisionResponse, 0, len(revisionsList))
	for _, rev := range revisionsList {
		dto, err := s.revisionDTO(rev)
		if err != nil {
			s.writeServerError(w, err)
			return
		}
		responses = append(responses, dto)
	}

	s.writeJSON(w, http.StatusOK, map[string]any{"revisions": responses})
}

func (s *Server) handlePayments(w http.ResponseWriter, r *http.Request) {
	switch r.Method {
	case http.MethodGet:
		s.listPayments(w, r)
	case http.MethodPost:
		s.createPayment(w, r)
	default:
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
	}
}

func (s *Server) listPayments(w http.ResponseWriter, r *http.Request) {
	var filter payments.Filter
	if so := strings.TrimSpace(r.URL.Query().Get("sales_order_id")); so != "" {
		id, err := strconv.ParseInt(so, 10, 64)
		if err != nil || id <= 0 {
			s.writeError(w, http.StatusBadRequest, errors.New("invalid sales_order_id"))
			return
		}
		filter.SalesOrderID = id
	}
	if from := strings.TrimSpace(r.URL.Query().Get("from")); from != "" {
		parsed, err := parseDate(from)
		if err != nil {
			s.writeError(w, http.StatusBadRequest, errors.New("invalid from date"))
			return
		}
		filter.From = &parsed
	}
	if to := strings.TrimSpace(r.URL.Query().Get("to")); to != "" {
		parsed, err := parseDate(to)
		if err != nil {
			s.writeError(w, http.StatusBadRequest, errors.New("invalid to date"))
			return
		}
		filter.To = &parsed
	}

	paymentsList, err := s.paymentSvc.List(r.Context(), filter)
	if err != nil {
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusOK, map[string]any{"payments": paymentsList})
}

func (s *Server) createPayment(w http.ResponseWriter, r *http.Request) {
	var input payments.CreateInput
	if err := json.NewDecoder(r.Body).Decode(&input); err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid JSON payload"))
		return
	}
	payment, err := s.paymentSvc.Create(r.Context(), input)
	if err != nil {
		if errors.Is(err, payments.ErrValidation) {
			s.writeError(w, http.StatusBadRequest, err)
			return
		}
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusCreated, map[string]any{"payment": payment})
}

func (s *Server) handleDocuments(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
		return
	}
	var input documents.CreateInput
	if err := json.NewDecoder(r.Body).Decode(&input); err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid JSON payload"))
		return
	}
	doc, err := s.documentSvc.Create(r.Context(), input)
	if err != nil {
		if errors.Is(err, documents.ErrValidation) {
			s.writeError(w, http.StatusBadRequest, err)
			return
		}
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusCreated, map[string]any{"document": doc})
}

func (s *Server) handleJobsRun(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
		return
	}
	if s.jobsSvc == nil {
		s.writeError(w, http.StatusServiceUnavailable, errors.New("job service unavailable"))
		return
	}
	claims := s.claimsFromContext(r.Context())
	if claims == nil || !strings.EqualFold(claims.Role, "admin") {
		s.writeError(w, http.StatusForbidden, errors.New("admin access required"))
		return
	}
	result, err := s.jobsSvc.Run(r.Context())
	if err != nil {
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusOK, map[string]any{"result": result})
}

func (s *Server) handleReportsKPIs(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
		return
	}
	if s.reportsSvc == nil {
		s.writeError(w, http.StatusServiceUnavailable, errors.New("reports unavailable"))
		return
	}
	start, err := parseOptionalTime(r.URL.Query().Get("start"))
	if err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid start date"))
		return
	}
	end, err := parseOptionalTime(r.URL.Query().Get("end"))
	if err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid end date"))
		return
	}
	kpi, err := s.reportsSvc.KPIs(r.Context(), start, end)
	if err != nil {
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusOK, map[string]any{"kpis": kpi})
}

func (s *Server) handleReportsExportCustomers(w http.ResponseWriter, r *http.Request) {
	s.handleReportsExport(w, r, "customers.csv", func(ctx context.Context) ([]byte, error) {
		return s.reportsSvc.CustomersCSV(ctx)
	})
}

func (s *Server) handleReportsExportOrders(w http.ResponseWriter, r *http.Request) {
	s.handleReportsExport(w, r, "orders.csv", func(ctx context.Context) ([]byte, error) {
		return s.reportsSvc.OrdersCSV(ctx)
	})
}

func (s *Server) handleReportsExportPayments(w http.ResponseWriter, r *http.Request) {
	s.handleReportsExport(w, r, "payments.csv", func(ctx context.Context) ([]byte, error) {
		return s.reportsSvc.PaymentsCSV(ctx)
	})
}

func (s *Server) handleReportsExport(w http.ResponseWriter, r *http.Request, filename string, fetch func(context.Context) ([]byte, error)) {
	if r.Method != http.MethodGet {
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
		return
	}
	if s.reportsSvc == nil {
		s.writeError(w, http.StatusServiceUnavailable, errors.New("reports unavailable"))
		return
	}
	data, err := fetch(r.Context())
	if err != nil {
		s.writeServerError(w, err)
		return
	}
	s.writeCSVResponse(w, filename, data)
}

func (s *Server) listSalesOrders(w http.ResponseWriter, r *http.Request) {
	var customerID int64
	if cid := strings.TrimSpace(r.URL.Query().Get("customer_id")); cid != "" {
		if parsed, err := strconv.ParseInt(cid, 10, 64); err == nil {
			customerID = parsed
		} else {
			s.writeError(w, http.StatusBadRequest, errors.New("invalid customer_id"))
			return
		}
	}
	status := r.URL.Query().Get("status")
	orders, err := s.orderSvc.List(r.Context(), customerID, status)
	if err != nil {
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusOK, map[string]any{"sales_orders": orders})
}

func (s *Server) createSalesOrder(w http.ResponseWriter, r *http.Request) {
	var input salesorders.Input
	if err := json.NewDecoder(r.Body).Decode(&input); err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid JSON payload"))
		return
	}
	order, err := s.orderSvc.Create(r.Context(), input)
	if err != nil {
		if isValidationError(err) {
			s.writeError(w, http.StatusBadRequest, err)
			return
		}
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusCreated, map[string]any{"sales_order": order})
}

func (s *Server) updateSalesOrder(w http.ResponseWriter, r *http.Request, id int64) {
	var payload struct {
		Status       *string    `json:"status"`
		Priority     *string    `json:"priority"`
		LeadTimeDays *int       `json:"lead_time_days"`
		StartedAt    *time.Time `json:"started_at"`
	}
	if err := json.NewDecoder(r.Body).Decode(&payload); err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid JSON payload"))
		return
	}

	input := salesorders.UpdateInput{
		Status:       payload.Status,
		Priority:     payload.Priority,
		LeadTimeDays: payload.LeadTimeDays,
	}
	if payload.StartedAt != nil {
		started := payload.StartedAt
		input.StartedAt = &started
	}

	order, err := s.orderSvc.Update(r.Context(), id, input)
	if err != nil {
		if errors.Is(err, salesorders.ErrNotFound) {
			s.writeError(w, http.StatusNotFound, err)
			return
		}
		if isValidationError(err) {
			s.writeError(w, http.StatusBadRequest, err)
			return
		}
		s.writeServerError(w, err)
		return
	}
	s.writeJSON(w, http.StatusOK, map[string]any{"sales_order": order})
}

func (s *Server) authorizeRequest(r *http.Request) (*auth.Claims, error) {
	header := r.Header.Get("Authorization")
	if header == "" {
		return nil, errors.New("missing authorization header")
	}
	parts := strings.SplitN(header, " ", 2)
	if len(parts) != 2 || !strings.EqualFold(parts[0], "Bearer") {
		return nil, errors.New("invalid authorization header")
	}
	token := strings.TrimSpace(parts[1])
	if token == "" {
		return nil, errors.New("empty token")
	}
	claims, err := s.authSvc.ParseToken(token)
	if err != nil {
		return nil, errors.New("invalid token")
	}
	return claims, nil
}

func (s *Server) claimsFromContext(ctx context.Context) *auth.Claims {
	if ctx == nil {
		return nil
	}
	if claims, ok := ctx.Value(contextKeyClaims).(*auth.Claims); ok {
		return claims
	}
	return nil
}

func (s *Server) writeJSON(w http.ResponseWriter, status int, payload any) {
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(status)
	_ = json.NewEncoder(w).Encode(payload)
}

func (s *Server) writeError(w http.ResponseWriter, status int, err error) {
	s.writeJSON(w, status, map[string]any{
		"error": err.Error(),
	})
}

func (s *Server) writeServerError(w http.ResponseWriter, err error) {
	s.logger.Error("request_failed", map[string]any{"error": err.Error()})
	s.writeError(w, http.StatusInternalServerError, fmt.Errorf("internal server error"))
}

func (s *Server) writeCSVResponse(w http.ResponseWriter, filename string, data []byte) {
	w.Header().Set("Content-Type", "text/csv")
	if filename != "" {
		w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", filename))
	}
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(data)
}
func isValidationError(err error) bool {
	if err == nil {
		return false
	}
	if errors.Is(err, customers.ErrValidation) ||
		errors.Is(err, salesorders.ErrValidation) ||
		errors.Is(err, revisions.ErrValidation) ||
		errors.Is(err, payments.ErrValidation) ||
		errors.Is(err, documents.ErrValidation) {
		return true
	}
	return false
}

func parseDate(value string) (time.Time, error) {
	trimmed := strings.TrimSpace(value)
	if trimmed == "" {
		return time.Time{}, fmt.Errorf("invalid date")
	}
	layouts := []string{time.RFC3339, "2006-01-02"}
	var lastErr error
	for _, layout := range layouts {
		if t, err := time.Parse(layout, trimmed); err == nil {
			return t.UTC(), nil
		} else {
			lastErr = err
		}
	}
	if lastErr != nil {
		return time.Time{}, lastErr
	}
	return time.Time{}, fmt.Errorf("invalid date")
}

func parseOptionalTime(raw string) (*time.Time, error) {
	trimmed := strings.TrimSpace(raw)
	if trimmed == "" {
		return nil, nil
	}
	parsed, err := parseDate(trimmed)
	if err != nil {
		return nil, err
	}
	return &parsed, nil
}

type revisionResponse struct {
	ID           int64     `json:"id"`
	SalesOrderID int64     `json:"sales_order_id"`
	Note         string    `json:"note"`
	FilePath     string    `json:"file_path"`
	Status       string    `json:"status"`
	CreatedAt    time.Time `json:"created_at"`
	FileURL      string    `json:"file_url"`
}

func (s *Server) revisionDTO(rev revisions.Revision) (revisionResponse, error) {
	if s.storage == nil {
		return revisionResponse{}, fmt.Errorf("storage adapter unavailable")
	}
	url, err := s.storage.URL(rev.FilePath)
	if err != nil {
		return revisionResponse{}, fmt.Errorf("revision url: %w", err)
	}
	return revisionResponse{
		ID:           rev.ID,
		SalesOrderID: rev.SalesOrderID,
		Note:         rev.Note,
		FilePath:     rev.FilePath,
		Status:       rev.Status,
		CreatedAt:    rev.CreatedAt,
		FileURL:      url,
	}, nil
}

// ClaimsFromContext extracts auth claims from request context when available.
func ClaimsFromContext(ctx context.Context) (*auth.Claims, bool) {
	val, ok := ctx.Value(contextKeyClaims).(*auth.Claims)
	return val, ok
}
