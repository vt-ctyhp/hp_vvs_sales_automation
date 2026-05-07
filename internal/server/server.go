package server

import (
	"context"
	"crypto/hmac"
	"crypto/sha256"
	"encoding/base64"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"

	"github.com/example/vvsapp/internal/auth"
	"github.com/example/vvsapp/internal/config"
	"github.com/example/vvsapp/internal/db"
	"github.com/example/vvsapp/internal/logging"
)

// contextKey helps avoid collisions when storing values in request contexts.
type contextKey string

const contextKeyClaims contextKey = "authClaims"

// Server wires HTTP handlers, authentication, and diagnostics.
type Server struct {
	cfg     *config.Config
	logger  *logging.Logger
	authSvc *auth.Service
	db      DB
	router  http.Handler
	events  externalEventAppender
}

// DB defines the subset of database/sql used by the HTTP server (for easier testing).
type DB interface {
	PingContext(ctx context.Context) error
}

// New constructs a server with routes and middleware applied.
func New(cfg *config.Config, logger *logging.Logger, database DB, authSvc *auth.Service) *Server {
	appender, err := newGoogleSheetsExternalEventAppender(cfg.ExternalBooking)
	if err != nil && logger != nil {
		logger.Error("external_booking_appender_init_failed", map[string]any{"error": err.Error()})
	}
	srv := &Server{
		cfg:     cfg,
		logger:  logger,
		authSvc: authSvc,
		db:      database,
		events:  appender,
	}
	srv.router = srv.routes()
	return srv
}

func (s *Server) routes() http.Handler {
	mux := http.NewServeMux()
	mux.HandleFunc("/api/health", s.handleHealth)
	mux.HandleFunc("/api/auth/login", s.handleLogin)
	mux.HandleFunc("/api/webhooks/acuity", s.handleAcuityWebhook)

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

func isPublicAPI(path string) bool {
	switch path {
	case "/api/health", "/api/auth/login", "/api/webhooks/acuity":
		return true
	default:
		return false
	}
}

func (s *Server) handleAcuityWebhook(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		s.writeError(w, http.StatusMethodNotAllowed, errors.New("method not allowed"))
		return
	}
	secret := strings.TrimSpace(s.cfg.ExternalBooking.AcuityWebhookSecret)
	if secret == "" {
		s.writeError(w, http.StatusServiceUnavailable, errors.New("acuity webhook secret is not configured"))
		return
	}
	body, err := io.ReadAll(io.LimitReader(r.Body, 1<<20))
	if err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("read webhook body failed"))
		return
	}
	if !verifyAcuitySignature(secret, body, r.Header.Get("x-acuity-signature")) {
		s.writeError(w, http.StatusUnauthorized, errors.New("invalid acuity signature"))
		return
	}
	form, err := url.ParseQuery(string(body))
	if err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid form payload"))
		return
	}
	action := normalizeAcuityAction(firstFormValue(form, "action"))
	if action == "changed" {
		s.writeJSON(w, http.StatusOK, map[string]any{"ok": true, "status": "ignored", "action": action})
		return
	}
	if action != "scheduled" && action != "rescheduled" && action != "canceled" {
		s.writeError(w, http.StatusBadRequest, errors.New("unsupported acuity action"))
		return
	}
	id := firstFormValue(form, "id")
	if strings.TrimSpace(id) == "" {
		s.writeError(w, http.StatusBadRequest, errors.New("missing appointment id"))
		return
	}
	if s.events == nil {
		s.writeError(w, http.StatusServiceUnavailable, errors.New("external booking sheet appender is not configured"))
		return
	}
	rawPayloadJSON, err := acuityWebhookPayloadJSON(form)
	if err != nil {
		s.writeError(w, http.StatusBadRequest, errors.New("invalid webhook payload"))
		return
	}
	event := externalBookingEvent{
		ReceivedAt:            time.Now().UTC(),
		Provider:              "acuity",
		Action:                action,
		ProviderAppointmentID: strings.TrimSpace(id),
		CalendarID:            firstFormValue(form, "calendarID"),
		AppointmentTypeID:     firstFormValue(form, "appointmentTypeID"),
		RawPayloadJSON:        rawPayloadJSON,
		SignatureVerified:     true,
	}
	if err := s.events.AppendExternalBookingEvent(r.Context(), event); err != nil {
		if s.logger != nil {
			s.logger.Error("acuity_webhook_enqueue_failed", map[string]any{"error": err.Error(), "action": action, "id": id})
		}
		s.writeError(w, http.StatusInternalServerError, errors.New("enqueue acuity webhook failed"))
		return
	}
	s.writeJSON(w, http.StatusOK, map[string]any{
		"ok":     true,
		"status": "queued",
		"action": action,
		"id":     strings.TrimSpace(id),
	})
}

func verifyAcuitySignature(secret string, body []byte, signature string) bool {
	signature = strings.TrimSpace(signature)
	if secret == "" || signature == "" {
		return false
	}
	mac := hmac.New(sha256.New, []byte(secret))
	_, _ = mac.Write(body)
	expected := base64.StdEncoding.EncodeToString(mac.Sum(nil))
	return hmac.Equal([]byte(expected), []byte(signature))
}

func normalizeAcuityAction(raw string) string {
	action := strings.ToLower(strings.TrimSpace(raw))
	action = strings.TrimPrefix(action, "appointment.")
	action = strings.TrimPrefix(action, "appointment_")
	action = strings.TrimPrefix(action, "appointment-")
	if action == "cancelled" {
		return "canceled"
	}
	if action == "schedule" {
		return "scheduled"
	}
	if action == "reschedule" {
		return "rescheduled"
	}
	return action
}

func firstFormValue(form url.Values, key string) string {
	values := form[key]
	if len(values) == 0 {
		return ""
	}
	return strings.TrimSpace(values[0])
}

func acuityWebhookPayloadJSON(form url.Values) (string, error) {
	payload := make(map[string]string, len(form))
	for key, values := range form {
		if len(values) == 0 {
			payload[key] = ""
			continue
		}
		payload[key] = values[0]
	}
	raw, err := json.Marshal(payload)
	if err != nil {
		return "", err
	}
	return string(raw), nil
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

// ClaimsFromContext extracts auth claims from request context when available.
func ClaimsFromContext(ctx context.Context) (*auth.Claims, bool) {
	val, ok := ctx.Value(contextKeyClaims).(*auth.Claims)
	return val, ok
}
