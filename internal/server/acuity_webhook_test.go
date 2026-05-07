package server

import (
	"context"
	"crypto/hmac"
	"crypto/sha256"
	"encoding/base64"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"github.com/example/vvsapp/internal/config"
)

type stubExternalEventAppender struct {
	events []externalBookingEvent
	err    error
}

func (s *stubExternalEventAppender) AppendExternalBookingEvent(ctx context.Context, event externalBookingEvent) error {
	s.events = append(s.events, event)
	return s.err
}

func TestAcuityWebhookQueuesVerifiedAppointment(t *testing.T) {
	body := "action=rescheduled&id=123&calendarID=7&appointmentTypeID=9"
	stub := &stubExternalEventAppender{}
	srv := testAcuityWebhookServer("secret", stub)

	req := httptest.NewRequest(http.MethodPost, "/api/webhooks/acuity", strings.NewReader(body))
	req.Header.Set("x-acuity-signature", testAcuitySignature("secret", body))
	rec := httptest.NewRecorder()

	srv.ServeHTTP(rec, req)

	if rec.Code != http.StatusOK {
		t.Fatalf("status=%d body=%s", rec.Code, rec.Body.String())
	}
	if len(stub.events) != 1 {
		t.Fatalf("queued events=%d", len(stub.events))
	}
	event := stub.events[0]
	if event.Provider != "acuity" || event.Action != "rescheduled" || event.ProviderAppointmentID != "123" {
		t.Fatalf("unexpected event: %+v", event)
	}
	if !event.SignatureVerified {
		t.Fatalf("event signature flag was false")
	}
	if !strings.Contains(event.RawPayloadJSON, `"calendarID":"7"`) {
		t.Fatalf("raw payload was not preserved: %s", event.RawPayloadJSON)
	}
}

func TestAcuityWebhookRejectsInvalidSignature(t *testing.T) {
	stub := &stubExternalEventAppender{}
	srv := testAcuityWebhookServer("secret", stub)

	req := httptest.NewRequest(http.MethodPost, "/api/webhooks/acuity", strings.NewReader("action=scheduled&id=123"))
	req.Header.Set("x-acuity-signature", "bad-signature")
	rec := httptest.NewRecorder()

	srv.ServeHTTP(rec, req)

	if rec.Code != http.StatusUnauthorized {
		t.Fatalf("status=%d body=%s", rec.Code, rec.Body.String())
	}
	if len(stub.events) != 0 {
		t.Fatalf("invalid signature queued %d events", len(stub.events))
	}
}

func TestAcuityWebhookIgnoresChangedAction(t *testing.T) {
	body := "action=changed&id=123&calendarID=7&appointmentTypeID=9"
	stub := &stubExternalEventAppender{}
	srv := testAcuityWebhookServer("secret", stub)

	req := httptest.NewRequest(http.MethodPost, "/api/webhooks/acuity", strings.NewReader(body))
	req.Header.Set("x-acuity-signature", testAcuitySignature("secret", body))
	rec := httptest.NewRecorder()

	srv.ServeHTTP(rec, req)

	if rec.Code != http.StatusOK {
		t.Fatalf("status=%d body=%s", rec.Code, rec.Body.String())
	}
	if len(stub.events) != 0 {
		t.Fatalf("changed action queued %d events", len(stub.events))
	}
}

func testAcuityWebhookServer(secret string, appender externalEventAppender) *Server {
	cfg := &config.Config{}
	cfg.ExternalBooking.HPAppAcuityWebhookSecret = secret
	srv := &Server{cfg: cfg, events: appender}
	srv.router = srv.routes()
	return srv
}

func testAcuitySignature(secret string, body string) string {
	mac := hmac.New(sha256.New, []byte(secret))
	_, _ = mac.Write([]byte(body))
	return base64.StdEncoding.EncodeToString(mac.Sum(nil))
}
