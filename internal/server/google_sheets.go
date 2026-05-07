package server

import (
	"bytes"
	"context"
	"crypto"
	"crypto/rand"
	"crypto/rsa"
	"crypto/sha256"
	"crypto/x509"
	"encoding/base64"
	"encoding/json"
	"encoding/pem"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"strings"
	"sync"
	"time"

	"github.com/example/vvsapp/internal/config"
)

const googleSheetsScope = "https://www.googleapis.com/auth/spreadsheets"

type externalBookingEvent struct {
	ReceivedAt            time.Time
	Provider              string
	Action                string
	ProviderAppointmentID string
	CalendarID            string
	AppointmentTypeID     string
	RawPayloadJSON        string
	SignatureVerified     bool
	TestRunID             string
}

type externalEventAppender interface {
	AppendExternalBookingEvent(ctx context.Context, event externalBookingEvent) error
}

type googleSheetsExternalEventAppender struct {
	spreadsheetID string
	queueRange    string
	sa            googleServiceAccount
	client        *http.Client
	mu            sync.Mutex
	token         string
	tokenExpiry   time.Time
}

type googleServiceAccount struct {
	ClientEmail string `json:"client_email"`
	PrivateKey  string `json:"private_key"`
	TokenURI    string `json:"token_uri"`
}

func newGoogleSheetsExternalEventAppender(cfg config.ExternalBookingConfig) (externalEventAppender, error) {
	if strings.TrimSpace(cfg.SpreadsheetID) == "" {
		return nil, nil
	}
	if strings.TrimSpace(cfg.GoogleServiceAccountJSON) == "" && strings.TrimSpace(cfg.GoogleServiceAccountFile) == "" {
		return nil, nil
	}
	sa, err := loadGoogleServiceAccount(cfg)
	if err != nil {
		return nil, err
	}
	if strings.TrimSpace(sa.ClientEmail) == "" || strings.TrimSpace(sa.PrivateKey) == "" {
		return nil, errors.New("google service account is missing client_email or private_key")
	}
	if strings.TrimSpace(sa.TokenURI) == "" {
		sa.TokenURI = "https://oauth2.googleapis.com/token"
	}
	queueRange := strings.TrimSpace(cfg.QueueRange)
	if queueRange == "" {
		queueRange = "'_ExternalBookingEvents'!A:P"
	}
	return &googleSheetsExternalEventAppender{
		spreadsheetID: strings.TrimSpace(cfg.SpreadsheetID),
		queueRange:    queueRange,
		sa:            sa,
		client:        &http.Client{Timeout: 15 * time.Second},
	}, nil
}

func loadGoogleServiceAccount(cfg config.ExternalBookingConfig) (googleServiceAccount, error) {
	raw := strings.TrimSpace(cfg.GoogleServiceAccountJSON)
	if raw == "" {
		data, err := os.ReadFile(strings.TrimSpace(cfg.GoogleServiceAccountFile))
		if err != nil {
			return googleServiceAccount{}, fmt.Errorf("read google service account file: %w", err)
		}
		raw = string(data)
	}
	if !strings.HasPrefix(raw, "{") {
		if decoded, err := base64.StdEncoding.DecodeString(raw); err == nil && len(decoded) > 0 {
			raw = string(decoded)
		}
	}
	var sa googleServiceAccount
	if err := json.Unmarshal([]byte(raw), &sa); err != nil {
		return googleServiceAccount{}, fmt.Errorf("parse google service account: %w", err)
	}
	return sa, nil
}

func (a *googleSheetsExternalEventAppender) AppendExternalBookingEvent(ctx context.Context, event externalBookingEvent) error {
	token, err := a.accessToken(ctx)
	if err != nil {
		return err
	}
	values := [][]any{{
		event.ReceivedAt.UTC().Format(time.RFC3339),
		event.Provider,
		event.Action,
		event.ProviderAppointmentID,
		event.CalendarID,
		event.AppointmentTypeID,
		event.RawPayloadJSON,
		boolForSheet(event.SignatureVerified),
		"PENDING",
		0,
		"",
		"",
		"",
		"",
		"",
		event.TestRunID,
	}}
	payload, _ := json.Marshal(map[string]any{"values": values})
	endpoint := fmt.Sprintf(
		"https://sheets.googleapis.com/v4/spreadsheets/%s/values/%s:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
		url.PathEscape(a.spreadsheetID),
		url.PathEscape(a.queueRange),
	)
	req, err := http.NewRequestWithContext(ctx, http.MethodPost, endpoint, bytes.NewReader(payload))
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")
	resp, err := a.client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	body, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return fmt.Errorf("sheets append failed: status=%d body=%s", resp.StatusCode, strings.TrimSpace(string(body)))
	}
	return nil
}

func (a *googleSheetsExternalEventAppender) accessToken(ctx context.Context) (string, error) {
	a.mu.Lock()
	if a.token != "" && time.Until(a.tokenExpiry) > 2*time.Minute {
		token := a.token
		a.mu.Unlock()
		return token, nil
	}
	a.mu.Unlock()

	token, expiry, err := a.fetchAccessToken(ctx)
	if err != nil {
		return "", err
	}

	a.mu.Lock()
	a.token = token
	a.tokenExpiry = expiry
	a.mu.Unlock()
	return token, nil
}

func (a *googleSheetsExternalEventAppender) fetchAccessToken(ctx context.Context) (string, time.Time, error) {
	assertion, err := a.jwtAssertion()
	if err != nil {
		return "", time.Time{}, err
	}
	form := url.Values{}
	form.Set("grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer")
	form.Set("assertion", assertion)
	req, err := http.NewRequestWithContext(ctx, http.MethodPost, a.sa.TokenURI, strings.NewReader(form.Encode()))
	if err != nil {
		return "", time.Time{}, err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	resp, err := a.client.Do(req)
	if err != nil {
		return "", time.Time{}, err
	}
	defer resp.Body.Close()
	body, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return "", time.Time{}, fmt.Errorf("google token request failed: status=%d body=%s", resp.StatusCode, strings.TrimSpace(string(body)))
	}
	var decoded struct {
		AccessToken string `json:"access_token"`
		ExpiresIn   int    `json:"expires_in"`
	}
	if err := json.Unmarshal(body, &decoded); err != nil {
		return "", time.Time{}, fmt.Errorf("parse google token: %w", err)
	}
	if decoded.AccessToken == "" {
		return "", time.Time{}, errors.New("google token response missing access_token")
	}
	expiresIn := decoded.ExpiresIn
	if expiresIn <= 0 {
		expiresIn = 3600
	}
	return decoded.AccessToken, time.Now().Add(time.Duration(expiresIn) * time.Second), nil
}

func (a *googleSheetsExternalEventAppender) jwtAssertion() (string, error) {
	key, err := parseRSAPrivateKey(a.sa.PrivateKey)
	if err != nil {
		return "", err
	}
	now := time.Now()
	header := base64.RawURLEncoding.EncodeToString([]byte(`{"alg":"RS256","typ":"JWT"}`))
	claims, _ := json.Marshal(map[string]any{
		"iss":   a.sa.ClientEmail,
		"scope": googleSheetsScope,
		"aud":   a.sa.TokenURI,
		"exp":   now.Add(55 * time.Minute).Unix(),
		"iat":   now.Unix(),
	})
	payload := base64.RawURLEncoding.EncodeToString(claims)
	signingInput := header + "." + payload
	sum := sha256.Sum256([]byte(signingInput))
	signature, err := rsa.SignPKCS1v15(rand.Reader, key, crypto.SHA256, sum[:])
	if err != nil {
		return "", fmt.Errorf("sign google jwt: %w", err)
	}
	return signingInput + "." + base64.RawURLEncoding.EncodeToString(signature), nil
}

func parseRSAPrivateKey(raw string) (*rsa.PrivateKey, error) {
	block, _ := pem.Decode([]byte(raw))
	if block == nil {
		return nil, errors.New("google private key PEM decode failed")
	}
	parsed, err := x509.ParsePKCS8PrivateKey(block.Bytes)
	if err == nil {
		key, ok := parsed.(*rsa.PrivateKey)
		if !ok {
			return nil, errors.New("google private key is not RSA")
		}
		return key, nil
	}
	if key, err2 := x509.ParsePKCS1PrivateKey(block.Bytes); err2 == nil {
		return key, nil
	}
	return nil, fmt.Errorf("parse google private key: %w", err)
}

func boolForSheet(v bool) string {
	if v {
		return "TRUE"
	}
	return "FALSE"
}
