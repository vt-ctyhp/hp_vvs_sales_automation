package auth

import (
	"context"
	"database/sql"
	"errors"
	"fmt"
	"strings"
	"time"

	"github.com/golang-jwt/jwt/v5"
	"golang.org/x/crypto/bcrypt"

	"github.com/example/vvsapp/internal/config"
	"github.com/example/vvsapp/internal/logging"
)

// Service provides authentication helpers for login and token validation.
type Service struct {
	db       *sql.DB
	secret   []byte
	tokenTTL time.Duration
	logger   *logging.Logger
}

// Claims describes the JWT payload returned to clients.
type Claims struct {
	Email string `json:"email"`
	Role  string `json:"role"`
	jwt.RegisteredClaims
}

// NewService constructs an auth service from config values.
func NewService(db *sql.DB, cfg config.AuthConfig, logger *logging.Logger) (*Service, error) {
	if len(cfg.JWTSecret) < 16 {
		return nil, fmt.Errorf("jwt secret must be at least 16 characters")
	}
	ttl := time.Duration(cfg.TokenTTLMinutes) * time.Minute
	if ttl <= 0 {
		ttl = 15 * time.Minute
	}
	return &Service{
		db:       db,
		secret:   []byte(cfg.JWTSecret),
		tokenTTL: ttl,
		logger:   logger,
	}, nil
}

// Authenticate verifies credentials and returns a signed JWT.
func (s *Service) Authenticate(ctx context.Context, email, password string) (string, *Claims, error) {
	email = strings.ToLower(strings.TrimSpace(email))
	const query = `SELECT id, password_hash, role FROM users WHERE email = ?`
	var (
		userID       int64
		passwordHash string
		role         string
	)
	err := s.db.QueryRowContext(ctx, query, email).Scan(&userID, &passwordHash, &role)
	if errors.Is(err, sql.ErrNoRows) {
		return "", nil, fmt.Errorf("invalid credentials")
	}
	if err != nil {
		return "", nil, fmt.Errorf("query user: %w", err)
	}

	if bcrypt.CompareHashAndPassword([]byte(passwordHash), []byte(password)) != nil {
		return "", nil, fmt.Errorf("invalid credentials")
	}

	claims := &Claims{
		Email: email,
		Role:  role,
		RegisteredClaims: jwt.RegisteredClaims{
			Subject:   fmt.Sprint(userID),
			ExpiresAt: jwt.NewNumericDate(time.Now().Add(s.tokenTTL)),
			IssuedAt:  jwt.NewNumericDate(time.Now()),
		},
	}

	token := jwt.NewWithClaims(jwt.SigningMethodHS256, claims)
	signed, err := token.SignedString(s.secret)
	if err != nil {
		return "", nil, fmt.Errorf("sign token: %w", err)
	}

	return signed, claims, nil
}

// ParseToken validates the JWT and returns claims when valid.
func (s *Service) ParseToken(tokenStr string) (*Claims, error) {
	token, err := jwt.ParseWithClaims(tokenStr, &Claims{}, func(token *jwt.Token) (any, error) {
		if token.Method != jwt.SigningMethodHS256 {
			return nil, fmt.Errorf("unexpected signing method")
		}
		return s.secret, nil
	})
	if err != nil {
		return nil, fmt.Errorf("parse token: %w", err)
	}
	claims, ok := token.Claims.(*Claims)
	if !ok || !token.Valid {
		return nil, fmt.Errorf("invalid token claims")
	}
	return claims, nil
}

// SeedAdmin ensures an admin account exists using the provided credentials.
func SeedAdmin(ctx context.Context, db *sql.DB, seedCfg config.SeedConfig, logger *logging.Logger) error {
	if seedCfg.AdminEmail == "" {
		return fmt.Errorf("seed admin email is required")
	}
	if seedCfg.AdminPassword == "" {
		return fmt.Errorf("seed admin password is required")
	}

	email := strings.ToLower(strings.TrimSpace(seedCfg.AdminEmail))

	const lookup = `SELECT id FROM users WHERE email = ?`
	var id int64
	err := db.QueryRowContext(ctx, lookup, email).Scan(&id)
	if err == nil {
		logger.Debug("admin user already present", map[string]any{"email": email})
		return nil
	}
	if !errors.Is(err, sql.ErrNoRows) {
		return fmt.Errorf("check admin existence: %w", err)
	}

	hashed, err := bcrypt.GenerateFromPassword([]byte(seedCfg.AdminPassword), bcrypt.DefaultCost)
	if err != nil {
		return fmt.Errorf("hash admin password: %w", err)
	}

	if seedCfg.AdminRole == "" {
		seedCfg.AdminRole = "admin"
	}

	const insert = `INSERT INTO users(email, password_hash, role) VALUES(?, ?, ?)`
	if _, err := db.ExecContext(ctx, insert, email, string(hashed), seedCfg.AdminRole); err != nil {
		return fmt.Errorf("insert admin: %w", err)
	}

	logger.Info("seeded admin user", map[string]any{"email": email, "role": seedCfg.AdminRole})
	return nil
}
