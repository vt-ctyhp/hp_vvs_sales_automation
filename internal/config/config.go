package config

import (
	"errors"
	"fmt"
	"os"

	"gopkg.in/yaml.v3"
)

// Config holds the full application configuration.
type Config struct {
	Server   ServerConfig   `yaml:"server"`
	Database DatabaseConfig `yaml:"database"`
	Logging  LoggingConfig  `yaml:"logging"`
	Auth     AuthConfig     `yaml:"auth"`
	Seed     SeedConfig     `yaml:"seed"`
}

// ServerConfig defines HTTP server settings.
type ServerConfig struct {
	Address string `yaml:"address"`
}

// DatabaseConfig defines persistence settings.
type DatabaseConfig struct {
	Path string `yaml:"path"`
}

// LoggingConfig defines log output behavior.
type LoggingConfig struct {
	Level string `yaml:"level"`
}

// AuthConfig controls JWT and session behavior.
type AuthConfig struct {
	JWTSecret       string `yaml:"jwt_secret"`
	TokenTTLMinutes int    `yaml:"token_ttl_minutes"`
}

// SeedConfig controls default data seeding.
type SeedConfig struct {
	AdminEmail    string `yaml:"admin_email"`
	AdminPassword string `yaml:"admin_password"`
	AdminRole     string `yaml:"admin_role"`
}

// Load reads configuration from disk and applies environment overrides.
func Load(path string) (*Config, error) {
	cfg := defaultConfig()

	if path != "" {
		if _, err := os.Stat(path); err == nil {
			data, readErr := os.ReadFile(path)
			if readErr != nil {
				return nil, fmt.Errorf("read config: %w", readErr)
			}
			if unmarshalErr := yaml.Unmarshal(data, &cfg); unmarshalErr != nil {
				return nil, fmt.Errorf("unmarshal config: %w", unmarshalErr)
			}
		} else if !errors.Is(err, os.ErrNotExist) {
			return nil, fmt.Errorf("stat config: %w", err)
		}
	}

	cfg.applyEnvOverrides()

	return &cfg, nil
}

func defaultConfig() Config {
	return Config{
		Server:   ServerConfig{Address: ":8080"},
		Database: DatabaseConfig{Path: "./local.db"},
		Logging:  LoggingConfig{Level: "info"},
		Auth: AuthConfig{
			JWTSecret:       "local-dev-secret-please-change",
			TokenTTLMinutes: 15,
		},
		Seed: SeedConfig{
			AdminEmail:    "admin@example.com",
			AdminPassword: "changeme123",
			AdminRole:     "admin",
		},
	}
}

func (c *Config) applyEnvOverrides() {
	if v := os.Getenv("VVSAPP_SERVER_ADDRESS"); v != "" {
		c.Server.Address = v
	}
	if v := os.Getenv("VVSAPP_DB_PATH"); v != "" {
		c.Database.Path = v
	}
	if v := os.Getenv("VVSAPP_LOG_LEVEL"); v != "" {
		c.Logging.Level = v
	}
	if v := os.Getenv("VVSAPP_JWT_SECRET"); v != "" {
		c.Auth.JWTSecret = v
	}
	if v := os.Getenv("VVSAPP_TOKEN_TTL_MINUTES"); v != "" {
		if ttl, err := parseIntEnv(v); err == nil {
			c.Auth.TokenTTLMinutes = ttl
		}
	}
	if v := os.Getenv("VVSAPP_ADMIN_EMAIL"); v != "" {
		c.Seed.AdminEmail = v
	}
	if v := os.Getenv("VVSAPP_ADMIN_PASSWORD"); v != "" {
		c.Seed.AdminPassword = v
	}
	if v := os.Getenv("VVSAPP_ADMIN_ROLE"); v != "" {
		c.Seed.AdminRole = v
	}
}

func parseIntEnv(raw string) (int, error) {
	var val int
	_, err := fmt.Sscanf(raw, "%d", &val)
	return val, err
}

// Summary returns a sanitized configuration snapshot suitable for logs.
func (c *Config) Summary() map[string]any {
	return map[string]any{
		"server": map[string]any{
			"address": c.Server.Address,
		},
		"database": map[string]any{
			"path": c.Database.Path,
		},
		"logging": map[string]any{
			"level": c.Logging.Level,
		},
		"auth": map[string]any{
			"token_ttl_minutes": c.Auth.TokenTTLMinutes,
		},
		"seed": map[string]any{
			"admin_email": c.Seed.AdminEmail,
			"admin_role":  c.Seed.AdminRole,
		},
	}
}
