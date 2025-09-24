package storage

import (
	"context"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/google/uuid"
)

var (
	// ErrInvalidPath indicates the provided path could not be resolved safely.
	ErrInvalidPath = errors.New("invalid storage path")
	// ErrNotFound indicates the requested file was not located on disk.
	ErrNotFound = errors.New("file not found")
)

const filesRoot = "files"

// Local implements Adapter by persisting files to the local filesystem.
type Local struct {
	basePath  string
	urlPrefix string
}

// NewLocal constructs a Local adapter backed by the provided directory.
func NewLocal(basePath, urlPrefix string) (*Local, error) {
	if strings.TrimSpace(basePath) == "" {
		return nil, fmt.Errorf("base path is required")
	}
	abs, err := filepath.Abs(basePath)
	if err != nil {
		return nil, fmt.Errorf("resolve base path: %w", err)
	}
	if err := os.MkdirAll(abs, 0o755); err != nil {
		return nil, fmt.Errorf("ensure base path: %w", err)
	}
	if urlPrefix == "" {
		urlPrefix = "/files/"
	}
	return &Local{basePath: abs, urlPrefix: urlPrefix}, nil
}

// Save writes the provided stream to disk and returns a relative storage path.
func (l *Local) Save(ctx context.Context, name string, r io.Reader) (string, error) {
	if r == nil {
		return "", fmt.Errorf("reader is required")
	}
	now := time.Now().UTC()
	year := fmt.Sprintf("%04d", now.Year())
	month := fmt.Sprintf("%02d", int(now.Month()))

	baseName, ext := sanitizeFileName(name)
	fileName := fmt.Sprintf("%s_%s%s", uuid.NewString(), baseName, ext)

	relDir := filepath.Join(year, month)
	absDir := filepath.Join(l.basePath, relDir)
	if err := os.MkdirAll(absDir, 0o755); err != nil {
		return "", fmt.Errorf("create storage directory: %w", err)
	}

	absPath := filepath.Join(absDir, fileName)
	file, err := os.Create(absPath)
	if err != nil {
		return "", fmt.Errorf("create file: %w", err)
	}
	defer file.Close()

	if _, err := io.Copy(file, r); err != nil {
		return "", fmt.Errorf("write file: %w", err)
	}

	relPath := filepath.ToSlash(filepath.Join(filesRoot, relDir, fileName))
	return relPath, nil
}

// Read opens a stored file for streaming back to callers.
func (l *Local) Read(ctx context.Context, path string) (io.ReadCloser, error) {
	absPath, _, err := l.resolve(path)
	if err != nil {
		return nil, err
	}
	file, err := os.Open(absPath)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return nil, ErrNotFound
		}
		return nil, fmt.Errorf("open file: %w", err)
	}
	return file, nil
}

// URL returns a browser-accessible path for the stored file.
func (l *Local) URL(path string) (string, error) {
	_, relWithRoot, err := l.resolve(path)
	if err != nil {
		return "", err
	}
	trimmed := strings.TrimPrefix(relWithRoot, filesRoot+"/")
	prefix := l.urlPrefix
	if !strings.HasSuffix(prefix, "/") {
		prefix += "/"
	}
	url := prefix + trimmed
	url = strings.ReplaceAll(url, "\\", "/")
	if !strings.HasPrefix(url, "/") {
		url = "/" + url
	}
	return url, nil
}

func (l *Local) resolve(path string) (string, string, error) {
	trimmed := strings.TrimSpace(path)
	if trimmed == "" {
		return "", "", ErrInvalidPath
	}
	cleaned := filepath.Clean(trimmed)
	cleaned = strings.TrimPrefix(cleaned, string(filepath.Separator))
	cleaned = strings.TrimPrefix(cleaned, filesRoot+string(filepath.Separator))
	cleaned = strings.TrimPrefix(cleaned, filesRoot+"/")
	if cleaned == "" {
		return "", "", ErrInvalidPath
	}
	for _, segment := range strings.Split(cleaned, string(filepath.Separator)) {
		if segment == ".." {
			return "", "", ErrInvalidPath
		}
	}
	absPath := filepath.Join(l.basePath, cleaned)
	rel, err := filepath.Rel(l.basePath, absPath)
	if err != nil {
		return "", "", ErrInvalidPath
	}
	if strings.HasPrefix(rel, "..") {
		return "", "", ErrInvalidPath
	}
	relWithRoot := filepath.ToSlash(filepath.Join(filesRoot, cleaned))
	return absPath, relWithRoot, nil
}

var invalidChars = regexp.MustCompile(`[^a-zA-Z0-9-_]+`)

func sanitizeFileName(name string) (string, string) {
	base := filepath.Base(strings.TrimSpace(name))
	if base == "." || base == "" {
		base = "file"
	}
	ext := strings.ToLower(filepath.Ext(base))
	base = strings.TrimSuffix(base, ext)
	base = strings.ToLower(base)
	base = invalidChars.ReplaceAllString(base, "_")
	base = strings.Trim(base, "_")
	if base == "" {
		base = "file"
	}
	return base, ext
}
