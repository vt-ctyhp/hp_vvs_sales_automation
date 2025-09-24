package storage

import (
	"context"
	"io"
)

// Adapter defines file persistence behaviors required by the application.
type Adapter interface {
	Save(ctx context.Context, name string, r io.Reader) (string, error)
	Read(ctx context.Context, path string) (io.ReadCloser, error)
	URL(path string) (string, error)
}
