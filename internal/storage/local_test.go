package storage

import (
	"bytes"
	"context"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestLocalSaveReadAndURL(t *testing.T) {
	tmp := t.TempDir()
	adapter, err := NewLocal(tmp, "/files/")
	if err != nil {
		t.Fatalf("new local: %v", err)
	}

	content := []byte("hello world")
	path, err := adapter.Save(context.Background(), " Invoice #1.PDF ", bytes.NewReader(content))
	if err != nil {
		t.Fatalf("save: %v", err)
	}
	if !strings.HasPrefix(path, "files/") {
		t.Fatalf("expected files/ prefix, got %s", path)
	}

	rc, err := adapter.Read(context.Background(), path)
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	defer rc.Close()
	data, err := io.ReadAll(rc)
	if err != nil {
		t.Fatalf("read all: %v", err)
	}
	if string(data) != string(content) {
		t.Fatalf("expected %q got %q", string(content), string(data))
	}

	url, err := adapter.URL(path)
	if err != nil {
		t.Fatalf("url: %v", err)
	}
	if !strings.HasPrefix(url, "/files/") {
		t.Fatalf("expected url prefix /files/, got %s", url)
	}

	// ensure file exists on disk where expected
	rel := strings.TrimPrefix(path, "files/")
	if rel == path {
		t.Fatalf("expected to trim files/ prefix")
	}
	if _, err := os.Stat(filepath.Join(tmp, rel)); err != nil {
		t.Fatalf("stat saved file: %v", err)
	}
}

func TestLocalRejectsInvalidPath(t *testing.T) {
	tmp := t.TempDir()
	adapter, err := NewLocal(tmp, "/files/")
	if err != nil {
		t.Fatalf("new local: %v", err)
	}
	if _, err := adapter.Read(context.Background(), "../../etc/passwd"); !errors.Is(err, ErrInvalidPath) {
		t.Fatalf("expected ErrInvalidPath, got %v", err)
	}
}
