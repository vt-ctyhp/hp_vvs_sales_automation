package logging

import (
	"encoding/json"
	"os"
	"strings"
	"sync"
	"time"
)

type Level int

const (
	LevelDebug Level = iota
	LevelInfo
	LevelError
)

type Logger struct {
	mu         sync.Mutex
	level      Level
	baseFields map[string]any
	encoder    *json.Encoder
}

// New creates a logger writing JSON to stdout.
func New(level string) *Logger {
	return &Logger{
		level:      parseLevel(level),
		baseFields: map[string]any{},
		encoder:    json.NewEncoder(os.Stdout),
	}
}

// With returns a child logger that includes the given fields on every entry.
func (l *Logger) With(fields map[string]any) *Logger {
	child := &Logger{
		level:   l.level,
		encoder: l.encoder,
	}
	merged := make(map[string]any, len(l.baseFields)+len(fields))
	for k, v := range l.baseFields {
		merged[k] = v
	}
	for k, v := range fields {
		merged[k] = v
	}
	child.baseFields = merged
	return child
}

func (l *Logger) log(level Level, msg string, fields map[string]any) {
	if level < l.level {
		return
	}
	entry := make(map[string]any, len(l.baseFields)+len(fields)+3)
	for k, v := range l.baseFields {
		entry[k] = v
	}
	for k, v := range fields {
		entry[k] = v
	}
	entry["ts"] = time.Now().UTC().Format(time.RFC3339Nano)
	entry["level"] = levelString(level)
	entry["msg"] = msg

	l.mu.Lock()
	defer l.mu.Unlock()
	_ = l.encoder.Encode(entry)
}

// Info logs an informational message.
func (l *Logger) Info(msg string, fields map[string]any) {
	l.log(LevelInfo, msg, fields)
}

// Debug logs a debug message.
func (l *Logger) Debug(msg string, fields map[string]any) {
	l.log(LevelDebug, msg, fields)
}

// Error logs an error message.
func (l *Logger) Error(msg string, fields map[string]any) {
	l.log(LevelError, msg, fields)
}

func parseLevel(level string) Level {
	switch strings.ToLower(level) {
	case "debug":
		return LevelDebug
	case "error":
		return LevelError
	default:
		return LevelInfo
	}
}

func levelString(level Level) string {
	switch level {
	case LevelDebug:
		return "debug"
	case LevelError:
		return "error"
	default:
		return "info"
	}
}
