package salesorders

import (
	"testing"
	"time"
)

func TestNormalizeInputDefaults(t *testing.T) {
	input := Input{CustomerID: 1, SOCode: " SO-1 ", Status: " open "}
	cleaned, err := normalizeInput(input)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if cleaned.Priority != "P2" {
		t.Fatalf("expected default priority P2, got %s", cleaned.Priority)
	}
	if cleaned.LeadTimeDays != 28 {
		t.Fatalf("expected default lead time 28, got %d", cleaned.LeadTimeDays)
	}
	if cleaned.SOCode != "SO-1" {
		t.Fatalf("expected trimmed so code, got %q", cleaned.SOCode)
	}
	if cleaned.Status != "open" {
		t.Fatalf("expected trimmed status, got %q", cleaned.Status)
	}
}

func TestNormalizeInputForUpdateRejectsNegativeLeadTime(t *testing.T) {
	lead := -1
	if _, err := normalizeInputForUpdate(UpdateInput{LeadTimeDays: &lead}); err == nil {
		t.Fatal("expected error for negative lead time")
	}
}

func TestNormalizeInputForUpdateTracksStartedAt(t *testing.T) {
	now := time.Now()
	ptr := &now
	res, err := normalizeInputForUpdate(UpdateInput{StartedAt: &ptr})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !res.HasStartedAt {
		t.Fatal("expected HasStartedAt to be true")
	}
	if res.StartedAt == nil {
		t.Fatal("expected StartedAt to be set")
	}
}

func TestNormalizeInputValidatesFields(t *testing.T) {
	if _, err := normalizeInput(Input{}); err == nil {
		t.Fatal("expected error for missing required fields")
	}
	if _, err := normalizeInput(Input{CustomerID: 1}); err == nil {
		t.Fatal("expected error for missing so_code")
	}
	if _, err := normalizeInput(Input{CustomerID: 1, SOCode: "abc"}); err == nil {
		t.Fatal("expected error for missing status")
	}
}

func TestNormalizeInputAllowsStartedAt(t *testing.T) {
	now := time.Now()
	input := Input{CustomerID: 1, SOCode: "so", Status: "open", StartedAt: &now}
	cleaned, err := normalizeInput(input)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if cleaned.StartedAt == nil || !cleaned.StartedAt.Equal(now) {
		t.Fatal("expected started at to remain unchanged")
	}
}
