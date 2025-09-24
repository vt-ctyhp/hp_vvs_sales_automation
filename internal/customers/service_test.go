package customers

import "testing"

func TestNormalizePhone(t *testing.T) {
	cases := map[string]string{
		"":                "",
		"123-456-7890":    "1234567890",
		" (555) 000-1111": "5550001111",
	}
	for input, want := range cases {
		if got := normalizePhone(input); got != want {
			t.Fatalf("normalizePhone(%q) = %q, want %q", input, got, want)
		}
	}
}

func TestNormalizeInputRequiresBusinessName(t *testing.T) {
	if _, err := normalizeInput(Input{}); err == nil {
		t.Fatal("expected error for missing business name")
	}
}
