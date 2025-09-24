package rules

import "testing"

func TestShippingForSubtotal(t *testing.T) {
	if got := ShippingForSubtotal(1500); got != 50 {
		t.Fatalf("expected 50, got %v", got)
	}
	if got := ShippingForSubtotal(2500); got != 0 {
		t.Fatalf("expected 0, got %v", got)
	}
	if got := ShippingForSubtotal(-10); got != 0 {
		t.Fatalf("negative subtotal should be 0 shipping, got %v", got)
	}
	if got := ShippingForSubtotal(0); got != 0 {
		t.Fatalf("zero subtotal should not charge shipping, got %v", got)
	}
}
