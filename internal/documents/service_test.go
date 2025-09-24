package documents

import "testing"

func TestShippingForSubtotal(t *testing.T) {
	if got := shippingForSubtotal(1500, "Sales Invoice"); got != 50 {
		t.Fatalf("expected shipping 50 got %v", got)
	}
	if got := shippingForSubtotal(2500, "Sales Invoice"); got != 0 {
		t.Fatalf("expected shipping 0 got %v", got)
	}
	if got := shippingForSubtotal(500, "Sales Receipt"); got != 0 {
		t.Fatalf("expected shipping 0 for receipts got %v", got)
	}
}
