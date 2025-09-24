package documents

import "testing"

func TestShippingForDocument(t *testing.T) {
        if got := shippingForDocument("Sales Invoice", 1500); got != 50 {
                t.Fatalf("expected shipping 50 got %v", got)
        }
        if got := shippingForDocument("Sales Invoice", 2500); got != 0 {
                t.Fatalf("expected shipping 0 got %v", got)
        }
        if got := shippingForDocument("Sales Receipt", 500); got != 0 {
                t.Fatalf("expected shipping 0 for receipts got %v", got)
        }
}
