package payments

import (
	"testing"
	"time"
)

func TestRoundCurrency(t *testing.T) {
	if got := roundCurrency(10.005); got != 10.01 {
		t.Fatalf("expected 10.01 got %0.2f", got)
	}
}

func TestAutoAllocateOldestFirst(t *testing.T) {
	outstanding := map[int64]OutstandingBalance{
		1: {SalesOrderID: 1, Outstanding: 100, CreatedAt: mustParse("2024-01-01T00:00:00Z")},
		2: {SalesOrderID: 2, Outstanding: 50, CreatedAt: mustParse("2024-01-05T00:00:00Z")},
		3: {SalesOrderID: 3, Outstanding: 75, CreatedAt: mustParse("2023-12-01T00:00:00Z")},
	}
	allocations := autoAllocate(180, outstanding)
	if len(allocations) != 3 {
		t.Fatalf("expected 3 allocations got %d", len(allocations))
	}
	if allocations[0].SalesOrderID != 3 {
		t.Fatalf("expected first allocation to order 3, got %d", allocations[0].SalesOrderID)
	}
	if allocations[0].Amount != 75 {
		t.Fatalf("expected allocation 75 got %0.2f", allocations[0].Amount)
	}
	if allocations[1].SalesOrderID != 1 || allocations[1].Amount != 100 {
		t.Fatalf("unexpected second allocation %+v", allocations[1])
	}
	if allocations[2].SalesOrderID != 2 || allocations[2].Amount != 5 {
		t.Fatalf("unexpected third allocation %+v", allocations[2])
	}
}

func mustParse(value string) time.Time {
	t, err := time.Parse(time.RFC3339, value)
	if err != nil {
		panic(err)
	}
	return t
}
