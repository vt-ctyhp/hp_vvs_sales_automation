package rules

// ShippingForSubtotal applies the default shipping rule: orders under 2000 incur
// a $50 charge, otherwise shipping is free. Negative or zero subtotals do not
// incur shipping.
func ShippingForSubtotal(subtotal float64) float64 {
	if subtotal <= 0 {
		return 0
	}
	if subtotal < 2000 {
		return 50
	}
	return 0
}
