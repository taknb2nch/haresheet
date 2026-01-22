package haresheet

import (
	"fmt"

	"google.golang.org/api/sheets/v4"
)

// RowBuilder
type RowBuilder struct {
	sb    *SheetBuilder
	start int
	count int
}

// Requests
func (rb *RowBuilder) Requests() ([]*sheets.Request, error) {
	return rb.sb.b.Requests()
}

// Hide makes the rows invisible.
func (rb *RowBuilder) Hide() *RowBuilder {
	rb.sb.SetRowsHidden(rb.start, rb.count, true)

	return rb
}

// Show makes the rows visible.
func (rb *RowBuilder) Show() *RowBuilder {
	rb.sb.SetRowsHidden(rb.start, rb.count, false)

	return rb
}

// SetHeight sets the height of the rows.
func (rb *RowBuilder) SetHeight(pixels int) *RowBuilder {
	if pixels < 0 {
		rb.sb.b.appendError(fmt.Errorf("SetHeight: pixels must be positive"))

		return rb
	}

	rb.sb.updateDimension("ROWS", rb.start, rb.count, &sheets.DimensionProperties{PixelSize: int64(pixels)}, "pixelSize")

	return rb
}
