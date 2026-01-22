package haresheet

import (
	"fmt"

	"google.golang.org/api/sheets/v4"
)

// ColumnBuilder
type ColumnBuilder struct {
	sb    *SheetBuilder
	start int
	count int
}

// Requests
func (cb *ColumnBuilder) Requests() ([]*sheets.Request, error) {
	return cb.sb.b.Requests()
}

// Hide makes the columns invisible.
func (cb *ColumnBuilder) Hide() *ColumnBuilder {
	cb.sb.SetColumnsHidden(cb.start, cb.count, true)

	return cb
}

// Show makes the columns visible.
func (cb *ColumnBuilder) Show() *ColumnBuilder {
	cb.sb.SetColumnsHidden(cb.start, cb.count, false)

	return cb
}

// SetWidth sets the width of the columns.
func (cb *ColumnBuilder) SetWidth(pixels int) *ColumnBuilder {
	if pixels < 0 {
		cb.sb.b.appendError(fmt.Errorf("SetHeight: pixels must be positive"))

		return cb
	}

	cb.sb.updateDimension("COLUMNS", cb.start, cb.count, &sheets.DimensionProperties{PixelSize: int64(pixels)}, "pixelSize")

	return cb
}
