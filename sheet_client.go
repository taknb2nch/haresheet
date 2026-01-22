package haresheet

import (
	"context"
	"fmt"

	"google.golang.org/api/sheets/v4"
)

type SheetClient struct {
	c       *Client
	sheetID int64
	err     error
}

func (sc *SheetClient) getValues(ctx context.Context, row int, col int, height int, width int) ([][]interface{}, error) {
	rng := &sheets.GridRange{
		SheetId:          sc.sheetID,
		StartRowIndex:    int64(row),
		StartColumnIndex: int64(col),
	}

	if height > 0 {
		rng.EndRowIndex = int64(row + height)
	}

	if width > 0 {
		rng.EndColumnIndex = int64(col + width)
	}

	req := &sheets.BatchGetValuesByDataFilterRequest{
		DataFilters: []*sheets.DataFilter{
			{GridRange: rng},
		},
	}

	resp, err := sc.c.service.Spreadsheets.Values.BatchGetByDataFilter(sc.c.spreadID, req).Context(ctx).Do()
	if err != nil {
		return nil, err
	}

	if len(resp.ValueRanges) == 0 ||
		resp.ValueRanges[0].ValueRange == nil ||
		resp.ValueRanges[0].ValueRange.Values == nil {
		return [][]interface{}{}, nil
	}

	return resp.ValueRanges[0].ValueRange.Values, nil
}

// GetRangeValues retrieves values using SheetID via DataFilter (ID直指定版)
func (sc *SheetClient) GetRangeValues(ctx context.Context, row int, col int, height int, width int) ([][]interface{}, error) {
	if sc.err != nil {
		return nil, sc.err
	}

	err := sc.checkRectInvalid(row, col, height, width, "GetRangeValues")
	if err != nil {
		return nil, err
	}

	return sc.getValues(ctx, row, col, height, width)
}

// GetColValues retrieves values from a specific column.
func (sc *SheetClient) GetColValues(ctx context.Context, col, width, skipRows int) ([][]interface{}, error) {
	if sc.err != nil {
		return nil, sc.err
	}

	if col < 0 {
		return nil, fmt.Errorf("GetColValues: invalid col: %d", col)
	}

	if width < 1 {
		return nil, fmt.Errorf("GetColValues: invalid width: %d", width)
	}

	if skipRows < 0 {
		return nil, fmt.Errorf("GetColValues: invalid skipRows: %d", skipRows)
	}

	return sc.getValues(ctx, skipRows, col, toEnd, width)
}

// GetRowValues retrieves values from a specific row.
func (sc *SheetClient) GetRowValues(ctx context.Context, row, height, skipCols int) ([][]interface{}, error) {
	if sc.err != nil {
		return nil, sc.err
	}

	if row < 0 {
		return nil, fmt.Errorf("GetRowValues: invalid row: %d", row)
	}

	if height < 1 {
		return nil, fmt.Errorf("GetRowValues: invalid height: %d", height)
	}

	if skipCols < 0 {
		return nil, fmt.Errorf("GetRowValues: invalid skipCols: %d", skipCols)
	}

	return sc.getValues(ctx, row, skipCols, height, toEnd)
}

func (sc *SheetClient) checkRectInvalid(row int, col int, height int, width int, label string) error {
	if row < 0 {
		return fmt.Errorf("%s: invalid row: %d", label, row)
	}

	if col < 0 {
		return fmt.Errorf("%s: invalid col: %d", label, col)
	}

	if height < 0 && height != toEnd {
		return fmt.Errorf("%s: invalid height: %d", label, height)
	}

	if width < 0 && width != toEnd {
		return fmt.Errorf("%s: invalid width: %d", label, width)
	}

	return nil
}
