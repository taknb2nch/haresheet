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

func (sc *SheetClient) getValues(ctx context.Context, row int, col int, height int, width int) ([][]any, error) {
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

	var rawValues [][]any

	if len(resp.ValueRanges) > 0 &&
		resp.ValueRanges[0].ValueRange != nil &&
		resp.ValueRanges[0].ValueRange.Values != nil {
		rawValues = resp.ValueRanges[0].ValueRange.Values
	} else {
		rawValues = [][]any{}
	}

	if height < 1 && width < 1 {
		return rawValues, nil
	}

	targetRows := height

	if targetRows < 1 {
		targetRows = len(rawValues)
	}

	result := make([][]any, targetRows)

	for r := 0; r < targetRows; r++ {
		targetCols := width

		if targetCols < 1 {
			if r < len(rawValues) {
				targetCols = len(rawValues[r])
			} else {
				targetCols = 0
			}
		}

		result[r] = make([]any, targetCols)

		for c := 0; c < targetCols; c++ {
			result[r][c] = ""

			if r < len(rawValues) && c < len(rawValues[r]) {
				if val := rawValues[r][c]; val != nil {
					result[r][c] = val
				}
			}
		}
	}

	return result, nil
}

// GetRangeValues retrieves values using SheetID via DataFilter (ID直指定版)
func (sc *SheetClient) GetRangeValues(ctx context.Context, row int, col int, height int, width int) ([][]any, error) {
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
func (sc *SheetClient) GetColValues(ctx context.Context, col, width, skipRows int) ([][]any, error) {
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

	return sc.getValues(ctx, skipRows, col, rangeUnset, width)
}

// GetRowValues retrieves values from a specific row.
func (sc *SheetClient) GetRowValues(ctx context.Context, row, height, skipCols int) ([][]any, error) {
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

	return sc.getValues(ctx, row, skipCols, height, rangeUnset)
}

func (sc *SheetClient) checkRectInvalid(row int, col int, height int, width int, label string) error {
	if row < 0 {
		return fmt.Errorf("%s: invalid row: %d", label, row)
	}

	if col < 0 {
		return fmt.Errorf("%s: invalid col: %d", label, col)
	}

	if height < 0 && height != rangeUnset {
		return fmt.Errorf("%s: invalid height: %d", label, height)
	}

	if width < 0 && width != rangeUnset {
		return fmt.Errorf("%s: invalid width: %d", label, width)
	}

	return nil
}

// GetGridSize returns the current grid dimensions (rows and columns) of this sheet.
func (s *SheetClient) GetGridSize(ctx context.Context) (rowCount, colCount int, err error) {
	resp, err := s.c.service.Spreadsheets.Get(s.c.spreadID).
		Fields("sheets(properties(sheetId,gridProperties))").
		Context(ctx).
		Do()
	if err != nil {
		return 0, 0, fmt.Errorf("GetGridSize: failed to fetch spreadsheet info: %w", err)
	}

	for _, sheet := range resp.Sheets {
		if sheet.Properties == nil {
			continue
		}

		if sheet.Properties.SheetId == s.sheetID {
			if sheet.Properties.GridProperties == nil {
				return 0, 0, fmt.Errorf("GetGridSize: grid properties are missing for sheet %d", s.sheetID)
			}

			props := sheet.Properties.GridProperties

			return int(props.RowCount), int(props.ColumnCount), nil
		}
	}

	return 0, 0, fmt.Errorf("GetGridSize: sheet %d not found", s.sheetID)
}
