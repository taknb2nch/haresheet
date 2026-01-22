package haresheet

import (
	"errors"
	"fmt"

	"google.golang.org/api/sheets/v4"
)

// SheetBuilder builds requests for a specific sheet.
type SheetBuilder struct {
	b       *Builder
	sheetID int64
}

// Requests
func (sb *SheetBuilder) Requests() ([]*sheets.Request, error) {
	return sb.b.Requests()
}

// Row returns a builder object for row operations.
func (sb *SheetBuilder) Row(startRow int, count int) *RowBuilder {
	rb := &RowBuilder{
		sb:    sb,
		start: startRow,
		count: count,
	}

	if startRow < 0 {
		sb.b.appendError(fmt.Errorf("Row: invalid start row: %d", startRow))

		return rb
	}

	if count < 1 {
		sb.b.appendError(fmt.Errorf("Row: invalid count: %d", count))

		return rb
	}

	return rb
}

// Column returns a builder object for column operations.
func (sb *SheetBuilder) Column(startCol int, count int) *ColumnBuilder {
	cb := &ColumnBuilder{
		sb:    sb,
		start: startCol,
		count: count,
	}

	if startCol < 0 {
		sb.b.appendError(fmt.Errorf("Row: invalid start col: %d", startCol))

		return cb
	}

	if count < 1 {
		sb.b.appendError(fmt.Errorf("Row: invalid count: %d", count))

		return cb
	}

	return cb
}

// CopyRange
func (sb *SheetBuilder) CopyRange(srcSheetID int64, src *Rect, dstR int, dstC int, pasteType PasteType) *SheetBuilder {
	if srcSheetID < 0 {
		sb.b.appendError(fmt.Errorf("CopyRange: invalid src sheet id: %d", srcSheetID))

		return sb
	}

	if sb.isRectInvalid(src, "CopyRange", "src") {
		return sb
	}

	if dstR < 0 {
		sb.b.appendError(fmt.Errorf("CopyRange: invalid dst row: %d", dstR))

		return sb
	}

	if dstC < 0 {
		sb.b.appendError(fmt.Errorf("CopyRange: invalid dst column: %d", dstR))

		return sb
	}

	if pasteType == "" {
		pasteType = PasteTypeNormal
	}

	srcRange := &sheets.GridRange{
		SheetId:          srcSheetID,
		StartRowIndex:    int64(src.Row),
		EndRowIndex:      int64(src.Row + src.Height),
		StartColumnIndex: int64(src.Col),
		EndColumnIndex:   int64(src.Col + src.Width),
	}

	dstRange := &sheets.GridRange{
		SheetId:          sb.sheetID,
		StartRowIndex:    int64(dstR),
		EndRowIndex:      int64(dstR + src.Height),
		StartColumnIndex: int64(dstC),
		EndColumnIndex:   int64(dstC + src.Width),
	}

	req := &sheets.Request{
		CopyPaste: &sheets.CopyPasteRequest{
			Source:           srcRange,
			Destination:      dstRange,
			PasteType:        string(pasteType),
			PasteOrientation: "NORMAL",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// FillDownFrom copies the content from srcRect and repeats it downwards to fill the specified height.
func (sb *SheetBuilder) FillDownFrom(srcSheetID int64, src *Rect, dstR int, dstC int, height int, pasteType PasteType) *SheetBuilder {
	if srcSheetID < 0 {
		sb.b.appendError(fmt.Errorf("FillDownFrom: invalid src sheet id: %d", srcSheetID))

		return sb
	}

	if sb.isRectInvalid(src, "FillDownFrom", "src") {
		return sb
	}

	if dstR < 0 {
		sb.b.appendError(fmt.Errorf("FillDownFrom: invalid dst row: %d", dstR))

		return sb
	}

	if dstC < 0 {
		sb.b.appendError(fmt.Errorf("FillDownFrom: invalid dst column: %d", dstR))

		return sb
	}

	if height <= 0 {
		sb.b.appendError(fmt.Errorf("FillDownFrom: invalid height: %d", height))

		return sb
	}

	srcRange := &sheets.GridRange{
		SheetId:          srcSheetID,
		StartRowIndex:    int64(src.Row),
		EndRowIndex:      int64(src.Row + src.Height),
		StartColumnIndex: int64(src.Col),
		EndColumnIndex:   int64(src.Col + src.Width),
	}

	// 貼り付け先 (開始位置 + 高さ)
	// Destinationの行幅を Source の倍数にすることで、APIが自動でリピート処理を行います
	dstRange := &sheets.GridRange{
		SheetId:          int64(sb.sheetID),
		StartRowIndex:    int64(dstR),
		EndRowIndex:      int64(dstR + height), // ★ここがポイント
		StartColumnIndex: int64(dstC),
		EndColumnIndex:   int64(dstC + src.Width),
	}

	if pasteType == "" {
		pasteType = PasteTypeNormal
	}

	req := &sheets.Request{
		CopyPaste: &sheets.CopyPasteRequest{
			Source:           srcRange,
			Destination:      dstRange,
			PasteType:        string(pasteType),
			PasteOrientation: "NORMAL",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// Merge
func (sb *SheetBuilder) Merge(rect *Rect, mergeType MergeType) *SheetBuilder {
	if rect == nil {
		sb.b.appendError(errors.New("Merge: rect should not be nil"))

		return sb
	}

	if mergeType == "" {
		mergeType = MergeTypeAll
	}

	req := &sheets.Request{
		MergeCells: &sheets.MergeCellsRequest{
			Range: &sheets.GridRange{
				SheetId:          sb.sheetID,
				StartRowIndex:    int64(rect.Row),
				EndRowIndex:      int64(rect.Row + rect.Height),
				StartColumnIndex: int64(rect.Col),
				EndColumnIndex:   int64(rect.Col + rect.Width),
			},
			MergeType: string(mergeType),
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// MergeRows
func (sb *SheetBuilder) MergeRows(rect *Rect) *SheetBuilder {
	return sb.Merge(rect, MergeTypeRows)
}

// MergeCols
func (sb *SheetBuilder) MergeCols(rect *Rect) *SheetBuilder {
	return sb.Merge(rect, MergeTypeCols)
}

// ExpandRows
func (sb *SheetBuilder) ExpandRows(count int) *SheetBuilder {
	if count < 0 {
		sb.b.appendError(fmt.Errorf("ExpandRows: invalid count: %d", count))

		return sb
	}

	req := &sheets.Request{
		AppendDimension: &sheets.AppendDimensionRequest{
			SheetId:   sb.sheetID,
			Dimension: "ROWS",
			Length:    int64(count),
		},
	}

	sb.b.PrependRequest(req)

	return sb
}

// ExpandColumns
func (sb *SheetBuilder) ExpandColumns(count int) *SheetBuilder {
	if count < 0 {
		sb.b.appendError(fmt.Errorf("ExpandRows: invalid count: %d", count))

		return sb
	}

	req := &sheets.Request{
		AppendDimension: &sheets.AppendDimensionRequest{
			SheetId:   sb.sheetID,
			Dimension: "COLUMNS",
			Length:    int64(count),
		},
	}

	sb.b.PrependRequest(req)

	return sb
}

// InsertColumns は指定したインデックスの位置に、指定した数の列を挿入します。
// startIndex: 0始まりのインデックス (A列=0, B列=1...)
// count: 挿入する列数
func (sb *SheetBuilder) InsertColumns(startIndex int, count int) *SheetBuilder {
	if startIndex < 0 {
		sb.b.appendError(fmt.Errorf("InsertColumns: invalid startIndex: %d", startIndex))

		return sb
	}

	if count <= 0 {
		sb.b.appendError(fmt.Errorf("InsertColumns: invalid count: %d", count))

		return sb
	}

	req := &sheets.Request{
		InsertDimension: &sheets.InsertDimensionRequest{
			Range: &sheets.DimensionRange{
				SheetId:    sb.sheetID,
				Dimension:  "COLUMNS",
				StartIndex: int64(startIndex),
				EndIndex:   int64(startIndex + count),
			},
		},
	}

	sb.b.PrependRequest(req)

	return sb
}

// SetCellValue
func (sb *SheetBuilder) SetCellValue(row int, col int, value any) *SheetBuilder {
	if row < 0 {
		sb.b.appendError(fmt.Errorf("SetCellValue: invalid row: %d", row))

		return sb
	}

	if col < 0 {
		sb.b.appendError(fmt.Errorf("SetCellValue: invalid col: %d", col))

		return sb
	}

	if value == nil {
		sb.b.appendError(errors.New("SetCellValue: value should not be nil"))

		return sb
	}

	cell := sb.toCellData(value)

	req := &sheets.Request{
		UpdateCells: &sheets.UpdateCellsRequest{
			Start: &sheets.GridCoordinate{
				SheetId:     sb.sheetID,
				RowIndex:    int64(row),
				ColumnIndex: int64(col),
			},
			Rows: []*sheets.RowData{
				{Values: []*sheets.CellData{cell}},
			},
			Fields: "userEnteredValue",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// SetRowValues
func (sb *SheetBuilder) SetRowValues(row int, col int, values []any) *SheetBuilder {
	if row < 0 {
		sb.b.appendError(fmt.Errorf("SetRowValues: invalid row: %d", row))

		return sb
	}

	if col < 0 {
		sb.b.appendError(fmt.Errorf("SetRowValues: invalid col: %d", col))

		return sb
	}

	if len(values) == 0 {
		sb.b.appendError(errors.New("SetRowValues: values should not be nil or empty"))

		return sb
	}

	cells := make([]*sheets.CellData, 0, len(values))

	for _, v := range values {
		cells = append(cells, sb.toCellData(v))
	}

	req := &sheets.Request{
		UpdateCells: &sheets.UpdateCellsRequest{
			Start: &sheets.GridCoordinate{
				SheetId:     sb.sheetID,
				RowIndex:    int64(row),
				ColumnIndex: int64(col),
			},
			Rows: []*sheets.RowData{
				{Values: cells},
			},
			Fields: "userEnteredValue",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// SetRangeValues
func (sb *SheetBuilder) SetRangeValues(row int, col int, values [][]any) *SheetBuilder {
	if row < 0 {
		sb.b.appendError(fmt.Errorf("SetRangeValues: invalid row: %d", row))

		return sb
	}

	if col < 0 {
		sb.b.appendError(fmt.Errorf("SetRangeValues: invalid col: %d", col))

		return sb
	}

	if len(values) == 0 {
		sb.b.appendError(errors.New("SetRangeValues: values should not be nil or empty"))

		return sb
	}

	rows := make([]*sheets.RowData, 0, len(values))

	for _, rowVals := range values {
		cells := make([]*sheets.CellData, 0, len(rowVals))

		for _, v := range rowVals {
			cells = append(cells, sb.toCellData(v))
		}

		rows = append(rows, &sheets.RowData{
			Values: cells,
		})
	}

	req := &sheets.Request{
		UpdateCells: &sheets.UpdateCellsRequest{
			Start: &sheets.GridCoordinate{
				SheetId:     sb.sheetID,
				RowIndex:    int64(row),
				ColumnIndex: int64(col),
			},
			Rows:   rows,
			Fields: "userEnteredValue",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// toCellData
func (sb *SheetBuilder) toCellData(v any) *sheets.CellData {
	cd := &sheets.CellData{}

	switch val := v.(type) {
	case string:
		// "=" で始まるなら数式として扱う
		if len(val) > 0 && val[0] == '=' {
			cd.UserEnteredValue = &sheets.ExtendedValue{FormulaValue: &val}
		} else {
			cd.UserEnteredValue = &sheets.ExtendedValue{StringValue: &val}
		}
	case int:
		f := float64(val)
		cd.UserEnteredValue = &sheets.ExtendedValue{NumberValue: &f}
	case float64:
		cd.UserEnteredValue = &sheets.ExtendedValue{NumberValue: &val}
	case bool:
		cd.UserEnteredValue = &sheets.ExtendedValue{BoolValue: &val}
	}

	return cd
}

// ProtectRange protects the specified area.
// If warningOnly is true, it shows a warning when editing but allows changes.
// If warningOnly is false, it restricts editing to the specified users (or owner only if users is empty).
func (sb *SheetBuilder) ProtectRange(rect *Rect, description string, users []string, warningOnly bool) *SheetBuilder {
	if sb.isRectInvalid(rect, "ProtectRange", "rect") {
		return sb
	}

	return sb.addProtectedRangeRequest(description, users, warningOnly, &sheets.GridRange{
		SheetId:          int64(sb.sheetID),
		StartRowIndex:    int64(rect.Row),
		EndRowIndex:      int64(rect.Row + rect.Height),
		StartColumnIndex: int64(rect.Col),
		EndColumnIndex:   int64(rect.Col + rect.Width),
	})
}

// ProtectSheet protects the entire sheet.
func (sb *SheetBuilder) ProtectSheet(description string, users []string, warningOnly bool) *SheetBuilder {
	return sb.addProtectedRangeRequest(description, users, warningOnly, &sheets.GridRange{
		SheetId: int64(sb.sheetID),
	})
}

// addProtectedRangeRequest
func (sb *SheetBuilder) addProtectedRangeRequest(desc string, users []string, warningOnly bool, rng *sheets.GridRange) *SheetBuilder {
	var editors *sheets.Editors

	if !warningOnly {
		if users == nil {
			users = []string{}
		}

		editors = &sheets.Editors{
			Users: users,
		}
	}

	req := &sheets.Request{
		AddProtectedRange: &sheets.AddProtectedRangeRequest{
			ProtectedRange: &sheets.ProtectedRange{
				Range:       rng,
				Description: desc,
				WarningOnly: warningOnly,
				Editors:     editors,
			},
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// SetForegroundColor sets the text color for the specified range.
func (sb *SheetBuilder) SetForegroundColor(rect *Rect, color *sheets.Color) *SheetBuilder {
	if sb.isRectInvalid(rect, "SetForegroundColor", "rect") {
		return sb
	}

	if color == nil {
		sb.b.appendError(errors.New("SetForegroundColor: color should not be nil"))

		return sb
	}

	req := &sheets.Request{
		RepeatCell: &sheets.RepeatCellRequest{
			Range: &sheets.GridRange{
				SheetId:          sb.sheetID,
				StartRowIndex:    int64(rect.Row),
				EndRowIndex:      int64(rect.Row + rect.Height),
				StartColumnIndex: int64(rect.Col),
				EndColumnIndex:   int64(rect.Col + rect.Width),
			},
			Cell: &sheets.CellData{
				UserEnteredFormat: &sheets.CellFormat{
					TextFormat: &sheets.TextFormat{
						ForegroundColor: color,
					},
				},
			},
			Fields: "userEnteredFormat.textFormat.foregroundColor",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// SetBackgroundColor
func (sb *SheetBuilder) SetBackgroundColor(rect *Rect, color *sheets.Color) *SheetBuilder {
	if sb.isRectInvalid(rect, "SetBackgroundColor", "rect") {
		return sb
	}

	if color == nil {
		sb.b.appendError(errors.New("SetBackgroundColor: color should not be nil"))

		return sb
	}

	req := &sheets.Request{
		RepeatCell: &sheets.RepeatCellRequest{
			Range: &sheets.GridRange{
				SheetId:          sb.sheetID,
				StartRowIndex:    int64(rect.Row),
				EndRowIndex:      int64(rect.Row + rect.Height),
				StartColumnIndex: int64(rect.Col),
				EndColumnIndex:   int64(rect.Col + rect.Width),
			},
			Cell: &sheets.CellData{
				UserEnteredFormat: &sheets.CellFormat{
					BackgroundColor: color,
				},
			},
			Fields: "userEnteredFormat.backgroundColor",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// SetTabColor sets the tab color using the API's Color struct.
func (sb *SheetBuilder) SetTabColor(color *sheets.Color) *SheetBuilder {
	if color == nil {
		sb.b.appendError(errors.New("SetTabColor: color should not be nil"))

		return sb
	}

	req := &sheets.Request{
		UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
			Properties: &sheets.SheetProperties{
				SheetId:  sb.sheetID,
				TabColor: color, // そのまま渡す
			},
			Fields: "tabColor",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// InsertRange inserts empty cells into the specified range and shifts existing cells.
func (sb *SheetBuilder) InsertRange(rect *Rect, shiftDimension ShiftDimensionType) *SheetBuilder {
	if sb.isRectInvalid(rect, "InsertRange", "rect") {
		return sb
	}

	if shiftDimension == "" {
		shiftDimension = ShiftDimensionTypeRows
	}

	req := &sheets.Request{
		InsertRange: &sheets.InsertRangeRequest{
			Range: &sheets.GridRange{
				SheetId:          sb.sheetID,
				StartRowIndex:    int64(rect.Row),
				EndRowIndex:      int64(rect.Row + rect.Height),
				StartColumnIndex: int64(rect.Col),
				EndColumnIndex:   int64(rect.Col + rect.Width),
			},
			ShiftDimension: string(shiftDimension),
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// InsertAndCopyRange inserts blank cells (shifting down) and then copies data from another sheet.
// It internally calls InsertRange and CopyRange.
func (sb *SheetBuilder) InsertAndCopyRange(srcSheetID int64, src *Rect, dstR int, dstC int, pasteType PasteType) *SheetBuilder {
	if srcSheetID < 0 {
		sb.b.appendError(fmt.Errorf("InsertAndCopyRange: invalid src sheet id: %d", srcSheetID))

		return sb
	}

	if sb.isRectInvalid(src, "InsertAndCopyRange", "src") {
		return sb
	}

	if dstR < 0 {
		sb.b.appendError(fmt.Errorf("InsertAndCopyRange: invalid dst row: %d", dstR))

		return sb
	}

	if dstC < 0 {
		sb.b.appendError(fmt.Errorf("InsertAndCopyRange: invalid dst column: %d", dstR))

		return sb
	}

	if pasteType == "" {
		pasteType = PasteTypeNormal
	}

	// 1. 挿入先の範囲を計算（コピー元と同じサイズ）
	dstRect := &Rect{
		Row:    dstR,
		Col:    dstC,
		Height: src.Height, // 高さはコピー元と同じ
		Width:  src.Width,  // 幅もコピー元と同じ
	}

	// 2. 空白を挿入（既存セルを下にずらす）
	sb.InsertRange(dstRect, ShiftDimensionTypeRows)

	// 3. コピー元の座標補正（※同一シートの場合のみ発動する安全装置）
	// 挿入によって自分が下にズレてしまった場合、コピー元の座標も追従させる
	actualSrc := *src

	if sb.sheetID == srcSheetID && dstR <= src.Row {
		actualSrc.Row += src.Height
	}

	sb.CopyRange(srcSheetID, &actualSrc, dstR, dstC, pasteType)

	return sb
}

// DeleteRange deletes the specified range and shifts other cells to fill the gap.
func (sb *SheetBuilder) DeleteRange(rect *Rect, shiftDimension ShiftDimensionType) *SheetBuilder {
	if sb.isRectInvalid(rect, "DeleteRange", "rect") {
		return sb
	}

	if shiftDimension == "" {
		shiftDimension = ShiftDimensionTypeRows
	}

	req := &sheets.Request{
		DeleteRange: &sheets.DeleteRangeRequest{
			Range: &sheets.GridRange{
				SheetId:          sb.sheetID,
				StartRowIndex:    int64(rect.Row),
				EndRowIndex:      int64(rect.Row + rect.Height),
				StartColumnIndex: int64(rect.Col),
				EndColumnIndex:   int64(rect.Col + rect.Width),
			},
			ShiftDimension: string(shiftDimension),
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// SetLink sets a hyperlink to the specified range.
// url: Can be an external URL (http://...) or an internal sheet link (#gid=...).
func (sb *SheetBuilder) SetLink(rect *Rect, url string) *SheetBuilder {
	if sb.isRectInvalid(rect, "SetLink", "rect") {
		return sb
	}

	if url == "" {
		sb.b.appendError(errors.New("SetLink: url should not be empty"))

		return sb
	}

	req := &sheets.Request{
		RepeatCell: &sheets.RepeatCellRequest{
			Range: &sheets.GridRange{
				SheetId:          sb.sheetID,
				StartRowIndex:    int64(rect.Row),
				EndRowIndex:      int64(rect.Row + rect.Height),
				StartColumnIndex: int64(rect.Col),
				EndColumnIndex:   int64(rect.Col + rect.Width),
			},
			Cell: &sheets.CellData{
				UserEnteredFormat: &sheets.CellFormat{
					TextFormat: &sheets.TextFormat{
						Link: &sheets.Link{
							Uri: url,
						},
						// 文字色は青、下線付きにするのが一般的ですが
						// 既存の書式を壊さないよう、ここではリンク属性のみ設定します
					},
				},
			},
			// リンク情報だけを更新する
			Fields: "userEnteredFormat.textFormat.link",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// SetLinkToSheet sets a link to jump to a specific sheet ID within the same spreadsheet.
func (sb *SheetBuilder) SetLinkToSheet(rect *Rect, targetSheetID int64) *SheetBuilder {
	if sb.isRectInvalid(rect, "SetLinkToSheet", "rect") {
		return sb
	}

	if targetSheetID <= 0 {
		sb.b.appendError(fmt.Errorf("Row: invalid target sheet id: %d", targetSheetID))

		return sb
	}

	// スプレッドシート内部リンクの形式 (#gid=ID) を生成
	internalLink := fmt.Sprintf("#gid=%d", targetSheetID)

	return sb.SetLink(rect, internalLink)
}

// SetHidden sets the visibility of the sheet.
// true to hide, false to show.
func (sb *SheetBuilder) SetHidden(hidden bool) *SheetBuilder {
	req := &sheets.Request{
		UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
			Properties: &sheets.SheetProperties{
				SheetId: sb.sheetID,
				Hidden:  hidden,
			},
			Fields: "hidden",
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// Hide hides the sheet.
func (sb *SheetBuilder) Hide() *SheetBuilder {
	return sb.SetHidden(true)
}

// Show shows the sheet.
func (sb *SheetBuilder) Show() *SheetBuilder {
	return sb.SetHidden(false)
}

// SetColumnsHidden sets the visibility of columns.
func (sb *SheetBuilder) SetColumnsHidden(startCol int, count int, hidden bool) *SheetBuilder {
	if startCol <= 0 {
		sb.b.appendError(fmt.Errorf("SetColumnsHidden: invalid start column: %d", startCol))

		return sb
	}

	if count <= 0 {
		sb.b.appendError(fmt.Errorf("SetColumnsHidden: invalid count: %d", count))

		return sb
	}

	return sb.updateDimension("COLUMNS", startCol, count, &sheets.DimensionProperties{HiddenByUser: hidden}, "hiddenByUser")
}

// SetRowsHidden sets the visibility of rows.
func (sb *SheetBuilder) SetRowsHidden(startRow int, count int, hidden bool) *SheetBuilder {
	if startRow <= 0 {
		sb.b.appendError(fmt.Errorf("SetColumnsHidden: invalid start row: %d", startRow))

		return sb
	}

	if count <= 0 {
		sb.b.appendError(fmt.Errorf("SetColumnsHidden: invalid count: %d", count))

		return sb
	}

	sb.updateDimension("ROWS", startRow, count, &sheets.DimensionProperties{HiddenByUser: hidden}, "hiddenByUser")

	return sb
}

// updateDimension is a generic helper to update row or column properties.
// It handles the tedious API request construction.
func (sb *SheetBuilder) updateDimension(dimension string, start int, count int, props *sheets.DimensionProperties, fields string) *SheetBuilder {
	req := &sheets.Request{
		UpdateDimensionProperties: &sheets.UpdateDimensionPropertiesRequest{
			Range: &sheets.DimensionRange{
				SheetId:    sb.sheetID,
				Dimension:  dimension,
				StartIndex: int64(start),
				EndIndex:   int64(start + count),
			},
			Properties: props,
			Fields:     fields,
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// toEnd represents a value indicating extension to the end of the dimension.
const toEnd = -1

// clearValues creates a request to clear values.
func (sb *SheetBuilder) clearValues(row int, col int, height int, width int) *SheetBuilder {
	rng := &sheets.GridRange{
		SheetId:          sb.sheetID,
		StartRowIndex:    int64(row),
		StartColumnIndex: int64(col),
	}

	if height > 0 {
		rng.EndRowIndex = int64(row + height)
	}

	if width > 0 {
		rng.EndColumnIndex = int64(col + width)
	}

	req := &sheets.Request{
		RepeatCell: &sheets.RepeatCellRequest{
			Range:  rng,
			Cell:   &sheets.CellData{},
			Fields: "userEnteredValue", // Only clear values
		},
	}

	sb.b.AppendRequest(req)

	return sb
}

// ClearRangeValues clears values in the specified rectangle.
func (sb *SheetBuilder) ClearRangeValues(rect *Rect) *SheetBuilder {
	if sb.isRectInvalid(rect, "ClearRangeValues", "rect") {
		return sb
	}

	return sb.clearValues(rect.Row, rect.Col, rect.Height, rect.Width)
}

// ClearColValues clears values in the specified columns.
func (sb *SheetBuilder) ClearColValues(col int, width int, skipRows int) *SheetBuilder {
	if col < 0 {
		sb.b.appendError(fmt.Errorf("ClearColValues: invalid col: %d", col))

		return sb
	}

	if width < 1 {
		sb.b.appendError(fmt.Errorf("ClearColValues: invalid width: %d", width))

		return sb
	}

	if skipRows < 0 {
		sb.b.appendError(fmt.Errorf("ClearColValues: invalid skipRows: %d", skipRows))

		return sb
	}

	return sb.clearValues(skipRows, col, toEnd, width)
}

// ClearRowValues clears values in the specified rows.
func (sb *SheetBuilder) ClearRowValues(row int, height int, skipCols int) *SheetBuilder {
	if row < 0 {
		sb.b.appendError(fmt.Errorf("ClearRowValues: invalid row: %d", row))

		return sb
	}

	if height < 1 {
		sb.b.appendError(fmt.Errorf("ClearRowValues: invalid height: %d", height))

		return sb
	}

	if skipCols < 0 {
		sb.b.appendError(fmt.Errorf("ClearRowValues: invalid skipCols: %d", skipCols))

		return sb
	}

	return sb.clearValues(row, skipCols, height, toEnd)
}

func (sb *SheetBuilder) isRectInvalid(rect *Rect, label string, name string) bool {
	if rect == nil {
		sb.b.appendError(fmt.Errorf("%s: %s should not be nil", label, name))

		return true
	}

	if rect.Row < 0 {
		sb.b.appendError(fmt.Errorf("%s: invalid %s.Row: %d", label, name, rect.Row))

		return true
	}

	if rect.Col < 0 {
		sb.b.appendError(fmt.Errorf("%s: invalid %s.Col: %d", label, name, rect.Col))

		return true
	}

	if rect.Height < 0 && rect.Height != toEnd {
		sb.b.appendError(fmt.Errorf("%s: invalid %s.Height: %d", label, name, rect.Height))

		return true
	}

	if rect.Width < 0 && rect.Width != toEnd {
		sb.b.appendError(fmt.Errorf("%s: invalid %s.Width: %d", label, name, rect.Width))

		return true
	}

	return false
}
