package haresheet

// PasteType defines the type of content to paste.
type PasteType string

const (
	PasteTypeNormal    PasteType = "PASTE_NORMAL"     // 値・式・書式すべて
	PasteTypeFormula   PasteType = "PASTE_FORMULA"    // 値・式のみ（書式なし）★今回使いたいのはこれ
	PasteTypeValues    PasteType = "PASTE_VALUES"     // 値のみ（式は計算結果になる）
	PasteTypeFormat    PasteType = "PASTE_FORMAT"     // 書式のみ
	PasteTypeNoBorders PasteType = "PASTE_NO_BORDERS" // 罫線以外すべて
)

type MergeType string

const (
	MergeTypeAll  MergeType = "MERGE_ALL"
	MergeTypeRows MergeType = "MERGE_ROWS"
	MergeTypeCols MergeType = "MERGE_COLS"
)

// ShiftDimension defines the direction of the operation (e.g. for DeleteRange).
type ShiftDimensionType string

const (
	ShiftDimensionTypeRows    ShiftDimensionType = "ROWS"    // 行方向（削除したら下から上に詰める）
	ShiftDimensionTypeColumns ShiftDimensionType = "COLUMNS" // 列方向（削除したら右から左に詰める）
)

type Rect struct {
	Row    int
	Col    int
	Height int
	Width  int
}

// SheetRow is a Row specialized for spreadsheet cells (values, formulas, etc).
type SheetRow = Row[any]

// NewSheetRow
func NewSheetRow(capacity int) *SheetRow {
	return NewRow[any](capacity)
}

// OffsetSheetRow is a OffsetRow specialized for spreadsheet cells (values, formulas, etc).
type OffsetSheetRow = OffsetRow[any]

// NewSheetRow
func NewOffsetSheetRow(offset int, capacity int) *OffsetSheetRow {
	return NewOffsetRow[any](offset, capacity)
}
