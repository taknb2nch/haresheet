package haresheet

import (
	"encoding/hex"
	"errors"
	"fmt"
	"strconv"

	"google.golang.org/api/sheets/v4"
)

var ErrNegativeIndex = errors.New("IndexToA1: negative index")

// AbsMode controls whether column/row are absolute ($) in A1 notation.
type AbsMode uint8

const (
	AbsNone AbsMode           = 0
	AbsCol  AbsMode           = 1 << iota // $A1
	AbsRow                                // A$1
	AbsBoth = AbsCol | AbsRow             // $A$1
)

// IndexToA1 converts a 0-based row/column index with offsets into A1 notation.
func IndexToA1(row int, col int, rowOffset int, colOffset int) (string, error) {
	r := row + rowOffset
	c := col + colOffset

	if r < 0 || c < 0 {
		return "", ErrNegativeIndex
	}

	// col letters (0-based) + row number (1-based)
	colLetters := ColIndexToLetters(c)

	// Build result with minimal allocations.
	// Max column letters for int range is small; allocate enough.
	b := make([]byte, 0, len(colLetters)+10)
	b = append(b, colLetters...)
	b = strconv.AppendInt(b, int64(r+1), 10)

	return string(b), nil
}

// MustIndexToA1 is the must variant of IndexToA1 and panics on invalid indexes.
func MustIndexToA1(row int, col int, rowOffset int, colOffset int) string {
	s, err := IndexToA1(row, col, rowOffset, colOffset)

	if err != nil {
		panic(err)
	}

	return s
}

// IndexToA1At converts a 0-based row/column index into A1 notation without offsets.
func IndexToA1At(row, col int) (string, error) {
	return IndexToA1(row, col, 0, 0)
}

// MustIndexToA1At is the must variant of IndexToA1At and panics on invalid indexes.
func MustIndexToA1At(row, col int) string {
	return MustIndexToA1(row, col, 0, 0)
}

// IndexToA1Abs converts a 0-based row/column index with offsets into A1 notation,
// optionally adding $ for absolute column/row.
func IndexToA1Abs(row, col, rowOffset, colOffset int, abs AbsMode) (string, error) {
	r := row + rowOffset
	c := col + colOffset

	if r < 0 || c < 0 {
		return "", ErrNegativeIndex
	}

	colLetters := ColIndexToLetters(c)

	// Build result with minimal allocations.
	// Capacity: [$] + col + [$] + row digits
	b := make([]byte, 0, 1+len(colLetters)+1+10)

	if abs&AbsCol != 0 {
		b = append(b, '$')
	}

	b = append(b, colLetters...)

	if abs&AbsRow != 0 {
		b = append(b, '$')
	}

	b = strconv.AppendInt(b, int64(r+1), 10)

	return string(b), nil
}

// MustIndexToA1Abs is the must variant of IndexToA1Abs and panics on invalid indexes.
func MustIndexToA1Abs(row, col, rowOffset, colOffset int, abs AbsMode) string {
	s, err := IndexToA1Abs(row, col, rowOffset, colOffset, abs)
	if err != nil {
		panic(err)
	}
	return s
}

// IndexToA1AtAbs converts a 0-based row/column index into A1 notation without offsets,
// optionally adding $ for absolute column/row.
func IndexToA1AtAbs(row, col int, abs AbsMode) (string, error) {
	return IndexToA1Abs(row, col, 0, 0, abs)
}

// MustIndexToA1AtAbs is the must variant of IndexToA1AtAbs and panics on invalid indexes.
func MustIndexToA1AtAbs(row, col int, abs AbsMode) string {
	return MustIndexToA1Abs(row, col, 0, 0, abs)
}

// ColIndexToLetters converts 0-based column index to letters (A..Z, AA..).
func ColIndexToLetters(col int) []byte {
	// We build in reverse then reverse once.
	// For typical sheet sizes this is tiny (<= 3-4 chars).
	var tmp [8]byte

	i := len(tmp)

	for col >= 0 {
		i--

		tmp[i] = byte('A' + (col % 26))

		col = col/26 - 1
	}

	// Copy to a new slice of exact size.
	out := make([]byte, len(tmp)-i)

	copy(out, tmp[i:])

	return out
}

// ParseHexColor parses a hex string (e.g. "#FFFFFF" or "FFFFFF") to *sheets.Color.
func ParseHexColor(s string) (*sheets.Color, error) {
	if len(s) > 0 && s[0] == '#' {
		s = s[1:]
	}

	if len(s) != 6 {
		return nil, fmt.Errorf("invalid hex color code: %s", s)
	}

	// 文字列を数値に変換
	rgb, err := hex.DecodeString(s)
	if err != nil {
		return nil, err
	}

	return &sheets.Color{
		Red:   float64(rgb[0]) / 255.0,
		Green: float64(rgb[1]) / 255.0,
		Blue:  float64(rgb[2]) / 255.0,
		Alpha: 1.0,
	}, nil
}

// MustParseHexColor parses a hex string and panics if invalid.
// Use this for hardcoded color constants.
func MustParseHexColor(s string) *sheets.Color {
	c, err := ParseHexColor(s)
	if err != nil {
		panic(fmt.Sprintf("invalid hex color: %s", s))
	}

	return c
}
