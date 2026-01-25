package haresheet

import (
	"errors"
	"fmt"
	"strings"

	"google.golang.org/api/sheets/v4"
)

// Builder builds a batch of requests for Google Sheets API.
type Builder struct {
	requests []*sheets.Request
	errs     []error

	props      *sheets.SpreadsheetProperties
	propFields []string // "title", "locale", etc...
}

// NewBuilder creates a new Builder instance.
func NewBuilder() *Builder {
	return &Builder{
		requests:   make([]*sheets.Request, 0, 100),
		errs:       make([]error, 0, 10),
		propFields: make([]string, 0, 5),
	}
}

// Requests
func (b *Builder) Requests() ([]*sheets.Request, error) {
	if len(b.errs) > 0 {
		return nil, errors.Join(b.errs...)
	}

	finalRequests := b.requests

	if b.props != nil && len(b.propFields) > 0 {
		req := &sheets.Request{
			UpdateSpreadsheetProperties: &sheets.UpdateSpreadsheetPropertiesRequest{
				Properties: b.props,
				Fields:     strings.Join(b.propFields, ","),
			},
		}

		// * create new slice
		finalRequests = append([]*sheets.Request{req}, finalRequests...)
	}

	return finalRequests, nil
}

// AppendRequest
func (b *Builder) AppendRequest(request *sheets.Request) {
	if request == nil {
		return
	}

	b.requests = append(b.requests, request)
}

// PrependRequest
func (b *Builder) PrependRequest(request *sheets.Request) {
	if request == nil {
		return
	}

	b.requests = append([]*sheets.Request{request}, b.requests...)
}

// addError appends an error to the list.
func (b *Builder) appendError(err error) {
	if err != nil {
		b.errs = append(b.errs, err)
	}
}

func (b *Builder) ensureProps() {
	if b.props == nil {
		b.props = &sheets.SpreadsheetProperties{}
	}
}

// Title sets the spreadsheet title.
func (b *Builder) Title(title string) *Builder {
	b.ensureProps()
	b.props.Title = title
	b.propFields = append(b.propFields, "title")

	return b
}

// Locale sets the locale (e.g., "ja_JP").
// Important for date formatting and currency symbols.
func (b *Builder) Locale(locale string) *Builder {
	b.ensureProps()
	b.props.Locale = locale
	b.propFields = append(b.propFields, "locale")

	return b
}

// TimeZone sets the time zone (e.g., "Asia/Tokyo").
func (b *Builder) TimeZone(tz string) *Builder {
	b.ensureProps()
	b.props.TimeZone = tz
	b.propFields = append(b.propFields, "timeZone")

	return b
}

// Sheet
func (b *Builder) Sheet(sheetID int64) *SheetBuilder {
	sb := &SheetBuilder{
		b:       b,
		sheetID: sheetID,
	}

	if sheetID < 0 {
		b.appendError(fmt.Errorf("Sheet: invalid sheet ID: %d", sheetID))

		return sb
	}

	return sb
}

// AddSheet adds a request to create a new sheet with a SPECIFIC ID and INDEX.
// Pass index: -1 to append to the end.
func (b *Builder) AddSheet(sheetID int64, title string, index int) *Builder {
	props := &sheets.SheetProperties{
		SheetId: sheetID,
		Title:   title,
	}

	if index >= 0 {
		props.Index = int64(index)
		props.ForceSendFields = []string{"Index"}
	}

	req := &sheets.Request{
		AddSheet: &sheets.AddSheetRequest{
			Properties: props,
		},
	}

	b.requests = append(b.requests, req)

	return b
}

// CopySheet copies a source sheet to create multiple new sheets.
func (b *Builder) CopySheet(srcSheetID int64, toIndex int, newNames ...string) *Builder {
	if srcSheetID < 0 {
		b.appendError(fmt.Errorf("CopySheet: invalid src sheet ID: %d", srcSheetID))

		return b
	}

	if toIndex < -1 {
		b.appendError(fmt.Errorf("CopySheet: invalid to index: %d", toIndex))

		return b
	}

	if len(newNames) == 0 {
		b.appendError(errors.New("CopySheet: no sheet names specified"))

		return b
	}

	currentIndex := toIndex

	for _, name := range newNames {
		dsRequest := &sheets.DuplicateSheetRequest{
			SourceSheetId: srcSheetID,
			NewSheetName:  name,
		}

		// -1 (末尾追加) 以外のときだけインデックスを指定してインクリメント
		if toIndex >= 0 {
			dsRequest.InsertSheetIndex = int64(currentIndex)

			currentIndex++
		}

		req := &sheets.Request{
			DuplicateSheet: dsRequest,
		}

		b.AppendRequest(req)
	}

	return b
}

// CopySheetWithID copies a sheet using a specific new ID.
// This allows you to reference the new sheet immediately in the same batch.
func (b *Builder) CopySheetWithID(srcSheetID int64, toIndex int, newSheetID int64, newName string) *Builder {
	if srcSheetID < 0 {
		b.appendError(fmt.Errorf("CopySheetWithID: invalid src sheet id: %d", srcSheetID))

		return b
	}

	if toIndex < -1 {
		b.appendError(fmt.Errorf("CopySheetWithID: invalid to index: %d", toIndex))

		return b
	}

	if newSheetID < 0 {
		b.appendError(fmt.Errorf("CopySheetWithID: invalid new sheet id: %d", toIndex))

		return b
	}

	if newName == "" {
		b.appendError(errors.New("CopySheetWithID: invalid new sheet name"))

		return b
	}

	dsRequest := &sheets.DuplicateSheetRequest{
		SourceSheetId: srcSheetID,
		NewSheetId:    newSheetID,
		NewSheetName:  newName,
	}

	if toIndex >= 0 {
		dsRequest.InsertSheetIndex = int64(toIndex)
	}

	req := &sheets.Request{
		DuplicateSheet: dsRequest,
	}

	b.AppendRequest(req)

	return b
}

// DeleteSheet deletes one or more sheets by their IDs.
func (b *Builder) DeleteSheet(sheetIDs ...int64) *Builder {
	if len(sheetIDs) == 0 {
		return b
	}

	for _, id := range sheetIDs {
		req := &sheets.Request{
			DeleteSheet: &sheets.DeleteSheetRequest{
				SheetId: id,
			},
		}

		b.AppendRequest(req)
	}

	return b
}
