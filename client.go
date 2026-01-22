package haresheet

import (
	"context"
	"fmt"

	"google.golang.org/api/sheets/v4"
)

// SheetInfo holds metadata for a specific sheet.
type SheetInfo struct {
	ID    int64
	Index int
	Title string
}

type Client struct {
	service  *sheets.Service
	spreadID string
}

func NewClient(service *sheets.Service, spreadsheetID string) *Client {
	return &Client{
		service:  service,
		spreadID: spreadsheetID,
	}
}

// GetSheetInfoMap retrieves a map of sheet names to their info (ID and Index).
// It fetches only the necessary properties to ensure high performance.
func (c *Client) GetSheetInfoMap(ctx context.Context) (map[string]*SheetInfo, error) {
	resp, err := c.service.Spreadsheets.Get(c.spreadID).
		Fields("sheets(properties(sheetId,title,index))").
		Context(ctx).
		Do()

	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet info: %w", err)
	}

	m := make(map[string]*SheetInfo)

	for _, sheet := range resp.Sheets {
		m[sheet.Properties.Title] = &SheetInfo{
			ID:    sheet.Properties.SheetId,
			Index: int(sheet.Properties.Index),
			Title: sheet.Properties.Title,
		}
	}

	return m, nil
}

// Sheet
func (c *Client) Sheet(sheetID int64) *SheetClient {
	sc := &SheetClient{
		c:       c,
		sheetID: sheetID,
		err:     nil,
	}

	if sheetID < 0 {
		sc.err = fmt.Errorf("Sheet: invalid sheet ID: %d", sheetID)

		return sc
	}

	return sc
}
