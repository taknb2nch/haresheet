package haresheet

import (
	"context"
	"fmt"
	"time"

	"google.golang.org/api/sheets/v4"
)

// RequestUnitInfo
type RequestUnitInfo struct {
	Label      string
	StartIndex int
	EndIndex   int
	Count      int
}

// BatchUpdateError
type BatchUpdateError struct {
	Err   error
	Units []*RequestUnitInfo
}

// Error
func (e *BatchUpdateError) Error() string {
	return fmt.Sprintf("haresheet: request execution failed: %v", e.Err)
}

// Unwrap
func (e *BatchUpdateError) Unwrap() error {
	return e.Err
}

// FlushStatus
type FlushStatus struct {
	RequestCount  int                // request count
	Units         []*RequestUnitInfo // units
	SpreadsheetID string             // spreadsheet id
}

// ClientTrace
type ClientTrace struct {
	//
	OnFlushStart func(ctx context.Context, status *FlushStatus)
	//
	OnFlushDone func(ctx context.Context, status *FlushStatus, duration time.Duration, err error)
}

// BatchUpdateExecutor
type BatchUpdateExecutor struct {
	service  *sheets.Service
	spreadID string
	requests []*sheets.Request
	units    []*RequestUnitInfo
	limit    int
	err      error
	Trace    *ClientTrace
}

// NewBatchUpdateExecutor
func NewBatchUpdateExecutor(service *sheets.Service, spreadsheetID string, limit int) *BatchUpdateExecutor {
	if limit <= 0 {
		limit = 100
	}

	return &BatchUpdateExecutor{
		service:  service,
		spreadID: spreadsheetID,
		requests: make([]*sheets.Request, 0, 100),
		units:    make([]*RequestUnitInfo, 0, 100),
		limit:    limit,
	}
}

// Queue
func (e *BatchUpdateExecutor) Queue(ctx context.Context, reqs []*sheets.Request, label string) *BatchUpdateExecutor {
	if e.err != nil {
		return e
	}

	currentCount := len(e.requests)
	newCount := len(reqs)

	if currentCount > 0 && (currentCount+newCount) > e.limit {
		err := e.Flush(ctx)
		if err != nil {
			e.err = err

			return e
		}

		currentCount = 0
	}

	startIndex := currentCount
	endIndex := currentCount + newCount - 1

	if label == "" {
		label = fmt.Sprintf("Unit-%d [%d - %d]", len(e.units)+1, startIndex, endIndex)
	}

	e.units = append(e.units, &RequestUnitInfo{
		Label:      label,
		StartIndex: startIndex,
		EndIndex:   endIndex,
		Count:      newCount,
	})

	e.requests = append(e.requests, reqs...)

	return e
}

// Flush
func (e *BatchUpdateExecutor) Flush(ctx context.Context) error {
	if len(e.requests) == 0 {
		return nil
	}

	labels := make([]string, 0, len(e.units))

	for _, u := range e.units {
		labels = append(labels, u.Label)
	}

	status := &FlushStatus{
		RequestCount:  len(e.requests),
		Units:         e.units,
		SpreadsheetID: e.spreadID,
	}

	if e.Trace != nil && e.Trace.OnFlushStart != nil {
		e.Trace.OnFlushStart(ctx, status)
	}

	start := time.Now()

	_, err := e.service.Spreadsheets.BatchUpdate(e.spreadID, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: e.requests,
	}).Context(ctx).Do()

	duration := time.Since(start)

	if e.Trace != nil && e.Trace.OnFlushDone != nil {
		// エラーも時間もここで渡す！
		e.Trace.OnFlushDone(ctx, status, duration, err)
	}

	if err != nil {
		return &BatchUpdateError{
			Err:   err,
			Units: e.units,
		}
	}

	// バッファが大きくなりすぎた場合に解放するためnilをいれる
	e.requests = nil
	e.units = nil

	return nil
}
