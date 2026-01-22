package haresheet

// OffsetRow is a Row wrapper that translates absolute indices by subtracting an offset.
type OffsetRow[T any] struct {
	r      *Row[T]
	offset int
}

// NewOffsetRow creates an OffsetRow with the given offset and optional initial capacity.
func NewOffsetRow[T any](offset int, cap int) *OffsetRow[T] {
	if offset < 0 {
		panic("offsetrow.NewOffsetRow: negative offset")
	}

	return &OffsetRow[T]{
		r:      NewRow[T](cap),
		offset: offset,
	}
}

// Set sets a value at index i (panics on out-of-range like a slice assignment).
func (or *OffsetRow[T]) Set(i int, v T) {
	j := i - or.offset

	if j < 0 {
		panic("offsetrow.Set: negative index")
	}

	or.r.Set(j, v)
}

// Sets writes values starting at start, growing the row if needed, and returns r for chaining.
func (or *OffsetRow[T]) Sets(start int, values ...T) *OffsetRow[T] {
	j := start - or.offset

	if j < 0 {
		panic("offsetrow.Sets: negative index")
	}

	or.r.Sets(j, values...)

	return or
}

// Put sets v at index i, growing the row to i+1 if needed.
func (or *OffsetRow[T]) Put(i int, v T) {
	j := i - or.offset

	if j < 0 {
		panic("offsetrow.Put: negative index")
	}

	or.r.Put(j, v)
}

// Puts writes values starting at start, growing as needed.
func (or *OffsetRow[T]) Puts(start int, values ...T) *OffsetRow[T] {
	j := start - or.offset

	if j < 0 {
		panic("offsetrow.Puts: negative index")
	}

	or.r.Puts(j, values...)

	return or
}

// At returns the value at index i (panics on out-of-range like a slice access).
func (or *OffsetRow[T]) At(i int) T {
	j := i - or.offset

	if j < 0 {
		panic("offsetrow.At: negative index")
	}

	return or.r.At(j)
}

// Get returns the value at index i and whether it exists.
func (or *OffsetRow[T]) Get(i int) (T, bool) {
	j := i - or.offset

	if j < 0 {
		var zero T

		return zero, false
	}

	return or.r.Get(j)
}

// Offset returns the offset used to translate external indices.
func (or *OffsetRow[T]) Offset() int {
	return or.offset
}

// --------------------------------------------------------------------------------
// safe passthroughs (no index arguments)
// --------------------------------------------------------------------------------

// Len returns the current length of the underlying row.
func (or *OffsetRow[T]) Len() int { return or.r.Len() }

// Cap returns the current capacity of the underlying row.
func (or *OffsetRow[T]) Cap() int { return or.r.Cap() }

// Grow extends the underlying row by n zero-value elements.
func (or *OffsetRow[T]) Grow(n int) int { return or.r.Grow(n) }

// Clear resets the underlying row length to zero, keeping capacity for reuse.
func (or *OffsetRow[T]) Clear() { or.r.Clear() }

// Append appends a value to the end of the underlying row.
func (or *OffsetRow[T]) Append(v T) *OffsetRow[T] { or.r.Append(v); return or }

// Slice returns the underlying slice of the row.
func (or *OffsetRow[T]) Slice() []T { return or.r.Slice() }

// Values returns a copy of the current contents of the row.
func (or *OffsetRow[T]) Values() []T { return or.r.Values() }
