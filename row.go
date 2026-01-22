package haresheet

// Row is a growable row builder that supports reserving slots and filling them later.
type Row[T any] struct {
	s []T
}

// NewRow returns a new Row with optional initial capacity (0 means no preallocation).
func NewRow[T any](capacity int) *Row[T] {
	if capacity < 0 {
		panic("row.NewRow: negative capacity")
	}

	r := &Row[T]{}

	if capacity > 0 {
		r.s = make([]T, 0, capacity)
	}

	return r
}

// Len returns the current length of the row.
func (r *Row[T]) Len() int {
	return len(r.s)
}

// Cap returns the current capacity of the row.
func (r *Row[T]) Cap() int {
	return cap(r.s)
}

// Grow extends the length by n zero-values; Grow(0) is a no-op and returns the current length.
func (r *Row[T]) Grow(n int) (start int) {
	if n < 0 {
		panic("row.Grow: negative n")
	}

	start = len(r.s)

	if n == 0 {
		return start
	}

	need := start + n

	if need > cap(r.s) {
		ns := make([]T, start, need)

		copy(ns, r.s)

		r.s = ns
	}

	r.s = r.s[:need] // new region is zero-values

	return start
}

// Clear resets the length to zero while keeping the underlying capacity for reuse.
func (r *Row[T]) Clear() {
	r.s = r.s[:0]
}

// Append appends one value to the end and returns r for chaining.
func (r *Row[T]) Append(v T) *Row[T] {
	r.s = append(r.s, v)

	return r
}

// Set sets a value at index i (panics on out-of-range like a slice assignment).
func (r *Row[T]) Set(i int, v T) {
	r.s[i] = v
}

// Sets writes values starting at start, growing the row if needed, and returns r for chaining.
func (r *Row[T]) Sets(start int, values ...T) *Row[T] {
	if start < 0 {
		panic("row.Sets: negative start")
	}

	if len(values) == 0 {
		return r
	}

	need := start + len(values)

	if need > len(r.s) {
		r.Grow(need - len(r.s))
	}

	copy(r.s[start:need], values)

	return r
}

// Put sets v at index i, growing the row to i+1 if needed.
func (r *Row[T]) Put(i int, v T) {
	if i < 0 {
		panic("row.Put: negative index")
	}

	if i >= len(r.s) {
		r.Grow(i + 1 - len(r.s))
	}

	r.s[i] = v
}

// Puts writes values starting at start, growing as needed.
func (r *Row[T]) Puts(start int, values ...T) *Row[T] {
	if start < 0 {
		panic("row.Puts: negative start")
	}

	if len(values) == 0 {
		return r
	}

	need := start + len(values)

	if need > len(r.s) {
		r.Grow(need - len(r.s))
	}

	copy(r.s[start:need], values)

	return r
}

// At returns the value at index i (panics on out-of-range like a slice access).
func (r *Row[T]) At(i int) T {
	return r.s[i]
}

// Get returns the value at index i and whether it exists.
func (r *Row[T]) Get(i int) (T, bool) {
	var zero T

	if i < 0 || i >= len(r.s) {
		return zero, false
	}

	return r.s[i], true
}

// Slice returns the underlying slice of the row.
//
// The returned slice aliases the Row's internal storage and uses relative indices
// starting at 0. Mutating the slice may break Row invariants.
// This method is intended for advanced or performance-critical use.
func (r *Row[T]) Slice() []T {
	return r.s
}

// Values returns a copy of the current contents of the row.
//
// The returned slice is safe to modify and uses relative indices starting at 0.
// Modifying the returned slice does not affect the Row.
func (r *Row[T]) Values() []T {
	out := make([]T, len(r.s))

	copy(out, r.s)

	return out
}
