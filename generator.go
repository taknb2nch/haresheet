package haresheet

import (
	"math/rand"
	"time"
)

// IDGenerator manages sheet IDs to avoid duplication.
type IDGenerator struct {
	usedIDs map[int64]bool
	rng     *rand.Rand
}

// NewIDGenerator creates a new generator with existing IDs marked as used.
func NewIDGenerator(existingIDs ...int64) *IDGenerator {
	g := &IDGenerator{
		usedIDs: make(map[int64]bool),
		rng:     rand.New(rand.NewSource(time.Now().UnixNano())),
	}

	// 既存のIDを使用済みとして登録
	for _, id := range existingIDs {
		g.usedIDs[id] = true
	}

	return g
}

// Next generates a unique, unused sheet ID.
func (g *IDGenerator) Next() int64 {
	const maxRetries = 10000

	for range maxRetries {
		id := int64(g.rng.Int31())

		if id != 0 && !g.usedIDs[id] {
			g.usedIDs[id] = true

			return id
		}
	}

	panic("failed to generate unique ID: all attempts failed")
}
