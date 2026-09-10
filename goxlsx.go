// Package goxlsx provides reading and writing of XLSX workbooks.
//
// The implementation lives in pkg/goxlsx. This package keeps the original
// module-root import path available for existing users.
package goxlsx

import (
	"github.com/SerKuzo/goxlsx.git/internal/xmltree"
	implementation "github.com/SerKuzo/goxlsx.git/pkg/goxlsx"
)

// ExcelDoc is the workbook handle used by the goxlsx API.
type ExcelDoc = implementation.ExcelDoc

// Sheet describes a worksheet exposed by GetSheets.
type Sheet = xmltree.Sheet

// MergeRange describes the zero-based bounds of a merged cell range.
type MergeRange = xmltree.MergeRange

// Open opens an XLSX workbook for reading or writing.
func Open(path string) (*ExcelDoc, error) {
	return implementation.Open(path)
}
