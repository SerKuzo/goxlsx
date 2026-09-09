package goxlsx

import (
	"fmt"

	"github.com/Puzzanis/goxlsx.git/internal/xmltree"
)

// GetCellValue returns the value of a cell addressed with A1 or R1C1 notation.
// Empty and missing cells are returned as an empty string without an error.
func (d *ExcelDoc) GetCellValue(sheetName, cellRef string) (string, error) {
	if d == nil || d.content == nil {
		return "", fmt.Errorf("document is not open")
	}

	column, row, err := xmltree.ParseCellRef(cellRef)
	if err != nil {
		return "", err
	}

	rows, err := d.GetRows(sheetName)
	if err != nil {
		return "", err
	}
	if row >= len(rows) || column >= len(rows[row]) {
		return "", nil
	}
	return rows[row][column], nil
}
