package xmltree

import (
	"fmt"
	"strconv"
	"strings"
)

type MergeRange struct {
	MinCol, MaxCol int
	MinRow, MaxRow int
	Value          string // Сюда мы скопируем значение из главной ячейки
}

// ParseRange переводит "A1:C3" в числовые границы
func ParseRange(ref string) (MergeRange, error) {
	parts := strings.Split(strings.TrimSpace(ref), ":")
	if len(parts) != 2 {
		return MergeRange{}, fmt.Errorf("неверный формат диапазона")
	}

	minC, minR, err := ParseCellRef(parts[0])
	if err != nil {
		return MergeRange{}, err
	}
	maxC, maxR, err := ParseCellRef(parts[1])
	if err != nil {
		return MergeRange{}, err
	}
	if minC > maxC || minR > maxR {
		return MergeRange{}, fmt.Errorf("границы диапазона указаны в обратном порядке")
	}

	return MergeRange{
		MinCol: minC, MaxCol: maxC,
		MinRow: minR, MaxRow: maxR,
	}, nil
}

func ColumnTitleToIndex(char string) int {
	index := 0
	for _, r := range char {
		index = index*26 + int(r-'A') + 1
	}
	return index - 1 // Возвращаем 0-based индекс
}

// Извлекает только буквы из адреса ячейки (например, "AC123" -> "AC")
func GetColumnName(cellRef string) string {
	for i, r := range cellRef {
		if r >= '0' && r <= '9' {
			return cellRef[:i]
		}
	}
	return cellRef
}

// ParseCellRef converts an A1 or R1C1 reference to zero-based column and row
// indexes. A1 references may optionally contain absolute markers ('$').
func ParseCellRef(ref string) (col, row int, err error) {
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return 0, 0, fmt.Errorf("пустой адрес ячейки")
	}

	if strings.HasPrefix(strings.ToUpper(ref), "R") {
		upper := strings.ToUpper(ref)
		if rIndex := strings.IndexByte(upper, 'C'); rIndex > 1 {
			rowText := upper[1:rIndex]
			colText := upper[rIndex+1:]
			parsedRow, rowErr := strconv.Atoi(rowText)
			parsedCol, colErr := strconv.Atoi(colText)
			if rowErr == nil && colErr == nil && parsedRow > 0 && parsedCol > 0 {
				return parsedCol - 1, parsedRow - 1, nil
			}
		}
	}

	ref = strings.TrimPrefix(ref, "$")
	separator := 0
	for separator < len(ref) && ((ref[separator] >= 'A' && ref[separator] <= 'Z') ||
		(ref[separator] >= 'a' && ref[separator] <= 'z')) {
		separator++
	}
	if separator == 0 || separator == len(ref) {
		return 0, 0, fmt.Errorf("неверный адрес ячейки %q", ref)
	}

	columnText := strings.ToUpper(ref[:separator])
	rowText := ref[separator:]
	if strings.Contains(rowText, "$") {
		rowText = strings.ReplaceAll(rowText, "$", "")
	}
	parsedRow, rowErr := strconv.Atoi(rowText)
	if rowErr != nil || parsedRow <= 0 {
		return 0, 0, fmt.Errorf("неверный номер строки в адресе %q", ref)
	}

	parsedCol := ColumnTitleToIndex(columnText)
	if parsedCol < 0 {
		return 0, 0, fmt.Errorf("неверный столбец в адресе %q", ref)
	}
	return parsedCol, parsedRow - 1, nil
}

// ColumnIndexToTitle converts a zero-based column index to an Excel title.
func ColumnIndexToTitle(index int) (string, error) {
	if index < 0 {
		return "", fmt.Errorf("индекс столбца не может быть отрицательным")
	}

	var result []byte
	for index >= 0 {
		result = append([]byte{byte('A' + index%26)}, result...)
		index = index/26 - 1
	}
	return string(result), nil
}
