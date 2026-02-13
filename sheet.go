package goxlsx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"goxlsx/internal/xmltree"
	"io"
	"strconv"
)

func (d *ExcelDoc) GetSheets() ([]xmltree.Sheet, error) {
	// Ищем файл во внутреннем индексе
	file, ok := d.files["xl/workbook.xml"]
	if !ok {
		return nil, fmt.Errorf("workbook.xml не найден в архиве")
	}

	// Открываем файл внутри ZIP
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	// Декодируем XML
	var wb xmltree.Workbook
	if err := xml.NewDecoder(rc).Decode(&wb); err != nil {
		return nil, fmt.Errorf("ошибка парсинга workbook: %w", err)
	}

	return wb.Sheets, nil
}

// GetRows возвращает все данные листа в виде двумерного слайса строк.
// На вход принимает путь к файлу листа в архиве (например, "xl/worksheets/sheet1.xml")
func (d *ExcelDoc) GetRows(sheetName string) ([][]string, error) {
	// Находим файл листа в нашем индексе (который создали в Open)
	file, ok := d.files[d.sheetMap[sheetName]]
	if !ok {
		return nil, fmt.Errorf("sheet not found")
	}

	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	data, _ := io.ReadAll(rc)
	var result [][]string

	// Ищем все строки <row ... >
	rowSplit := bytes.Split(data, []byte("<row"))
	for i, rowData := range rowSplit {
		if i == 0 {
			continue
		} // Пропускаем всё до первой строки

		var currentRow []string
		// Ищем все ячейки <c ... > внутри строки
		cellSplit := bytes.Split(rowData, []byte("<c "))
		for j, cellData := range cellSplit {
			if j == 0 {
				continue
			}

			// 1. Проверяем, является ли значение индексом строки (t="s")
			isShared := bytes.Contains(cellData, []byte(`t="s"`))

			// 2. Ищем адрес ячейки (r="A1"), чтобы не было смещения колонок
			refIdx := bytes.Index(cellData, []byte(`r="`))
			if refIdx != -1 {
				refEnd := bytes.Index(cellData[refIdx+3:], []byte(`"`))
				cellRef := string(cellData[refIdx+3 : refIdx+3+refEnd])

				colIdx := columnTitleToIndex(getColumnName(cellRef))
				for len(currentRow) < colIdx {
					currentRow = append(currentRow, "")
				}
			}

			// 3. Достаем значение из <v>...</v>
			vStart := bytes.Index(cellData, []byte("<v>"))
			if vStart != -1 {
				vEnd := bytes.Index(cellData[vStart:], []byte("</v>"))
				val := string(cellData[vStart+3 : vStart+vEnd])

				if isShared {
					idx, _ := strconv.Atoi(val)
					if idx >= 0 && idx < len(d.sharedStrings) {
						val = d.sharedStrings[idx]
					}
				}
				currentRow = append(currentRow, val)
			}
		}
		if len(currentRow) > 0 {
			result = append(result, currentRow)
		}
	}
	return result, nil
}

func columnTitleToIndex(char string) int {
	index := 0
	for _, r := range char {
		index = index*26 + int(r-'A') + 1
	}
	return index - 1 // Возвращаем 0-based индекс
}

// Извлекает только буквы из адреса ячейки (например, "AC123" -> "AC")
func getColumnName(cellRef string) string {
	for i, r := range cellRef {
		if r >= '0' && r <= '9' {
			return cellRef[:i]
		}
	}
	return cellRef
}
