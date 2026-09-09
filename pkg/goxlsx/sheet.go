package goxlsx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"html"
	"strconv"
	"strings"

	"github.com/Puzzanis/goxlsx.git/internal/xmltree"
)

func (d *ExcelDoc) GetSheets() ([]xmltree.Sheet, error) {
	if d == nil || d.content == nil {
		return nil, fmt.Errorf("документ не открыт")
	}
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
	if d == nil || d.content == nil {
		return nil, fmt.Errorf("документ не открыт")
	}
	path, ok := d.sheetMap[sheetName]
	if !ok {
		return nil, fmt.Errorf("sheet not found")
	}

	data, err := d.sheetBytes(path)
	if err != nil {
		return nil, err
	}

	return parseRows(data, d.sharedStrings)
}

func parseRows(data []byte, sharedStrings []string) ([][]string, error) {
	var result [][]string
	position := 0
	lastRowIndex := -1
	for {
		rowStart, rowOpenEnd := findStartTag(data, position, "row")
		if rowStart < 0 {
			break
		}
		rowClose := rowOpenEnd
		rowIsSelfClosing := bytes.HasSuffix(bytes.TrimSpace(data[rowStart:rowOpenEnd]), []byte("/>"))
		if !rowIsSelfClosing {
			rowClose = findClosingTag(data, rowOpenEnd, "row")
			if rowClose < 0 {
				return nil, fmt.Errorf("незакрытая строка листа")
			}
		}

		rowNumber, _ := strconv.Atoi(attributeValueRaw(data[rowStart:rowOpenEnd], "r"))
		rowIndex := rowNumber - 1
		if rowIndex < 0 {
			rowIndex = lastRowIndex + 1
		}
		for len(result) <= rowIndex {
			result = append(result, nil)
		}
		if rowIndex > lastRowIndex {
			lastRowIndex = rowIndex
		}

		parseCells(data[rowOpenEnd:rowClose], sharedStrings, &result[rowIndex])
		if rowIsSelfClosing {
			position = rowOpenEnd
		} else {
			position = findTagEnd(data, rowClose) + 1
		}
	}
	return result, nil
}

func parseCells(row []byte, sharedStrings []string, output *[]string) {
	position := 0
	for {
		cellStart, cellOpenEnd := findStartTag(row, position, "c")
		if cellStart < 0 {
			return
		}
		cellEnd := findTagEnd(row, cellStart)
		if cellEnd < 0 {
			return
		}
		cellDataEnd := cellEnd + 1
		cellData := row[cellOpenEnd:cellDataEnd]
		if !bytes.HasSuffix(bytes.TrimSpace(row[cellStart:cellOpenEnd]), []byte("/>")) {
			closing := findClosingTag(row, cellOpenEnd, "c")
			if closing < 0 {
				return
			}
			cellData = row[cellOpenEnd:closing]
			cellDataEnd = findTagEnd(row, closing) + 1
		}

		columnIndex, _, err := xmltree.ParseCellRef(attributeValueRaw(row[cellStart:cellOpenEnd], "r"))
		if err == nil {
			for len(*output) <= columnIndex {
				*output = append(*output, "")
			}
			value := cellValueRaw(cellData, attributeValueRaw(row[cellStart:cellOpenEnd], "t"), sharedStrings)
			(*output)[columnIndex] = value
		}
		position = cellDataEnd
	}
}

func cellValueRaw(data []byte, cellType string, sharedStrings []string) string {
	value := tagText(data, "v")
	if cellType == "inlineStr" {
		value = inlineText(data)
	}
	value = html.UnescapeString(value)
	if cellType == "s" {
		index, err := strconv.Atoi(value)
		if err == nil && index >= 0 && index < len(sharedStrings) {
			return sharedStrings[index]
		}
	}
	return value
}

func inlineText(data []byte) string {
	var result strings.Builder
	position := 0
	for {
		start, openEnd := findStartTag(data, position, "t")
		if start < 0 {
			return result.String()
		}
		close := findClosingTag(data, openEnd, "t")
		if close < 0 {
			return result.String()
		}
		result.WriteString(html.UnescapeString(string(data[openEnd:close])))
		position = findTagEnd(data, close) + 1
	}
}

func tagText(data []byte, name string) string {
	start, openEnd := findStartTag(data, 0, name)
	if start < 0 {
		return ""
	}
	close := findClosingTag(data, openEnd, name)
	if close < 0 {
		return ""
	}
	return string(data[openEnd:close])
}

func findStartTag(data []byte, from int, name string) (start, end int) {
	needle := []byte("<" + name)
	for from < len(data) {
		relative := bytes.Index(data[from:], needle)
		if relative < 0 {
			return -1, -1
		}
		start = from + relative
		boundary := start + len(needle)
		if boundary < len(data) && isTagBoundary(data[boundary]) {
			end = findTagEnd(data, start)
			if end >= 0 {
				return start, end + 1
			}
			return -1, -1
		}
		from = boundary
	}
	return -1, -1
}

func findClosingTag(data []byte, from int, name string) int {
	needle := []byte("</" + name)
	for from < len(data) {
		relative := bytes.Index(data[from:], needle)
		if relative < 0 {
			return -1
		}
		start := from + relative
		boundary := start + len(needle)
		if boundary < len(data) && (data[boundary] == '>' || data[boundary] == ' ' || data[boundary] == '\t' || data[boundary] == '\r' || data[boundary] == '\n') {
			return start
		}
		from = boundary
	}
	return -1
}

func findTagEnd(data []byte, from int) int {
	quote := byte(0)
	for index := from; index < len(data); index++ {
		switch data[index] {
		case '\'', '"':
			if quote == 0 {
				quote = data[index]
			} else if quote == data[index] {
				quote = 0
			}
		case '>':
			if quote == 0 {
				return index
			}
		}
	}
	return -1
}

func isTagBoundary(value byte) bool {
	return value == '>' || value == '/' || value == ' ' || value == '\t' || value == '\r' || value == '\n'
}

func attributeValueRaw(tag []byte, name string) string {
	for position := 0; position < len(tag); {
		for position < len(tag) && (tag[position] == '<' || tag[position] == '>' || tag[position] == '/' || tag[position] == ' ' || tag[position] == '\t' || tag[position] == '\r' || tag[position] == '\n') {
			position++
		}
		start := position
		for position < len(tag) && tag[position] != '=' && tag[position] != ' ' && tag[position] != '\t' && tag[position] != '\r' && tag[position] != '\n' && tag[position] != '>' && tag[position] != '/' {
			position++
		}
		attributeName := string(tag[start:position])
		for position < len(tag) && (tag[position] == ' ' || tag[position] == '\t' || tag[position] == '\r' || tag[position] == '\n') {
			position++
		}
		if position >= len(tag) || tag[position] != '=' {
			continue
		}
		position++
		for position < len(tag) && (tag[position] == ' ' || tag[position] == '\t' || tag[position] == '\r' || tag[position] == '\n') {
			position++
		}
		if position >= len(tag) || (tag[position] != '\'' && tag[position] != '"') {
			continue
		}
		quote := tag[position]
		position++
		valueStart := position
		for position < len(tag) && tag[position] != quote {
			position++
		}
		if attributeName == name {
			return html.UnescapeString(string(tag[valueStart:position]))
		}
		if position < len(tag) {
			position++
		}
	}
	return ""
}
