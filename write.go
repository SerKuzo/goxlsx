package goxlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/Puzzanis/goxlsx.git/internal/xmltree"
)

// SetCellString changes a cell to a literal string value. The change is kept
// in memory until Save is called. cellRef accepts A1 (for example, F48) and
// R1C1 (for example, R48C6) notation.
func (d *ExcelDoc) SetCellString(sheetName, cellRef, value string) error {
	if d == nil || d.content == nil {
		return fmt.Errorf("документ не открыт")
	}

	path, ok := d.sheetMap[sheetName]
	if !ok {
		return fmt.Errorf("лист %s не найден", sheetName)
	}

	column, row, err := xmltree.ParseCellRef(cellRef)
	if err != nil {
		return err
	}
	columnTitle, err := xmltree.ColumnIndexToTitle(column)
	if err != nil {
		return err
	}
	canonicalRef := fmt.Sprintf("%s%d", columnTitle, row+1)

	data, err := d.sheetBytes(path)
	if err != nil {
		return err
	}
	updated, err := setCellStringXML(data, canonicalRef, row+1, value)
	if err != nil {
		return fmt.Errorf("ячейка %s: %w", canonicalRef, err)
	}

	if d.modifiedFiles == nil {
		d.modifiedFiles = make(map[string][]byte)
	}
	d.modifiedFiles[path] = updated
	return nil
}

// SetCellValue is an alias for SetCellString for callers that use generic
// cell-value terminology.
func (d *ExcelDoc) SetCellValue(sheetName, cellRef, value string) error {
	return d.SetCellString(sheetName, cellRef, value)
}

// SetCellStringAt changes a cell using one-based row and column numbers.
func (d *ExcelDoc) SetCellStringAt(sheetName string, row, column int, value string) error {
	if row <= 0 || column <= 0 {
		return fmt.Errorf("номер строки и столбца должны быть положительными")
	}
	columnTitle, err := xmltree.ColumnIndexToTitle(column - 1)
	if err != nil {
		return err
	}
	return d.SetCellString(sheetName, fmt.Sprintf("%s%d", columnTitle, row), value)
}

// SetRowStrings writes values into adjacent cells in one row. Row and column
// numbers are one-based.
func (d *ExcelDoc) SetRowStrings(sheetName string, row, startColumn int, values []string) error {
	if row <= 0 || startColumn <= 0 {
		return fmt.Errorf("номер строки и столбца должны быть положительными")
	}
	for offset, value := range values {
		if err := d.SetCellStringAt(sheetName, row, startColumn+offset, value); err != nil {
			return err
		}
	}
	return nil
}

func (d *ExcelDoc) sheetBytes(path string) ([]byte, error) {
	if data, ok := d.modifiedFiles[path]; ok {
		return append([]byte(nil), data...), nil
	}

	file, ok := d.files[path]
	if !ok {
		return nil, fmt.Errorf("файл листа %s не найден в архиве", path)
	}
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	data, err := io.ReadAll(rc)
	if err != nil {
		return nil, err
	}
	return data, nil
}

func setCellStringXML(data []byte, cellRef string, rowNumber int, value string) ([]byte, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	rowInsertionOffset := -1
	rowStartOffset := -1
	rowStartEndOffset := -1
	rowIsSelfClosing := false
	sheetDataStartOffset := -1
	sheetDataStartEndOffset := -1
	sheetDataIsSelfClosing := false
	sheetDataInsertionOffset := -1

	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("ошибка парсинга листа: %w", err)
		}

		switch element := token.(type) {
		case xml.StartElement:
			switch element.Name.Local {
			case "sheetData":
				sheetDataStartOffset = before
				sheetDataStartEndOffset = int(decoder.InputOffset())
				sheetDataIsSelfClosing = bytes.HasSuffix(bytes.TrimSpace(data[sheetDataStartOffset:sheetDataStartEndOffset]), []byte("/>"))
				sheetDataInsertionOffset = -1
			case "row":
				if attributeValue(element.Attr, "r") == fmt.Sprint(rowNumber) {
					rowStartOffset = before
					rowStartEndOffset = int(decoder.InputOffset())
					rowIsSelfClosing = bytes.HasSuffix(bytes.TrimSpace(data[rowStartOffset:rowStartEndOffset]), []byte("/>"))
				}
			case "c":
				if strings.EqualFold(attributeValue(element.Attr, "r"), cellRef) {
					end, err := consumeElement(decoder)
					if err != nil {
						return nil, fmt.Errorf("ошибка чтения ячейки %s: %w", cellRef, err)
					}
					cellXML := literalStringCell(element, cellRef, value)
					return spliceBytes(data, before, end, cellXML), nil
				}
			}

		case xml.EndElement:
			switch element.Name.Local {
			case "row":
				if rowStartOffset >= 0 {
					if rowIsSelfClosing {
						cellXML := literalStringCell(xml.StartElement{Name: xml.Name{Local: "c"}}, cellRef, value)
						rowXML := append([]byte{}, data[rowStartOffset:rowStartEndOffset]...)
						rowXML = bytes.TrimSuffix(bytes.TrimSpace(rowXML), []byte("/>"))
						rowXML = append(rowXML, '>')
						rowXML = append(rowXML, cellXML...)
						rowXML = append(rowXML, []byte("</row>")...)
						return spliceBytes(data, rowStartOffset, rowStartEndOffset, rowXML), nil
					}
					rowInsertionOffset = before
					rowStartOffset = -1
				}
			case "sheetData":
				if sheetDataIsSelfClosing {
					rowXML := newRowXML(rowNumber, cellRef, value)
					sheetDataXML := append([]byte{}, data[sheetDataStartOffset:sheetDataStartEndOffset]...)
					sheetDataXML = bytes.TrimSuffix(bytes.TrimSpace(sheetDataXML), []byte("/>"))
					sheetDataXML = append(sheetDataXML, '>')
					sheetDataXML = append(sheetDataXML, rowXML...)
					sheetDataXML = append(sheetDataXML, []byte("</sheetData>")...)
					return spliceBytes(data, sheetDataStartOffset, sheetDataStartEndOffset, sheetDataXML), nil
				}
				if sheetDataInsertionOffset == -1 {
					sheetDataInsertionOffset = before
				}
			}
		}
	}

	cellXML := literalStringCell(xml.StartElement{Name: xml.Name{Local: "c"}}, cellRef, value)
	if sheetDataIsSelfClosing && sheetDataStartOffset >= 0 {
		rowXML := newRowXML(rowNumber, cellRef, value)
		sheetDataXML := append([]byte{}, data[sheetDataStartOffset:sheetDataStartEndOffset]...)
		sheetDataXML = bytes.TrimSuffix(bytes.TrimSpace(sheetDataXML), []byte("/>"))
		sheetDataXML = append(sheetDataXML, '>')
		sheetDataXML = append(sheetDataXML, rowXML...)
		sheetDataXML = append(sheetDataXML, []byte("</sheetData>")...)
		return spliceBytes(data, sheetDataStartOffset, sheetDataStartEndOffset, sheetDataXML), nil
	}
	if rowInsertionOffset >= 0 {
		return spliceBytes(data, rowInsertionOffset, rowInsertionOffset, cellXML), nil
	}
	if sheetDataInsertionOffset >= 0 {
		rowXML := newRowXML(rowNumber, cellRef, value)
		return spliceBytes(data, sheetDataInsertionOffset, sheetDataInsertionOffset, rowXML), nil
	}
	return nil, fmt.Errorf("элемент sheetData не найден")
}

func newRowXML(rowNumber int, cellRef, value string) []byte {
	cellXML := literalStringCell(xml.StartElement{Name: xml.Name{Local: "c"}}, cellRef, value)
	rowXML := []byte(fmt.Sprintf("<row r=\"%d\">", rowNumber))
	rowXML = append(rowXML, cellXML...)
	rowXML = append(rowXML, []byte("</row>")...)
	return rowXML
}

func consumeElement(decoder *xml.Decoder) (int, error) {
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return 0, err
		}
		switch token.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return int(decoder.InputOffset()), nil
}

func spliceBytes(data []byte, start, end int, replacement []byte) []byte {
	result := make([]byte, 0, len(data)-(end-start)+len(replacement))
	result = append(result, data[:start]...)
	result = append(result, replacement...)
	result = append(result, data[end:]...)
	return result
}

func literalStringCell(start xml.StartElement, cellRef, value string) []byte {
	var result bytes.Buffer
	result.WriteByte('<')
	result.WriteString(start.Name.Local)
	hasReference := false
	for _, attribute := range start.Attr {
		if attribute.Name.Local == "t" {
			continue
		}
		if attribute.Name.Local == "r" {
			hasReference = true
		}
		result.WriteByte(' ')
		result.WriteString(xmlAttributeName(attribute.Name))
		result.WriteString("=\"")
		_ = xml.EscapeText(&result, []byte(attribute.Value))
		result.WriteString("\"")
	}
	if !hasReference {
		result.WriteString(" r=\"")
		result.WriteString(cellRef)
		result.WriteString("\"")
	}
	result.WriteString(" t=\"inlineStr\">")
	result.WriteString("<is><t")
	if strings.TrimSpace(value) != value {
		result.WriteString(" xml:space=\"preserve\"")
	}
	result.WriteString(">")
	_ = xml.EscapeText(&result, []byte(value))
	result.WriteString("</t></is></")
	result.WriteString(start.Name.Local)
	result.WriteByte('>')
	return result.Bytes()
}

func xmlAttributeName(name xml.Name) string {
	if name.Space == "http://www.w3.org/XML/1998/namespace" {
		return "xml:" + name.Local
	}
	return name.Local
}

func attributeValue(attributes []xml.Attr, name string) string {
	for _, attribute := range attributes {
		if attribute.Name.Local == name {
			return attribute.Value
		}
	}
	return ""
}

// Save writes the current document to path while preserving every archive
// entry that was not changed through SetCellString.
func (d *ExcelDoc) Save(path string) error {
	if d == nil || d.content == nil {
		return fmt.Errorf("документ не открыт")
	}
	if strings.TrimSpace(path) == "" {
		return fmt.Errorf("путь сохранения не задан")
	}

	absPath, err := filepath.Abs(path)
	if err != nil {
		return err
	}
	temp, err := os.CreateTemp(filepath.Dir(absPath), ".goxlsx-*.tmp")
	if err != nil {
		return err
	}
	tempPath := temp.Name()
	defer os.Remove(tempPath)

	archive := zip.NewWriter(temp)
	for _, file := range d.content.File {
		header := file.FileHeader
		header.Name = file.Name
		writer, err := archive.CreateHeader(&header)
		if err != nil {
			_ = archive.Close()
			_ = temp.Close()
			return err
		}

		if data, ok := d.modifiedFiles[file.Name]; ok {
			if _, err := writer.Write(data); err != nil {
				_ = archive.Close()
				_ = temp.Close()
				return err
			}
			continue
		}

		reader, err := file.Open()
		if err != nil {
			_ = archive.Close()
			_ = temp.Close()
			return err
		}
		_, copyErr := io.Copy(writer, reader)
		closeErr := reader.Close()
		if copyErr != nil {
			_ = archive.Close()
			_ = temp.Close()
			return copyErr
		}
		if closeErr != nil {
			_ = archive.Close()
			_ = temp.Close()
			return closeErr
		}
	}

	if err := archive.Close(); err != nil {
		_ = temp.Close()
		return err
	}
	if err := temp.Close(); err != nil {
		return err
	}

	sameSource := d.sourcePath != "" && strings.EqualFold(filepath.Clean(d.sourcePath), filepath.Clean(absPath))
	if sameSource {
		if err := d.content.Close(); err != nil {
			return err
		}
		d.content = nil
	}
	if _, err := os.Stat(absPath); err == nil {
		if err := os.Remove(absPath); err != nil {
			return err
		}
	} else if !os.IsNotExist(err) {
		return err
	}
	if err := os.Rename(tempPath, absPath); err != nil {
		return err
	}

	if sameSource {
		reader, err := zip.OpenReader(absPath)
		if err != nil {
			return err
		}
		d.content = reader
		d.files = make(map[string]*zip.File, len(reader.File))
		for _, file := range reader.File {
			d.files[file.Name] = file
		}
	}
	return nil
}

// SaveAs is an explicit alias for Save.
func (d *ExcelDoc) SaveAs(path string) error {
	return d.Save(path)
}
