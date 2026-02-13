package goxlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"goxlsx/internal/xmltree"
	"strings"
)

// ExcelDoc — основная структура для работы с документом
type ExcelDoc struct {
	content *zip.ReadCloser
	// Карта для быстрого поиска файлов внутри архива (например, "xl/sharedStrings.xml")
	files         map[string]*zip.File
	sharedStrings []string
	sheetMap      map[string]string // Карта: "Лист1" -> "xl/worksheets/sheet1.xml"
}

// Open открывает xlsx файл и подготавливает его к чтению
func Open(path string) (*ExcelDoc, error) {
	// 1. Открываем ZIP-архив
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("ошибка открытия файла: %w", err)
	}

	doc := &ExcelDoc{
		content: reader,
		files:   make(map[string]*zip.File),
	}

	// 2. Индексируем файлы внутри архива для удобного доступа
	for _, f := range reader.File {
		doc.files[f.Name] = f
	}

	err1 := doc.loadSheetMap()
	if err1 != nil {
		return nil, err1
	}

	if len(doc.sharedStrings) == 0 {
		doc.loadSharedStrings()
	}

	return doc, nil
}

// Close закрывает дескриптор файла
func (d *ExcelDoc) Close() error {
	return d.content.Close()
}

func (d *ExcelDoc) loadSheetMap() error {
	// Читаем workbook.xml (Имена -> rId)
	file, ok := d.files["xl/workbook.xml"]
	if !ok {
		return fmt.Errorf("workbook.xml не найден в архиве")
	}

	// Открываем файл внутри ZIP
	rc, err := file.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	var wb xmltree.Workbook
	if err := xml.NewDecoder(rc).Decode(&wb); err != nil {
		return fmt.Errorf("ошибка парсинга workbook: %w", err)
	}

	// Читаем workbook.xml.rels (rId -> Путь)
	relFile, _ := d.files["xl/_rels/workbook.xml.rels"]
	rc, _ = relFile.Open()
	var rels xmltree.Relationships
	_ = xml.NewDecoder(rc).Decode(&rels)
	rc.Close()

	// Создаем временную карту rId -> Target
	idToPath := make(map[string]string)
	for _, r := range rels.Relationship {
		idToPath[r.ID] = r.Target
	}

	// 3. Сопоставляем Имя листа с полным Путем
	d.sheetMap = make(map[string]string)
	for _, s := range wb.Sheets {
		targetPath := idToPath[s.RID]
		// Важно: в rels пути относительные, добавляем префикс "xl/"
		if !strings.HasPrefix(targetPath, "xl/") {
			targetPath = "xl/" + targetPath
		}
		d.sheetMap[s.Name] = targetPath
	}
	return nil
}

// loadSharedStrings читает файл из архива и возвращает список строк
func (d *ExcelDoc) loadSharedStrings() error {
	file, ok := d.files["xl/sharedStrings.xml"]
	if !ok {
		// Файла может не быть, если в таблице только числа
		return nil
	}

	rc, err := file.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	var sst xmltree.SST
	if err := xml.NewDecoder(rc).Decode(&sst); err != nil {
		return fmt.Errorf("ошибка парсинга sharedStrings: %w", err)
	}

	// Вытаскиваем только текст из структур в плоский слайс
	d.sharedStrings = make([]string, len(sst.SI))
	for i, si := range sst.SI {
		d.sharedStrings[i] = si.T
	}
	return nil
}
