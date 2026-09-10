package goxlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	pathpkg "path"
	"path/filepath"
	"strings"

	"github.com/Puzzanis/goxlsx.git/internal/xmltree"
)

// ExcelDoc — основная структура для работы с документом
type ExcelDoc struct {
	content    *zip.ReadCloser
	sourcePath string
	// Карта для быстрого поиска файлов внутри архива (например, "xl/sharedStrings.xml")
	files         map[string]*zip.File
	sharedStrings []string
	sheetMap      map[string]string // Карта: "Лист1" -> "xl/worksheets/sheet1.xml"
	modifiedFiles map[string][]byte
	sheetData     map[string][]byte // Распакованные листы, чтобы не читать ZIP повторно
}

// Open открывает xlsx файл и подготавливает его к чтению
func Open(path string) (*ExcelDoc, error) {
	// 1. Открываем ZIP-архив
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("ошибка открытия файла: %w", err)
	}

	doc := &ExcelDoc{
		content:       reader,
		files:         make(map[string]*zip.File),
		modifiedFiles: make(map[string][]byte),
		sheetData:     make(map[string][]byte),
	}
	doc.sourcePath, _ = filepath.Abs(path)

	// 2. Индексируем файлы внутри архива для удобного доступа
	for _, f := range reader.File {
		doc.files[f.Name] = f
	}

	if err := doc.loadSheetMap(); err != nil {
		_ = reader.Close()
		return nil, err
	}

	if err := doc.loadSharedStrings(); err != nil {
		_ = reader.Close()
		return nil, err
	}

	return doc, nil
}

// Close закрывает дескриптор файла
func (d *ExcelDoc) Close() error {
	if d == nil || d.content == nil {
		return nil
	}
	content := d.content
	d.content = nil
	return content.Close()
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
	relFile, ok := d.files["xl/_rels/workbook.xml.rels"]
	if !ok {
		return fmt.Errorf("workbook.xml.rels не найден в архиве")
	}
	rc, err = relFile.Open()
	if err != nil {
		return err
	}
	var rels xmltree.Relationships
	if err := xml.NewDecoder(rc).Decode(&rels); err != nil {
		_ = rc.Close()
		return fmt.Errorf("ошибка парсинга workbook relationships: %w", err)
	}
	if err := rc.Close(); err != nil {
		return err
	}

	// Создаем временную карту rId -> Target
	idToPath := make(map[string]string)
	for _, r := range rels.Relationship {
		idToPath[r.ID] = r.Target
	}

	// 3. Сопоставляем Имя листа с полным Путем
	d.sheetMap = make(map[string]string)
	for _, s := range wb.Sheets {
		targetPath, ok := idToPath[s.RID]
		if !ok || strings.TrimSpace(targetPath) == "" {
			return fmt.Errorf("путь листа %s не найден в relationships", s.Name)
		}
		targetPath = strings.ReplaceAll(targetPath, "\\", "/")
		targetPath = strings.TrimPrefix(targetPath, "/")
		if !strings.HasPrefix(targetPath, "xl/") {
			targetPath = pathpkg.Join("xl", targetPath)
		} else {
			targetPath = pathpkg.Clean(targetPath)
		}
		if _, ok := d.files[targetPath]; !ok {
			return fmt.Errorf("файл листа %s не найден в архиве", targetPath)
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
		d.sharedStrings[i] = si.String()
	}
	return nil
}

func (d *ExcelDoc) GetMergedCells(sheetName string) ([]string, error) {
	if d == nil || d.content == nil {
		return nil, fmt.Errorf("документ не открыт")
	}
	path, ok := d.sheetMap[sheetName]
	if !ok {
		return nil, fmt.Errorf("лист %s не найден", sheetName)
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

	var xmlSheet xmltree.XMLWorksheet
	if err := xml.NewDecoder(rc).Decode(&xmlSheet); err != nil {
		return nil, err
	}

	// Собираем все диапазоны в простой слайс строк
	var merged []string
	for _, mc := range xmlSheet.MergeCells.Cells {
		merged = append(merged, mc.Ref)
	}

	return merged, nil
}

// метод, который при чтении строк будет проверять каждую ячейку на принадлежность к «мерджу».
func (d *ExcelDoc) GetRowsWithMerged(sheetName string) ([][]string, error) {
	rawMerged, err := d.GetMergedCells(sheetName)
	if err != nil {
		return nil, err
	}

	// Используем срез структур из пакета xmltree
	var ranges []xmltree.MergeRange
	for _, r := range rawMerged {
		mRange, err := xmltree.ParseRange(r)
		if err != nil {
			return nil, fmt.Errorf("объединённый диапазон %s: %w", r, err)
		}
		ranges = append(ranges, mRange)
	}

	rows, err := d.GetRows(sheetName)
	if err != nil {
		return nil, err
	}
	for _, m := range ranges {
		for len(rows) <= m.MaxRow {
			rows = append(rows, nil)
		}
		for rowIndex := m.MinRow; rowIndex <= m.MaxRow; rowIndex++ {
			for len(rows[rowIndex]) <= m.MaxCol {
				rows[rowIndex] = append(rows[rowIndex], "")
			}
		}
	}

	for rIdx := range rows {
		for cIdx := range rows[rIdx] {
			if rows[rIdx][cIdx] == "" {
				for _, m := range ranges {
					// Проверяем вхождение в диапазон
					if rIdx >= m.MinRow && rIdx <= m.MaxRow &&
						cIdx >= m.MinCol && cIdx <= m.MaxCol {

						// Защита от выхода за границы при обращении к главной ячейке
						if m.MinRow < len(rows) && m.MinCol < len(rows[m.MinRow]) {
							rows[rIdx][cIdx] = rows[m.MinRow][m.MinCol]
						}
						break
					}
				}
			}
		}
	}
	return rows, nil
}
