package xmltree

import "encoding/xml"

// Workbook описывает структуру файла xl/workbook.xml
type Workbook struct {
	XMLName xml.Name `xml:"workbook"`
	Sheets  []Sheet  `xml:"sheets>sheet"`
}

// Sheet содержит информацию о конкретном листе
type Sheet struct {
	Name    string `xml:"name,attr"`    // Видимое имя (на ярлычке в Excel)
	SheetID string `xml:"sheetId,attr"` // Внутренний ID
	RID     string `xml:"id,attr"`      // Ссылка на файл
}

type Relationships struct {
	Relationship []Relationship `xml:"Relationship"`
}

type Relationship struct {
	ID     string `xml:"Id,attr"`
	Target string `xml:"Target,attr"` // Путь к файлу, например: worksheets/sheet1.xml
}

// SST (Shared String Table) — корневой элемент файла sharedStrings.xml
type SST struct {
	XMLName     xml.Name `xml:"sst"`
	Count       int      `xml:"count,attr"`       // Общее кол-во упоминаний
	UniqueCount int      `xml:"uniqueCount,attr"` // Кол-во уникальных строк
	SI          []SI     `xml:"si"`               // Сами элементы строк
}

// SI (String Item) — отдельная строка
type SI struct {
	T string `xml:"t"` // Простой текст в теге <t>
	// В будущем здесь можно добавить поддержку форматированного текста (RichText)
}

type XMLWorksheet struct {
	SheetData struct {
		Rows []XMLRow `xml:"row"`
	} `xml:"sheetData"`
}

type XMLRow struct {
	R     int       `xml:"r,attr"` // Номер строки
	Cells []XMLCell `xml:"c"`      // Ячейки
}

type XMLCell struct {
	R     string `xml:"r,attr"` // Координаты (например, "A1")
	T     string `xml:"t,attr"` // Тип (s = shared string, n = number)
	Value string `xml:"v"`      // Значение или индекс
}
