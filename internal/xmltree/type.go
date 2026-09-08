package xmltree

import (
	"encoding/xml"
	"strings"
)

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
	T    string        `xml:"t"` // Простой текст в теге <t>
	Runs []RichTextRun `xml:"r"` // Форматированный текст
}

func (s SI) String() string {
	if s.T != "" || len(s.Runs) == 0 {
		return s.T
	}
	var result strings.Builder
	for _, run := range s.Runs {
		result.WriteString(run.Text)
	}
	return result.String()
}

type XMLWorksheet struct {
	SheetData struct {
		Rows []XMLRow `xml:"row"`
	} `xml:"sheetData"`

	MergeCells struct {
		Count int            `xml:"count,attr"`
		Cells []XMLMergeCell `xml:"mergeCell"`
	} `xml:"mergeCells"`
}

type XMLRow struct {
	R     int       `xml:"r,attr"` // Номер строки
	Cells []XMLCell `xml:"c"`      // Ячейки
}

type XMLCell struct {
	R            string        `xml:"r,attr"` // Координаты (например, "A1")
	T            string        `xml:"t,attr"` // Тип (s = shared string, n = number)
	Value        string        `xml:"v"`      // Значение или индекс
	InlineString *InlineString `xml:"is"`     // Строка, хранящаяся внутри ячейки
}

// InlineString contains the text of an inline string cell. Runs are included
// for files that store rich text as a sequence of <r><t> elements.
type InlineString struct {
	Text string        `xml:"t"`
	Runs []RichTextRun `xml:"r"`
}

type RichTextRun struct {
	Text string `xml:"t"`
}

func (s *InlineString) String() string {
	if s == nil {
		return ""
	}
	if s.Text != "" || len(s.Runs) == 0 {
		return s.Text
	}
	var result strings.Builder
	for _, run := range s.Runs {
		result.WriteString(run.Text)
	}
	return result.String()
}

type XMLMergeCell struct {
	Ref string `xml:"ref,attr"` // Например: "A1:C3"
}
