package goxlsx

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"testing"
)

func TestSetCellStringAndSave(t *testing.T) {
	doc, err := Open(filepath.Join("testdata", "Ex.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	const sheet = "Временные уставки"
	const value = "строка <тест> & сохранена"
	if err := doc.SetCellString(sheet, "R2C3", value); err != nil {
		t.Fatal(err)
	}
	rows, err := doc.GetRows(sheet)
	if err != nil {
		t.Fatal(err)
	}
	if got := rows[1][2]; got != value {
		t.Fatalf("in-memory value = %q, want %q", got, value)
	}

	output := filepath.Join(t.TempDir(), "result.xlsx")
	if err := doc.Save(output); err != nil {
		t.Fatal(err)
	}

	reopened, err := Open(output)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	rows, err = reopened.GetRows(sheet)
	if err != nil {
		t.Fatal(err)
	}
	if got := rows[1][2]; got != value {
		t.Fatalf("saved value = %q, want %q", got, value)
	}
}

func TestGetRowsWithMergedUsesExcelCoordinates(t *testing.T) {
	doc, err := Open(filepath.Join("testdata", "Ex.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	rows, err := doc.GetRows("Временные уставки")
	if err != nil {
		t.Fatal(err)
	}
	if got := rows[1][2]; got != "теск" {
		t.Fatalf("formula string = %q, want %q", got, "теск")
	}

	rows, err = doc.GetRowsWithMerged("Временные уставки")
	if err != nil {
		t.Fatal(err)
	}
	if got := rows[1][0]; got != "USO.KC" {
		t.Fatalf("merged A2 = %q", got)
	}
	for row := 0; row <= 2; row++ {
		if got := rows[row][3]; got != "3" {
			t.Fatalf("merged D%d = %q", row+1, got)
		}
	}
}

func TestSetCellStringRejectsInvalidReferences(t *testing.T) {
	doc, err := Open(filepath.Join("testdata", "Ex.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	for _, ref := range []string{"", "A", "A0", "A1:B2", "R0C1", "R1C0"} {
		if err := doc.SetCellString("Временные уставки", ref, "value"); err == nil {
			t.Fatalf("reference %q was accepted", ref)
		}
	}
}

func TestSetCellStringXMLAddsCellToEmptySheetData(t *testing.T) {
	input := []byte(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`)
	updated, err := setCellStringXML(input, "B2", 2, "значение")
	if err != nil {
		t.Fatal(err)
	}
	rows, err := parseRows(updated, nil)
	if err != nil {
		t.Fatal(err)
	}
	if got := rows[1][1]; got != "значение" {
		t.Fatalf("empty sheet value = %q", got)
	}
}

func TestSetCellStringXMLReplacesSelfClosingCell(t *testing.T) {
	input := []byte(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"/></row></sheetData></worksheet>`)
	updated, err := setCellStringXML(input, "A1", 1, "значение")
	if err != nil {
		t.Fatal(err)
	}
	rows, err := parseRows(updated, nil)
	if err != nil {
		t.Fatal(err)
	}
	if got := rows[0][0]; got != "значение" {
		t.Fatalf("self-closing cell value = %q", got)
	}
}

func TestSetCellStringAtAddsCellAndSaveSamePath(t *testing.T) {
	input := filepath.Join("testdata", "Ex.xlsx")
	path := filepath.Join(t.TempDir(), "result.xlsx")
	source, err := os.Open(input)
	if err != nil {
		t.Fatal(err)
	}
	destination, err := os.Create(path)
	if err != nil {
		source.Close()
		t.Fatal(err)
	}
	if _, err := io.Copy(destination, source); err != nil {
		source.Close()
		destination.Close()
		t.Fatal(err)
	}
	source.Close()
	destination.Close()

	doc, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := doc.SetCellStringAt("Временные уставки", 100, 11, "добавленная строка"); err != nil {
		doc.Close()
		t.Fatal(err)
	}
	if err := doc.Save(path); err != nil {
		doc.Close()
		t.Fatal(err)
	}
	defer doc.Close()

	rows, err := doc.GetRows("Временные уставки")
	if err != nil {
		t.Fatal(err)
	}
	if got := rows[99][10]; got != "добавленная строка" {
		t.Fatalf("same-path value = %q", got)
	}
}

func TestSetRowStrings(t *testing.T) {
	doc, err := Open(filepath.Join("testdata", "Ex.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	if err := doc.SetRowStrings("Временные уставки", 2, 6, []string{"первое", "второе"}); err != nil {
		t.Fatal(err)
	}
	rows, err := doc.GetRows("Временные уставки")
	if err != nil {
		t.Fatal(err)
	}
	if rows[1][5] != "первое" || rows[1][6] != "второе" {
		t.Fatalf("row values = %q, %q", rows[1][5], rows[1][6])
	}
}

func TestSavePreservesArchiveEntries(t *testing.T) {
	doc, err := Open(filepath.Join("testdata", "Ex1.xlsm"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	if err := doc.SetCellString("Таблица сигналов", "F48", "обновлённое значение"); err != nil {
		t.Fatal(err)
	}
	output := filepath.Join(t.TempDir(), "result.xlsm")
	if err := doc.Save(output); err != nil {
		t.Fatal(err)
	}
	reopened, err := Open(output)
	if err != nil {
		t.Fatal(err)
	}
	rows, err := reopened.GetRows("Таблица сигналов")
	if err != nil {
		reopened.Close()
		t.Fatal(err)
	}
	if got := rows[47][5]; got != "обновлённое значение" {
		reopened.Close()
		t.Fatalf("saved XLSM value = %q", got)
	}
	reopened.Close()

	archive, err := zip.OpenReader(output)
	if err != nil {
		t.Fatal(err)
	}
	defer archive.Close()
	for _, name := range []string{"xl/vbaProject.bin", "xl/media/image12.emf", "xl/styles.xml"} {
		found := false
		for _, file := range archive.File {
			if file.Name == name {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("archive entry %q was not preserved", name)
		}
	}
}
