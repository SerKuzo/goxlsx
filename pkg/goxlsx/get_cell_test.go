package goxlsx

import (
	"path/filepath"
	"testing"
)

func TestGetCellValue(t *testing.T) {
	doc, err := Open(filepath.Join("..", "..", "testdata", "Ex.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	const sheet = "\u0412\u0440\u0435\u043c\u0435\u043d\u043d\u044b\u0435 \u0443\u0441\u0442\u0430\u0432\u043a\u0438"
	for _, test := range []struct {
		name string
		ref  string
		want string
	}{
		{name: "A1", ref: "C2", want: "\u0442\u0435\u0441\u043a"},
		{name: "R1C1", ref: "R2C3", want: "\u0442\u0435\u0441\u043a"},
		{name: "missing", ref: "Z100", want: ""},
	} {
		t.Run(test.name, func(t *testing.T) {
			got, err := doc.GetCellValue(sheet, test.ref)
			if err != nil {
				t.Fatal(err)
			}
			if got != test.want {
				t.Fatalf("GetCellValue(%q) = %q, want %q", test.ref, got, test.want)
			}
		})
	}

	if _, err := doc.GetCellValue(sheet, "A0"); err == nil {
		t.Fatal("invalid cell reference was accepted")
	}
}
