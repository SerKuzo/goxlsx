package test

import (
	"goxlsx"
	"testing"
)

//go test -bench . -benchmem .\test\

func BenchmarkMyFunction(b *testing.B) {
	doc, err := goxlsx.Open("E:\\Go\\goxlsx\\testdata\\Ex1.xlsm")
	if err != nil {
		b.Fatalf("Failed to open file: %v", err) // Остановит бенчмарк и покажет реальную причину
	}
	b.ResetTimer() // Очищаем время, потраченное на чтение файла
	for i := 0; i < b.N; i++ {
		sheets, _ := doc.GetSheets()
		for _, sheet := range sheets {
			_, _ = doc.GetRows(sheet.Name)
		}
	}

	// Рассчитываем среднее время в миллисекундах за одну операцию
	// b.Elapsed() — общее время, b.N — количество итераций
	msPerOp := float64(b.Elapsed().Milliseconds()) / float64(b.N)
	b.ReportMetric(msPerOp, "ms/op")
}
