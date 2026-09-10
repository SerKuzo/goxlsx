package main

import (
	"fmt"
	"regexp"
	"slices"
	"strconv"

	"github.com/SerKuzo/goxlsx.git"
)

// вырезаем число после буквы 'A' и сравниваем его
func sortExcelRanges(m []string) {
	re := regexp.MustCompile(`[A-Z]+(\d+)`)

	slices.SortFunc(m, func(a, b string) int {
		// Извлекаем числа из строк, например "A1:H1" -> 1
		matchA := re.FindStringSubmatch(a)
		matchB := re.FindStringSubmatch(b)

		if len(matchA) > 1 && len(matchB) > 1 {
			valA, _ := strconv.Atoi(matchA[1])
			valB, _ := strconv.Atoi(matchB[1])
			if valA != valB {
				return valA - valB
			}
		}
		// Если числа равны или не найдены, сравниваем строки целиком
		if a < b {
			return -1
		}
		if a > b {
			return 1
		}
		return 0
	})
}

func main() {
	doc, err := goxlsx.Open("testdata/Ex1.xlsm")
	if err != nil {
		fmt.Println(err)
	}

	defer func(doc *goxlsx.ExcelDoc) {
		err := doc.Close()
		if err != nil {

		}
	}(doc)

	sheets, _ := doc.GetSheets()
	//for _, sheet := range sheets {
	//	//_, _ := doc.GetRows(sheet.Name)
	//	m, _ := doc.GetMergedCells(sheet.Name)
	fmt.Println(sheets)
	//}

	m, _ := doc.GetRows("Таблица сигналов")
	//m, _ := doc.GetMergedCells("Таблица сигналов")
	//sortExcelRanges(m)
	fmt.Println(m)

	//for _, sheet := range sheets {
	//	fmt.Println(sheet.Name)
	//}

}
