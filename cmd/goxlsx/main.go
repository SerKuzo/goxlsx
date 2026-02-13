package main

import (
	"fmt"
	"goxlsx"
)

func main() {
	doc, err := goxlsx.Open("testdata/Ex1.xlsm")
	if err != nil {
		fmt.Println(err)
	}
	sheets, _ := doc.GetSheets()
	for _, sheet := range sheets {
		row, _ := doc.GetRows(sheet.Name)
		fmt.Println(row)
	}
	//for _, sheet := range sheets {
	//	fmt.Println(sheet.Name)
	//}

}
