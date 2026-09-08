## Запись строк в ячейки

```go
doc, err := goxlsx.Open("input.xlsx")
if err != nil {
	return err
}
defer doc.Close()

// Поддерживаются адреса A1 (F48) и R1C1 (R48C6).
if err := doc.SetCellString("Таблица сигналов", "F48", "Новое значение"); err != nil {
	return err
}

// Альтернатива: строка и столбец, нумерация с единицы.
if err := doc.SetCellStringAt("Таблица сигналов", 48, 6, "Новое значение"); err != nil {
	return err
}

if err := doc.SetRowStrings("Таблица сигналов", 48, 6, []string{"Колонка 6", "Колонка 7"}); err != nil {
	return err
}

return doc.Save("output.xlsx")
```

`SetCellValue` и `SaveAs` являются синонимами соответственно `SetCellString` и `Save`.
