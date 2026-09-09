## Структура проекта

```text
cmd/goxlsx/          CLI-пример
pkg/goxlsx/          реализация библиотеки и unit-тесты
internal/xmltree/    внутренние модели XLSX и разбор координат
tests/benchmark/     benchmark чтения книги
testdata/            локальные Excel fixtures
```

Публичный импорт из корня модуля сохранён для обратной совместимости. Для нового кода рекомендуется использовать `pkg/goxlsx`.

## Чтение значения ячейки

```go
value, err := doc.GetCellValue("Таблица сигналов", "C2")
if err != nil {
	return err
}
fmt.Println(value)
```

Поддерживаются адреса `A1` (`C2`) и `R1C1` (`R2C3`). Пустая или отсутствующая ячейка возвращает пустую строку.

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
