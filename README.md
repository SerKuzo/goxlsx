# goxlsx

`goxlsx` - библиотека для чтения и изменения книг Excel в форматах `.xlsx` и `.xlsm`.
Она работает с XLSX как с ZIP-архивом, читает листы и значения ячеек, поддерживает
shared strings, inline strings, объединённые ячейки и сохраняет неизменённые записи
архива, включая VBA-проект и медиафайлы.

Модуль требует Go 1.25 или новее.

## Возможности

- открытие и закрытие `.xlsx`/`.xlsm` файлов;
- получение списка листов;
- чтение всей таблицы листа в виде `[][]string`;
- чтение отдельной ячейки;
- адреса ячеек в форматах `A1` и `R1C1`;
- абсолютные маркеры `$` в A1-адресах, например `$C$2`;
- чтение списка объединённых диапазонов;
- чтение строк с заполнением объединённых ячеек значением верхней левой ячейки;
- изменение строковых значений ячеек;
- пакетная запись соседних ячеек одной строки;
- сохранение в новый файл или перезапись исходного файла;
- сохранение неизменённых записей XLSX-архива.

## Установка

```bash
go get github.com/SerKuzo/goxlsx.git
```

Основной публичный импорт:

```go
import goxlsx "github.com/SerKuzo/goxlsx.git"
```

Реализация также находится в `pkg/goxlsx`. Корневой пакет предоставляет те же
публичные типы и функции и удобен для совместимости с существующим кодом.

## Быстрый старт

```go
package main

import (
	"fmt"
	"log"

	goxlsx "github.com/SerKuzo/goxlsx.git"
)

func main() {
	doc, err := goxlsx.Open("input.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Close()

	sheets, err := doc.GetSheets()
	if err != nil {
		log.Fatal(err)
	}
	for _, sheet := range sheets {
		fmt.Println(sheet.Name)
	}

	value, err := doc.GetCellValue("Лист1", "C2")
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("C2:", value)
}
```

## Открытие и закрытие

`Open` открывает ZIP-архив, индексирует его записи, загружает карту листов и
таблицу shared strings.

```go
doc, err := goxlsx.Open("input.xlsm")
if err != nil {
	return err
}
defer doc.Close()
```

`Close` закрывает файл. Метод безопасно повторно вызывать и можно вызывать для
`nil`-указателя.

После `Close` методы чтения и записи возвращают ошибку о закрытом документе.

## Список листов

```go
sheets, err := doc.GetSheets()
if err != nil {
	return err
}

for _, sheet := range sheets {
	fmt.Printf("name=%q sheetId=%q relationship=%q\n",
		sheet.Name, sheet.SheetID, sheet.RID)
}
```

Каждый элемент имеет поля:

| Поле | Описание |
| --- | --- |
| `Name` | Отображаемое имя листа в Excel. Используется в остальных методах. |
| `SheetID` | Внутренний идентификатор листа из `workbook.xml`. |
| `RID` | Идентификатор связи листа с XML-файлом внутри архива. |

## Чтение таблицы

`GetRows` возвращает лист как двумерный срез строк:

```go
rows, err := doc.GetRows("Лист1")
if err != nil {
	return err
}

for rowIndex, row := range rows {
	for columnIndex, value := range row {
		fmt.Printf("[%d,%d] = %q\n", rowIndex, columnIndex, value)
	}
}
```

Индексы результата начинаются с нуля. Поэтому Excel-ячейка `C2` находится в
`rows[1][2]`. Отсутствующие ячейки представлены пустыми строками. Размер
каждой строки определяется максимальным столбцом с данными в этой строке, а
пропущенные строки сохраняются как пустые элементы результата.

Пример чтения заголовков и данных:

```go
rows, err := doc.GetRows("Таблица сигналов")
if err != nil {
	return err
}
if len(rows) == 0 {
	return nil
}

headers := rows[0]
for _, row := range rows[1:] {
	if len(row) == 0 {
		continue
	}
	fmt.Printf("%q: %q\n", headers[0], row[0])
}
```

Каждый вызов `GetRows` строит новый результат. Если таблица большая и нужна
только отдельная ячейка, используйте `GetCellValue`, чтобы не обходить таблицу
на стороне приложения.

## Чтение отдельной ячейки

Поддерживаются A1-адреса:

```go
value, err := doc.GetCellValue("Таблица сигналов", "C2")
if err != nil {
	return err
}
fmt.Println(value)
```

И R1C1-адреса:

```go
value, err := doc.GetCellValue("Таблица сигналов", "R2C3")
if err != nil {
	return err
}
```

Пустая или отсутствующая ячейка возвращает пустую строку без ошибки. Ошибка
возвращается для неизвестного листа, закрытого документа или некорректного
адреса, например `A0`, `A` или `R0C1`.

## Объединённые ячейки

`GetMergedCells` возвращает исходные ссылки диапазонов из листа:

```go
merged, err := doc.GetMergedCells("Лист1")
if err != nil {
	return err
}

for _, ref := range merged {
	fmt.Println(ref) // например: A1:C3
}
```

`GetRowsWithMerged` дополнительно заполняет каждую ячейку объединённого
диапазона значением его верхней левой ячейки:

```go
rows, err := doc.GetRowsWithMerged("Лист1")
if err != nil {
	return err
}

// Для диапазона C2:E3 значение rows[1][2] будет продублировано
// в rows[1][3], rows[1][4], rows[2][2], rows[2][3] и rows[2][4].
fmt.Println(rows[1][2])
```

## Запись значений

Изменения хранятся в памяти до вызова `Save` или `SaveAs`.

### A1 и R1C1

```go
if err := doc.SetCellString("Лист1", "F48", "Новое значение"); err != nil {
	return err
}

if err := doc.SetCellString("Лист1", "R49C7", "Ещё одно значение"); err != nil {
	return err
}
```

`SetCellValue` - синоним `SetCellString`:

```go
if err := doc.SetCellValue("Лист1", "A1", "Заголовок"); err != nil {
	return err
}
```

Запись по номерам строки и столбца использует нумерацию с единицы:

```go
if err := doc.SetCellStringAt("Лист1", 48, 6, "Новое значение"); err != nil {
	return err
}
```

### Запись нескольких ячеек

`SetRowStrings` записывает значения в соседние ячейки одной строки. Номера
строки и начального столбца начинаются с единицы:

```go
values := []string{"Колонка 6", "Колонка 7", "Колонка 8"}
if err := doc.SetRowStrings("Лист1", 48, 6, values); err != nil {
	return err
}
```

Запись сейчас предназначена для строковых значений. Формулы, стили и другие
свойства ячейки этим API не изменяются.

## Сохранение

`Save` записывает результат в указанный путь:

```go
if err := doc.Save("output.xlsx"); err != nil {
	return err
}
```

`SaveAs` - явный синоним `Save`:

```go
if err := doc.SaveAs("output.xlsm"); err != nil {
	return err
}
```

При сохранении неизменённые записи исходного архива копируются без изменений.
Это позволяет сохранять `.xlsm` вместе с `vbaProject.bin`, изображения,
стили и прочие записи, которые библиотека не редактирует.

Полный пример изменения и повторного открытия файла:

```go
package main

import (
	"fmt"
	"log"

	goxlsx "github.com/SerKuzo/goxlsx.git"
)

func main() {
	doc, err := goxlsx.Open("input.xlsm")
	if err != nil {
		log.Fatal(err)
	}

	if err := doc.SetCellString("Таблица сигналов", "F48", "Обновлено"); err != nil {
		doc.Close()
		log.Fatal(err)
	}
	if err := doc.SetRowStrings("Таблица сигналов", 49, 6, []string{"A", "B"}); err != nil {
		doc.Close()
		log.Fatal(err)
	}

	if err := doc.Save("output.xlsm"); err != nil {
		doc.Close()
		log.Fatal(err)
	}
	if err := doc.Close(); err != nil {
		log.Fatal(err)
	}

	check, err := goxlsx.Open("output.xlsm")
	if err != nil {
		log.Fatal(err)
	}
	defer check.Close()

	value, err := check.GetCellValue("Таблица сигналов", "F48")
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println(value)
}
```

Если путь `Save` совпадает с исходным файлом, библиотека безопасно закрывает
старый архив, заменяет файл и повторно открывает его для дальнейшей работы.

## Обработка ошибок

Все операции открытия, чтения, записи и сохранения возвращают ошибку. Не
игнорируйте её в прикладном коде:

```go
doc, err := goxlsx.Open(path)
if err != nil {
	return fmt.Errorf("открытие книги: %w", err)
}
defer doc.Close()

rows, err := doc.GetRows(sheetName)
if err != nil {
	return fmt.Errorf("чтение листа %q: %w", sheetName, err)
}
```

Типичные причины ошибки:

- файл не существует или не является корректным ZIP/XLSX;
- в книге отсутствуют `workbook.xml`, relationships или XML-файл листа;
- указан неизвестный лист;
- документ уже закрыт;
- передан некорректный адрес ячейки;
- путь сохранения пустой или недоступен для записи.

## Ограничения

- чтение возвращает строковые значения; отдельного API для числового типа,
  дат или формул нет;
- формулы не вычисляются, читается сохранённое в XML значение ячейки;
- запись изменяет строковое содержимое ячейки и не редактирует стили,
  формулы, комментарии, гиперссылки и другие свойства;
- `GetRows` материализует весь лист в памяти;
- документ и его методы не предназначены для одновременного использования
  из нескольких goroutine без внешней синхронизации;
- для `.xlsm` сохраняется существующий VBA-проект, но библиотека не изменяет
  макросы.

## Структура проекта

```text
cmd/goxlsx/          пример приложения
pkg/goxlsx/          реализация библиотеки и unit-тесты
internal/xmltree/    внутренние модели XLSX и разбор координат
tests/alignment/     тест выравнивания структур
tests/benchmark/     benchmark чтения книги
testdata/            локальные Excel-файлы для тестов
```

## Тесты и benchmark

Запуск всех тестов:

```bash
go test ./...
```

Benchmark чтения книги с отчётом об аллокациях:

```bash
go test -bench=BenchmarkWorkbookRead -benchmem -count=5 ./tests/benchmark
```

Одна операция benchmark читает все листы, перечисленные в `Ex1.xlsm`. Параметр
`allocs/op` показывает число аллокаций на одну такую операцию, а `B/op` - объём
выделенной памяти.
