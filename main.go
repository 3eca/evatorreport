package main

import (
	"fmt"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"
	"github.com/xuri/excelize/v2"
)

type DayReport struct {
	date time.Time
	cash int
	card int
	total int
}

type Report struct {
	name string
	period string
	cash int
	card int
	total int
}

func readFile(s string) [][]string {
	f, err := excelize.OpenFile(s)
    if err != nil {
        panic(err)
    }
    defer func() {
        if err := f.Close(); err != nil {
            panic(err)
        }
    }()

    rows, err := f.GetRows("Sheet1")
    if err != nil {
        panic(err)
    }

	return rows
}

func mutateData(array *[][]string) (*Report, *[]DayReport) {
	var result Report
	dReport := make([]DayReport, 0)

	for index, rows := range *array {
		str := strings.Join(rows, "")
		clearS := strings.TrimRight(str, ";")
		
		switch index {
		case 0:
			result.name = clearS
		case 1:
			result.period = strings.Split(clearS, "¦")[1]
		case 2:
			cash, err := strconv.ParseFloat(strings.Fields(clearS)[1], 64)
			if err != nil {
				fmt.Println(err)
			}
			result.cash = int(cash)
		case 3:
			card, err := strconv.ParseFloat(strings.Fields(clearS)[2], 64)
			if err != nil {
				fmt.Println(err)
			}
			result.card = int(card)
		case 4:
			total, err := strconv.ParseFloat(strings.Fields(clearS)[1], 64)
			if err != nil {
				fmt.Println(err)
			}
			result.total = int(total)
		default:
			re := regexp.MustCompile(`\d+\/\d+\/\d+;\d+.\d+;\d+.\d+;\d+.\d+`)
			match := re.FindAllString(clearS, -1)

			if match != nil {
				row := strings.Split(strings.Join(rows, ""), ";")
				dateX, err := time.Parse("02/01/2006", row[0])
				if err != nil {
					fmt.Println(err)
				}
				cashX, err := strconv.ParseFloat(row[1], 64)
				if err != nil {
					fmt.Println(err)
				}
				cardX, err := strconv.ParseFloat(row[2], 64)
				if err != nil {
					fmt.Println(err)
				}
				totalX, err := strconv.ParseFloat(row[3], 64)
				if err != nil {
					fmt.Println(err)
				}

				if cashX == 0 || cardX == 0 || totalX == 0 {
					continue
				}
				dReport = append(dReport, DayReport{dateX, int(cashX), int(cardX), int(totalX)})
			}
		}
	}
	return &result, &dReport
}

func writeFile(s string, r *Report, dr *[]DayReport) {
	var max, min, avr int
	var maxDay, minDay time.Time

	f := excelize.NewFile()

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	f.MergeCell("Sheet1", "A1", "C1")
	f.MergeCell("Sheet1", "A2", "C2")
    f.SetCellValue("Sheet1", "A1", r.name)
	f.SetCellValue("Sheet1", "A2", r.period)
	f.SetCellValue("Sheet1", "A3", "Нал")
	f.SetCellValue("Sheet1", "B3", "Безнал")
	f.SetCellValue("Sheet1", "C3", "Итого")
	f.SetCellValue("Sheet1", "A4", r.cash)
	f.SetCellValue("Sheet1", "B4", r.card)
	f.SetCellValue("Sheet1", "C4", r.total)

	for index, value := range *dr {
		if index == 0 {
			min = value.total
		}
		avr += value.total
		
		switch {
		case value.total > max:
			max = value.total
			maxDay = value.date
		case value.total < min:
			min = value.total
			minDay = value.date
		}
	}

	f.MergeCell("Sheet1", "A6", "A7")
	f.MergeCell("Sheet1", "A8", "A9")
	f.SetCellValue("Sheet1", "A6", "Макс")
	f.SetCellValue("Sheet1", "A8", "Мин")
	f.SetCellValue("Sheet1", "A10", "Среднее")
	f.SetCellValue("Sheet1", "B6", maxDay)
	f.SetCellValue("Sheet1", "B7", max)
	f.SetCellValue("Sheet1", "B8", minDay)
	f.SetCellValue("Sheet1", "B9", min)
	f.SetCellValue("Sheet1", "B10", avr/len(*dr))

    f.SetActiveSheet(0)

    if err := f.SaveAs(s); err != nil {
        fmt.Println(err)
    }
}

func main() {
	var src, dst string
	fmt.Println("Введите полный путь до файла: ")
	fmt.Scanln(&src)
	t := strings.Split(src, "\\")
	dst = strings.Join(t[:len(t)-1], "\\") + "\\EvatorReport.xlsx"
	data := readFile(src)
	report, dReports := mutateData(&data)
	writeFile(dst, report, dReports)
	fmt.Printf("Результат сохранен в файл %s\nЗавершение через %d секунд", dst, 10)
	time.Sleep(10 * time.Second)
	os.Exit(0)
}