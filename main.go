package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/weeee9/excelclaim/excel"
)

func main() {
	url := "./test.xlsx"
	xlsx := excelize.NewFile()

	mergeRow(xlsx)

	err := xlsx.SaveAs(url)
	if err != nil {
		fmt.Println(err)
		return
	}
}

func mergeRow(xlsx *excelize.File) {
	sheet := excel.NewSheet(xlsx, "test", 8, 15)
	sheet.SetAllColsWidth(12, 12, 12, 9, 9, 9, 9, 9)
	sheet.WriteRow("|", "|", "|", "-", "-", "-", "-", "Column4")
	sheet.WriteRow("|", "|", "|", "-", "-", "A", "|", "|")
	sheet.WriteRow("|", "|", "|", "-", "A1", "A2", "|", "|")
	sheet.WriteRow("Column1", "Column2", "Column3", "A1.1", "A1.2", "A2.1", "B", "C")
	sheet.MergeRow()
	sheet.SetColor(1, 3, "#ffff00")
	sheet.SetColor(4, 6, "#E0EBF5")
	sheet.SetColor(7, 8, "#e4b001")
	sheet.WriteRow("data 1", "data 2", "data 3", "data 4", "data 5", "data 6", "data 7", "data 8")
}
