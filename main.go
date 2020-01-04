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
	sheet.Apply(excel.NewExcelStyle(12, 0, false))
}
