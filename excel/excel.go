package excel

import (
	"fmt"
	"log"
	"regexp"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const BaseStyle2 = `{"border":[{"type":"left","color":"#808080","style":1},{"type":"top","color":"#808080","style":1},
{"type":"right","color":"#808080","style":1},{"type":"bottom","color":"#808080","style":1}],
"font":{"bold":%t, "family":"calibri", "fontSize":%d,"color":"#000000"},"alignment":{"horizontal":"%s", "vertical":"distributed"}}`

// 由於 excelize 本身的限制所以要設置顏色又覆蓋掉其他 style 設置，只能用這樣的方法
const baseStyle = `
{
	"border":[
		{"type":"left","color":"#808080","style":1},
		{"type":"top","color":"#808080","style":1},
		{"type":"right","color":"#808080","style":1},
		{"type":"bottom","color":"#808080","style":1}
	],
	"font": {"family":"calibri", "fontSize":12, "color":"#000000"},
	"alignment": {"horizontal":"center", "vertical":"distributed"},
	"fill": {"type":"pattern","color":["%s"],"pattern":1}}
`

type ExcelSheet struct {
	xlsx      *excelize.File
	name      string
	colNum    int
	rowHeight float64
	rowNum    int
}

var firstSheet = true

func NewSheet(xlsx *excelize.File, sheet string, colNum int, rowHeight float64) *ExcelSheet {
	excelSheet := &ExcelSheet{xlsx: xlsx, name: sheet, rowHeight: rowHeight, colNum: colNum}

	if firstSheet {
		xlsx.SetSheetName("Sheet1", sheet)
		firstSheet = false
	} else {
		xlsx.NewSheet(sheet)
	}
	return excelSheet
}

func (p *ExcelSheet) SetColWidth(col int, width float64) {
	colText := toExcelColName(col)
	p.xlsx.SetColWidth(p.name, colText, colText, width)
}

func (p *ExcelSheet) SetAllColsWidth(widths ...float64) {
	for i, v := range widths {
		colText := toExcelColName(i + 1)
		p.xlsx.SetColWidth(p.name, colText, colText, v)
	}
}

func (p *ExcelSheet) SetCellValue(col int, row int, v interface{}) {
	colText := toExcelColName(col)
	index := fmt.Sprintf("%s%d", colText, row)
	p.xlsx.SetCellValue(p.name, index, v)
}

func (p *ExcelSheet) WriteRow(cols ...string) *ExcelSheetRow {
	row := NewExcelSheetRow(p)
	row.WriteRow(cols...)
	return row
}

func (p *ExcelSheet) Apply(excelStyle *ExcelStyle) {
	p.ApplyRows(excelStyle, p.rowNum)
}

func (p *ExcelSheet) ApplyRows(excelStyle *ExcelStyle, rowNum int) {
	p.ApplyRowsRange(excelStyle, 1, rowNum)
}

func (p *ExcelSheet) ApplyRowsRange(excelStyle *ExcelStyle, rowStart int, rowEnd int) {
	alignText := "center"
	if excelStyle.align == -1 {
		alignText = "left"
	} else if excelStyle.align == 1 {
		alignText = "right"
	}
	txt := fmt.Sprintf(BaseStyle2, excelStyle.fontBold, excelStyle.fontSize, alignText)
	style, _ := p.xlsx.NewStyle(txt)
	s := makeFormatter(rowStart, rowStart)
	print(s)
	p.xlsx.SetCellStyle(p.name, s, makeFormatter(p.colNum, rowEnd), style)
}

func makeFormatter(col int, row int) string {
	colText := toExcelColName(col)
	index := fmt.Sprintf("%s%d", colText, row)
	return index
}

func (p *ExcelSheet) MergeCell(colStart int, rowStart int, colEnd int, rowEnd int) {
	p.xlsx.MergeCell(p.name, makeFormatter(colStart, rowStart), makeFormatter(colEnd, rowEnd))
}

type ExcelSheetRow struct {
	sheet    *ExcelSheet
	row      int
	boldCols []int
}

func NewExcelSheetRow(sheet *ExcelSheet) *ExcelSheetRow {
	sheet.rowNum++
	excelSheetRow := &ExcelSheetRow{sheet: sheet, row: sheet.rowNum}
	excelSheetRow.SetRowHeight(sheet.rowHeight)
	return excelSheetRow
}

func (p *ExcelSheetRow) SetRowHeight(height float64) *ExcelSheetRow {
	p.sheet.xlsx.SetRowHeight(p.sheet.name, p.row, height)
	return p
}

func (p *ExcelSheetRow) MergeCell(colStart int, colEnd int) *ExcelSheetRow {
	p.sheet.xlsx.MergeCell(p.sheet.name, makeFormatter(colStart, p.row), makeFormatter(colEnd, p.row))
	return p
}

func (p *ExcelSheetRow) MergeRowLine() *ExcelSheetRow {
	p.sheet.xlsx.MergeCell(p.sheet.name, makeFormatter(1, p.row), makeFormatter(p.sheet.colNum, p.row))
	return p
}

func (p *ExcelSheetRow) SetCellValue(col int, v interface{}) *ExcelSheetRow {
	p.sheet.SetCellValue(col, p.row, v)
	return p
}

func (p *ExcelSheetRow) SetBold(cols ...int) *ExcelSheetRow {
	p.boldCols = cols
	return p
}

func (p *ExcelSheetRow) WriteRow(cols ...string) *ExcelSheetRow {
	if len(cols) == 1 && p.sheet.colNum > 1 {
		p.MergeRowLine()
		p.sheet.SetCellValue(1, p.row, cols[0])
		return p
	}

	num := 0
	for i, v := range cols {
		if v == "-" {
			num++
			if i < len(cols)-1 && cols[i+1] != "-" {
				p.MergeCell(i+2-num, i+2)
				num = 0
			}
		}
		p.SetCellValue(i+1, v)
	}
	return p
}

func (p *ExcelSheetRow) Apply(excelStyle *ExcelStyle) *ExcelSheetRow {
	txt := fmt.Sprintf(BaseStyle2, excelStyle.fontBold, excelStyle.fontSize, excelStyle.alignText())
	style, _ := p.sheet.xlsx.NewStyle(txt)
	beg := makeFormatter(1, p.row)
	end := makeFormatter(p.sheet.colNum, p.row)
	p.sheet.xlsx.SetCellStyle(p.sheet.name, beg, end, style)
	return p
}

func (p *ExcelSheetRow) ApplyItem(col int, excelStyle *ExcelStyle) *ExcelSheetRow {
	txt := fmt.Sprintf(BaseStyle2, excelStyle.fontBold, excelStyle.fontSize, excelStyle.alignText())
	style, _ := p.sheet.xlsx.NewStyle(txt)

	beg := makeFormatter(col, p.row)
	p.sheet.xlsx.SetCellStyle(p.sheet.name, beg, beg, style)
	return p
}

type ExcelStyle struct {
	fontSize int
	//0 left, 1center, 2right
	align    int
	fontBold bool
}

func NewExcelStyle(size int, align int, bold bool) *ExcelStyle {
	excelStyle := &ExcelStyle{size, align, bold}
	return excelStyle
}

func (p *ExcelSheet) SetColor(startCol int, endCol int, colorHex string) *ExcelSheet {
	match, _ := regexp.MatchString("^#[0-9A-Fa-f]{6}$", colorHex)
	if !match {
		log.Println("wrong hex code")
		return p
	}
	styleStr := fmt.Sprintf(baseStyle, colorHex)
	style, err := p.xlsx.NewStyle(styleStr)
	if err != nil {
		log.Println(err)
		return p
	}
	for i := startCol; i <= endCol; i++ {
		for j := 1; j <= p.rowNum; j++ {
			index := makeFormatter(i, j)
			p.xlsx.SetCellStyle(p.name, index, index, style)
		}
	}
	return p
}

func (p *ExcelSheet) MergeRow() *ExcelSheet {
	var i, j int
	var index, nextRow string
	for i = 1; i <= p.colNum; i++ {
		num := 0
		merged := ""
		for j = 1; j <= p.rowNum; j++ {
			index = makeFormatter(i, j)
			if p.xlsx.GetCellValue(p.name, index) == "|" {
				num++
				nextRow = makeFormatter(i, j+1)
				if j+1 <= p.rowNum && p.xlsx.GetCellValue(p.name, nextRow) != "|" {
					merged = p.xlsx.GetCellValue(p.name, nextRow)
					p.MergeRowCell(j+1-num, j+1, i)
					num = 0
				}
			}
		}
		if len(merged) != 0 {
			p.SetCellValue(i, j-1, merged)
		}
	}
	return p
}

func (p *ExcelSheet) MergeRowCell(rowHead int, rowEnd int, col int) *ExcelSheet {
	p.xlsx.MergeCell(p.name, makeFormatter(col, rowHead), makeFormatter(col, rowEnd))
	return p
}

func toExcelColName(colNum int) string {
	result := ""

	for colNum > 0 {
		m := (colNum - 1) % 26
		result = string('A'+m) + result
		colNum = (colNum - m) / 26
	}
	return result
}

func (p *ExcelStyle) alignText() string {
	alignText := "center"
	if p.align == -1 {
		alignText = "left"
	} else if p.align == 1 {
		alignText = "right"
	}
	return alignText
}
