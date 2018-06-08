package main

import (
	"os"

	"github.com/tealeg/xlsx"
)

var md string

func check(e error) {
	if e != nil {
		panic(e)
	}
}

func main() {

	// 檔案路徑
	originPath := os.Args[1]
	// 寫出位置
	destinationPath := os.Args[2]

	// 開啟 excel 檔案
	xlFile, err := xlsx.OpenFile(originPath)
	// xlFile, err := xlsx.OpenFile("resources/example.xlsx")
	// 檢查檔案是否有誤
	check(err)
	// Excel sheet
	for _, sheet := range xlFile.Sheets {
		md += "# " + sheet.Name + "\r\n\r\n"
		// i : 0 ， Header
		for i := 0; i < sheet.MaxRow; i++ {
			md += "|"
			row := sheet.Rows[i]
			cellsLen := len(row.Cells)
			if i == 1 {
				for j := 0; j < cellsLen; j++ {
					md += markdownAppend(j, cellsLen, " ---- ")
				}
				md += "|"
			}
			for j := 0; j < cellsLen; j++ {
				cell := row.Cells[j]
				text := " " + cell.String() + " "
				md += markdownAppend(j, cellsLen, text)
			}

		}
		md += "\r\n"
	}

	// 檔案寫入

	f, err2 := os.Create(destinationPath)
	// f, err2 := os.Create("C:\\tmp\\tmp.md")
	check(err2)
	f.WriteString(md)
	f.Sync()
}

func markdownAppend(count, sum int, text string) (result string) {
	if count == (sum - 1) {
		result += text + "| \r\n"
	} else {
		result += text + "|"
	}
	return result
}
