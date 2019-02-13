// attempt to recreate account_entry_checker_V2 (from python) to go
package main

import (
	"bytes"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/ledongthuc/pdf"
	"strconv"
	"strings"
)

func main() {
	pdfText, err := readPdf("employees.pdf")
	if err != nil {
		fmt.Printf("While reading PDF: %v\n", err)
	}
	cellBook, err := excelize.OpenFile("./cellular.xlsx")
	rows := cellBook.GetRows("Sheet1")
	if err != nil {
		fmt.Printf("While opening cellular file: %v", err)
	}

	for i := 2; i < len(rows); i++ {
		name := strings.Split(cellBook.GetCellValue("sheet1", "M"+strconv.Itoa(i)), " ")
		nameFormattedToPDF := name[1] + ", " + name[0]
		nameWithInitial := name[1] + ", " + string(name[0][0])
		if strings.Contains(pdfText, nameFormattedToPDF) {
			cellBook.SetCellValue("Sheet1", "M"+strconv.Itoa(i), cellBook.GetCellValue("sheet1", "M"+strconv.Itoa(i))+" $")
		} else if strings.Contains(pdfText, nameWithInitial) {
			cellBook.SetCellValue("Sheet1", "M"+strconv.Itoa(i), cellBook.GetCellValue("sheet1", "M"+strconv.Itoa(i))+" $$")
		} else {
			cellBook.SetCellValue("Sheet1", "M"+strconv.Itoa(i), cellBook.GetCellValue("sheet1", "M"+strconv.Itoa(i))+" @")
		}
	}
}
func readPdf(path string) (string, error) {
	f, r, err := pdf.Open(path)
	// remember to close the pdf file
	defer f.Close()
	if err != nil {
		return "", err
	}
	var buf bytes.Buffer
	b, err := r.GetPlainText()
	if err != nil {
		return "", err
	}
	buf.ReadFrom(b)
	return buf.String(), nil
}
