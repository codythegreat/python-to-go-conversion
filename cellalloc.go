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
	// parse the PDF file to a string variable
	pdfText, err := readPdf("employees.pdf")
	if err != nil {
		fmt.Printf("While reading PDF: %v\n", err)
	}
	// open the existing excel file
	cellBook, err := excelize.OpenFile("./cellular.xlsx")
	// get the maximum row number in the sheet (used in for loop)
	rows := cellBook.GetRows("Sheet1")
	if err != nil {
		fmt.Printf("While opening cellular file: %v", err)
	}
	// starting at row 2 and moving down to the end of the book
	var name []string
	var commaName bool
	var nameFormattedToPDF string
	var nameWithInitial string
	for i := 2; i < len(rows); i++ {
		// take the original name and split at the space
		if strings.Contains(cellBook.GetCellValue("sheet1", "M"+strconv.Itoa(i)), ",") {
			commaName = true
			name = strings.Split(cellBook.GetCellValue("sheet1", "M"+strconv.Itoa(i)), ", ")
		} else {
			commaName = false
			name = strings.Split(cellBook.GetCellValue("sheet1", "M"+strconv.Itoa(i)), " ")
		}
		// avoid "staff", "managers", or other single word names
		if len(name) < 2 {
			continue
		}
		// reverse the order to match the PDF's formatting
		if commaName {
			nameFormattedToPDF = name[0] + ", " + name[1]
			nameWithInitial = name[0] + ", " + string(name[1][0])
		} else {
			nameFormattedToPDF = name[1] + ", " + name[0]
			nameWithInitial = name[1] + ", " + string(name[0][0])
		}
		fmt.Printf("searching for name at row %d:\t%s\n", i, nameFormattedToPDF)

		if strings.Contains(pdfText, nameFormattedToPDF) {
			cellBook.SetCellValue("Sheet1", "N"+strconv.Itoa(i), "PERFECT MATCH")
			fmt.Println("PERFECT MATCH")
		} else if strings.Contains(pdfText, nameWithInitial) {
			cellBook.SetCellValue("Sheet1", "N"+strconv.Itoa(i), "PARTIAL MATCH")
			fmt.Println("PARTIAL MATCH")
		} else {
			cellBook.SetCellValue("Sheet1", "N"+strconv.Itoa(i), "NONMATCH")
			fmt.Println("NONMATCH")
		}
	}
	// save the edited book
	err = cellBook.SaveAs("./cellular_complete.xlsx")
	if err != nil {
		fmt.Printf("While saving cellular.xlsx: %v", err)
	}
}

// function for parsing the PDF file
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
