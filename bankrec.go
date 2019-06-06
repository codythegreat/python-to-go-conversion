// app that analyzes bank statement PDF and bank rec excel sheet for matching amounts
package main

import (
	"bufio"
	"code.sajari.com/docconv"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"regexp"
	"strconv"
	"strings"
)

// struct to hold information on each entry
type entry struct {
	amount       float64
	date         string
	explaination string
}

// slice that holds all entries
var entries []entry
var bankAmounts []string
var bankAmtReg = regexp.MustCompile(`\d+.\d{2}\-?`)

func getFileName() string {
	reader := bufio.NewReader(os.Stdin)
	fmt.Print("Enter the book name containing this months entries: ")
	text, _ := reader.ReadString('\n')
	return text
}

func extractEntries() {
	// get the name of the book from the user
	fileString := getFileName()
	// initialize the workbook
	xlsx, err := excelize.OpenFile("./" + fileString)
	if err != nil {
		fmt.Println(err)
	}
	// get maximum row of JDE sheet
	rows := xlsx.GetRows("JDE")
	// extract entries from JDE sheet
	for currentRow := 9; currentRow <= len(rows); currentRow++ {
		floatAmount, err := strconv.ParseFloat(xlsx.GetCellValue("JDE", "E"+strconv.Itoa(currentRow)), 64)
		if err != nil {
			fmt.Printf("%v", err)
		}
		entries = append(entries, entry{
			amount:       floatAmount,
			date:         xlsx.GetCellValue("JDE", "C"+strconv.Itoa(currentRow)),
			explaination: xlsx.GetCellValue("JDE", "D"+strconv.Itoa(currentRow))})
	}
}

func pullPDFAmounts() []string {
	//reopen exce
	xlsx, err := excelize.OpenFile("./" + fileString)
	if err != nil {
		fmt.Println(err)
	}
	// pull all text from pdf doc
	res, err := docconv.ConvertPath("pdf.PDF")
	if err != nil {
		fmt.Printf("%v", err)
	}
	// use a regex to sort and find only dollar amounts
	lineAmounts := bankAmtReg.FindAllString(res, -1)
	// print the resuls and original text
	fmt.Println(lineAmounts)
	fmt.Println(res)
	return lineAmounts
}
func compareEntries() {
	xlsx, err := excelize.OpenFile("./" + fileString)
	if err != nil {
		fmt.Println(err)
	}
	for i, item := range entries {
		if strings.Contains(lineAmounts, fmt.Sprintf("%f", item.amount)) {
			xlsx.SetCellValue("JDE", "F"+strconv.Itoa(9+i), "match")
		}
	}
}

func main() {
	lineAmounts := pullPDFAmounts()
	extractEntries()
	compareEntries()
}
