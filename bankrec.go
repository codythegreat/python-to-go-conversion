// app that analyzes bank statement PDF and bank rec excel sheet for matching amounts
// todo : fix hardcoded file names and work on pulling pdf values into go
package main

import (
	"bufio"
	"code.sajari.com/docconv"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"regexp"
	"strconv"
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

func extractEntries(name string) {
	// initialize the workbook
	xlsx, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	// get maximum row of JDE sheet
	rows, err := xlsx.GetRows("JDE")
	if err != nil {
		fmt.Println(err)
	}
	// extract entries from JDE sheet
	for currentRow := 9; currentRow <= len(rows); currentRow++ {
		something, err := xlsx.GetCellValue("JDE", "E"+strconv.Itoa(currentRow))
		if err != nil {
			fmt.Println(err)
		}
		floatAmount, err := strconv.ParseFloat(something, 64)
		if err != nil {
			fmt.Printf("%v", err)
		}
		dt, err := xlsx.GetCellValue("JDE", "C"+strconv.Itoa(currentRow))
		if err != nil {
			fmt.Printf("%v", err)
		}
		exp, err := xlsx.GetCellValue("JDE", "D"+strconv.Itoa(currentRow))
		if err != nil {
			fmt.Printf("%v", err)
		}
		entries = append(entries, entry{
			amount:       floatAmount,
			date:         dt,
			explaination: exp})
	}
}

func pullPDFAmounts() []string {
	// pull all text from pdf doc
	res, err := docconv.ConvertPath("bank-statement.pdf")
	if err != nil {
		fmt.Printf("%v", err)
	}
	// use a regex to sort and find only dollar amounts
	lineAmounts := bankAmtReg.FindAllString(fmt.Sprintf("%s", res), -1)
	// print the resuls and original text
	fmt.Println(lineAmounts)
	fmt.Println(res)
	return lineAmounts
}
func compareEntries(name string, lines []string) {
	xlsx, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	for i, item := range entries {
		for _, line := range lines {
			if line == fmt.Sprintf("%s", item.amount) {
				xlsx.SetCellValue("JDE", "F"+strconv.Itoa(9+i), "match")
			}
		}
	}
}

func main() {
	// get the name of the book from the user
	fileString := getFileName()
	lineAmounts := pullPDFAmounts()
	extractEntries(fileString)
	compareEntries(fileString, lineAmounts)
}
