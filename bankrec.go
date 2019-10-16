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
var bankAmtReg = regexp.MustCompile(`\d*,?\d+\.\d{2}\-?`)
var dateDescReg = regexp.MustCompile(`\d{1}\d?\/\d{2}\b[\w\s]*\b`)

//todo add ability to code in pdf doc name and return [2]string of these names
func getFileName() string {
	scanner := bufio.NewScanner(os.Stdin)
	fmt.Println("Enter the book name containing this months entries")
	scanner.Scan()
	text := scanner.Text()
	return text
}

func extractEntries(name string) {
	// initialize the workbook
	xlsx, err := excelize.OpenFile(name)
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

func pullPDFAmounts() [2][]string {
	scanner := bufio.NewScanner(os.Stdin)
	// ask user for name of pdf file
	fmt.Println("Enter the name of the bank statement pdf")
	scanner.Scan()
	text := scanner.Text()
	// pull all text from pdf doc
	res, err := docconv.ConvertPath(text)
	if err != nil {
		fmt.Printf("%v", err)
	}
	var slicesRegex [2][]string
	// use a regex to sort and find only dollar amounts
	slicesRegex[0] = dateDescReg.FindAllString(fmt.Sprintf("%s", res), -1)
	slicesRegex[1] = bankAmtReg.FindAllString(fmt.Sprintf("%s", res), -1)
	// print the resuls and original text
	return slicesRegex
}
func pdfAmountsToExcel(strAmts []string) {
	// convert amounts to float64
	var pdfAmounts []float64
	for i, _ := range strAmts {
		floatAmt, err := strconv.ParseFloat(strAmts[i], 64)
		if err != nil {
			fmt.Println(err)
		}
		pdfAmounts[i] = floatAmt 
	}
	// initialize a new excel sheet

	// input amounts into the excel sheet

	// save the sheet
}

//todo expand function to look at near matches, dates, description matching, etc
func compareEntries(name string, lines []string) {
	xlsx, err := excelize.OpenFile(name)
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
	//todo ask if user needs to extract pdf amounts, otherwise simply perform comparison
	lineAmounts := pullPDFAmounts()
	//todo pdfAmountsToExcel(lineAmounts)
	//todo prompt user to continue after pdf cleanup
	extractEntries(fileString)
	compareEntries(fileString, lineAmounts[1])
}
