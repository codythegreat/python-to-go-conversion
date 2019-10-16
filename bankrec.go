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
var bankAmtReg = regexp.MustCompile(`\d*,?\d+\.\d{2}\-?`)
var dateDescReg = regexp.MustCompile(`\d{1}\d?\/\d{2}\n[\w ]*`)

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
	// get maximum row of 605 sheet
	rows, err := xlsx.GetRows("605")
	if err != nil {
		fmt.Println(err)
	}
	// extract entries from 605 sheet
	for currentRow := 9; currentRow <= len(rows); currentRow++ {
		something, err := xlsx.GetCellValue("605", "E"+strconv.Itoa(currentRow))
		if err != nil {
			fmt.Println(err)
		}
		floatAmount, err := strconv.ParseFloat(something, 64)
		if err != nil {
			fmt.Printf("%v", err)
		}
		dt, err := xlsx.GetCellValue("605", "C"+strconv.Itoa(currentRow))
		if err != nil {
			fmt.Printf("%v", err)
		}
		exp, err := xlsx.GetCellValue("605", "D"+strconv.Itoa(currentRow))
		if err != nil {
			fmt.Printf("%v", err)
		}
		entries = append(entries, entry{
			amount:       floatAmount,
			date:         dt,
			explaination: exp})
	}
	fmt.Println(entries)
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
	// print the results and original text
	return slicesRegex
}
func pdfAmountsToExcel(data [2][]string) {
	var pdfAmounts []float64
	for i, _ := range data[1] {
		floatAmt, err := strconv.ParseFloat(strings.Replace(data[1][i], ",", "", -1), 64)
		if err != nil {
			fmt.Println(err)
		}
		pdfAmounts = append(pdfAmounts, floatAmt)
	}
	f := excelize.NewFile()
	for i, _ := range data[0] {
		f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), strings.Replace(data[0][i], "\n", " - ", -1))
	}
	for i, _ := range pdfAmounts {
		f.SetCellValue("Sheet1", "B"+strconv.Itoa(i), pdfAmounts[i])
	}
	err := f.SaveAs("./Statement.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

//todo expand function to look at near matches, dates, description matching, etc
//todo give a marging of error of 5 cents for matches
func compareEntries(name string, lines []string) {
	xlsx, err := excelize.OpenFile(name)
	if err != nil {
		fmt.Println(err)
	}
	for i, item := range entries {
		for _, line := range lines {
			if line == fmt.Sprintf("%s", item.amount) {
				xlsx.SetCellValue("605", "F"+strconv.Itoa(9+i), "match")
			}
		}
	}
}

func main() {
	// get the name of the book from the user
	fileString := getFileName()
	//todo ask if user needs to extract pdf amounts, otherwise simply perform comparison
	lineAmounts := pullPDFAmounts()
	pdfAmountsToExcel(lineAmounts)
	//todo prompt user to continue after pdf cleanup
	extractEntries(fileString)
	compareEntries(fileString, lineAmounts[1])
}
