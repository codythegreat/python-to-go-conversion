// app that analyzes bank statement PDF and bank rec excel sheet for matching amounts
package main

import (
	"bufio"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
)

// struct to hold information on each entry
type entry struct {
	amount       float64
	date         string
	explaination string
}

// slice that holds all entries
var entries []entry

func getFileName() {
	reader := bufio.NewReader(os.Stdin)
	fmt.Print("Enter the book name containing this months entries: ")
	text, _ := reader.ReadString('\n')
	return text
}

func extractEntries() {
	// get the name of the book from the user
	fileString := getFileName()
	// initialize the workbook
	xlsx, err := excelize.OpenFile("./" + FileString)
	if err != nil {
		fmt.Println(err)
	}
	// get maximum row of JDE sheet
	rows := xlsx.GetRows("JDE")
	// extract entries from JDE sheet
	for currentRow := 9; currentRow <= len(rows); current++ {
		floatAmount, err := srtconv.ParseFloat(xlsx.GetCellValue("JDE", "E"+strconv.Itoa(currentRow)), 64)
		if err != nil {
			fmt.Printf("%v", err)
		}
		entries = append(entries, entry{
			amount:       floatAmount,
			date:         xlsx.GetCellValue("JDE", "C"+strconv.Itoa(currentRow)),
			explaination: xlsx.GetCellValue("JDE", "D"+strconv.Itoa(currentRow))})
	}
}
