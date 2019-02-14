// attempt to recreate account_entry_checker_V2 (from python) to go
package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type outstandingAmount struct {
	amount      float64
	date        string
	description string
}

var matches []outstandingAmount

func del_matching_data() {
	matches = matches[:0]
}

func extractAmounts() {
	// initialize the workbook
	xlsx, err := excelize.OpenFile("./Book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	rows := xlsx.GetRows("Sheet1")
	for currentRow := 2; currentRow <= len(rows)-4; currentRow++ {
		floatAmount, err := strconv.ParseFloat(xlsx.GetCellValue("sheet1", "H"+strconv.Itoa(currentRow)), 64)
		if err != nil {
			fmt.Printf("%v", err)
		}
		matches = append(matches, outstandingAmount{
			amount:      floatAmount,
			date:        xlsx.GetCellValue("sheet1", "F"+strconv.Itoa(currentRow)),
			description: xlsx.GetCellValue("sheet1", "G"+strconv.Itoa(currentRow))})
	}
}

func reduceAmounts() {
	for i, _ := range matches {
		for j, _ := range matches {
			if matches[i].amount+matches[j].amount > -.01 && matches[i].amount+matches[j].amount < .01 {
				matches[i].amount = 0
				matches[j].amount = 0
			}
		}
	}
}

func printMatches() {
	var total float64
	for _, match := range matches {
		total += match.amount
		if match.amount != 0 {
			fmt.Printf("%f\t%s\t%s\n", match.amount, match.description, match.date)
		}
	}
	fmt.Printf("\nTotal: %f\n", total)
}

func appendMatches(name string) {
	masterBook, err := excelize.OpenFile("./account_recs.xlsx")
	if err != nil {
		fmt.Printf("While opening master file: %v", err)
	}
	masterBook.NewSheet(name)
	rowNumber := 0
	for _, match := range matches {
		if match.amount != 0 {
			masterBook.SetCellValue(name, "A"+strconv.Itoa(rowNumber+1), match.amount)
			masterBook.SetCellValue(name, "B"+strconv.Itoa(rowNumber+1), match.description)
			masterBook.SetCellValue(name, "C"+strconv.Itoa(rowNumber+1), match.date)
			rowNumber++
		}
	}
	err = masterBook.SaveAs("./account_recs.xlsx")
	if err != nil {
		fmt.Printf("While saving master excel: %v", err)
	}

}

func programLoop() {
	scanner := bufio.NewScanner(os.Stdin)
	fmt.Println(`type "begin" to start the program.`)
	scanner.Scan()
	text := scanner.Text()
	if text == "begin" {
		fmt.Println("working...\n")
		extractAmounts()
		reduceAmounts()
		printMatches()
		fmt.Println("complete\n")
	}
	fmt.Println("Input sheet name:")
	scanner.Scan()
	text = scanner.Text()
	fmt.Println(text)
	if text == "" {
		del_matching_data()
		programLoop()
	} else {
		appendMatches(text)
		del_matching_data()
		programLoop()
	}
}

func main() {
	programLoop()
}
