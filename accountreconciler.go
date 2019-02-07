// attempt to recreate account_entry_checker_V2 (from python) to go
package main

import (
	"bufio"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"strconv"
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
		floatAmount, err := strconv.ParseFloat(xlsx.GetCellValue("sheet1", "F"+strconv.Itoa(currentRow)), 64)
		if err != nil {
			fmt.Printf("%v", err)
		}
		matches = append(matches, outstandingAmount{
			amount:      floatAmount,
			date:        xlsx.GetCellValue("sheet1", "F"+strconv.Itoa(currentRow)),
			description: xlsx.GetCellValue("sheet1", "F"+strconv.Itoa(currentRow))})
		fmt.Println(xlsx.GetCellValue("sheet1", "F"+strconv.Itoa(currentRow)))
		fmt.Println(xlsx.GetCellValue("sheet1", "G"+strconv.Itoa(currentRow)))
		fmt.Println(xlsx.GetCellValue("sheet1", "H"+strconv.Itoa(currentRow)))
	}
}

func reduceAmounts() {
	for i, matchX := range matches {
		for j, matchY := range matches {
			if matchX.amount+matchY.amount == 0 {
				matches = append(matches[:i], matches[i+1:]...)
				matches = append(matches[:j], matches[j+1:]...)
			}
		}
	}
}

func appendMatches() {
	for _, match := range matches {
		fmt.Printf("%d\t%s\t%s\n", match.amount, match.description, match.date)
	}
}

func programLoop() {
	scanner := bufio.NewScanner(os.Stdin)
	fmt.Println(`type "begin" to start the program.`)
	scanner.Scan()
	text := scanner.Text()
	if text == "begin" {
		fmt.Println("working...")
		extractAmounts()
		reduceAmounts()
		fmt.Println("complete\n\n")
	}
	for i, amount := range matches {
		fmt.Printf("%d\t%d", i, amount)
	}
	fmt.Println("Input sheet name:")
	scanner.Scan()
	text = scanner.Text()
	if text != "" {
		del_matching_data()
		programLoop()
	} else {
		appendMatches()
		del_matching_data()
		programLoop()
	}
}

func main() {
	programLoop()
}
