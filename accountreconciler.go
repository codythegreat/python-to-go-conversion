// attempt to recreate account_entry_checker_V2 (from python) to go
package main

import (
	"bufio"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"time"
)

type outstandingAmount struct {
	amount      int64
	date        time.Time
	description string
}

var matches []outstandingAmount

func del_matching_data() {
	matches = matches[:0]
}

func compareAmounts() {
	// initialize the workbook
	xlsx, err := excelize.OpenFile("./Book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	// grab all rows
	rows := xlsx.GetRows("sheet1")
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Println(colCell, "\t")
		}
	}
}

func appendMatches() {

}

func programLoop() {
	scanner := bufio.NewScanner(os.Stdin)
	fmt.Println(`type "begin" to start the program.`)
	scanner.Scan()
	text := scanner.Text()
	if text == "begin" {
		fmt.Println("working...")
		compareAmounts()
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
