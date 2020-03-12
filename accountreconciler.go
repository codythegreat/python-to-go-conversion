// attempt to recreate account_entry_checker_V2 (from python) to go
package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// each "match" will assume this struct
type outstandingAmount struct {
	amount      float64
	date        string
	description string
	remark      string
	batchNumb   string
}

// initialize slice to hold all matches from each book
var matches []outstandingAmount

func del_matching_data() {
	// clear the matches variable completely
	matches = matches[:0]
}

func extractAmounts() {
	// initialize the workbook
	xlsx, err := excelize.OpenFile("./Book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	// get the maximum row of the sheet
	rows := xlsx.GetRows("Sheet1")
	// append wanted information from each row to matches
	for currentRow := 2; currentRow <= len(rows)-4; currentRow++ {
		// if line has no amount, continue
		if xlsx.GetCellValue("sheet1", "H"+strconv.Itoa(currentRow)) == "" {
			continue
			// otherwise, append to our list of match instances
		} else {
			floatAmount, err := strconv.ParseFloat(xlsx.GetCellValue("sheet1", "H"+strconv.Itoa(currentRow)), 64)
			if err != nil {
				fmt.Printf("%v", err)
			}
			matches = append(matches, outstandingAmount{
				amount:      floatAmount,
				date:        xlsx.GetCellValue("sheet1", "F"+strconv.Itoa(currentRow)),
				description: xlsx.GetCellValue("sheet1", "G"+strconv.Itoa(currentRow)),
				remark:      xlsx.GetCellValue("sheet1", "AJ"+strconv.Itoa(currentRow)),
				batchNumb:   xlsx.GetCellValue("sheet1", "R"+strconv.Itoa(currentRow))})
		}
	}
}

func reduceAmounts() {
	// checks all amounts agains one another to see if they zero out (what we want)
	// if they do, zero out their amounts
	for i, _ := range matches {
		for j, _ := range matches {
			if matches[i].amount+matches[j].amount > -.005 && matches[i].amount+matches[j].amount < .005 {
				matches[i].amount = 0
				matches[j].amount = 0
			}
		}
	}
}

func printMatches() {
	// create a total variable
	var total float64
	// loop over and print all matches that do not equal zero. append each to total
	for _, match := range matches {
		total += match.amount
		if match.amount != 0 {
			fmt.Printf("%0.2f\t%s\t%s\n", match.amount, match.description, match.date)
		}
	}
	// print total
	fmt.Printf("\nTotal: %f\n", total)
	// create another total for up to 10 newest entries
	var recentEntriesTotal float64
	// starting from the last item, append to recent total. test if recent total == total
	// if it does, print the # of amounts from the bottom that make the total.
	for i := len(matches) - 1; i > len(matches)-10; i-- {
		recentEntriesTotal += matches[i].amount
		if recentEntriesTotal < total+.05 && recentEntriesTotal > total-.05 {
			fmt.Printf("Bottom %d matches make up amount.\n", len(matches)-i)
			break
		}
	}
}

func appendMatches(name string) {
	// open the book holding all account data
	masterBook, err := excelize.OpenFile("./account_recs.xlsx")
	if err != nil {
		fmt.Printf("While opening master file: %v", err)
	}
	// add a new sheet where the name is the user's input
	masterBook.NewSheet(name)
	// initialize a row counter starting at 1
	rowNumber := 1
	// write all matches that don't equal zero to the master book
	for _, match := range matches {
		if match.amount != 0 {
			masterBook.SetCellValue(name, "A"+strconv.Itoa(rowNumber), match.amount)
			masterBook.SetCellValue(name, "B"+strconv.Itoa(rowNumber), match.description)
			masterBook.SetCellValue(name, "C"+strconv.Itoa(rowNumber), match.remark)
			masterBook.SetCellValue(name, "D"+strconv.Itoa(rowNumber), match.date)
			masterBook.SetCellValue(name, "E"+strconv.Itoa(rowNumber), match.batchNumb)
			rowNumber++
		}
	}
	// save the book
	err = masterBook.SaveAs("./account_recs.xlsx")
	if err != nil {
		fmt.Printf("While saving master excel: %v", err)
	}

}

func programLoop() {
	// initialize scanner that will read user input
	scanner := bufio.NewScanner(os.Stdin)
	fmt.Println(`type "begin" to start the program.`)
	// grab the input and assign it to text
	scanner.Scan()
	text := scanner.Text()
	// if begin, find and print all matches from Book1
	if text == "begin" {
		fmt.Println("working...\n")
		extractAmounts()
		reduceAmounts()
		printMatches()
		fmt.Println("complete\n")
	}
	fmt.Println("Input sheet name:")
	scanner.Scan()
	fmt.Println("")
	text = scanner.Text()
	// if no sheet name, delete values and start over
	// else create a new sheet in the master book and write the values to it
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
	// run the loop
	programLoop()
}
