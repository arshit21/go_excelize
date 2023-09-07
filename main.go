package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	xlsxFilePath := "data.xlsx"

	f, err := excelize.OpenFile(xlsxFilePath)
	if err != nil {
		log.Fatalf("Error opening XLSX file: %v", err)
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		log.Fatalf("%v", err)
	}
	var branches [22]string

	for i := 0; i < 22; i++ {
		ID := rows[i+1][0]
		if rune(ID[4]) == 'A' {
			if ID[5] == 49 {
				branch := "Chemical"
				branches[i] = branch
			} else if ID[5] == 50 {
				branch := "Civil"
				branches[i] = branch
			} else if ID[5] == 51 {
				branch := "EEE"
				branches[i] = branch
			} else if ID[5] == 52 {
				branch := "Mechanical"
				branches[i] = branch
			} else if ID[5] == 53 {
				branch := "B.Pharma"
				branches[i] = branch
			} else if ID[5] == 55 {
				branch := "Computer Science"
				branches[i] = branch
			} else if ID[5] == 56 {
				branch := "ENI"
				branches[i] = branch
			} else if ID[5] == 65 {
				branch := "ECE"
				branches[i] = branch
			} else if ID[5] == 66 {
				branch := "Manufacturing"
				branches[i] = branch
			}
		} else {
			if ID[5] == 49 {
				branch := "Msc Biology"
				branches[i] = branch
			} else if ID[5] == 50 {
				branch := "Msc Chemistry"
				branches[i] = branch
			} else if ID[5] == 51 {
				branch := "Msc Economics"
				branches[i] = branch
			} else if ID[5] == 52 {
				branch := "Msc Mathematics"
				branches[i] = branch
			} else if ID[5] == 53 {
				branch := "Msc Physics"
				branches[i] = branch
			}
		}
	}
	var emails [22]string
	for i := 0; i < 22; i++ {
		ID := rows[i+1][0]
		var ID_no string = ID[8:12]
		email := fmt.Sprintf("f2022%s@pilani.bits-pilani.ac.in", ID_no)
		emails[i] = email
	}

	f.SetCellValue("Sheet1", "C1", "Branch")
	f.SetCellValue("Sheet1", "D1", "Bits Mail")

	for i, value := range branches {
		cellname := fmt.Sprintf("C%d", i+2)
		f.SetCellValue("Sheet1", cellname, value)
	}

	for i, value := range emails {
		cellname := fmt.Sprintf("D%d", i+2)
		f.SetCellValue("Sheet1", cellname, value)
	}

	if err := f.SaveAs("output.xlsx"); err != nil {
		fmt.Println(err)
	}

}
