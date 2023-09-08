package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func output_excel() {
	xlsxFilePath := "data.xlsx"

	f, err := excelize.OpenFile(xlsxFilePath)
	if err != nil {
		log.Fatalf("Error opening XLSX file: %v", err)
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		log.Fatalf("%v", err)
	}
	length := len(rows)
	branches := make([]string, length-1)

	for i := 0; i < length-1; i++ {
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
	emails := make([]string, length-1)
	for i := 0; i < length-1; i++ {
		ID := rows[i+1][0]
		var ID_no string = ID[8:12]
		email := fmt.Sprintf("f2022%s@pilani.bits-pilani.ac.in", ID_no)
		emails[i] = email
	}

	f.SetCellValue("Sheet1", "C1", "BRANCH")
	f.SetCellValue("Sheet1", "D1", "BITS MAIL")

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

func student_details() {
	xlsxFilePath := "output.xlsx"
	f, err := excelize.OpenFile(xlsxFilePath)
	if err != nil {
		log.Fatalf("Error opening XLSX file: %v", err)
	}

	type student struct {
		name   string
		id     string
		branch string
		email  string
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		log.Fatalf("%v", err)
	}
	length := len(rows)
	students := make([]student, length-1)
	for i := 0; i < length-1; i++ {
		id := rows[i+1][0]
		name := rows[i+1][1]
		branch := rows[i+1][2]
		email := rows[i+1][3]
		var Student student
		Student.name = name
		Student.id = id
		Student.branch = branch
		Student.email = email

		students[i] = Student
	}

	for i := 0; i < length-1; i++ {
		fmt.Println(students[i])
	}
}

func main() {
	//to create output.xlsx file with branch and email
	output_excel()

	//to create and print student details using student structure
	student_details()
}
