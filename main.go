package main

import (
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	src, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	dst := excelize.NewFile()
	src.CopySheetFrom()

	// // Create a new sheet.
	// index := f.NewSheet("Sheet2")
	// // Set value of a cell.
	// f.SetCellValue("Sheet2", "A2", "Hello world.")
	// f.SetCellValue("Sheet1", "B2", 100)
	// // Set active sheet of the workbook.
	// f.SetActiveSheet(index)
	// // Save spreadsheet by the given path.
	// if err := f.SaveAs("Book1.xlsx"); err != nil {
	// 	log.Fatal(err)
	// }

	// f, err := excelize.OpenFile("Book1.xlsx")
	// if err != nil {
	// 	log.Fatal(err)
	// }
	// //defer f.Close()

	// // Get value from cell by given worksheet name and cell reference.
	// cell := f.GetCellValue("Sheet1", "B2")
	// fmt.Println(cell)

	// // Get all the rows in the Sheet1.
	// rows := f.GetRows("Sheet1")

	// for _, row := range rows {
	// 	for _, colCell := range row {
	// 		fmt.Print(colCell, "\t")
	// 	}
	// 	fmt.Println()
	// }
}
