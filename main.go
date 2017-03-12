package main

import (
	"errors"
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

var xlsxPath = flag.String("f", "", "Path to an XLSX file")
var sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")
var delimiter = flag.String("d", ";", "Delimiter to use between fields")
var skipheader = flag.Int("h", 0, "Number of top rows to skip")

type outputer func(s string)

func generateCSVFromXLSXFile(excelFileName string, sheetIndex int, outputf outputer) error {
	xlFile, error := xlsx.OpenFile(excelFileName)
	if error != nil {
		return error
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}
	sheet := xlFile.Sheets[sheetIndex]
	for rownum, row := range sheet.Rows {
		var vals []string
		if row != nil && rownum >= *skipheader {
			for _, cell := range row.Cells {
				str := cell.Value
				vals = append(vals, fmt.Sprintf("%q", str))
			}
			//Skip empty rows
			allempty := true
			for col := range vals {
				if vals[col] != "\"\"" {
					allempty = false
				}
			}
			if allempty == false {
				outputf(strings.Join(vals, *delimiter) + "\n")
			}
		}
	}
	return nil
}

func main() {
	flag.Parse()
	if len(os.Args) < 4 {
		flag.PrintDefaults()
		return
	}
	flag.Parse()
	printer := func(s string) { fmt.Printf("%s", s) }
	if err := generateCSVFromXLSXFile(*xlsxPath, *sheetIndex, printer); err != nil {
		fmt.Println(err)
	}
}
