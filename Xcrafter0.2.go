/*
 * --------------------------------------------------------------------------------
 * "SUSHIWARE LICENSE" (Revision 01):
 * As long as you retain this notice, you can do whatever you want with this code.
 * If we meet someday around the universe, and you think this code was useful,
 * if you want, you can pay me a sushi round in return.
 * Ewerson Guimaraes a.k.a Crash
 * --------------------------------------------------------------------------------
 */

package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"strconv"
	"strings"
)

const Xversion = "0.2"

func main() {
	var argCells string
	flag.StringVar(&argCells, "s", "", "Single cells where the payload will be placed. Use a comma as a separator. Ex: A1,H4,D20")
	var payload string
	flag.StringVar(&payload, "p", "", "Any payload to be written in the file.")
	var argRow string
	flag.StringVar(&argRow, "l", "", "Line range of a column where the payload will be placed. Use a colon as a range separator and a comma to add a new range. Ex. A1:A10,B1:B10")
	var argCol string
	flag.StringVar(&argCol, "c", "", "Column as a range where the payload will be placed. Use a colon as a range separator and a comma to add a new range. Ex: C1:F1,J7:N7,H1:K1")
	var outPut string
	flag.StringVar(&outPut, "o", "Xcrafter.xlsx", "Crafted file name output.")
	var workSheet string
	flag.StringVar(&workSheet, "w", "Sheet1", "Set the worksheet name.")
	var specialPayload string
	flag.StringVar(&specialPayload, "e", "", "Use this option to set a different payload from -p option. For single cells only.")
	version := flag.Bool("v", false, "Prints the current version and exit.")
	flag.Parse()
	//fmt.Println("Cells:", argCells) //debug
	//fmt.Println("Worksheet name:", workSheet) //debug
	if *version {
		fmt.Println(Xversion)
		fmt.Println("Under the SushiWare license")
		os.Exit(0)
	}

	xlsx := excelize.NewFile()
	if workSheet != "Sheet1" {
		xlsx.SetSheetName("Sheet1",workSheet)
	}
	cells := strings.Split(argCells, ",")
	//fmt.Println(cells[0]) // debug
	//fmt.Println("Array size:",len(cells)) //debug
	if argCells == "" {
	}else {
			for i := 0; i < len(cells); i++ {
				fmt.Println("Writing the cell:", cells[i])
				if specialPayload == "" {
					xlsx.SetCellValue(workSheet, cells[i], payload)
				} else {
					xlsx.SetCellValue(workSheet, cells[i], specialPayload)
				}
			}

		}

	if argRow == "" {
	}else {
		loopRow := strings.Split(argRow, ",")
				for a := 0; a < len(loopRow); a++ {
					row := strings.Split(loopRow[a], ":")
					//fmt.Println("Row: ",row) //debug
					var row_line_start, row_number_start, _ = excelize.SplitCellName(row[0])
					var _, row_number_end, _ = excelize.SplitCellName(row[1])
					//fmt.Println("Letra de inicio: ", row_line_start) //debug
					//fmt.Println("Numero de inicio: ", row_number_start) //debug
					//fmt.Println("Numero final: ", row_number_end) //debug
						for i := row_number_start; i <= row_number_end; i++ {
							fmt.Println("Writing the cell:", row_line_start+strconv.Itoa(i))
							xlsx.SetCellValue(workSheet, row_line_start+strconv.Itoa(i), payload)
						}
				}
		}
		if argCol == "" {
		}else {
			colsLoop := strings.Split(argCol, ",")
			//fmt.Println("Argumentos de coluna:", colsLoop) // debug
			for m := 0; m < len(colsLoop); m++ {
				colsRange := strings.Split(colsLoop[m], ":")
				// fmt.Println("Range separado: ", colsRange) //debug
				// fmt.Println("Coluna de Inicio: ", colsRange[0]) //debug
				xlsx.SetCellValue(workSheet, colsRange[0], payload)
				//fmt.Println("Coluna Final: ", colsRange[1]) //debug
				coordCol, coordLn,_ := excelize.CellNameToCoordinates(colsRange[1])
				coordColIni, _,_ := excelize.CellNameToCoordinates(colsRange[0])
				// fmt.Println("Cordenadas da coluna inicial - Letra  ", coordColIni) //debug
				// fmt.Println("Cordenadas da coluna inicial - Linha ", coordLnIni) //debug
				// fmt.Println("Cordenadas da coluna final - Letra  ", coordCol) //debug
				// fmt.Println("Cordenadas da coluna final - Linha ", coordLn) //debug

					for p := coordColIni; p <= coordCol; p++ {
						getCel, _ := excelize.CoordinatesToCellName(p,coordLn)
						xlsx.SetCellValue(workSheet, getCel, payload)
						fmt.Println("Writing the cell:", getCel)
						}
				}
		}
	err := xlsx.SaveAs(outPut)
	if err != nil {
		fmt.Println(err)
	}  else {
		fmt.Println("The crafted Excel file " + outPut + " has been created!")
	}
	}
