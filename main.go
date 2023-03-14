package main

import (
	"fmt"
	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
	"strings"
)

func main() {
	if len(os.Args) == 1 {
		fmt.Println("Specify last working row of file")
		os.Exit(1)
	}
	lastRealRow, err := strconv.Atoi(os.Args[1])
	if err != nil {
		fmt.Println("Unable to get last working row")
		os.Exit(1)
	}
	lastFakeRow, err := strconv.Atoi(os.Args[2])
	entries, err := os.ReadDir("./")
	if err != nil {
		fmt.Println(errors.Wrap(err, "Unable to read working directory"))
		os.Exit(1)
	}
	var fileName string
	for _, e := range entries {
		if strings.HasSuffix(e.Name(), ".xlsx") {
			fileName = e.Name()
			break
		}
	}
	if fileName == "" {
		fmt.Println("Excel files not found")
		os.Exit(1)
	}
	fmt.Println("Found file", fileName)
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(errors.Wrap(err, "Unable to load excel file"))
		os.Exit(1)
	}
	sheetIndex := f.GetActiveSheetIndex()
	sheetName := f.GetSheetName(sheetIndex)
	deleted := 0
	total := float32(lastFakeRow - lastRealRow)
	for i := lastFakeRow; i > lastRealRow; i-- {
		err = f.RemoveRow(sheetName, i)
		if err != nil {
			fmt.Println(errors.Wrap(err, "Unable to remove row "+strconv.Itoa(i)))
			os.Exit(1)
		}
		deleted++
		completed := float32(deleted) / total * float32(100)
		check := float32(int(completed))
		fmt.Println("Deleted", deleted)
		if completed == check && int(check)%1 == 0 {
			fmt.Printf("Completed %v%%\n", completed)
		}
	}
	fmt.Printf("Successfully deleted %v rows from sheet %v\n", deleted, sheetName)
	err = f.Save()
	if err != nil {
		fmt.Println(errors.Wrap(err, "Unable to save file"))
		os.Exit(1)
	}
}
