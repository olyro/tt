package main

import (
	"log"
	"os"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/xuri/excelize/v2"
)

func main() {
	argsWithoutProg := os.Args[1:]

	var f *excelize.File

	if len(argsWithoutProg) < 1 {
		f = excelize.NewFile()
	} else {
		file, err := excelize.OpenFile(argsWithoutProg[0])
		if err != nil {
			log.Println(err)
			return
		}

		f = file
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			log.Println(err)
		}
	}()

	sheet := getFirstSheetOrCreate(f)

	if _, err := tea.NewProgram(initialModel(f, sheet), tea.WithAltScreen(), tea.WithMouseAllMotion()).Run(); err != nil {
		os.Exit(1)
	}
}
