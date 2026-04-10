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
	filePath := ""

	if len(argsWithoutProg) < 1 {
		f = excelize.NewFile()
	} else {
		filePath = argsWithoutProg[0]
		file, err := excelize.OpenFile(filePath)
		if err != nil {
			log.Println(err)
			return
		}

		f = file
	}

	sheet := getFirstSheetOrCreate(f)

	finalModel, err := tea.NewProgram(initialModel(f, sheet, filePath), tea.WithAltScreen(), tea.WithMouseAllMotion()).Run()
	if err != nil {
		if closeErr := f.Close(); closeErr != nil {
			log.Println(closeErr)
		}
		os.Exit(1)
	}

	if m, ok := finalModel.(model); ok && m.excelFile != nil {
		if err := m.excelFile.Close(); err != nil {
			log.Println(err)
		}
	}
}
