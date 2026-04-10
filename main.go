package main

import (
	"fmt"
	"log"
	"os"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/xuri/excelize/v2"
)

const version = "v0.1.5"

func printHelp() {
	fmt.Printf(`tt %s

Usage:
  tt [file.xlsx]
  tt --help
  tt --version

Options:
  -h, --help     Show this help
  -v, --version  Show version
`, version)
}

func main() {
	argsWithoutProg := os.Args[1:]

	for _, arg := range argsWithoutProg {
		switch arg {
		case "-h", "--help":
			printHelp()
			return
		case "-v", "--version":
			fmt.Println(version)
			return
		}
	}

	if len(argsWithoutProg) > 1 {
		fmt.Fprintf(os.Stderr, "tt: expected at most one file argument\n\n")
		printHelp()
		os.Exit(1)
	}

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
