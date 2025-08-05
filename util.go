package main

import (
	"strings"

	"github.com/charmbracelet/bubbles/textinput"
	"github.com/charmbracelet/lipgloss"
	"github.com/xuri/excelize/v2"
)

var (
	headerStyle = lipgloss.NewStyle().Bold(true).Border(lipgloss.Border{
		Top:         "─",
		Bottom:      "─",
		Left:        "│",
		Right:       "│",
		TopLeft:     "┬",
		TopRight:    "┬",
		BottomLeft:  "┼",
		BottomRight: "┼",
	}).BorderRight(false)
	headerTopLeft = lipgloss.NewStyle().Bold(true).Border(lipgloss.Border{
		Top:         "─",
		Bottom:      "─",
		Left:        "│",
		Right:       "│",
		TopLeft:     "┌",
		TopRight:    "┬",
		BottomLeft:  "├",
		BottomRight: "┼",
	}).BorderRight(false)
	headerTopRight = lipgloss.NewStyle().Bold(true).Border(lipgloss.Border{
		Top:         "─",
		Bottom:      "─",
		Left:        "│",
		Right:       "│",
		TopLeft:     "┬",
		TopRight:    "┐",
		BottomLeft:  "┼",
		BottomRight: "┤",
	})
	mergedStyle    = lipgloss.NewStyle().Background(lipgloss.Color("248")).Foreground(lipgloss.Color("230"))
	selectedStyle  = lipgloss.NewStyle().Background(lipgloss.Color("244")).Foreground(lipgloss.Color("230"))
	collapsedStyle = lipgloss.NewStyle().
			Border(lipgloss.Border{
			Top:         "─",
			Bottom:      "─",
			Left:        "│",
			Right:       "│",
			TopLeft:     "┼",
			TopRight:    "┼",
			BottomLeft:  "┼",
			BottomRight: "┼",
		}).
		BorderTop(false).
		BorderRight(false)
	collapsedLeftStyle = lipgloss.NewStyle().
				Border(lipgloss.Border{
			Top:         "─",
			Bottom:      "─",
			Left:        "│",
			Right:       "│",
			TopLeft:     "├",
			TopRight:    "┼",
			BottomLeft:  "├",
			BottomRight: "┼",
		}).
		BorderTop(false).
		BorderRight(false)
	collapsedRightStyle = lipgloss.NewStyle().
				Border(lipgloss.Border{
			Top:         "─",
			Bottom:      "─",
			Left:        "│",
			Right:       "│",
			TopLeft:     "┼",
			TopRight:    "┤",
			BottomLeft:  "┼",
			BottomRight: "┤",
		}).
		BorderTop(false)
	collapsedBottomRightStyle = lipgloss.NewStyle().
					Border(lipgloss.Border{
			Top:         "─",
			Bottom:      "─",
			Left:        "│",
			Right:       "│",
			TopLeft:     "┼",
			TopRight:    "┤",
			BottomLeft:  "┴",
			BottomRight: "┘",
		}).
		BorderTop(false)
	collapsedBottomLeftStyle = lipgloss.NewStyle().
					Border(lipgloss.Border{
			Top:         "─",
			Bottom:      "─",
			Left:        "│",
			Right:       "│",
			TopLeft:     "├",
			TopRight:    "┤",
			BottomLeft:  "└",
			BottomRight: "┴",
		}).
		BorderTop(false).
		BorderRight(false)
	collapsedBottomStyle = lipgloss.NewStyle().
				Border(lipgloss.Border{
			Top:         "─",
			Bottom:      "─",
			Left:        "│",
			Right:       "│",
			TopLeft:     "┼",
			TopRight:    "┼",
			BottomLeft:  "┴",
			BottomRight: "┴",
		}).
		BorderTop(false).
		BorderRight(false)
)

func getFirstSheetOrCreate(f *excelize.File) string {
	sheets := f.GetSheetList()
	sheet := ""

	if len(sheets) > 0 {
		sheet = sheets[0]
	} else {
		sheet = "Sheet1"
		f.NewSheet(sheet)
	}

	return sheet
}

func initialModel(f *excelize.File, sheet string) model {
	m := model{}
	ti := textinput.New()
	ti.Blur() // start unfocused
	ti.Prompt = ""

	startValue, err := f.GetCellValue(sheet, "A1")

	if err == nil {
		ti.SetValue(startValue)
	}

	m.input = ti
	m.excelFile = f
	m.sheetName = sheet
	m.columnWidth = 5
	m.selection = selection{
		kind: CellSelect,
		cell: &cellSelect{
			x: 0,
			y: 0,
		},
	}
	m.opStack = make([]operation, 0)
	m.opStackPointer = -1

	return m
}

func getNumberPrefix(input string) int {
	if len(input) == 0 {
		return 0
	}
	prefix := 0
	for i := 0; i < len(input); i++ {
		if input[i] < '0' || input[i] > '9' {
			break
		}
		prefix = prefix*10 + int(input[i]-'0')
	}
	if prefix == 0 && len(input) > 0 && input[0] == '-' {
		return -1
	}
	return prefix
}

func replaceNewLineWithWhiteSpace(s string) string {
	return strings.ReplaceAll(s, "\n", " ")
}

func limitString(s string, maxLength int) string {
	runes := []rune(s)
	if len(runes) > maxLength {
		return string(runes[:maxLength])
	} else if len(runes) < maxLength {
		return string(runes) + strings.Repeat(" ", maxLength-len(runes))
	}

	return s
}

// function that returns the length of a number in digits
func getNumberLength(n int) int {
	if n == 0 {
		return 1
	}
	length := 0
	if n < 0 {
		n = -n
		length++
	}
	for n > 0 {
		n /= 10
		length++
	}
	return length
}
