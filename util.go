package main

import (
	"os"
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

func initialModel(f *excelize.File, sheet string, filePath string) model {
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
	m.filePath = filePath
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

func (m model) withWorkbook(f *excelize.File, filePath string) model {
	m.excelFile = f
	m.filePath = filePath
	m.sheetName = getFirstSheetOrCreate(f)
	m.offsetX = 0
	m.offsetY = 0
	m.cursorX = 0
	m.cursorY = 0
	m.currentOp = nil
	m.opStack = make([]operation, 0)
	m.opStackPointer = -1
	m.normalInput = ""
	m.searchQuery = ""
	m.copy = nil
	m.useInput = false
	m.mode = Normal
	m.input.Prompt = ""
	m.input.Blur()
	m.resetToCellSelection()
	m.UpdateValuePrompt()
	return m
}

func (m model) hasUndoHistory() bool {
	return len(m.opStack) > 0
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

func expandHomeDir(path string) string {
	if path == "~" {
		homeDir, err := os.UserHomeDir()
		if err == nil {
			return homeDir
		}
		return path
	}

	if strings.HasPrefix(path, "~/") {
		homeDir, err := os.UserHomeDir()
		if err == nil {
			return homeDir + path[1:]
		}
	}

	return path
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

// isDateNumFmt returns true for well-known built-in Excel date/time number format IDs.
// This covers common IDs: 14-22 (dates/times) and 45-47 (time durations).
func isDateNumFmt(numFmtID int) bool {
	switch numFmtID {
	case 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47:
		return true
	default:
		return false
	}
}

// prefix with = if the string value has not already a =
func prefixWithEqual(s string) string {
	if strings.HasPrefix(s, "=") {
		return s
	}
	return "=" + s
}
