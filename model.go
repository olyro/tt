package main

import (
	"math"
	"strconv"
	"strings"

	"github.com/charmbracelet/bubbles/textinput"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
	"github.com/xuri/excelize/v2"
)

type mode int

const (
	Normal mode = iota
	Visual
	Command
	Search
)

type copy struct {
	selection selection
	values    map[string]string
	styles    map[string]int
	types     map[string]excelize.CellType
}

type model struct {
	offsetX        int
	offsetY        int
	cursorX        int
	cursorY        int
	width          int
	height         int
	input          textinput.Model
	useInput       bool
	excelFile      *excelize.File
	sheetName      string
	selection      selection
	columnWidth    int
	currentOp      operation
	opStack        []operation
	opStackPointer int
	normalInput    string
	mode           mode
	searchQuery    string
	copy           *copy
}

func (m model) GetNrOfVisibleRows() int {
	return (m.height - 2) / 2
}

func (m model) GetRowNrColumnWidth() int {
	return getNumberLength(m.offsetY + m.GetNrOfVisibleRows() - 1)
}

func (m model) GetNrOfVisibleColumns() int {
	rowNrColumnWidth := m.GetRowNrColumnWidth()
	return int(math.Ceil(float64(m.width-rowNrColumnWidth-1) / float64(m.columnWidth+1)))
}

func (m model) Init() tea.Cmd {
	return nil
}

func (m model) GetCellValue() (string, error) {
	address, err := excelize.CoordinatesToCellName(m.cursorX+1, m.cursorY+1)

	if err != nil {
		return "", err
	}

	value, err := m.excelFile.GetCellValue(m.sheetName, address)
	if err != nil {
		return "", err
	}

	return value, nil
}

func (m *model) UpdateValuePrompt() {
	address, err := excelize.CoordinatesToCellName(m.cursorX+1, m.cursorY+1)

	if err == nil {
		value, err := m.excelFile.GetCellValue(m.sheetName, address)

		if err == nil {
			m.input.SetValue(value)
		}
	}
}

func (m model) GetMaxColumn(rowIndex int) int {
	rows, err := m.excelFile.Rows(m.sheetName)

	if err != nil {
		return 0
	}

	index := 0

	for rows.Next() {
		if rowIndex == index {
			columns, err := rows.Columns()
			if err == nil {
				return len(columns)
			}
		}
		index++
	}

	return 0
}

func (m *model) pushOp(operation operation) {
	if m.opStackPointer < len(m.opStack)-1 {
		m.opStack = m.opStack[:m.opStackPointer+1]
	}
	m.opStack = append(m.opStack, operation)
	m.opStackPointer++
}

func (m model) IsPartOfMergeCell(x, y int) (bool, bool) {
	mergeCells, err := m.excelFile.GetMergeCells(m.sheetName)
	if err != nil {
		return false, false
	}

	for _, mergeCell := range mergeCells {
		startX, startY, err := excelize.CellNameToCoordinates(mergeCell.GetStartAxis())

		if err != nil {
			return false, false
		}

		endX, endY, err := excelize.CellNameToCoordinates(mergeCell.GetEndAxis())

		if err != nil {
			return false, false
		}

		if startX == (x+1) && startY == (y+1) {
			return true, true // This is the first cell of a merged range
		}

		if startX <= (x+1) && endX >= (x+1) && startY <= (y+1) && endY >= (y+1) {
			return true, false
		}
	}

	return false, false
}

func (m model) setCellValue(sheetName, address, value string) error {
	if value == "" {
		return m.excelFile.SetCellValue(sheetName, address, nil)
	}

	if parsedFloat, err := strconv.ParseFloat(value, 64); err == nil {
		return m.excelFile.SetCellValue(sheetName, address, parsedFloat)
	}

	if parsedInt, err := strconv.Atoi(value); err == nil {
		return m.excelFile.SetCellValue(sheetName, address, parsedInt)
	}

	return m.excelFile.SetCellValue(sheetName, address, value)
}

func (m model) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	var cmd tea.Cmd
	switch msg := msg.(type) {
	case tea.MouseMsg:
		m = handleMouseEvent(m, msg)
	case tea.KeyMsg:
		m, cmd = handleKeyboardEvent(m, msg)
	case tea.WindowSizeMsg:
		m.width = msg.Width
		m.height = msg.Height
	}
	switch m.selection.kind {
	case CellSelect:
		m.selection.cell.x = m.cursorX
		m.selection.cell.y = m.cursorY
	case BlockSelect:
		m.selection.block.endX = m.cursorX
		m.selection.block.endY = m.cursorY
	case RowsSelect:
		m.selection.rows.end = m.cursorY
	case ColumnsSelect:
		m.selection.columns.end = m.cursorX
	}

	return m, cmd
}

func (m model) isSelected(row int, column int) bool {
	switch m.selection.kind {
	case CellSelect:
		if m.selection.cell.x == column && m.selection.cell.y == row {
			return true
		}
	case BlockSelect:
		if min(m.selection.block.startX, m.selection.block.endX) <= column && max(m.selection.block.startX, m.selection.block.endX) >= column &&
			min(m.selection.block.startY, m.selection.block.endY) <= row && max(m.selection.block.startY, m.selection.block.endY) >= row {
			return true
		}
	case RowsSelect:
		if min(m.selection.rows.start, m.selection.rows.end) <= row && max(m.selection.rows.start, m.selection.rows.end) >= row {
			return true
		}
	case ColumnsSelect:
		if min(m.selection.columns.start, m.selection.columns.end) <= column && max(m.selection.columns.start, m.selection.columns.end) >= column {
			return true
		}
	}

	return false
}

func (m model) View() string {
	var b strings.Builder

	height := m.GetNrOfVisibleRows()

	rowNumberWidth := m.GetRowNrColumnWidth()
	visibleColumns := m.GetNrOfVisibleColumns()

	for i := range height {
		i = i + m.offsetY
		widths := make([]int, m.GetNrOfVisibleColumns()+1)

		for j := 0; j <= visibleColumns; j++ {
			switch j {
			case 0:
				widths[j] = rowNumberWidth
			case visibleColumns:
				widths[j] = max(m.width-rowNumberWidth-1-(visibleColumns-1)*(m.columnWidth+1)-2, 1) // last column
			default:
				widths[j] = m.columnWidth // data columns
			}
		}

		switch i {
		case m.offsetY:
			labels := make([]string, m.GetNrOfVisibleColumns()+1)
			labels[0] = ""
			for j := 1; j <= m.GetNrOfVisibleColumns(); j++ {
				label, err := excelize.ColumnNumberToName(j + m.offsetX)

				if err == nil {
					labels[j] = label
				}
			}

			line := m.formatRow(labels, widths, i)
			b.WriteString(line + "\n")
		default:
			labels := make([]string, m.GetNrOfVisibleColumns()+1)
			labels[0] = strconv.Itoa(i)

			for j := 1; j <= m.GetNrOfVisibleColumns(); j++ {
				address, err := excelize.CoordinatesToCellName(j+m.offsetX, i)
				if err == nil {

					value, err := m.excelFile.GetCellValue(m.sheetName, address)
					if err == nil {
						labels[j] = value
					}
				}
			}

			line := m.formatRow(labels, widths, i)
			b.WriteString(line + "\n")
		}
	}

	leftWidth := lipgloss.Width(m.input.View())
	rightWidth := lipgloss.Width(m.normalInput)
	space := max(m.width-leftWidth-rightWidth, 0)
	bottomLine := m.input.View() + strings.Repeat(" ", space) + m.normalInput
	b.WriteString(bottomLine)

	return b.String()
}

func (m model) formatRow(row []string, widths []int, rowIndex int) string {
	var rendered []string
	for i, cell := range row {
		value := limitString(replaceNewLineWithWhiteSpace(cell), widths[i])

		isMergeCell, isFirst := m.IsPartOfMergeCell(i+m.offsetX-1, rowIndex-1)
		isSelected := m.isSelected(rowIndex-1, i+m.offsetX-1)

		if isMergeCell && !isFirst {
			value = strings.Repeat(" ", widths[i])
		}

		if isMergeCell && !isSelected {
			value = mergedStyle.Render(value)
		}

		if isSelected {
			value = selectedStyle.Render(value)
		}

		if rowIndex == m.offsetY {
			if i == 0 {
				rendered = append(rendered, headerTopLeft.Render(value))
			} else if i == len(row)-1 {
				rendered = append(rendered, headerTopRight.Render(value))
			} else {
				rendered = append(rendered, headerStyle.Render(value))
			}
		} else if rowIndex == m.offsetY+m.GetNrOfVisibleRows()-1 {
			if i == 0 {
				rendered = append(rendered, collapsedBottomLeftStyle.Bold(true).Render(value))
			} else if i == len(row)-1 {
				rendered = append(rendered, collapsedBottomRightStyle.Render(value))
			} else {
				rendered = append(rendered, collapsedBottomStyle.Render(value))
			}
		} else if i == len(row)-1 {
			rendered = append(rendered, collapsedRightStyle.Render(value))
		} else if i == 0 {
			rendered = append(rendered, collapsedLeftStyle.Bold(true).Render(value))
		} else {
			rendered = append(rendered, collapsedStyle.Render(value))
		}
	}
	return lipgloss.JoinHorizontal(lipgloss.Top, rendered...)
}
