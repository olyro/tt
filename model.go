package main

import (
	"strconv"
	"strings"
	"time"

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
	cells     map[string]cellSnapshot
}

type operationHistory struct {
	stack   []operationRecord
	pointer int
}

type operationRecord struct {
	operation     operation
	beforeStateID uint64
	afterStateID  uint64
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
	filePath       string
	sheetName      string
	selection      selection
	currentOp      operation
	opStack        []operationRecord
	opStackPointer int
	histories      map[string]operationHistory
	sheetStates    map[string]uint64
	savedStates    map[string]uint64
	nextStateID    uint64
	metadataState  uint64
	savedMetadata  uint64
	normalInput    string
	mode           mode
	searchQuery    string
	copy           *copy
	dirty          bool
	clipboard      clipboardService
	displayCache   map[string]string
}

const maxExcelRows = 1048576

func (m model) GetNrOfVisibleRows() int {
	return max((m.height-2)/2, 1)
}

func (m model) GetRowNrColumnWidth() int {
	return getNumberLength(m.offsetY + m.GetNrOfVisibleRows() - 1)
}

func (m model) GetNrOfVisibleColumns() int {
	return len(m.visibleColumnLayout())
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
		if cellType, err := m.excelFile.GetCellType(m.sheetName, address); err == nil && cellType == excelize.CellTypeFormula {
			formula, err := m.excelFile.GetCellFormula(m.sheetName, address)
			if err == nil {
				m.input.SetValue(prefixWithEqual(formula))
			}
		} else {
			value, err := m.excelFile.GetCellValue(m.sheetName, address)

			if err == nil {
				m.input.SetValue(value)
			}
		}
	}
}

func (m model) GetMaxColumn(rowIndex int) int {
	rows, err := m.excelFile.Rows(m.sheetName)

	if err != nil {
		return 0
	}
	defer rows.Close()

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
	beforeStateID := m.sheetStates[m.sheetName]
	m.nextStateID++
	afterStateID := m.nextStateID
	m.opStack = append(m.opStack, operationRecord{
		operation:     operation,
		beforeStateID: beforeStateID,
		afterStateID:  afterStateID,
	})
	m.opStackPointer++
	m.sheetStates[m.sheetName] = afterStateID
	m.updateDirtyState()
}

func (m *model) invalidateDisplayCache() {
	clear(m.displayCache)
}

func cloneStates(states map[string]uint64) map[string]uint64 {
	cloned := make(map[string]uint64, len(states))
	for sheetName, state := range states {
		cloned[sheetName] = state
	}
	return cloned
}

func (m *model) updateDirtyState() {
	if m.metadataState != m.savedMetadata || len(m.sheetStates) != len(m.savedStates) {
		m.dirty = true
		return
	}
	for sheetName, state := range m.sheetStates {
		if m.savedStates[sheetName] != state {
			m.dirty = true
			return
		}
	}
	m.dirty = false
}

func (m *model) markSaved() {
	m.savedStates = cloneStates(m.sheetStates)
	m.savedMetadata = m.metadataState
	m.dirty = false
}

func (m *model) markMetadataMutation() {
	m.metadataState++
	m.updateDirtyState()
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

	// if starts with ' remove prefix and set string
	if strings.HasPrefix(value, "'") {
		value = strings.TrimPrefix(value, "'")
		return m.excelFile.SetCellValue(sheetName, address, value)
	}

	if parsedInt, err := strconv.Atoi(value); err == nil {
		return m.excelFile.SetCellValue(sheetName, address, parsedInt)
	}

	if parsedFloat, err := strconv.ParseFloat(value, 64); err == nil {
		return m.excelFile.SetCellValue(sheetName, address, parsedFloat)
	}

	if parsedTime, err := time.Parse(time.DateTime, value); err == nil {
		return m.excelFile.SetCellValue(sheetName, address, parsedTime)
	}

	if parsedDuration, err := time.ParseDuration(value); err == nil {
		return m.excelFile.SetCellValue(sheetName, address, parsedDuration)
	}

	// if the value starts with =, set as formula
	if strings.HasPrefix(value, "=") {
		return m.excelFile.SetCellFormula(sheetName, address, value)
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
		m.ensureCursorColumnVisible()
	case clipboardWriteMsg:
		if msg.err != nil {
			m.input.SetValue("Clipboard error: " + msg.err.Error())
		} else {
			m.input.SetValue("copied to system clipboard")
		}
	case clipboardReadMsg:
		if msg.err != nil {
			m.input.SetValue("Clipboard error: " + msg.err.Error())
		} else {
			m = m.importClipboard(msg.text)
		}
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
	visibleColumns := m.visibleColumnLayout()
	merges := m.mergedCellLookup()

	for i := range height {
		i = i + m.offsetY
		widths := make([]int, len(visibleColumns)+1)
		widths[0] = rowNumberWidth
		for columnOffset, column := range visibleColumns {
			widths[columnOffset+1] = column.width
		}

		switch i {
		case m.offsetY:
			labels := make([]string, len(visibleColumns)+1)
			labels[0] = ""
			for columnOffset, column := range visibleColumns {
				label, err := excelize.ColumnNumberToName(column.index + 1)

				if err == nil {
					labels[columnOffset+1] = label
				}
			}

			line := m.formatRow(labels, widths, i, merges)
			b.WriteString(line + "\n")
		default:
			labels := make([]string, len(visibleColumns)+1)
			labels[0] = strconv.Itoa(i)

			for columnOffset, column := range visibleColumns {
				address, err := excelize.CoordinatesToCellName(column.index+1, i)
				if err == nil {

					// if the cell is a formula, get the value of calculated value
					cellType, err := m.excelFile.GetCellType(m.sheetName, address)

					if err == nil && cellType == excelize.CellTypeFormula {
						cacheKey := m.sheetName + "\x00" + address
						result, ok := m.displayCache[cacheKey]
						if !ok {
							result, err = m.excelFile.CalcCellValue(m.sheetName, address)
							if err == nil && m.displayCache != nil {
								m.displayCache[cacheKey] = result
							}
						}
						if err == nil {
							labels[columnOffset+1] = result
						}
					} else {
						value, err := m.excelFile.GetCellValue(m.sheetName, address)
						if err == nil {
							labels[columnOffset+1] = value
						}
					}
				}
			}

			line := m.formatRow(labels, widths, i, merges)
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

type mergePosition struct {
	merged bool
	first  bool
}

func (m model) mergedCellLookup() map[coordinate]mergePosition {
	lookup := make(map[coordinate]mergePosition)
	mergeCells, err := m.excelFile.GetMergeCells(m.sheetName)
	if err != nil {
		return lookup
	}
	columns := m.visibleColumnLayout()
	visibleStartX := columns[0].index
	visibleEndX := columns[len(columns)-1].index
	for _, mergeCell := range mergeCells {
		startX, startY, startErr := excelize.CellNameToCoordinates(mergeCell.GetStartAxis())
		endX, endY, endErr := excelize.CellNameToCoordinates(mergeCell.GetEndAxis())
		if startErr != nil || endErr != nil {
			continue
		}
		visibleStartY := m.offsetY
		visibleEndY := m.offsetY + m.GetNrOfVisibleRows() - 2
		for y := max(startY-1, visibleStartY); y <= min(endY-1, visibleEndY); y++ {
			for x := max(startX-1, visibleStartX); x <= min(endX-1, visibleEndX); x++ {
				lookup[coordinate{x: x, y: y}] = mergePosition{merged: true, first: x == startX-1 && y == startY-1}
			}
		}
	}
	return lookup
}

func (m model) formatRow(row []string, widths []int, rowIndex int, merges map[coordinate]mergePosition) string {
	var rendered []string
	for i, cell := range row {
		value := limitString(replaceNewLineWithWhiteSpace(cell), widths[i])

		merge := merges[coordinate{x: i + m.offsetX - 1, y: rowIndex - 1}]
		isMergeCell, isFirst := merge.merged, merge.first
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
