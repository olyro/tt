package main

import (
	"fmt"
	"math"
	"strings"

	"github.com/charmbracelet/lipgloss"
	"github.com/xuri/excelize/v2"
)

type visibleColumn struct {
	index int
	width int
}

func (m model) columnDisplayWidth(columnIndex int) int {
	columnName, err := excelize.ColumnNumberToName(columnIndex + 1)
	if err != nil {
		return 1
	}
	width, err := m.excelFile.GetColWidth(m.sheetName, columnName)
	if err != nil {
		return 1
	}
	return max(int(math.Round(width)), 1)
}

func (m model) visibleColumnLayout() []visibleColumn {
	availableWidth := max(m.width-m.GetRowNrColumnWidth()-2, 1)
	columns := make([]visibleColumn, 0)
	usedWidth := 0

	for columnIndex := m.offsetX; columnIndex < excelize.MaxColumns; columnIndex++ {
		remainingWidth := availableWidth - usedWidth
		if remainingWidth <= 1 && len(columns) > 0 {
			break
		}
		width := m.columnDisplayWidth(columnIndex)
		if width+1 > remainingWidth {
			width = max(remainingWidth-1, 1)
			columns = append(columns, visibleColumn{index: columnIndex, width: width})
			break
		}
		columns = append(columns, visibleColumn{index: columnIndex, width: width})
		usedWidth += width + 1
	}

	if len(columns) == 0 {
		columns = append(columns, visibleColumn{index: m.offsetX, width: 1})
	}
	return columns
}

func (m model) columnAtScreenX(screenX int) int {
	relativeX := max(screenX-(m.GetRowNrColumnWidth()+1), 0)
	columns := m.visibleColumnLayout()
	usedWidth := 0
	for _, column := range columns {
		usedWidth += column.width + 1
		if relativeX < usedWidth {
			return column.index
		}
	}
	return columns[len(columns)-1].index
}

func (m model) offsetWithColumnAtRight(columnIndex int) int {
	availableWidth := max(m.width-m.GetRowNrColumnWidth()-2, 1)
	startColumn := columnIndex
	usedWidth := m.columnDisplayWidth(columnIndex) + 1
	for previousColumn := columnIndex - 1; previousColumn >= 0; previousColumn-- {
		previousWidth := m.columnDisplayWidth(previousColumn) + 1
		if usedWidth+previousWidth > availableWidth {
			break
		}
		usedWidth += previousWidth
		startColumn = previousColumn
	}
	return startColumn
}

func (m *model) ensureCursorColumnVisible() {
	if m.cursorX < m.offsetX {
		m.offsetX = m.cursorX
		return
	}
	columns := m.visibleColumnLayout()
	if m.cursorX > columns[len(columns)-1].index {
		m.offsetX = m.offsetWithColumnAtRight(m.cursorX)
	}
}

func maxLineWidth(value string) int {
	width := 0
	for line := range strings.SplitSeq(value, "\n") {
		width = max(width, lipgloss.Width(line))
	}
	return width
}

func (m model) autoFitColumns(columnName string) (model, int, error) {
	rows, err := m.excelFile.GetRows(m.sheetName)
	if err != nil {
		return m, 0, err
	}

	startColumn, endColumn := 1, 0
	if columnName != "" {
		columnNumber, err := excelize.ColumnNameToNumber(strings.ToUpper(columnName))
		if err != nil {
			return m, 0, fmt.Errorf("invalid column %q: %w", columnName, err)
		}
		startColumn, endColumn = columnNumber, columnNumber
	} else {
		for _, row := range rows {
			endColumn = max(endColumn, len(row))
		}
		if endColumn == 0 {
			return m, 0, nil
		}
	}

	type widthChange struct {
		columnName string
		oldWidth   float64
		newWidth   float64
	}
	changes := make([]widthChange, 0, endColumn-startColumn+1)
	for columnNumber := startColumn; columnNumber <= endColumn; columnNumber++ {
		name, err := excelize.ColumnNumberToName(columnNumber)
		if err != nil {
			return m, 0, err
		}
		width := lipgloss.Width(name)
		for rowIndex, row := range rows {
			if columnNumber > len(row) {
				continue
			}
			value := row[columnNumber-1]
			address, err := excelize.CoordinatesToCellName(columnNumber, rowIndex+1)
			if err != nil {
				return m, 0, err
			}
			cellType, err := m.excelFile.GetCellType(m.sheetName, address)
			if err != nil {
				return m, 0, err
			}
			if cellType == excelize.CellTypeFormula {
				value, err = m.excelFile.CalcCellValue(m.sheetName, address)
				if err != nil {
					return m, 0, err
				}
			}
			width = max(width, maxLineWidth(value))
		}
		oldWidth, err := m.excelFile.GetColWidth(m.sheetName, name)
		if err != nil {
			return m, 0, err
		}
		changes = append(changes, widthChange{
			columnName: name,
			oldWidth:   oldWidth,
			newWidth:   float64(min(max(width, 1), excelize.MaxColumnWidth)),
		})
	}

	for index, change := range changes {
		if err := m.excelFile.SetColWidth(m.sheetName, change.columnName, change.columnName, change.newWidth); err != nil {
			for rollbackIndex := 0; rollbackIndex < index; rollbackIndex++ {
				rollback := changes[rollbackIndex]
				_ = m.excelFile.SetColWidth(m.sheetName, rollback.columnName, rollback.columnName, rollback.oldWidth)
			}
			return m, 0, err
		}
	}

	if len(changes) > 0 {
		m.markMetadataMutation()
	}
	return m, len(changes), nil
}
