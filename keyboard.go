package main

import (
	tea "github.com/charmbracelet/bubbletea"
	"github.com/xuri/excelize/v2"
)

func handleKeyboardEvent(m model, msg tea.KeyMsg) (model, tea.Cmd) {
	var cmd tea.Cmd
	switch msg.String() {
	case "0":
		if !m.useInput {
			if m.normalInput != "" {
				m.normalInput += msg.String()
			} else {
				m.cursorX = 0
				m.offsetX = 0
				m.UpdateValuePrompt()
			}
		}
	case "1", "2", "3", "4", "5", "6", "7", "8", "9":
		if !m.useInput {
			m.normalInput += msg.String()
		}
	case "$":
		if !m.useInput {
			maxCol := m.GetMaxColumn(m.cursorY)
			m.cursorX = max(maxCol-1, 0)
			m.ensureCursorColumnVisible()
			m.UpdateValuePrompt()
		}
	case ":":
		if !m.useInput {
			m.currentOp = nil
			m.input.Prompt = ":"
			m.useInput = true
			m.input.Reset()
			m.enableSheetCompletion()
			m.input.Focus()
			m.mode = Command
			return m, nil
		}
	case "/":
		if !m.useInput {
			m.currentOp = nil
			m.input.Prompt = "/"
			m.useInput = true
			m.input.Reset()
			m.disableInputCompletion()
			m.input.Focus()
			m.mode = Search
			return m, nil
		}
	case "enter":
		if m.useInput {
			m.useInput = false
			var message string
			newModel := m
			var cmd tea.Cmd
			var err error
			if m.currentOp != nil {
				op := m.currentOp
				newModel, err = executeOperation(m, op)
				newModel.currentOp = nil
				if err != nil {
					message = err.Error()
					newModel.input.SetValue(message)
				} else {
					newModel.UpdateValuePrompt()
				}
			} else if m.mode == Command {
				message, newModel, cmd = m.evaluateInput(m.input.Value())
				newModel.input.SetValue(message)
			} else if m.mode == Search {
				searchQuery := m.input.Value()
				newModel = m.SearchIterator(searchQuery, true)
				newModel.searchQuery = searchQuery
			}
			m = newModel
			m.input.Prompt = ""
			m.disableInputCompletion()
			m.resetToCellSelection()
			m.input.Blur()
			m.mode = Normal
			return m, cmd
		}
	case "up", "k":
		if !m.useInput && m.cursorY > 0 {
			prefix := getNumberPrefix(m.normalInput)

			if prefix <= 0 {
				prefix = 1
			}

			newCursor := max(m.cursorY-prefix, 0)
			if newCursor < m.offsetY {
				m.offsetY = newCursor
			}
			m.cursorY = newCursor
			m.normalInput = ""
			m.UpdateValuePrompt()
		}
	case "down", "j":
		if !m.useInput {
			prefix := getNumberPrefix(m.normalInput)

			if prefix <= 0 {
				prefix = 1
			}

			newCursor := min(m.cursorY+prefix, maxExcelRows-1)
			if newCursor > m.GetNrOfVisibleRows()-2+m.offsetY {
				m.offsetY = max(newCursor-m.GetNrOfVisibleRows()+2, 0)
			}
			m.cursorY = newCursor
			m.normalInput = ""
			m.UpdateValuePrompt()
		}
	case "left", "h", "b":
		if !m.useInput && m.cursorX > 0 {
			prefix := getNumberPrefix(m.normalInput)

			if prefix <= 0 {
				prefix = 1
			}

			newCursor := max(m.cursorX-prefix, 0)
			if newCursor < m.offsetX {
				m.offsetX = newCursor
			}
			m.cursorX = newCursor
			m.normalInput = ""
			m.UpdateValuePrompt()
		}
	case "right", "l", "w":
		if !m.useInput {
			prefix := getNumberPrefix(m.normalInput)

			if prefix <= 0 {
				prefix = 1
			}

			newCursor := min(m.cursorX+prefix, excelize.MaxColumns-1)
			m.cursorX = newCursor
			m.ensureCursorColumnVisible()
			m.normalInput = ""
			m.UpdateValuePrompt()
		}
	case "ctrl+r":
		{
			if !m.useInput && m.opStackPointer < len(m.opStack)-1 {
				m.opStackPointer++
				nextOp := m.opStack[m.opStackPointer]
				newModel, err := nextOp.operation.Do(m)
				if err != nil {
					m.input.SetValue(err.Error())
					m.opStackPointer--
				} else {
					m = newModel
					m.sheetStates[m.sheetName] = nextOp.afterStateID
					m.updateDirtyState()
					m.invalidateDisplayCache()
				}
				return m, nil
			}
		}
	case "u":
		{
			if !m.useInput && m.opStackPointer >= 0 {
				lastOp := m.opStack[m.opStackPointer]
				newModel, err := lastOp.operation.Undo(m)
				if err != nil {
					m.input.SetValue(err.Error())
				} else {
					m = newModel
					m.opStackPointer--
					m.sheetStates[m.sheetName] = lastOp.beforeStateID
					m.updateDirtyState()
					m.invalidateDisplayCache()
				}
				return m, nil
			}
		}
	case "c":
		if !m.useInput {
			m.currentOp = &changeOperation{}
			m.useInput = true
			m.input.Reset()
			m.input.Focus()
			return m, nil
		}
	case "i":
		if !m.useInput {
			m.currentOp = &changeOperation{}

			address, err := excelize.CoordinatesToCellName(m.cursorX+1, m.cursorY+1)
			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}

			cellType, err := m.excelFile.GetCellType(m.sheetName, address)

			value := ""

			if err == nil && cellType == excelize.CellTypeFormula {
				result, err := m.excelFile.GetCellFormula(m.sheetName, address)
				if err == nil {
					value = prefixWithEqual(result)
				}
			} else {
				val, err := m.excelFile.GetCellValue(m.sheetName, address)
				if err == nil {
					value = val
				}
			}

			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}

			m.useInput = true
			m.input.SetValue(value)
			m.input.Focus()
			m.input.SetCursor(0)
			return m, nil
		}
	case "a":
		if !m.useInput {
			m.currentOp = &changeOperation{}
			address, err := excelize.CoordinatesToCellName(m.cursorX+1, m.cursorY+1)
			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}
			cellType, err := m.excelFile.GetCellType(m.sheetName, address)

			value := ""

			if err == nil && cellType == excelize.CellTypeFormula {
				result, err := m.excelFile.GetCellFormula(m.sheetName, address)
				if err == nil {
					value = prefixWithEqual(result)
				}
			} else {
				val, err := m.excelFile.GetCellValue(m.sheetName, address)
				if err == nil {
					value = val
				}
			}

			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}

			m.useInput = true
			m.input.Focus()
			m.input.SetValue(value)
			m.input.SetCursor(len(value))
			return m, nil
		}
	case "I":
		if !m.useInput {
			prefix := getNumberPrefix(m.normalInput)
			amount := 1

			if prefix > 0 {
				amount = prefix
			}

			op := &columnInsertOperation{colIndex: max(m.cursorX+1, 0), amount: amount}
			newModel, err := executeOperation(m, op)

			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}

			newModel.normalInput = ""
			m = newModel
		}
	case "A":
		if !m.useInput {
			prefix := getNumberPrefix(m.normalInput)
			amount := 1

			if prefix > 0 {
				amount = prefix
			}

			op := &columnInsertOperation{colIndex: m.cursorX + 2, amount: amount}
			newModel, err := executeOperation(m, op)

			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}

			newModel.normalInput = ""
			m = newModel
		}
	case "O":
		if !m.useInput {
			prefix := getNumberPrefix(m.normalInput)
			amount := 1

			if prefix > 0 {
				amount = prefix
			}

			op := &rowInsertOperation{rowIndex: max(m.cursorY+1, 0), amount: amount}
			newModel, err := executeOperation(m, op)

			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}

			newModel.normalInput = ""
			m = newModel
		}
	case "o":
		if !m.useInput {
			prefix := getNumberPrefix(m.normalInput)
			amount := 1

			if prefix > 0 {
				amount = prefix
			}

			op := &rowInsertOperation{rowIndex: m.cursorY + 2, amount: amount}
			newModel, err := executeOperation(m, op)

			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}

			newModel.normalInput = ""
			m = newModel
		}
	case "v":
		if !m.useInput {
			m.selection.kind = ColumnsSelect
			m.selection.columns = &rowColumnSelect{
				start: m.cursorX,
				end:   m.cursorX,
			}
			m.selection.cell = nil
			m.selection.rows = nil
			m.selection.block = nil
		}
	case "V":
		if !m.useInput {
			m.selection.kind = RowsSelect
			m.selection.rows = &rowColumnSelect{
				start: m.cursorY,
				end:   m.cursorY,
			}
			m.selection.cell = nil
			m.selection.columns = nil
			m.selection.block = nil
		}
	case "ctrl+v":
		if !m.useInput {
			m.selection.kind = BlockSelect
			m.selection.block = &blockSelect{
				startX: m.cursorX,
				startY: m.cursorY,
				endX:   m.cursorX,
				endY:   m.cursorY,
			}
			m.selection.cell = nil
			m.selection.columns = nil
			m.selection.rows = nil
		}
	case "esc":
		if m.useInput {
			m.useInput = false
			m.currentOp = nil
			m.input.Blur()
			m.input.Prompt = ""
			m.disableInputCompletion()
		} else {
			m.normalInput = ""
			m.resetToCellSelection()
		}
		m.UpdateValuePrompt()
		m.mode = Normal
	case "g":
		if !m.useInput {
			// get last char of m.normalInput
			if len(m.normalInput) > 0 {
				lastChar := m.normalInput[len(m.normalInput)-1:]
				prefix := getNumberPrefix(m.normalInput)
				if lastChar == "g" {
					if prefix > 0 {
						m.cursorY = min(prefix-1, maxExcelRows-1)
						m.offsetY = max(m.cursorY-m.GetNrOfVisibleRows()+2, 0)
					} else {
						m.offsetY = 0
						m.cursorY = 0
					}
					m.UpdateValuePrompt()
					m.normalInput = ""
				} else if prefix > 0 {
					m.normalInput += "g"
				}
			} else {
				m.normalInput += "g"
			}
		}
	case "G":
		if !m.useInput {
			prefix := getNumberPrefix(m.normalInput)
			if prefix > 0 {
				m.cursorY = min(prefix-1, maxExcelRows-1)
				m.offsetY = max(m.cursorY-m.GetNrOfVisibleRows()+2, 0)
			} else {
				rows, err := m.excelFile.Rows(m.sheetName)
				if err == nil {
					defer rows.Close()
					lastRow := 0
					for rows.Next() {
						lastRow++
					}
					m.offsetY = max(lastRow-m.GetNrOfVisibleRows()+1, 0)
					m.cursorY = max(lastRow-1, 0)
				}
			}
			m.normalInput = ""
			m.UpdateValuePrompt()
		}
	case "N":
		if !m.useInput && m.searchQuery != "" {
			m = m.SearchIterator(m.searchQuery, false)
		}
	case "n":
		if !m.useInput && m.searchQuery != "" {
			m = m.SearchIterator(m.searchQuery, true)
		}
	case "M":
		if !m.useInput {
			op := &unmergeOperation{}
			newModel, err := executeOperation(m, op)
			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}
			newModel.resetToCellSelection()
			m = newModel
		}
	case "m":
		if !m.useInput {
			op := &mergeOperation{}
			newModel, err := executeOperation(m, op)
			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}
			newModel.resetToCellSelection()
			m = newModel
		}
	case "y":
		if !m.useInput {
			if err := m.makeCopy(); err != nil {
				m.input.SetValue(err.Error())
			}
		}
	case "Y":
		if !m.useInput {
			if m.clipboard == nil {
				m.clipboard = systemClipboard{}
			}
			value, err := m.exportSelectionTSV()
			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}
			return m, writeClipboardCmd(m.clipboard, value)
		}
	case "P":
		if !m.useInput {
			if m.clipboard == nil {
				m.clipboard = systemClipboard{}
			}
			return m, readClipboardCmd(m.clipboard)
		}
	case "p":
		if !m.useInput && m.copy != nil {
			op := &pasteOperation{}
			newModel, err := executeOperation(m, op)

			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}

			newModel.normalInput = ""
			m = newModel
		}
	case "d":
		if !m.useInput {
			op := &deleteOperation{}
			newModel, err := executeOperation(m, op)
			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}
			newModel.resetToCellSelection()
			m = newModel
		}
	case "x":
		if !m.useInput {
			op := &clearOperation{}
			newModel, err := executeOperation(m, op)
			if err != nil {
				m.input.SetValue(err.Error())
				return m, nil
			}
			newModel.resetToCellSelection()
			m = newModel
		}
	case "ctrl+d":
		if !m.useInput {
			amount := max(m.GetNrOfVisibleRows()-1, 1)
			m.cursorY = min(m.cursorY+amount, maxExcelRows-1)
			m.offsetY = max(m.cursorY-m.GetNrOfVisibleRows()+2, 0)
		}
	case "ctrl+u":
		if !m.useInput {
			amount := max(m.GetNrOfVisibleRows()-1, 1)
			m.cursorY = max(m.cursorY-amount, 0)
			m.offsetY = min(m.offsetY, m.cursorY)
		}
	case "ctrl+c":
		if m.dirty {
			m.input.SetValue("Workbook has unsaved changes, use :q! to force quit")
			return m, nil
		}
		return m, tea.Quit
	}

	if m.useInput {
		m.normalInput = ""
		m.input, cmd = m.input.Update(msg)
	}

	return m, cmd
}
