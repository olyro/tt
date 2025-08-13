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
			m.offsetX = max(maxCol-m.GetNrOfVisibleColumns(), 0)
			m.UpdateValuePrompt()
		}
	case ":":
		if !m.useInput {
			m.input.Prompt = ":"
			m.useInput = true
			m.input.Reset()
			m.input.Focus()
			m.mode = Command
			return m, nil
		}
	case "/":
		if !m.useInput {
			m.input.Prompt = "/"
			m.useInput = true
			m.input.Reset()
			m.input.Focus()
			m.mode = Search
			return m, nil
		}
	case "enter":
		if m.useInput {
			m.useInput = false
			var message string
			var newModel model
			var cmd tea.Cmd
			if m.currentOp != nil {
				err := m.currentOp.Init(m)
				newM, err := m.currentOp.Do(m)
				newModel = newM
				newModel.currentOp = nil
				if err != nil {
					message = err.Error()
					newModel.input.SetValue(message)
				} else {
					newModel.pushOp(m.currentOp)
					m.UpdateValuePrompt()
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

			if m.cursorY-prefix < m.offsetY {
				m.offsetY -= prefix
			}
			m.cursorY -= prefix
			m.normalInput = ""
			m.UpdateValuePrompt()
		}
	case "down", "j":
		if !m.useInput {
			prefix := getNumberPrefix(m.normalInput)

			if prefix <= 0 {
				prefix = 1
			}

			if m.cursorY+prefix > m.GetNrOfVisibleRows()-2+m.offsetY {
				m.offsetY += prefix
			}
			m.cursorY += prefix
			m.normalInput = ""
			m.UpdateValuePrompt()
		}
	case "left", "h", "b":
		if !m.useInput && m.cursorX > 0 {
			prefix := getNumberPrefix(m.normalInput)

			if prefix <= 0 {
				prefix = 1
			}

			if m.cursorX-prefix < m.offsetX {
				m.offsetX -= prefix
			}
			m.cursorX -= prefix
			m.normalInput = ""
			m.UpdateValuePrompt()
		}
	case "right", "l", "w":
		if !m.useInput {
			prefix := getNumberPrefix(m.normalInput)

			if prefix <= 0 {
				prefix = 1
			}

			if m.cursorX+prefix > m.GetNrOfVisibleColumns()+m.offsetX-1 {
				m.offsetX += prefix
			}
			m.cursorX += prefix
			m.normalInput = ""
			m.UpdateValuePrompt()
		}
	case "ctrl+r":
		{
			if !m.useInput && m.opStackPointer < len(m.opStack)-1 {
				m.opStackPointer++
				nextOp := m.opStack[m.opStackPointer]
				m, err := nextOp.Do(m)
				if err != nil {
					m.input.SetValue(err.Error())
				}
				return m, nil
			}
		}
	case "u":
		{
			if !m.useInput && m.opStackPointer >= 0 {
				lastOp := m.opStack[m.opStackPointer]
				m, err := lastOp.Undo(m)
				m.opStackPointer--
				if err != nil {
					m.input.SetValue(err.Error())
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
			op.Init(m)
			newModel, err := op.Do(m)

			if err != nil {
				newModel.input.SetValue(err.Error())
			}

			newModel.normalInput = ""
			newModel.pushOp(op)
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
			op.Init(m)
			newModel, err := op.Do(m)

			if err != nil {
				newModel.input.SetValue(err.Error())
			}

			newModel.normalInput = ""
			newModel.pushOp(op)
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
			op.Init(m)
			newModel, err := op.Do(m)

			if err != nil {
				newModel.input.SetValue(err.Error())
			}

			newModel.normalInput = ""
			newModel.pushOp(op)
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
			op.Init(m)
			newModel, err := op.Do(m)

			if err != nil {
				newModel.input.SetValue(err.Error())
			}

			newModel.normalInput = ""
			newModel.pushOp(op)
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
			m.input.Blur()
			m.input.Prompt = ""
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
						m.cursorY = prefix - 1
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
				m.cursorY = prefix - 1
				m.offsetY = max(m.cursorY-m.GetNrOfVisibleRows()+2, 0)
			} else {
				rows, err := m.excelFile.Rows(m.sheetName)
				if err == nil {
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
			op.Init(m)
			newModel, err := op.Do(m)
			if err != nil {
				newModel.input.SetValue(err.Error())
			} else {
				newModel.pushOp(op)
			}
			newModel.resetToCellSelection()
			m = newModel
		}
	case "m":
		if !m.useInput {
			op := &mergeOperation{}
			op.Init(m)
			newModel, err := op.Do(m)
			if err != nil {
				newModel.input.SetValue(err.Error())
			} else {
				newModel.pushOp(op)
			}
			newModel.resetToCellSelection()
			m = newModel
		}
	case "y":
		if !m.useInput {
			m.makeCopy()
		}
	case "p":
		if !m.useInput && m.copy != nil {
			op := &pasteOperation{}
			op.Init(m)
			newModel, err := op.Do(m)

			if err != nil {
				newModel.input.SetValue(err.Error())
			}

			newModel.normalInput = ""
			newModel.pushOp(op)
			m = newModel
		}
	case "d":
		if !m.useInput {
			op := &deleteOperation{}
			op.Init(m)
			newModel, err := op.Do(m)
			if err != nil {
				newModel.input.SetValue(err.Error())
			} else {
				newModel.pushOp(op)
			}
			newModel.resetToCellSelection()
			m = newModel
		}
	case "x":
		if !m.useInput {
			op := &clearOperation{}
			op.Init(m)
			newModel, err := op.Do(m)
			if err != nil {
				newModel.input.SetValue(err.Error())
			} else {
				newModel.pushOp(op)
			}
			newModel.resetToCellSelection()
			m = newModel
		}
	case "ctrl+d":
		if !m.useInput {
			m.offsetY += m.GetNrOfVisibleRows() - 1
			m.cursorY += m.GetNrOfVisibleRows() - 1
		}
	case "ctrl+u":
		if !m.useInput {
			m.offsetY = max(m.offsetY-(m.GetNrOfVisibleRows()-1), 0)
			m.cursorY = max(m.cursorY-(m.GetNrOfVisibleRows()-1), 0)
		}
	case "ctrl+c":
		return m, tea.Quit
	}

	if m.useInput {
		m.normalInput = ""
		m.input, cmd = m.input.Update(msg)
	}

	return m, cmd
}
