package main

import (
	"fmt"
	"strconv"
	"strings"

	tea "github.com/charmbracelet/bubbletea"
)

func (m model) evaluateInput(input string) (string, model, tea.Cmd) {
	inputs := strings.Split(input, " ")

	if len(inputs) < 1 {
		return "", m, nil
	}

	if inputs[0] == "b" || inputs[0] == "sheet" {
		if len(inputs) < 2 {
			return m.sheetName, m, nil
		} else {
			sheetName := inputs[1]
			index, err := m.excelFile.GetSheetIndex(sheetName)
			if err != nil {
				return "Error setting active sheet: " + err.Error(), m, nil
			} else if index == -1 {
				return "Sheet not found: " + sheetName, m, nil
			}
			m.excelFile.SetActiveSheet(index)
			m.sheetName = sheetName
			return fmt.Sprintf("Active sheet set to %s", sheetName), m, nil
		}
	}

	if inputs[0] == "bn" || inputs[0] == "nextSheet" {
		index := m.excelFile.GetActiveSheetIndex()
		if index == -1 {
			return "No active sheet found", m, nil
		}
		sheetCount := len(m.excelFile.GetSheetList())
		nextIndex := (index + 1) % sheetCount
		m.excelFile.SetActiveSheet(nextIndex)
		m.sheetName = m.excelFile.GetSheetName(nextIndex)
		return fmt.Sprintf("Active sheet set to %s", m.sheetName), m, nil
	}

	if inputs[0] == "bp" || inputs[0] == "previousSheet" {
		index := m.excelFile.GetActiveSheetIndex()
		if index == -1 {
			return "No active sheet found", m, nil
		}
		sheetCount := len(m.excelFile.GetSheetList())
		prevIndex := (index - 1 + sheetCount) % sheetCount
		m.excelFile.SetActiveSheet(prevIndex)
		m.sheetName = m.excelFile.GetSheetName(prevIndex)
		return fmt.Sprintf("Active sheet set to %s", m.sheetName), m, nil
	}

	if inputs[0] == "bd" || inputs[0] == "deleteSheet" {
		sheetName := ""
		if len(inputs) < 2 {
			sheetName = m.sheetName
		} else {
			sheetName = inputs[1]
		}

		sheetCount := len(m.excelFile.GetSheetList())

		if sheetCount <= 1 {
			return "Cannot delete the only sheet", m, nil
		}

		if err := m.excelFile.DeleteSheet(sheetName); err != nil {
			return "Error deleting sheet: " + err.Error(), m, nil
		}

		if m.sheetName == sheetName {
			m.sheetName = m.excelFile.GetSheetName(m.excelFile.GetActiveSheetIndex())
		}

		return fmt.Sprintf("Sheet %s deleted", sheetName), m, nil
	}

	if inputs[0] == "badd" || inputs[0] == "addSheet" {
		if len(inputs) < 2 {
			return "Please provide a name for the new sheet", m, nil
		}
		sheetName := inputs[1]
		if _, err := m.excelFile.NewSheet(sheetName); err != nil {
			return "Error adding sheet: " + err.Error(), m, nil
		}

		return fmt.Sprintf("Sheet %s added", sheetName), m, nil
	}

	if inputs[0] == "q" || inputs[0] == "quit" {
		return "", m, tea.Quit
	}

	if inputs[0] == "cw" || inputs[0] == "columnWidth" {
		if len(inputs) < 2 {
			return strconv.Itoa(m.columnWidth), m, nil
		}
		width, err := strconv.Atoi(inputs[1])
		if err != nil || width <= 0 {
			return "Invalid column width", m, nil
		}
		m.columnWidth = width
		return fmt.Sprintf("Column width set to %d", m.columnWidth), m, nil
	}

	if inputs[0] == "w" || inputs[0] == "write" {
		if len(inputs) < 2 {
			if err := m.excelFile.Save(); err != nil {
				return "Error saving file: " + err.Error(), m, nil
			}
		} else {
			fileName := inputs[1]
			if err := m.excelFile.SaveAs(fileName); err != nil {
				return "Error saving file: " + err.Error(), m, nil
			}
		}

		return "written", m, nil
	}

	return "Unknown Operation: " + inputs[0], m, nil
}
