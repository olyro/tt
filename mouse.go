package main

import (
	tea "github.com/charmbracelet/bubbletea"
	"github.com/xuri/excelize/v2"
)

func handleMouseEvent(m model, msg tea.MouseMsg) model {
	if msg.Action == tea.MouseActionPress {
		switch msg.Button {
		case tea.MouseButtonWheelUp:
			if !m.useInput && m.cursorY > 0 {
				if m.cursorY == m.offsetY {
					m.offsetY--
				}
				m.cursorY--
				m.UpdateValuePrompt()
			}
		case tea.MouseButtonWheelDown:
			if !m.useInput {
				if m.cursorY == m.GetNrOfVisibleRows()-2+m.offsetY {
					m.offsetY++
				}
				m.cursorY = min(m.cursorY+1, maxExcelRows-1)
				m.offsetY = min(m.offsetY, m.cursorY)
				m.UpdateValuePrompt()
			}
		case tea.MouseButtonWheelLeft:
			if !m.useInput {
				columns := m.visibleColumnLayout()
				if m.cursorX == columns[len(columns)-1].index {
					m.offsetX++
				}
				m.cursorX = min(m.cursorX+1, excelize.MaxColumns-1)
				m.offsetX = min(m.offsetX, m.cursorX)
				m.ensureCursorColumnVisible()
				m.UpdateValuePrompt()
			}
		case tea.MouseButtonWheelRight:
			if !m.useInput && m.cursorX > 0 {
				if m.cursorX == m.offsetX {
					m.offsetX--
				}
				m.cursorX--
				m.UpdateValuePrompt()
			}
		case tea.MouseButtonLeft:
			if !m.useInput {
				m.cursorY = min(max(msg.Y/2+m.offsetY-1, 0), maxExcelRows-1)
				m.cursorX = min(m.columnAtScreenX(msg.X), excelize.MaxColumns-1)
				m.UpdateValuePrompt()
			}
		}
	}

	return m
}
