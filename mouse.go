package main

import (
	tea "github.com/charmbracelet/bubbletea"
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
				m.cursorY++
				m.UpdateValuePrompt()
			}
		case tea.MouseButtonWheelLeft:
			if !m.useInput {
				if m.cursorX == m.GetNrOfVisibleColumns()+m.offsetX-1 {
					m.offsetX++
				}
				m.cursorX++
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
				m.cursorY = max(msg.Y/2+m.offsetY-1, 0)
				m.cursorX = max((msg.X-4)/(m.columnWidth+1)+m.offsetX, 0)
				m.UpdateValuePrompt()
			}
		}
	}

	return m
}
