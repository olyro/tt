package main

import (
	"strings"
)

func (m model) Search(searchTerm string, forward bool) model {
	rows, err := m.excelFile.GetRows(m.sheetName)

	if err != nil {
		return m
	}

	startIndex := 0
	endIndex := len(rows) - 1

	if forward {
		startIndex = m.cursorY
	} else {
		endIndex = m.cursorY
	}

	var index = m.cursorY

	for index >= startIndex && index <= endIndex {
		columns := rows[index]

		for i, column := range columns {
			if index == m.cursorY && ((forward && m.cursorX >= i) || (!forward && m.cursorX <= i)) {
				continue
			}

			if strings.Contains(strings.ToLower(column), strings.ToLower(searchTerm)) {
				m.cursorX = i
				m.cursorY = index
				m.offsetX = max(m.cursorX-m.GetNrOfVisibleColumns()+1, 0)
				m.offsetY = max(m.cursorY-m.GetNrOfVisibleRows()+2, 0)
				m.UpdateValuePrompt()
				return m
			}
		}

		if forward {
			index++
		} else {
			index--
		}
	}

	return m
}

type coordinate struct {
	x int
	y int
}

func (m model) SearchIterator(searchTerm string, forward bool) model {
	rows, err := m.excelFile.Rows(m.sheetName)

	if err != nil {
		return m
	}

	var index = 0

	results := make([]coordinate, 0)

Outer:
	for rows.Next() {
		if forward && index < m.cursorY {
			index++
			continue
		} else if !forward && index > m.cursorY {
			break
		}

		columns, err := rows.Columns()

		if err != nil {
			return m
		}

		for i, column := range columns {
			if index == m.cursorY && ((forward && m.cursorX >= i) || (!forward && m.cursorX <= i)) {
				continue
			}

			if strings.Contains(strings.ToLower(column), strings.ToLower(searchTerm)) {
				results = append(results, coordinate{x: i, y: index})

				if forward {
					break Outer
				}
			}
		}

		index++
	}

	if len(results) > 0 {
		if forward {
			m.cursorX = results[0].x
			m.cursorY = results[0].y
		} else {
			m.cursorX = results[len(results)-1].x
			m.cursorY = results[len(results)-1].y
		}
		m.offsetX = max(m.cursorX-m.GetNrOfVisibleColumns()+1, 0)
		m.offsetY = max(m.cursorY-m.GetNrOfVisibleRows()+2, 0)
		m.UpdateValuePrompt()
	}

	return m
}
