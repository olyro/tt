package main

import "github.com/xuri/excelize/v2"

const (
	CellSelect selectionType = iota
	BlockSelect
	RowsSelect
	ColumnsSelect
)

type selectionType int

type blockSelect struct {
	startX int
	startY int
	endX   int
	endY   int
}

type rowColumnSelect struct {
	start int
	end   int
}

type cellSelect struct {
	x int
	y int
}

type selection struct {
	kind    selectionType
	block   *blockSelect
	cell    *cellSelect
	rows    *rowColumnSelect
	columns *rowColumnSelect
}

func (m model) getSelectedCellAddresses() []string {
	adresses := make([]string, 0)
	switch m.selection.kind {
	case CellSelect:
		address, err := excelize.CoordinatesToCellName(m.selection.cell.x+1, m.selection.cell.y+1)
		if err != nil {
			return nil
		}
		adresses = append(adresses, address)
	case BlockSelect:
		for x := min(m.selection.block.startX, m.selection.block.endX); x <= max(m.selection.block.endX, m.selection.block.startX); x++ {
			for y := min(m.selection.block.startY, m.selection.block.endY); y <= max(m.selection.block.endY, m.selection.block.startY); y++ {
				address, err := excelize.CoordinatesToCellName(x+1, y+1)
				if err != nil {
					return nil
				}
				adresses = append(adresses, address)
			}
		}
	case RowsSelect:
		rows, err := m.excelFile.Rows(m.sheetName)

		if err != nil {
			return nil
		}

		index := 0

		for rows.Next() {
			columns, err := rows.Columns()

			if index >= min(m.selection.rows.start, m.selection.rows.end) && index <= max(m.selection.rows.start, m.selection.rows.end) {
				if err != nil {
					return nil
				}

				for i := 0; i < len(columns); i++ {
					address, err := excelize.CoordinatesToCellName(i+1, index+1)

					if err != nil {
						return nil
					}

					adresses = append(adresses, address)
				}
			}

			index++
		}
	case ColumnsSelect:
		rows, err := m.excelFile.Rows(m.sheetName)

		if err != nil {
			return nil
		}

		index := 0

		for rows.Next() {
			for i := min(m.selection.columns.start, m.selection.columns.end); i <= max(m.selection.columns.start, m.selection.columns.end); i++ {
				address, err := excelize.CoordinatesToCellName(i+1, index+1)

				if err != nil {
					return nil
				}

				adresses = append(adresses, address)
			}

			index++
		}
	}

	return adresses
}

// function to reset the selection to a cell
func (m *model) resetToCellSelection() {
	m.selection.kind = CellSelect
	m.selection.cell = &cellSelect{
		x: m.cursorX,
		y: m.cursorY,
	}
	m.selection.block = nil
}
