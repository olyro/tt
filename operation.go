package main

import "github.com/xuri/excelize/v2"

type operation interface {
	Init(model model) error
	Do(m model) (model, error)
	Undo(m model) (model, error)
}

type rowInsertOperation struct {
	rowIndex  int
	amount    int
	sheetName string
}

func (r *rowInsertOperation) Init(m model) error {
	r.sheetName = m.sheetName
	return nil

}

func (r *rowInsertOperation) Do(m model) (model, error) {
	if err := m.excelFile.InsertRows(r.sheetName, r.rowIndex, r.amount); err != nil {
		return m, err
	}

	m.cursorY = r.rowIndex + r.amount - 2

	return m, nil
}

func (r *rowInsertOperation) Undo(m model) (model, error) {
	for i := 0; i < r.amount; i++ {
		if err := m.excelFile.RemoveRow(r.sheetName, r.rowIndex); err != nil {
			return m, err
		}
	}

	return m, nil
}

type columnInsertOperation struct {
	colIndex  int
	amount    int
	sheetName string
}

func (r *columnInsertOperation) Init(m model) error {
	r.sheetName = m.sheetName
	return nil
}

func (r *columnInsertOperation) Do(m model) (model, error) {
	colName, err := excelize.ColumnNumberToName(r.colIndex)

	if err != nil {
		return m, err
	}

	if err := m.excelFile.InsertCols(r.sheetName, colName, r.amount); err != nil {
		return m, err
	}

	m.cursorX = r.colIndex + r.amount - 2

	return m, nil
}

func (r *columnInsertOperation) Undo(m model) (model, error) {
	colName, err := excelize.ColumnNumberToName(r.colIndex)

	if err != nil {
		return m, err
	}

	for i := 0; i < r.amount; i++ {
		if err := m.excelFile.RemoveCol(r.sheetName, colName); err != nil {
			return m, err
		}
	}

	return m, nil
}

type changeOperation struct {
	oldValues map[string]string
	newValues map[string]string
	sheetName string
}

func (c *changeOperation) Init(m model) error {
	c.oldValues = make(map[string]string)
	c.newValues = make(map[string]string)
	c.sheetName = m.sheetName
	for _, address := range m.getSelectedCellAddresses() {
		oldValue, err := m.excelFile.GetCellValue(c.sheetName, address, excelize.Options{
			RawCellValue: true,
		})
		if err != nil {
			return err
		}
		newValue := m.input.Value()
		c.oldValues[address] = oldValue
		c.newValues[address] = newValue
	}

	return nil
}

func (c *changeOperation) Do(m model) (model, error) {
	for address, newValue := range c.newValues {
		if err := m.setCellValue(c.sheetName, address, newValue); err != nil {
			return m, err
		}
	}

	return m, nil
}

func (c *changeOperation) Undo(m model) (model, error) {
	for address, oldValue := range c.oldValues {
		if err := m.setCellValue(c.sheetName, address, oldValue); err != nil {
			return m, err
		}
	}

	return m, nil
}

type pasteOperation struct {
	copy      copy
	oldValues map[string]string
	oldStyles map[string]int
	oldTypes  map[string]excelize.CellType
	cursorX   int
	cursorY   int
	sheetName string
}

func (c *pasteOperation) Init(m model) error {
	c.copy = *m.copy
	c.oldValues = make(map[string]string)
	c.oldStyles = make(map[string]int)
	c.oldTypes = make(map[string]excelize.CellType)
	c.cursorX = m.cursorX
	c.cursorY = m.cursorY
	c.sheetName = m.sheetName
	return nil
}

func (c *pasteOperation) Do(m model) (model, error) {
	if c.copy.selection.rows != nil {
		height := max(c.copy.selection.rows.start, c.copy.selection.rows.end) - min(c.copy.selection.rows.start, c.copy.selection.rows.end) + 1
		m.excelFile.InsertRows(c.sheetName, c.cursorY+2, height)
	} else if c.copy.selection.columns != nil {
		width := max(c.copy.selection.columns.start, c.copy.selection.columns.end) - min(c.copy.selection.columns.start, c.copy.selection.columns.end) + 1
		colName, err := excelize.ColumnNumberToName(c.cursorX + 2)

		if err != nil {
			return m, err
		}

		m.excelFile.InsertCols(c.sheetName, colName, width)
	}

	for address, newValue := range c.copy.values {
		startX := 0
		startY := 0

		offsetX := 0
		offsetY := 0

		if c.copy.selection.block != nil {
			startX = c.copy.selection.block.startX
			startY = c.copy.selection.block.startY
			offsetX = c.cursorX
			offsetY = c.cursorY
		} else if c.copy.selection.cell != nil {
			startX = c.copy.selection.cell.x
			startY = c.copy.selection.cell.y
			offsetX = c.cursorX
			offsetY = c.cursorY
		} else if c.copy.selection.rows != nil {
			startY = c.copy.selection.rows.start
			offsetY = c.cursorY + 1
		} else if c.copy.selection.columns != nil {
			startX = c.copy.selection.columns.start
			offsetX = c.cursorX + 1
		}

		x, y, err := excelize.CellNameToCoordinates(address)

		if err != nil {
			return m, err
		}

		newAddress, err := excelize.CoordinatesToCellName(x-startX+offsetX, y-startY+offsetY)

		if err != nil {
			return m, err
		}

		c.oldValues[newAddress], err = m.excelFile.GetCellValue(c.sheetName, newAddress, excelize.Options{
			RawCellValue: true,
		})

		if err != nil {
			return m, err
		}

		c.oldStyles[newAddress], err = m.excelFile.GetCellStyle(c.sheetName, newAddress)

		if err != nil {
			return m, err
		}

		c.oldTypes[newAddress], err = m.excelFile.GetCellType(c.sheetName, newAddress)
		if err != nil {
			return m, err
		}

		if err := m.setCellValue(c.sheetName, newAddress, newValue); err != nil {
			return m, err
		}

		// check if the style exists in the copy
		if style, ok := c.copy.styles[address]; ok {
			if err := m.excelFile.SetCellStyle(c.sheetName, newAddress, newAddress, style); err != nil {
				return m, err
			}
		}
	}

	return m, nil
}

func (c *pasteOperation) Undo(m model) (model, error) {
	switch c.copy.selection.kind {
	case BlockSelect, CellSelect:
		for address, oldValue := range c.oldValues {
			if err := m.setCellValue(c.sheetName, address, oldValue); err != nil {
				return m, err
			}

			if err := m.excelFile.SetCellStyle(c.sheetName, address, address, c.oldStyles[address]); err != nil {
				return m, err
			}
		}
	case RowsSelect:
		height := max(c.copy.selection.rows.start, c.copy.selection.rows.end) - min(c.copy.selection.rows.start, c.copy.selection.rows.end) + 1
		for range height {
			if err := m.excelFile.RemoveRow(m.sheetName, c.cursorY+2); err != nil {
				return m, err
			}
		}
	case ColumnsSelect:
		width := max(c.copy.selection.columns.start, c.copy.selection.columns.end) - min(c.copy.selection.columns.start, c.copy.selection.columns.end) + 1
		for range width {
			colName, err := excelize.ColumnNumberToName(c.cursorX + 2)

			if err != nil {
				return m, err
			}

			if err := m.excelFile.RemoveCol(m.sheetName, colName); err != nil {
				return m, err
			}
		}
	}

	return m, nil
}

type mergeOperation struct {
	blockSelect blockSelect
	oldValues   map[string]string
	sheetName   string
}

func (m *mergeOperation) Init(model model) error {
	m.sheetName = model.sheetName
	m.oldValues = make(map[string]string)

	if model.selection.block == nil {
		return nil
	}

	m.blockSelect = *model.selection.block
	addresses := model.getSelectedCellAddresses()

	for _, address := range addresses {
		value, err := model.excelFile.GetCellValue(m.sheetName, address, excelize.Options{
			RawCellValue: true,
		})
		if err != nil {
			return err
		}
		m.oldValues[address] = value
	}

	return nil
}

func (m *mergeOperation) Do(model model) (model, error) {
	startX := min(m.blockSelect.startX, m.blockSelect.endX)
	startY := min(m.blockSelect.startY, m.blockSelect.endY)
	endX := max(m.blockSelect.startX, m.blockSelect.endX)
	endY := max(m.blockSelect.startY, m.blockSelect.endY)

	startAddress, err := excelize.CoordinatesToCellName(startX+1, startY+1)
	if err != nil {
		return model, err
	}

	endAddress, err := excelize.CoordinatesToCellName(endX+1, endY+1)

	if err := model.excelFile.MergeCell(m.sheetName, startAddress, endAddress); err != nil {
		return model, err
	}

	return model, nil
}

func (m *mergeOperation) Undo(model model) (model, error) {
	startX := min(m.blockSelect.startX, m.blockSelect.endX)
	startY := min(m.blockSelect.startY, m.blockSelect.endY)
	endX := max(m.blockSelect.startX, m.blockSelect.endX)
	endY := max(m.blockSelect.startY, m.blockSelect.endY)

	startAddress, err := excelize.CoordinatesToCellName(startX+1, startY+1)
	if err != nil {
		return model, err
	}

	endAddress, err := excelize.CoordinatesToCellName(endX+1, endY+1)

	if err := model.excelFile.UnmergeCell(model.sheetName, startAddress, endAddress); err != nil {
		return model, err
	}

	for address, oldValue := range m.oldValues {
		if err := model.setCellValue(m.sheetName, address, oldValue); err != nil {
			return model, err
		}
	}

	return model, nil
}

type unmergeOperation struct {
	blockSelect blockSelect
	sheetName   string
}

func (m *unmergeOperation) Init(model model) error {
	m.sheetName = model.sheetName

	if model.selection.block == nil {
		return nil
	}

	m.blockSelect = *model.selection.block

	return nil
}

func (m *unmergeOperation) Do(model model) (model, error) {
	startX := min(m.blockSelect.startX, m.blockSelect.endX)
	startY := min(m.blockSelect.startY, m.blockSelect.endY)
	endX := max(m.blockSelect.startX, m.blockSelect.endX)
	endY := max(m.blockSelect.startY, m.blockSelect.endY)

	startAddress, err := excelize.CoordinatesToCellName(startX+1, startY+1)
	if err != nil {
		return model, err
	}

	endAddress, err := excelize.CoordinatesToCellName(endX+1, endY+1)

	if err := model.excelFile.UnmergeCell(m.sheetName, startAddress, endAddress); err != nil {
		return model, err
	}

	return model, nil
}

func (m *unmergeOperation) Undo(model model) (model, error) {
	startX := min(m.blockSelect.startX, m.blockSelect.endX)
	startY := min(m.blockSelect.startY, m.blockSelect.endY)
	endX := max(m.blockSelect.startX, m.blockSelect.endX)
	endY := max(m.blockSelect.startY, m.blockSelect.endY)

	startAddress, err := excelize.CoordinatesToCellName(startX+1, startY+1)
	if err != nil {
		return model, err
	}

	endAddress, err := excelize.CoordinatesToCellName(endX+1, endY+1)

	if err := model.excelFile.MergeCell(model.sheetName, startAddress, endAddress); err != nil {
		return model, err
	}

	return model, nil
}

type clearOperation struct {
	oldValues map[string]string
	sheetName string
}

func (c *clearOperation) Init(m model) error {
	c.sheetName = m.sheetName
	c.oldValues = make(map[string]string)

	for _, address := range m.getSelectedCellAddresses() {
		value, err := m.excelFile.GetCellValue(c.sheetName, address, excelize.Options{RawCellValue: true})
		if err != nil {
			return err
		}
		c.oldValues[address] = value
	}

	return nil
}

func (c *clearOperation) Do(m model) (model, error) {
	for address := range c.oldValues {
		if err := m.excelFile.SetCellValue(c.sheetName, address, nil); err != nil {
			return m, err
		}
	}

	return m, nil
}

func (c *clearOperation) Undo(m model) (model, error) {
	for address, oldValue := range c.oldValues {
		if err := m.setCellValue(c.sheetName, address, oldValue); err != nil {
			return m, err
		}
	}

	return m, nil
}
