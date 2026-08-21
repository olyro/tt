package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

type operation interface {
	Init(model model) error
	Do(m model) (model, error)
	Undo(m model) (model, error)
}

type cellSnapshot struct {
	rawValue string
	formula  string
	cellType excelize.CellType
	style    int
}

func captureCell(file *excelize.File, sheetName, address string) (cellSnapshot, error) {
	rawValue, err := file.GetCellValue(sheetName, address, excelize.Options{RawCellValue: true})
	if err != nil {
		return cellSnapshot{}, err
	}
	cellType, err := file.GetCellType(sheetName, address)
	if err != nil {
		return cellSnapshot{}, err
	}
	style, err := file.GetCellStyle(sheetName, address)
	if err != nil {
		return cellSnapshot{}, err
	}
	formula := ""
	if cellType == excelize.CellTypeFormula {
		formula, err = file.GetCellFormula(sheetName, address)
		if err != nil {
			return cellSnapshot{}, err
		}
	}
	return cellSnapshot{rawValue: rawValue, formula: formula, cellType: cellType, style: style}, nil
}

func restoreCell(file *excelize.File, sheetName, address string, snapshot cellSnapshot) error {
	var err error
	switch snapshot.cellType {
	case excelize.CellTypeFormula:
		err = file.SetCellFormula(sheetName, address, snapshot.formula)
	case excelize.CellTypeBool:
		err = file.SetCellBool(sheetName, address, snapshot.rawValue == "1" || snapshot.rawValue == "TRUE")
	case excelize.CellTypeInlineString, excelize.CellTypeSharedString:
		err = file.SetCellStr(sheetName, address, snapshot.rawValue)
	default:
		err = file.SetCellDefault(sheetName, address, snapshot.rawValue)
	}
	if err != nil {
		return err
	}
	return file.SetCellStyle(sheetName, address, address, snapshot.style)
}

func executeOperation(m model, op operation) (model, error) {
	if err := op.Init(m); err != nil {
		return m, err
	}
	newModel, err := op.Do(m)
	if err != nil {
		return m, err
	}
	newModel.pushOp(op)
	newModel.invalidateDisplayCache()
	return newModel, nil
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
	oldCells  map[string]cellSnapshot
	newValues map[string]string
	sheetName string
}

func (c *changeOperation) Init(m model) error {
	c.oldCells = make(map[string]cellSnapshot)
	c.newValues = make(map[string]string)
	c.sheetName = m.sheetName
	for _, address := range m.getSelectedCellAddresses() {
		oldValue, err := captureCell(m.excelFile, c.sheetName, address)
		if err != nil {
			return err
		}
		newValue := m.input.Value()
		c.oldCells[address] = oldValue
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
	for address, oldValue := range c.oldCells {
		if err := restoreCell(m.excelFile, c.sheetName, address, oldValue); err != nil {
			return m, err
		}
	}

	return m, nil
}

type pasteOperation struct {
	copy      copy
	oldCells  map[string]cellSnapshot
	cursorX   int
	cursorY   int
	sheetName string
}

func (c *pasteOperation) Init(m model) error {
	c.copy = *m.copy
	c.oldCells = make(map[string]cellSnapshot)
	c.cursorX = m.cursorX
	c.cursorY = m.cursorY
	c.sheetName = m.sheetName
	return nil
}

func (c *pasteOperation) Do(m model) (model, error) {
	switch c.copy.selection.kind {
	case RowsSelect:
		height := max(c.copy.selection.rows.start, c.copy.selection.rows.end) - min(c.copy.selection.rows.start, c.copy.selection.rows.end) + 1
		if err := m.excelFile.InsertRows(c.sheetName, c.cursorY+2, height); err != nil {
			return m, err
		}
	case ColumnsSelect:
		width := max(c.copy.selection.columns.start, c.copy.selection.columns.end) - min(c.copy.selection.columns.start, c.copy.selection.columns.end) + 1
		colName, err := excelize.ColumnNumberToName(c.cursorX + 2)

		if err != nil {
			return m, err
		}

		if err := m.excelFile.InsertCols(c.sheetName, colName, width); err != nil {
			return m, err
		}
	}

	for address, newCell := range c.copy.cells {
		startX := 0
		startY := 0

		offsetX := 0
		offsetY := 0

		switch c.copy.selection.kind {
		case BlockSelect:
			startX = min(c.copy.selection.block.startX, c.copy.selection.block.endX)
			startY = min(c.copy.selection.block.startY, c.copy.selection.block.endY)
			offsetX = c.cursorX
			offsetY = c.cursorY
		case CellSelect:
			startX = c.copy.selection.cell.x
			startY = c.copy.selection.cell.y
			offsetX = c.cursorX
			offsetY = c.cursorY
		case RowsSelect:
			startY = min(c.copy.selection.rows.start, c.copy.selection.rows.end)
			offsetY = c.cursorY + 1
		case ColumnsSelect:
			startX = min(c.copy.selection.columns.start, c.copy.selection.columns.end)
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

		c.oldCells[newAddress], err = captureCell(m.excelFile, c.sheetName, newAddress)
		if err != nil {
			return m, err
		}

		if err := restoreCell(m.excelFile, c.sheetName, newAddress, newCell); err != nil {
			return m, err
		}
	}

	return m, nil
}

func (c *pasteOperation) Undo(m model) (model, error) {
	switch c.copy.selection.kind {
	case BlockSelect, CellSelect:
		for address, oldValue := range c.oldCells {
			if err := restoreCell(m.excelFile, c.sheetName, address, oldValue); err != nil {
				return m, err
			}
		}
	case RowsSelect:
		height := max(c.copy.selection.rows.start, c.copy.selection.rows.end) - min(c.copy.selection.rows.start, c.copy.selection.rows.end) + 1
		for range height {
			if err := m.excelFile.RemoveRow(c.sheetName, c.cursorY+2); err != nil {
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

			if err := m.excelFile.RemoveCol(c.sheetName, colName); err != nil {
				return m, err
			}
		}
	}

	return m, nil
}

type mergeOperation struct {
	blockSelect blockSelect
	oldCells    map[string]cellSnapshot
	sheetName   string
}

func (m *mergeOperation) Init(model model) error {
	m.sheetName = model.sheetName
	m.oldCells = make(map[string]cellSnapshot)

	if model.selection.block == nil {
		return fmt.Errorf("merge requires a block selection")
	}

	m.blockSelect = *model.selection.block
	addresses := model.getSelectedCellAddresses()

	for _, address := range addresses {
		value, err := captureCell(model.excelFile, m.sheetName, address)
		if err != nil {
			return err
		}
		m.oldCells[address] = value
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

	if err := model.excelFile.UnmergeCell(m.sheetName, startAddress, endAddress); err != nil {
		return model, err
	}

	for address, oldValue := range m.oldCells {
		if err := restoreCell(model.excelFile, m.sheetName, address, oldValue); err != nil {
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
		return fmt.Errorf("unmerge requires a block selection")
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

	if err := model.excelFile.MergeCell(m.sheetName, startAddress, endAddress); err != nil {
		return model, err
	}

	return model, nil
}

type clearOperation struct {
	oldCells  map[string]cellSnapshot
	sheetName string
}

func (c *clearOperation) Init(m model) error {
	c.sheetName = m.sheetName
	c.oldCells = make(map[string]cellSnapshot)

	for _, address := range m.getSelectedCellAddresses() {
		value, err := captureCell(m.excelFile, c.sheetName, address)
		if err != nil {
			return err
		}
		c.oldCells[address] = value
	}

	return nil
}

func (c *clearOperation) Do(m model) (model, error) {
	if err := m.makeCopy(); err != nil {
		return m, err
	}

	for address := range c.oldCells {
		if err := m.excelFile.SetCellValue(c.sheetName, address, nil); err != nil {
			return m, err
		}
	}

	return m, nil
}

func (c *clearOperation) Undo(m model) (model, error) {
	for address, oldValue := range c.oldCells {
		if err := restoreCell(m.excelFile, c.sheetName, address, oldValue); err != nil {
			return m, err
		}
	}

	return m, nil
}

type deleteOperation struct {
	sheetName     string
	selectionKind selectionType

	// For Cell/Block selections
	oldCells map[string]cellSnapshot

	// For Rows/Columns selections
	rowsStart  int
	rowsEnd    int
	colsStart  int
	colsEnd    int
	savedCells map[string]cellSnapshot
}

func (d *deleteOperation) Init(m model) error {
	d.sheetName = m.sheetName
	d.selectionKind = m.selection.kind

	switch m.selection.kind {
	case CellSelect, BlockSelect:
		d.oldCells = make(map[string]cellSnapshot)
		for _, address := range m.getSelectedCellAddresses() {
			value, err := captureCell(m.excelFile, d.sheetName, address)
			if err != nil {
				return err
			}
			d.oldCells[address] = value
		}
	case RowsSelect:
		d.rowsStart = min(m.selection.rows.start, m.selection.rows.end)
		d.rowsEnd = max(m.selection.rows.start, m.selection.rows.end)
		d.savedCells = make(map[string]cellSnapshot)
		for _, address := range m.getSelectedCellAddresses() {
			value, err := captureCell(m.excelFile, d.sheetName, address)
			if err != nil {
				return err
			}
			d.savedCells[address] = value
		}
	case ColumnsSelect:
		d.colsStart = min(m.selection.columns.start, m.selection.columns.end)
		d.colsEnd = max(m.selection.columns.start, m.selection.columns.end)
		d.savedCells = make(map[string]cellSnapshot)
		for _, address := range m.getSelectedCellAddresses() {
			value, err := captureCell(m.excelFile, d.sheetName, address)
			if err != nil {
				return err
			}
			d.savedCells[address] = value
		}
	}

	return nil
}

func (d *deleteOperation) Do(m model) (model, error) {
	if err := m.makeCopy(); err != nil {
		return m, err
	}

	switch d.selectionKind {
	case CellSelect, BlockSelect:
		for address := range d.oldCells {
			if err := m.excelFile.SetCellValue(d.sheetName, address, nil); err != nil {
				return m, err
			}
		}
	case RowsSelect:
		count := d.rowsEnd - d.rowsStart + 1
		for i := 0; i < count; i++ {
			if err := m.excelFile.RemoveRow(d.sheetName, d.rowsStart+1); err != nil {
				return m, err
			}
		}
	case ColumnsSelect:
		count := d.colsEnd - d.colsStart + 1
		for i := 0; i < count; i++ {
			colName, err := excelize.ColumnNumberToName(d.colsStart + 1)
			if err != nil {
				return m, err
			}
			if err := m.excelFile.RemoveCol(d.sheetName, colName); err != nil {
				return m, err
			}
		}
	}

	return m, nil
}

func (d *deleteOperation) Undo(m model) (model, error) {
	switch d.selectionKind {
	case CellSelect, BlockSelect:
		for address, oldValue := range d.oldCells {
			if err := restoreCell(m.excelFile, d.sheetName, address, oldValue); err != nil {
				return m, err
			}
		}
	case RowsSelect:
		count := d.rowsEnd - d.rowsStart + 1
		if err := m.excelFile.InsertRows(d.sheetName, d.rowsStart+1, count); err != nil {
			return m, err
		}
		for address, value := range d.savedCells {
			if err := restoreCell(m.excelFile, d.sheetName, address, value); err != nil {
				return m, err
			}
		}
	case ColumnsSelect:
		count := d.colsEnd - d.colsStart + 1
		colName, err := excelize.ColumnNumberToName(d.colsStart + 1)
		if err != nil {
			return m, err
		}
		if err := m.excelFile.InsertCols(d.sheetName, colName, count); err != nil {
			return m, err
		}
		for address, value := range d.savedCells {
			if err := restoreCell(m.excelFile, d.sheetName, address, value); err != nil {
				return m, err
			}
		}
	}

	return m, nil
}

type clipboardCell struct {
	address string
	value   string
}

type clipboardPasteOperation struct {
	records   [][]string
	cursorX   int
	cursorY   int
	sheetName string
	cells     []clipboardCell
	oldCells  map[string]cellSnapshot
}

func (c *clipboardPasteOperation) Init(m model) error {
	c.sheetName = m.sheetName
	c.cells = make([]clipboardCell, 0)
	c.oldCells = make(map[string]cellSnapshot)
	for rowOffset, record := range c.records {
		for columnOffset, value := range record {
			address, err := excelize.CoordinatesToCellName(c.cursorX+columnOffset+1, c.cursorY+rowOffset+1)
			if err != nil {
				return err
			}
			snapshot, err := captureCell(m.excelFile, c.sheetName, address)
			if err != nil {
				return err
			}
			c.cells = append(c.cells, clipboardCell{address: address, value: value})
			c.oldCells[address] = snapshot
		}
	}
	return nil
}

func (c *clipboardPasteOperation) Do(m model) (model, error) {
	applied := make([]string, 0, len(c.cells))
	for _, cell := range c.cells {
		if err := m.setCellValue(c.sheetName, cell.address, cell.value); err != nil {
			for _, address := range applied {
				_ = restoreCell(m.excelFile, c.sheetName, address, c.oldCells[address])
			}
			return m, fmt.Errorf("paste %s: %w", cell.address, err)
		}
		applied = append(applied, cell.address)
	}
	return m, nil
}

func (c *clipboardPasteOperation) Undo(m model) (model, error) {
	for _, cell := range c.cells {
		if err := restoreCell(m.excelFile, c.sheetName, cell.address, c.oldCells[cell.address]); err != nil {
			return m, err
		}
	}
	return m, nil
}
