package main

import "github.com/xuri/excelize/v2"

func (m *model) makeCopy() {
	adresses := m.getSelectedCellAddresses()
	values := make(map[string]string)
	styles := make(map[string]int)

	for _, address := range adresses {
		cellType, err := m.excelFile.GetCellType(m.sheetName, address)
		if err == nil && cellType == excelize.CellTypeFormula {
			formula, err := m.excelFile.GetCellFormula(m.sheetName, address)
			if err == nil {
				values[address] = prefixWithEqual(formula)
			}
		} else {
			value, err := m.excelFile.GetCellValue(m.sheetName, address, excelize.Options{
				RawCellValue: true,
			})
			if err != nil {
				continue
			}
			values[address] = value
		}

		style, err := m.excelFile.GetCellStyle(m.sheetName, address)
		if err == nil {
			styles[address] = style
		}
	}

	m.copy = &copy{
		selection: m.selection,
		values:    values,
		styles:    styles,
	}

	m.selection = selection{
		kind: CellSelect,
		cell: &cellSelect{
			x: m.cursorX,
			y: m.cursorY,
		},
	}
}
