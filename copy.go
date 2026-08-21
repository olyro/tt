package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"strings"

	"github.com/atotto/clipboard"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/xuri/excelize/v2"
)

type clipboardService interface {
	ReadAll() (string, error)
	WriteAll(string) error
}

type systemClipboard struct{}

func (systemClipboard) ReadAll() (string, error)   { return clipboard.ReadAll() }
func (systemClipboard) WriteAll(text string) error { return clipboard.WriteAll(text) }

type clipboardWriteMsg struct{ err error }

type clipboardReadMsg struct {
	text string
	err  error
}

func writeClipboardCmd(service clipboardService, value string) tea.Cmd {
	return func() tea.Msg { return clipboardWriteMsg{err: service.WriteAll(value)} }
}

func readClipboardCmd(service clipboardService) tea.Cmd {
	return func() tea.Msg {
		value, err := service.ReadAll()
		return clipboardReadMsg{text: value, err: err}
	}
}

func (m *model) makeCopy() error {
	addresses := m.getSelectedCellAddresses()
	cells := make(map[string]cellSnapshot, len(addresses))

	for _, address := range addresses {
		snapshot, err := captureCell(m.excelFile, m.sheetName, address)
		if err != nil {
			return err
		}
		cells[address] = snapshot
	}

	m.copy = &copy{selection: m.selection, cells: cells}
	m.resetToCellSelection()
	return nil
}

func (m model) selectionBounds() (startX, startY, endX, endY int, err error) {
	switch m.selection.kind {
	case CellSelect:
		return m.selection.cell.x, m.selection.cell.y, m.selection.cell.x, m.selection.cell.y, nil
	case BlockSelect:
		return min(m.selection.block.startX, m.selection.block.endX),
			min(m.selection.block.startY, m.selection.block.endY),
			max(m.selection.block.startX, m.selection.block.endX),
			max(m.selection.block.startY, m.selection.block.endY), nil
	case RowsSelect:
		rows, rowsErr := m.excelFile.GetRows(m.sheetName)
		if rowsErr != nil {
			return 0, 0, 0, 0, rowsErr
		}
		maxColumns := 1
		for _, row := range rows {
			maxColumns = max(maxColumns, len(row))
		}
		return 0, min(m.selection.rows.start, m.selection.rows.end), maxColumns - 1,
			max(m.selection.rows.start, m.selection.rows.end), nil
	case ColumnsSelect:
		rows, rowsErr := m.excelFile.GetRows(m.sheetName)
		if rowsErr != nil {
			return 0, 0, 0, 0, rowsErr
		}
		return min(m.selection.columns.start, m.selection.columns.end), 0,
			max(m.selection.columns.start, m.selection.columns.end), max(len(rows)-1, 0), nil
	default:
		return 0, 0, 0, 0, fmt.Errorf("unsupported selection")
	}
}

func (m model) exportSelectionTSV() (string, error) {
	startX, startY, endX, endY, err := m.selectionBounds()
	if err != nil {
		return "", err
	}

	var output strings.Builder
	for y := startY; y <= endY; y++ {
		record := make([]string, endX-startX+1)
		for x := startX; x <= endX; x++ {
			address, addressErr := excelize.CoordinatesToCellName(x+1, y+1)
			if addressErr != nil {
				return "", addressErr
			}
			cellType, typeErr := m.excelFile.GetCellType(m.sheetName, address)
			if typeErr != nil {
				return "", typeErr
			}
			if cellType == excelize.CellTypeFormula {
				record[x-startX], err = m.excelFile.CalcCellValue(m.sheetName, address)
			} else {
				record[x-startX], err = m.excelFile.GetCellValue(m.sheetName, address)
			}
			if err != nil {
				return "", err
			}
		}
		if y > startY {
			output.WriteByte('\n')
		}
		for index, field := range record {
			if index > 0 {
				output.WriteByte('\t')
			}
			output.WriteString(encodeTSVField(field))
		}
	}
	return output.String(), nil
}

func encodeTSVField(value string) string {
	if value != "" && !strings.ContainsAny(value, "\t\r\n\"") {
		return value
	}
	return `"` + strings.ReplaceAll(value, `"`, `""`) + `"`
}

func parseClipboardTSV(value string) ([][]string, error) {
	if value == "" {
		return nil, fmt.Errorf("clipboard is empty")
	}
	reader := csv.NewReader(strings.NewReader(value))
	reader.Comma = '\t'
	reader.FieldsPerRecord = -1
	records := make([][]string, 0)
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("invalid TSV: %w", err)
		}
		records = append(records, record)
	}
	if len(records) == 0 {
		return nil, fmt.Errorf("clipboard is empty")
	}
	return records, nil
}

func (m model) importClipboard(value string) model {
	records, err := parseClipboardTSV(value)
	if err != nil {
		m.input.SetValue("Clipboard error: " + err.Error())
		return m
	}
	maxWidth := 0
	for _, record := range records {
		maxWidth = max(maxWidth, len(record))
	}
	if m.cursorY+len(records) > maxExcelRows || m.cursorX+maxWidth > excelize.MaxColumns {
		m.input.SetValue("Clipboard data exceeds Excel worksheet limits")
		return m
	}
	op := &clipboardPasteOperation{records: records, cursorX: m.cursorX, cursorY: m.cursorY}
	newModel, err := executeOperation(m, op)
	if err != nil {
		m.input.SetValue(err.Error())
		return m
	}
	newModel.resetToCellSelection()
	newModel.input.SetValue(fmt.Sprintf("pasted %d row(s) from system clipboard", len(records)))
	return newModel
}
