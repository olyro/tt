package main

import (
	"errors"
	"path/filepath"
	"testing"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/xuri/excelize/v2"
)

func newTestModel(t *testing.T) model {
	t.Helper()
	file := excelize.NewFile()
	t.Cleanup(func() { _ = file.Close() })
	m := initialModel(file, "Sheet1", "")
	m.width = 80
	m.height = 24
	return m
}

func key(value string) tea.KeyMsg {
	switch value {
	case "esc":
		return tea.KeyMsg{Type: tea.KeyEsc}
	case "enter":
		return tea.KeyMsg{Type: tea.KeyEnter}
	default:
		return tea.KeyMsg{Type: tea.KeyRunes, Runes: []rune(value)}
	}
}

func TestCancelClearsPendingCellEdit(t *testing.T) {
	m := newTestModel(t)
	m, _ = handleKeyboardEvent(m, key("c"))
	if m.currentOp == nil || !m.useInput {
		t.Fatal("change command did not start an edit")
	}
	m, _ = handleKeyboardEvent(m, key("esc"))
	if m.currentOp != nil {
		t.Fatal("escape left a canceled operation armed")
	}
	m, _ = handleKeyboardEvent(m, key(":"))
	m.input.SetValue("type")
	m, _ = handleKeyboardEvent(m, key("enter"))
	value, err := m.excelFile.GetCellValue(m.sheetName, "A1")
	if err != nil {
		t.Fatal(err)
	}
	if value != "" {
		t.Fatalf("command text was written into A1: %q", value)
	}
}

func TestPrefixedNavigationClampsToWorksheetBounds(t *testing.T) {
	m := newTestModel(t)
	m.cursorX, m.offsetX, m.normalInput = 3, 3, "10"
	m, _ = handleKeyboardEvent(m, key("h"))
	if m.cursorX != 0 || m.offsetX != 0 {
		t.Fatalf("left movement was not clamped: cursor=%d offset=%d", m.cursorX, m.offsetX)
	}

	m.cursorY, m.offsetY, m.normalInput = 2, 2, "10"
	m, _ = handleKeyboardEvent(m, key("k"))
	if m.cursorY != 0 || m.offsetY != 0 {
		t.Fatalf("up movement was not clamped: cursor=%d offset=%d", m.cursorY, m.offsetY)
	}

	m.cursorX, m.normalInput = excelize.MaxColumns-2, "10"
	m, _ = handleKeyboardEvent(m, key("l"))
	if m.cursorX != excelize.MaxColumns-1 {
		t.Fatalf("right movement exceeded worksheet: %d", m.cursorX)
	}

	m.cursorY, m.normalInput = maxExcelRows-2, "10"
	m, _ = handleKeyboardEvent(m, key("j"))
	if m.cursorY != maxExcelRows-1 {
		t.Fatalf("down movement exceeded worksheet: %d", m.cursorY)
	}
}

func TestUndoRestoresFormulaAndStringType(t *testing.T) {
	t.Run("formula", func(t *testing.T) {
		m := newTestModel(t)
		if err := m.excelFile.SetCellFormula(m.sheetName, "A1", "=1+1"); err != nil {
			t.Fatal(err)
		}
		m.input.SetValue("replacement")
		var err error
		m, err = executeOperation(m, &changeOperation{})
		if err != nil {
			t.Fatal(err)
		}
		m, _ = handleKeyboardEvent(m, key("u"))
		formula, err := m.excelFile.GetCellFormula(m.sheetName, "A1")
		if err != nil {
			t.Fatal(err)
		}
		if formula != "=1+1" {
			t.Fatalf("formula was not restored: %q", formula)
		}
	})

	t.Run("leading zero string", func(t *testing.T) {
		m := newTestModel(t)
		if err := m.excelFile.SetCellStr(m.sheetName, "A1", "0012"); err != nil {
			t.Fatal(err)
		}
		m.input.SetValue("replacement")
		var err error
		m, err = executeOperation(m, &changeOperation{})
		if err != nil {
			t.Fatal(err)
		}
		m, _ = handleKeyboardEvent(m, key("u"))
		value, err := m.excelFile.GetCellValue(m.sheetName, "A1", excelize.Options{RawCellValue: true})
		if err != nil {
			t.Fatal(err)
		}
		cellType, err := m.excelFile.GetCellType(m.sheetName, "A1")
		if err != nil {
			t.Fatal(err)
		}
		if value != "0012" || (cellType != excelize.CellTypeSharedString && cellType != excelize.CellTypeInlineString) {
			t.Fatalf("string was not restored exactly: value=%q type=%v", value, cellType)
		}
	})
}

func TestReverseBlockPasteUsesNormalizedOrigin(t *testing.T) {
	m := newTestModel(t)
	values := map[string]string{"A1": "a", "B1": "b", "A2": "c", "B2": "d"}
	for address, value := range values {
		if err := m.excelFile.SetCellValue(m.sheetName, address, value); err != nil {
			t.Fatal(err)
		}
	}
	m.selection = selection{kind: BlockSelect, block: &blockSelect{startX: 1, startY: 1, endX: 0, endY: 0}}
	m.makeCopy()
	m.cursorX, m.cursorY = 2, 2

	var err error
	m, err = executeOperation(m, &pasteOperation{})
	if err != nil {
		t.Fatal(err)
	}
	for address, want := range map[string]string{"C3": "a", "D3": "b", "C4": "c", "D4": "d"} {
		got, err := m.excelFile.GetCellValue(m.sheetName, address)
		if err != nil {
			t.Fatal(err)
		}
		if got != want {
			t.Errorf("%s = %q, want %q", address, got, want)
		}
	}
}

func TestResetToCellSelectionClearsStaleRangePointers(t *testing.T) {
	m := newTestModel(t)
	m.selection = selection{
		kind:    RowsSelect,
		rows:    &rowColumnSelect{start: 1, end: 2},
		columns: &rowColumnSelect{start: 3, end: 4},
		block:   &blockSelect{startX: 0, startY: 0, endX: 1, endY: 1},
	}
	m.resetToCellSelection()
	if m.selection.kind != CellSelect || m.selection.cell == nil || m.selection.rows != nil || m.selection.columns != nil || m.selection.block != nil {
		t.Fatalf("selection retained stale pointers: %#v", m.selection)
	}
}

type failingOperation struct {
	initErr error
	doErr   error
}

func (o *failingOperation) Init(model) error            { return o.initErr }
func (o *failingOperation) Do(m model) (model, error)   { return m, o.doErr }
func (o *failingOperation) Undo(m model) (model, error) { return m, nil }

func TestFailedOperationsDoNotEnterHistory(t *testing.T) {
	for _, op := range []operation{
		&failingOperation{initErr: errors.New("init")},
		&failingOperation{doErr: errors.New("do")},
	} {
		m := newTestModel(t)
		if _, err := executeOperation(m, op); err == nil {
			t.Fatal("expected operation to fail")
		}
		if len(m.opStack) != 0 || m.opStackPointer != -1 || m.dirty {
			t.Fatal("failed operation changed history or dirty state")
		}
	}
}

func TestDirtyStateTracksMutationsAndSave(t *testing.T) {
	m := newTestModel(t)
	m.input.SetValue("changed")
	var err error
	m, err = executeOperation(m, &changeOperation{})
	if err != nil {
		t.Fatal(err)
	}
	if !m.dirty {
		t.Fatal("successful edit did not mark workbook dirty")
	}
	if _, _, cmd := m.evaluateInput("q"); cmd != nil {
		t.Fatal("quit was allowed with unsaved changes")
	}
	m, _ = handleKeyboardEvent(m, key("u"))
	if m.dirty {
		t.Fatal("undoing back to the loaded state did not clear dirty state")
	}
	m, _ = handleKeyboardEvent(m, tea.KeyMsg{Type: tea.KeyCtrlR})
	if !m.dirty {
		t.Fatal("redo did not restore dirty state")
	}

	path := filepath.Join(t.TempDir(), "saved.xlsx")
	_, m, _ = m.evaluateInput("write " + path)
	if m.dirty {
		t.Fatal("successful save did not clear dirty state")
	}
	if _, _, cmd := m.evaluateInput("q"); cmd == nil {
		t.Fatal("quit was blocked after saving")
	}
	m, _ = handleKeyboardEvent(m, key("u"))
	if !m.dirty {
		t.Fatal("undoing away from the saved state did not mark the workbook dirty")
	}
	m, _ = handleKeyboardEvent(m, tea.KeyMsg{Type: tea.KeyCtrlR})
	if m.dirty {
		t.Fatal("redoing back to the saved state did not clear dirty state")
	}

	_, m, _ = m.evaluateInput("addSheet Other")
	if !m.dirty {
		t.Fatal("adding a sheet did not mark workbook dirty")
	}
}

func TestUndoHistoryIsScopedPerSheet(t *testing.T) {
	m := newTestModel(t)
	m.input.SetValue("one")
	var err error
	m, err = executeOperation(m, &changeOperation{})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := m.excelFile.NewSheet("Other"); err != nil {
		t.Fatal(err)
	}
	index, _ := m.excelFile.GetSheetIndex("Other")
	m.excelFile.SetActiveSheet(index)
	m = m.activateSheet("Other")
	m.input.SetValue("two")
	m, err = executeOperation(m, &changeOperation{})
	if err != nil {
		t.Fatal(err)
	}
	m, _ = handleKeyboardEvent(m, key("u"))
	other, _ := m.excelFile.GetCellValue("Other", "A1")
	one, _ := m.excelFile.GetCellValue("Sheet1", "A1")
	if other != "" || one != "one" {
		t.Fatalf("sheet undo affected wrong sheet: Sheet1=%q Other=%q", one, other)
	}
	m = m.activateSheet("Sheet1")
	m, _ = handleKeyboardEvent(m, key("u"))
	one, _ = m.excelFile.GetCellValue("Sheet1", "A1")
	if one != "" {
		t.Fatalf("Sheet1 history was not restored: %q", one)
	}
}
