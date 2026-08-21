package main

import (
	"errors"
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
)

type fakeClipboard struct {
	readValue string
	readErr   error
	written   string
	writeErr  error
}

func (c *fakeClipboard) ReadAll() (string, error) { return c.readValue, c.readErr }
func (c *fakeClipboard) WriteAll(value string) error {
	c.written = value
	return c.writeErr
}

func TestClipboardTSVRoundTrip(t *testing.T) {
	m := newTestModel(t)
	values := map[string]string{
		"A1": "plain",
		"B1": "with\ttab",
		"A2": "with\nnewline",
		"B2": "Grüße 👋",
	}
	for address, value := range values {
		if err := m.excelFile.SetCellStr(m.sheetName, address, value); err != nil {
			t.Fatal(err)
		}
	}
	m.selection = selection{kind: BlockSelect, block: &blockSelect{startX: 0, startY: 0, endX: 1, endY: 2}}

	encoded, err := m.exportSelectionTSV()
	if err != nil {
		t.Fatal(err)
	}
	decoded, err := parseClipboardTSV(encoded)
	if err != nil {
		t.Fatal(err)
	}
	want := [][]string{{"plain", "with\ttab"}, {"with\nnewline", "Grüße 👋"}, {"", ""}}
	if !reflect.DeepEqual(decoded, want) {
		t.Fatalf("TSV round trip = %#v, want %#v", decoded, want)
	}
}

func TestSingleEmptyCellExportsAsTSVField(t *testing.T) {
	m := newTestModel(t)
	encoded, err := m.exportSelectionTSV()
	if err != nil {
		t.Fatal(err)
	}
	if encoded != `""` {
		t.Fatalf("empty cell encoded as %q", encoded)
	}
	records, err := parseClipboardTSV(encoded)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(records, [][]string{{""}}) {
		t.Fatalf("empty cell decoded as %#v", records)
	}
}

func TestParseClipboardTSVHandlesCRLFAndRaggedRows(t *testing.T) {
	records, err := parseClipboardTSV("a\tb\r\nc\r\n")
	if err != nil {
		t.Fatal(err)
	}
	want := [][]string{{"a", "b"}, {"c"}}
	if !reflect.DeepEqual(records, want) {
		t.Fatalf("records = %#v, want %#v", records, want)
	}
}

func TestClipboardImportIsOneUndoableOperation(t *testing.T) {
	m := newTestModel(t)
	if err := m.excelFile.SetCellFormula(m.sheetName, "A1", "=2+2"); err != nil {
		t.Fatal(err)
	}
	style, err := m.excelFile.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true}})
	if err != nil {
		t.Fatal(err)
	}
	if err := m.excelFile.SetCellStyle(m.sheetName, "A1", "A1", style); err != nil {
		t.Fatal(err)
	}

	m = m.importClipboard("12\t=1+1\ntext\t")
	if len(m.opStack) != 1 || m.opStackPointer != 0 || !m.dirty {
		t.Fatalf("import was not one dirty operation: len=%d pointer=%d dirty=%v", len(m.opStack), m.opStackPointer, m.dirty)
	}
	for address, want := range map[string]string{"A1": "12", "A2": "text", "B2": ""} {
		got, err := m.excelFile.GetCellValue(m.sheetName, address)
		if err != nil {
			t.Fatal(err)
		}
		if got != want {
			t.Errorf("%s = %q, want %q", address, got, want)
		}
	}
	importedFormula, err := m.excelFile.GetCellFormula(m.sheetName, "B1")
	if err != nil {
		t.Fatal(err)
	}
	if importedFormula != "=1+1" {
		t.Fatalf("B1 formula = %q", importedFormula)
	}

	m, _ = handleKeyboardEvent(m, key("u"))
	formula, err := m.excelFile.GetCellFormula(m.sheetName, "A1")
	if err != nil {
		t.Fatal(err)
	}
	restoredStyle, err := m.excelFile.GetCellStyle(m.sheetName, "A1")
	if err != nil {
		t.Fatal(err)
	}
	if formula != "=2+2" || restoredStyle != style {
		t.Fatalf("undo did not restore formula/style: formula=%q style=%d", formula, restoredStyle)
	}
}

func TestClipboardImportRejectsWorksheetOverflow(t *testing.T) {
	m := newTestModel(t)
	m.cursorX = excelize.MaxColumns - 1
	m = m.importClipboard("a\tb")
	if m.dirty || len(m.opStack) != 0 {
		t.Fatal("overflowing import changed the workbook")
	}
	if got := m.input.Value(); got != "Clipboard data exceeds Excel worksheet limits" {
		t.Fatalf("unexpected overflow message: %q", got)
	}
}

func TestSystemClipboardKeyBindings(t *testing.T) {
	m := newTestModel(t)
	if err := m.excelFile.SetCellStr(m.sheetName, "A1", "exported"); err != nil {
		t.Fatal(err)
	}
	fake := &fakeClipboard{readValue: "imported"}
	m.clipboard = fake

	updated, writeCmd := handleKeyboardEvent(m, key("Y"))
	if writeCmd == nil {
		t.Fatal("Y did not return a clipboard command")
	}
	writeMsg := writeCmd()
	if fake.written != "exported" {
		t.Fatalf("Y wrote %q", fake.written)
	}
	modelAfterWrite, _ := updated.Update(writeMsg)
	if got := modelAfterWrite.(model).input.Value(); got != "copied to system clipboard" {
		t.Fatalf("unexpected write status: %q", got)
	}

	updated, readCmd := handleKeyboardEvent(updated, key("P"))
	if readCmd == nil {
		t.Fatal("P did not return a clipboard command")
	}
	readMsg := readCmd()
	modelAfterRead, _ := updated.Update(readMsg)
	value, err := modelAfterRead.(model).excelFile.GetCellValue("Sheet1", "A1")
	if err != nil {
		t.Fatal(err)
	}
	if value != "imported" {
		t.Fatalf("P imported %q", value)
	}
}

func TestClipboardErrorsAreReportedWithoutMutation(t *testing.T) {
	m := newTestModel(t)
	m.clipboard = &fakeClipboard{readErr: errors.New("unavailable")}
	updated, cmd := handleKeyboardEvent(m, key("P"))
	if cmd == nil {
		t.Fatal("P did not return a clipboard command")
	}
	result, _ := updated.Update(cmd())
	got := result.(model)
	if got.dirty || got.input.Value() != "Clipboard error: unavailable" {
		t.Fatalf("clipboard error mutated model or was not reported: dirty=%v message=%q", got.dirty, got.input.Value())
	}
}
