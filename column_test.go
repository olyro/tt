package main

import (
	"math"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
	"github.com/xuri/excelize/v2"
)

func TestVisibleColumnLayoutUsesWorksheetWidths(t *testing.T) {
	m := newTestModel(t)
	for column, width := range map[string]float64{"A": 3, "B": 7, "C": 2} {
		if err := m.excelFile.SetColWidth(m.sheetName, column, column, width); err != nil {
			t.Fatal(err)
		}
	}
	m.width = 20
	layout := m.visibleColumnLayout()
	want := []visibleColumn{{index: 0, width: 3}, {index: 1, width: 7}, {index: 2, width: 2}}
	if !reflect.DeepEqual(layout, want) {
		t.Fatalf("layout = %#v, want %#v", layout, want)
	}
	header := strings.SplitN(m.View(), "\n", 2)[0]
	if width := lipgloss.Width(header); width > m.width {
		t.Fatalf("rendered header width = %d, terminal width = %d", width, m.width)
	}

	if column := m.columnAtScreenX(2); column != 0 {
		t.Fatalf("screen position 2 selected column %d", column)
	}
	secondColumnX := m.GetRowNrColumnWidth() + 1 + layout[0].width + 1
	if column := m.columnAtScreenX(secondColumnX); column != 1 {
		t.Fatalf("screen position %d selected column %d", secondColumnX, column)
	}
}

func TestAutoWidthWritesActiveWorksheetWidths(t *testing.T) {
	m := newTestModel(t)
	if _, err := m.excelFile.NewSheet("Other"); err != nil {
		t.Fatal(err)
	}
	for address, value := range map[string]string{
		"A1": "abc",
		"A2": "12345",
		"B1": "short",
		"B2": "line\nlongest",
	} {
		if err := m.excelFile.SetCellStr(m.sheetName, address, value); err != nil {
			t.Fatal(err)
		}
	}
	if err := m.excelFile.SetCellFormula(m.sheetName, "C1", `="12345678"`); err != nil {
		t.Fatal(err)
	}

	message, m, _ := m.evaluateInput("autoWidth")
	if message != "Auto-fitted 3 column(s) on Sheet1" {
		t.Fatalf("unexpected command result: %q", message)
	}
	for column, want := range map[string]float64{"A": 5, "B": 7, "C": 8} {
		got, err := m.excelFile.GetColWidth("Sheet1", column)
		if err != nil {
			t.Fatal(err)
		}
		if math.Abs(got-want) > 0.001 {
			t.Errorf("Sheet1 %s width = %v, want %v", column, got, want)
		}
	}
	otherWidth, err := m.excelFile.GetColWidth("Other", "A")
	if err != nil {
		t.Fatal(err)
	}
	if math.Abs(otherWidth-9.140625) > 0.001 {
		t.Fatalf("auto-width changed another sheet: %v", otherWidth)
	}
	if !m.dirty {
		t.Fatal("auto-width did not mark the workbook dirty")
	}
}

func TestAutoWidthCanTargetOneColumnAndPersists(t *testing.T) {
	m := newTestModel(t)
	if err := m.excelFile.SetCellStr(m.sheetName, "A1", "unchanged"); err != nil {
		t.Fatal(err)
	}
	if err := m.excelFile.SetCellStr(m.sheetName, "B1", "sixsix"); err != nil {
		t.Fatal(err)
	}
	originalA, err := m.excelFile.GetColWidth(m.sheetName, "A")
	if err != nil {
		t.Fatal(err)
	}
	message, m, _ := m.evaluateInput("aw B")
	if message != "Auto-fitted 1 column(s) on Sheet1" {
		t.Fatalf("unexpected command result: %q", message)
	}
	widthA, _ := m.excelFile.GetColWidth(m.sheetName, "A")
	widthB, _ := m.excelFile.GetColWidth(m.sheetName, "B")
	if widthA != originalA || widthB != 6 {
		t.Fatalf("targeted auto-width changed wrong columns: A=%v B=%v", widthA, widthB)
	}

	path := filepath.Join(t.TempDir(), "widths.xlsx")
	message, m, _ = m.evaluateInput("write " + path)
	if message != "written" || m.dirty {
		t.Fatalf("save failed: message=%q dirty=%v", message, m.dirty)
	}
	loaded, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer loaded.Close()
	persisted, err := loaded.GetColWidth("Sheet1", "B")
	if err != nil {
		t.Fatal(err)
	}
	if persisted != 6 {
		t.Fatalf("saved width = %v, want 6", persisted)
	}
}

func TestAutoWidthEmptySheetIsNoOp(t *testing.T) {
	m := newTestModel(t)
	message, m, _ := m.evaluateInput("aw")
	if message != "No populated columns to auto-fit" || m.dirty {
		t.Fatalf("empty auto-width result: message=%q dirty=%v", message, m.dirty)
	}
}

func TestSheetCommandTabCompletionSupportsSpaces(t *testing.T) {
	m := newTestModel(t)
	for _, sheetName := range []string{"Data 2026", "Dashboard"} {
		if _, err := m.excelFile.NewSheet(sheetName); err != nil {
			t.Fatal(err)
		}
	}

	m, _ = handleKeyboardEvent(m, key(":"))
	if !m.input.ShowSuggestions {
		t.Fatal("command mode did not enable sheet completion")
	}
	for _, character := range "b Data" {
		m, _ = handleKeyboardEvent(m, key(string(character)))
	}
	m, _ = handleKeyboardEvent(m, tea.KeyMsg{Type: tea.KeyTab})
	if got := m.input.Value(); got != "b Data 2026" {
		t.Fatalf("tab completed to %q", got)
	}
	m, _ = handleKeyboardEvent(m, key("enter"))
	if m.sheetName != "Data 2026" {
		t.Fatalf("completed sheet command opened %q", m.sheetName)
	}
	if m.input.ShowSuggestions {
		t.Fatal("completion remained enabled outside command mode")
	}
}
