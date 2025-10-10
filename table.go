package docx

import (
	"fmt"
	"strings"
)

// Table represents a table in a Word document
type Table struct {
	rows  []*TableRow
	owner *DocumentPart
}

// TableRow represents a row in a table
type TableRow struct {
	table *Table
	cells []*TableCell
}

// TableCell represents a cell in a table
type TableCell struct {
	row        *TableRow
	paragraphs []*Paragraph
	width      int // width in twentieths of a point
}

// NewTable creates a new table with the specified number of rows and columns
func NewTable(rows, cols int) *Table {
	table := &Table{
		rows: make([]*TableRow, rows),
	}

	for i := 0; i < rows; i++ {
		row := &TableRow{
			table: table,
			cells: make([]*TableCell, cols),
		}

		for j := 0; j < cols; j++ {
			cell := &TableCell{
				row:        row,
				paragraphs: []*Paragraph{NewParagraph()},
				width:      1440, // 1 inch default
			}
			row.cells[j] = cell
		}

		table.rows[i] = row
	}

	return table
}

// Rows returns all rows in the table
func (t *Table) Rows() []*TableRow {
	return t.rows
}

// Row returns the row at the specified index
func (t *Table) Row(index int) *TableRow {
	if index < 0 || index >= len(t.rows) {
		return nil
	}
	return t.rows[index]
}

// AddRow adds a new row to the table
func (t *Table) AddRow() *TableRow {
	// Determine number of columns from first row
	cols := 0
	if len(t.rows) > 0 {
		cols = len(t.rows[0].cells)
	}

	row := &TableRow{
		table: t,
		cells: make([]*TableCell, cols),
	}

	for i := 0; i < cols; i++ {
		cell := &TableCell{
			row:        row,
			paragraphs: []*Paragraph{NewParagraph()},
			width:      1440,
		}
		if len(cell.paragraphs) > 0 && cell.paragraphs[0] != nil {
			cell.paragraphs[0].owner = t.owner
		}
		row.cells[i] = cell
	}

	t.rows = append(t.rows, row)
	return row
}

// ToXML converts the table to WordprocessingML XML
func (t *Table) ToXML() string {
	var rowsXML strings.Builder

	for _, row := range t.rows {
		rowsXML.WriteString(row.ToXML())
	}

	return fmt.Sprintf(`<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders>
  </w:tblPr>
  %s
</w:tbl>`, rowsXML.String())
}

// Cells returns all cells in the row
func (tr *TableRow) Cells() []*TableCell {
	return tr.cells
}

// Cell returns the cell at the specified index
func (tr *TableRow) Cell(index int) *TableCell {
	if index < 0 || index >= len(tr.cells) {
		return nil
	}
	return tr.cells[index]
}

// ToXML converts the table row to WordprocessingML XML
func (tr *TableRow) ToXML() string {
	var cellsXML strings.Builder

	for _, cell := range tr.cells {
		cellsXML.WriteString(cell.ToXML())
	}

	return fmt.Sprintf(`<w:tr>%s</w:tr>`, cellsXML.String())
}

// Paragraphs returns all paragraphs in the cell
func (tc *TableCell) Paragraphs() []*Paragraph {
	return tc.paragraphs
}

// AddParagraph adds a new paragraph to the cell
func (tc *TableCell) AddParagraph(text ...string) *Paragraph {
	paragraph := NewParagraph()
	if tc.row != nil && tc.row.table != nil {
		paragraph.owner = tc.row.table.owner
	}

	for _, t := range text {
		paragraph.AddRun(t)
	}

	tc.paragraphs = append(tc.paragraphs, paragraph)
	return paragraph
}

func (t *Table) setOwner(owner *DocumentPart) {
	t.owner = owner
	for _, row := range t.rows {
		row.table = t
		for _, cell := range row.cells {
			cell.row = row
			for _, paragraph := range cell.paragraphs {
				paragraph.owner = owner
			}
		}
	}
}

// Text returns the combined text of all paragraphs in the cell
func (tc *TableCell) Text() string {
	var text strings.Builder
	for i, paragraph := range tc.paragraphs {
		if i > 0 {
			text.WriteString("\n")
		}
		text.WriteString(paragraph.Text())
	}
	return text.String()
}

// SetText clears the cell and sets it to contain a single paragraph with the given text
func (tc *TableCell) SetText(text string) {
	paragraph := NewParagraph()
	if tc.row != nil && tc.row.table != nil {
		paragraph.owner = tc.row.table.owner
	}
	paragraph.AddRun(text)
	tc.paragraphs = []*Paragraph{paragraph}
}

// SetWidth sets the width of the cell in twentieths of a point
func (tc *TableCell) SetWidth(width int) {
	tc.width = width
}

// Width returns the width of the cell in twentieths of a point
func (tc *TableCell) Width() int {
	return tc.width
}

// ToXML converts the table cell to WordprocessingML XML
func (tc *TableCell) ToXML() string {
	var paragraphsXML strings.Builder

	for _, paragraph := range tc.paragraphs {
		paragraphsXML.WriteString(paragraph.ToXML())
	}

	return fmt.Sprintf(`<w:tc>
  <w:tcPr>
    <w:tcW w:w="%d" w:type="dxa"/>
  </w:tcPr>
  %s
</w:tc>`, tc.width, paragraphsXML.String())
}
