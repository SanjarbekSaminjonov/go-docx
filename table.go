package docx

import (
	"errors"
	"fmt"
	"strings"
)

// Table represents a table in a Word document
type Table struct {
	rows        []*TableRow
	owner       *DocumentPart
	gridColumns int
	borders     map[TableBorderSide]*TableBorder
	shading     *Shading
	cellMargins *TableCellMargins
}

// Shading describes background shading for table elements.
type Shading struct {
	Pattern string
	Fill    string
	Color   string
}

// TableRow represents a row in a table
type TableRow struct {
	table *Table
	cells []*TableCell
}

// TableCell represents a cell in a table
type TableCell struct {
	row           *TableRow
	paragraphs    []*Paragraph
	width         int // width in twentieths of a point
	gridSpan      int
	verticalMerge TableVerticalMerge
	borders       map[TableBorderSide]*TableBorder
	shading       *Shading
}

// TableBorderSide identifies borders on tables and cells.
type TableBorderSide string

const (
	TableBorderTop     TableBorderSide = "top"
	TableBorderLeft    TableBorderSide = "left"
	TableBorderBottom  TableBorderSide = "bottom"
	TableBorderRight   TableBorderSide = "right"
	TableBorderInsideH TableBorderSide = "insideH"
	TableBorderInsideV TableBorderSide = "insideV"
)

// TableBorder describes border appearance.
type TableBorder struct {
	Style string // e.g. "single", "double"
	Color string // Hex color (without #) or "auto"
	Size  int    // Width in eighths of a point
	Space int    // Spacing in twips between border and content
}

// TableCellMargins represents default margins for table cells.
type TableCellMargins struct {
	Top, Left, Bottom, Right *int
}

// TableVerticalMerge represents vertical merge state of a cell.
type TableVerticalMerge string

const (
	TableVerticalMergeNone     TableVerticalMerge = ""
	TableVerticalMergeRestart  TableVerticalMerge = "restart"
	TableVerticalMergeContinue TableVerticalMerge = "continue"
)

// NewTable creates a new table with the specified number of rows and columns
func NewTable(rows, cols int) *Table {
	table := &Table{
		rows:        make([]*TableRow, rows),
		gridColumns: cols,
		borders:     make(map[TableBorderSide]*TableBorder),
	}

	for _, side := range []TableBorderSide{TableBorderTop, TableBorderLeft, TableBorderBottom, TableBorderRight, TableBorderInsideH, TableBorderInsideV} {
		border := TableBorder{Style: "single", Color: "auto", Size: 4, Space: 0}
		copy := border
		table.borders[side] = &copy
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
				borders:    make(map[TableBorderSide]*TableBorder),
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
	cols := t.gridColumns
	row := &TableRow{
		table: t,
		cells: make([]*TableCell, cols),
	}

	for i := 0; i < cols; i++ {
		cell := &TableCell{
			row:        row,
			paragraphs: []*Paragraph{NewParagraph()},
			width:      1440,
			borders:    make(map[TableBorderSide]*TableBorder),
		}
		if len(cell.paragraphs) > 0 && cell.paragraphs[0] != nil {
			cell.paragraphs[0].owner = t.owner
		}
		row.cells[i] = cell
	}

	t.rows = append(t.rows, row)
	return row
}

// InsertRowAt inserts a new empty row at the given index (0..len). If index is out of range,
// it will be clamped to the valid range. The new row will have the same number of columns as the table grid.
func (t *Table) InsertRowAt(index int) *TableRow {
	if index < 0 {
		index = 0
	}
	if index > len(t.rows) {
		index = len(t.rows)
	}
	cols := t.gridColumns
	row := &TableRow{
		table: t,
		cells: make([]*TableCell, cols),
	}
	for i := 0; i < cols; i++ {
		cell := &TableCell{
			row:        row,
			paragraphs: []*Paragraph{NewParagraph()},
			width:      1440,
			borders:    make(map[TableBorderSide]*TableBorder),
		}
		if len(cell.paragraphs) > 0 && cell.paragraphs[0] != nil {
			cell.paragraphs[0].owner = t.owner
		}
		row.cells[i] = cell
	}
	t.rows = append(t.rows, nil)
	copy(t.rows[index+1:], t.rows[index:])
	t.rows[index] = row
	return row
}

// SetBorder configures the border for the specified table side. An empty style clears the border.
func (t *Table) SetBorder(side TableBorderSide, border TableBorder) {
	if side == "" {
		return
	}
	if t.borders == nil {
		t.borders = make(map[TableBorderSide]*TableBorder)
	}
	if border.Style == "" {
		delete(t.borders, side)
		return
	}
	copy := border
	t.borders[side] = &copy
}

// Border returns the configured border for the specified side.
func (t *Table) Border(side TableBorderSide) (*TableBorder, bool) {
	if t.borders == nil {
		return nil, false
	}
	border, ok := t.borders[side]
	return border, ok
}

// ClearBorders removes all table border definitions.
func (t *Table) ClearBorders() {
	t.borders = make(map[TableBorderSide]*TableBorder)
}

// SetShading configures table-level shading.
func (t *Table) SetShading(pattern, fill, color string) {
	t.shading = &Shading{Pattern: pattern, Fill: fill, Color: color}
}

// Shading returns table-level shading if configured.
func (t *Table) Shading() (*Shading, bool) {
	if t.shading == nil {
		return nil, false
	}
	return t.shading, true
}

// ClearShading removes table-level shading.
func (t *Table) ClearShading() {
	t.shading = nil
}

// SetCellMargins configures table-wide default cell margins (twentieths of a point).
func (t *Table) SetCellMargins(top, left, bottom, right int) {
	if t.cellMargins == nil {
		t.cellMargins = &TableCellMargins{}
	}
	t.cellMargins.Top = intPtr(top)
	t.cellMargins.Left = intPtr(left)
	t.cellMargins.Bottom = intPtr(bottom)
	t.cellMargins.Right = intPtr(right)
}

// CellMargins returns the configured cell margins if present.
func (t *Table) CellMargins() (*TableCellMargins, bool) {
	if t.cellMargins == nil {
		return nil, false
	}
	return t.cellMargins, true
}

// ClearCellMargins removes table-wide cell margin settings.
func (t *Table) ClearCellMargins() {
	t.cellMargins = nil
}

// MergeCellsHorizontally merges cells in the specified row between start and end inclusive.
func (t *Table) MergeCellsHorizontally(rowIndex, start, end int) error {
	if rowIndex < 0 || rowIndex >= len(t.rows) {
		return fmt.Errorf("row index %d out of range", rowIndex)
	}
	if start < 0 || end < start {
		return errors.New("invalid start/end for horizontal merge")
	}
	row := t.rows[rowIndex]
	if end >= len(row.cells) {
		return fmt.Errorf("end index %d out of range", end)
	}
	span := end - start + 1
	if span <= 1 {
		return nil
	}
	cell := row.cells[start]
	totalWidth := cell.width
	for i := end; i > start; i-- {
		totalWidth += row.cells[i].width
		row.cells = append(row.cells[:i], row.cells[i+1:]...)
	}
	cell.width = totalWidth
	cell.gridSpan = span
	return nil
}

// MergeCellsVertically merges cells in the specified column between the start and end row (inclusive).
func (t *Table) MergeCellsVertically(column, startRow, endRow int) error {
	if column < 0 {
		return errors.New("column index must be non-negative")
	}
	if startRow < 0 || endRow < startRow || endRow >= len(t.rows) {
		return errors.New("invalid start/end row for vertical merge")
	}
	for rowIndex := startRow; rowIndex <= endRow; rowIndex++ {
		row := t.rows[rowIndex]
		if column >= len(row.cells) {
			return fmt.Errorf("column index %d out of range for row %d", column, rowIndex)
		}
		cell := row.cells[column]
		if rowIndex == startRow {
			cell.verticalMerge = TableVerticalMergeRestart
		} else {
			cell.verticalMerge = TableVerticalMergeContinue
			if len(cell.paragraphs) == 0 {
				continue
			}
		}
	}
	return nil
}

func (t *Table) tblPropertiesXML() string {
	var builder strings.Builder
	builder.WriteString("<w:tblPr>")
	builder.WriteString(`<w:tblW w:w="0" w:type="auto"/>`)
	if t.hasBorders() {
		builder.WriteString(t.bordersXML())
	}
	if t.hasShading() {
		builder.WriteString(shadingElement(t.shading))
	}
	if t.hasCellMargins() {
		builder.WriteString(t.cellMarginsXML())
	}
	builder.WriteString("</w:tblPr>")
	return builder.String()
}

func (t *Table) hasBorders() bool {
	if len(t.borders) == 0 {
		return false
	}
	for _, border := range t.borders {
		if border != nil && border.Style != "" {
			return true
		}
	}
	return false
}

func (t *Table) bordersXML() string {
	if !t.hasBorders() {
		return ""
	}
	var builder strings.Builder
	builder.WriteString("<w:tblBorders>")
	for _, side := range []TableBorderSide{TableBorderTop, TableBorderLeft, TableBorderBottom, TableBorderRight, TableBorderInsideH, TableBorderInsideV} {
		border, ok := t.borders[side]
		if !ok || border == nil || border.Style == "" {
			continue
		}
		builder.WriteString(borderElement(string(side), border))
	}
	builder.WriteString("</w:tblBorders>")
	return builder.String()
}

func (t *Table) hasShading() bool {
	return t.shading != nil
}

func (t *Table) hasCellMargins() bool {
	if t.cellMargins == nil {
		return false
	}
	return t.cellMargins.Top != nil || t.cellMargins.Left != nil || t.cellMargins.Bottom != nil || t.cellMargins.Right != nil
}

func (t *Table) cellMarginsXML() string {
	if t.cellMargins == nil {
		return ""
	}
	var builder strings.Builder
	builder.WriteString("<w:tblCellMar>")
	if t.cellMargins.Top != nil {
		builder.WriteString(fmt.Sprintf(`<w:top w:w="%d" w:type="dxa"/>`, *t.cellMargins.Top))
	}
	if t.cellMargins.Left != nil {
		builder.WriteString(fmt.Sprintf(`<w:left w:w="%d" w:type="dxa"/>`, *t.cellMargins.Left))
	}
	if t.cellMargins.Bottom != nil {
		builder.WriteString(fmt.Sprintf(`<w:bottom w:w="%d" w:type="dxa"/>`, *t.cellMargins.Bottom))
	}
	if t.cellMargins.Right != nil {
		builder.WriteString(fmt.Sprintf(`<w:right w:w="%d" w:type="dxa"/>`, *t.cellMargins.Right))
	}
	builder.WriteString("</w:tblCellMar>")
	return builder.String()
}

func (tc *TableCell) hasBorders() bool {
	if len(tc.borders) == 0 {
		return false
	}
	for _, border := range tc.borders {
		if border != nil && border.Style != "" {
			return true
		}
	}
	return false
}

func (tc *TableCell) bordersXML() string {
	if !tc.hasBorders() {
		return ""
	}
	var builder strings.Builder
	builder.WriteString("<w:tcBorders>")
	for _, side := range []TableBorderSide{TableBorderTop, TableBorderLeft, TableBorderBottom, TableBorderRight, TableBorderInsideH, TableBorderInsideV} {
		border, ok := tc.borders[side]
		if !ok || border == nil || border.Style == "" {
			continue
		}
		builder.WriteString(borderElement(string(side), border))
	}
	builder.WriteString("</w:tcBorders>")
	return builder.String()
}

func borderElement(tag string, border *TableBorder) string {
	attrs := []string{fmt.Sprintf(`w:val="%s"`, border.Style)}
	if border.Size > 0 {
		attrs = append(attrs, fmt.Sprintf(`w:sz="%d"`, border.Size))
	}
	if border.Space > 0 {
		attrs = append(attrs, fmt.Sprintf(`w:space="%d"`, border.Space))
	}
	color := border.Color
	if color == "" {
		color = "auto"
	}
	attrs = append(attrs, fmt.Sprintf(`w:color="%s"`, color))
	return fmt.Sprintf(`<w:%s %s/>`, tag, strings.Join(attrs, " "))
}

func shadingElement(shading *Shading) string {
	if shading == nil {
		return ""
	}
	pattern := shading.Pattern
	if pattern == "" {
		pattern = "clear"
	}
	fill := shading.Fill
	if fill == "" {
		fill = "auto"
	}
	color := shading.Color
	if color == "" {
		color = "auto"
	}
	return fmt.Sprintf(`<w:shd w:val="%s" w:color="%s" w:fill="%s"/>`, pattern, color, fill)
}

// ToXML converts the table to WordprocessingML XML
func (t *Table) ToXML() string {
	var rowsXML strings.Builder

	for _, row := range t.rows {
		rowsXML.WriteString(row.ToXML())
	}

	return fmt.Sprintf(`<w:tbl>%s%s</w:tbl>`, t.tblPropertiesXML(), rowsXML.String())
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

// SetGridSpan configures the number of grid columns spanned by this cell.
func (tc *TableCell) SetGridSpan(span int) {
	if span <= 1 {
		tc.gridSpan = 0
		return
	}
	tc.gridSpan = span
}

// GridSpan returns the number of columns this cell spans (default 1).
func (tc *TableCell) GridSpan() int {
	if tc.gridSpan <= 1 {
		return 1
	}
	return tc.gridSpan
}

// SetVerticalMerge sets the vertical merge behavior for this cell.
func (tc *TableCell) SetVerticalMerge(state TableVerticalMerge) {
	tc.verticalMerge = state
}

// VerticalMerge returns the vertical merge state for this cell.
func (tc *TableCell) VerticalMerge() TableVerticalMerge {
	return tc.verticalMerge
}

// ClearVerticalMerge removes the vertical merge setting.
func (tc *TableCell) ClearVerticalMerge() {
	tc.verticalMerge = TableVerticalMergeNone
}

// SetShading configures the cell shading.
func (tc *TableCell) SetShading(pattern, fill, color string) {
	tc.shading = &Shading{Pattern: pattern, Fill: fill, Color: color}
}

// Shading returns the cell shading if set.
func (tc *TableCell) Shading() (*Shading, bool) {
	if tc.shading == nil {
		return nil, false
	}
	return tc.shading, true
}

// ClearShading removes cell shading.
func (tc *TableCell) ClearShading() {
	tc.shading = nil
}

// SetBorder configures a border for the cell.
func (tc *TableCell) SetBorder(side TableBorderSide, border TableBorder) {
	if side == "" {
		return
	}
	if tc.borders == nil {
		tc.borders = make(map[TableBorderSide]*TableBorder)
	}
	if border.Style == "" {
		delete(tc.borders, side)
		return
	}
	copy := border
	tc.borders[side] = &copy
}

// Border returns the border definition for the specified side.
func (tc *TableCell) Border(side TableBorderSide) (*TableBorder, bool) {
	if tc.borders == nil {
		return nil, false
	}
	border, ok := tc.borders[side]
	return border, ok
}

// ClearBorders removes all borders from the cell.
func (tc *TableCell) ClearBorders() {
	tc.borders = make(map[TableBorderSide]*TableBorder)
}

func (tc *TableCell) hasShading() bool {
	return tc.shading != nil
}

func (tc *TableCell) tcPropertiesXML() string {
	var builder strings.Builder
	builder.WriteString("<w:tcPr>")
	builder.WriteString(fmt.Sprintf(`<w:tcW w:w="%d" w:type="dxa"/>`, tc.width))
	if tc.gridSpan > 1 {
		builder.WriteString(fmt.Sprintf(`<w:gridSpan w:val="%d"/>`, tc.gridSpan))
	}
	switch tc.verticalMerge {
	case TableVerticalMergeRestart:
		builder.WriteString(`<w:vMerge w:val="restart"/>`)
	case TableVerticalMergeContinue:
		builder.WriteString(`<w:vMerge w:val="continue"/>`)
	}
	if tc.hasBorders() {
		builder.WriteString(tc.bordersXML())
	}
	if tc.hasShading() {
		builder.WriteString(shadingElement(tc.shading))
	}
	builder.WriteString("</w:tcPr>")
	return builder.String()
}

// ToXML converts the table cell to WordprocessingML XML
func (tc *TableCell) ToXML() string {
	var paragraphsXML strings.Builder

	for _, paragraph := range tc.paragraphs {
		paragraphsXML.WriteString(paragraph.ToXML())
	}

	return fmt.Sprintf(`<w:tc>%s%s</w:tc>`, tc.tcPropertiesXML(), paragraphsXML.String())
}
