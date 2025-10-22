package docx

import (
	"errors"
	"fmt"
	"strings"
)

// Table represents a table in a Word document
type Table struct {
	rows           []*TableRow
	owner          *DocumentPart
	gridColumns    int
	grid           []int
	width          int // table width in twentieths of a point (0 for auto)
	widthType      string
	indent         int
	indentType     string
	indentSet      bool
	style          string
	layout         string
	look           *TableLook
	alignment      TableAlignment
	bordersDefined bool
	borders        map[TableBorderSide]*TableBorder
	shading        *Shading
	cellMargins    *TableCellMargins
}

var xmlAttrEscaper = strings.NewReplacer(
	"&", "&amp;",
	"\"", "&quot;",
	"<", "&lt;",
	">", "&gt;",
	"'", "&apos;",
)

func xmlEscapeAttribute(value string) string {
	return xmlAttrEscaper.Replace(value)
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
	verticalAlign WDVerticalAlignment // vertical alignment in cell
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

// TableLook describes visual flags applied to a table style.
type TableLook struct {
	Val         string
	FirstRow    bool
	LastRow     bool
	FirstColumn bool
	LastColumn  bool
	NoHBand     bool
	NoVBand     bool
}

// TableAlignment represents table justification.
type TableAlignment string

const (
	TableAlignmentLeft   TableAlignment = "left"
	TableAlignmentCenter TableAlignment = "center"
	TableAlignmentRight  TableAlignment = "right"
	TableAlignmentBoth   TableAlignment = "both"
	TableAlignmentStart  TableAlignment = "start"
	TableAlignmentEnd    TableAlignment = "end"
)

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

// WDVerticalAlignment represents vertical alignment options for table cells
type WDVerticalAlignment string

const (
	WDVerticalAlignmentTop    WDVerticalAlignment = "top"
	WDVerticalAlignmentCenter WDVerticalAlignment = "center"
	WDVerticalAlignmentBottom WDVerticalAlignment = "bottom"
)

// NewTable creates a new table with the specified number of rows and columns
func NewTable(rows, cols int) *Table {
	table := &Table{
		rows:        make([]*TableRow, rows),
		gridColumns: cols,
		grid:        make([]int, cols),
		widthType:   "auto",
		borders:     make(map[TableBorderSide]*TableBorder),
	}

	for _, side := range []TableBorderSide{TableBorderTop, TableBorderLeft, TableBorderBottom, TableBorderRight, TableBorderInsideH, TableBorderInsideV} {
		table.SetBorder(side, TableBorder{Style: "single", Color: "auto", Size: 4, Space: 0})
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

	for i := range table.grid {
		table.grid[i] = 1440
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

// GetRow is an alias for Row() - returns the row at the specified index
func (t *Table) GetRow(index int) *TableRow {
	return t.Row(index)
}

// AddRow adds a new row to the table
func (t *Table) AddRow() *TableRow {
	// Agar table'da qatorlar mavjud bo'lsa, oxirgi qatordagi ustunlar sonidan foydalanish
	cols := t.gridColumns
	if len(t.rows) > 0 && t.rows[len(t.rows)-1] != nil {
		cols = len(t.rows[len(t.rows)-1].cells)
	}
	if cols <= 0 {
		cols = len(t.grid)
	}
	if cols <= 0 {
		cols = 1
	}

	t.ensureGridLength(cols)
	if cols > t.gridColumns {
		t.gridColumns = cols
	}

	row := &TableRow{
		table: t,
		cells: make([]*TableCell, cols),
	}

	for i := 0; i < cols; i++ {
		width := 1440
		if i < len(t.grid) && t.grid[i] > 0 {
			width = t.grid[i]
		}
		cell := &TableCell{
			row:        row,
			paragraphs: []*Paragraph{NewParagraph()},
			width:      width,
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

	// Agar table'da qatorlar mavjud bo'lsa, avvalgi qatordagi ustunlar sonidan foydalanish
	cols := t.gridColumns
	if len(t.rows) > 0 {
		// Yaqin qatordagi ustunlar sonini olish
		refIndex := index - 1
		if refIndex < 0 && len(t.rows) > 0 {

			if cols <= 0 {
				cols = len(t.grid)
			}
			if cols <= 0 {
				cols = 1
			}

			t.ensureGridLength(cols)
			if cols > t.gridColumns {
				t.gridColumns = cols
			}
			refIndex = 0 // Birinchi qatorga qo'shyotganda, birinchi qatorni reference qilish
		}
		if refIndex >= 0 && refIndex < len(t.rows) && t.rows[refIndex] != nil {
			cols = len(t.rows[refIndex].cells)
		}
	}

	if cols <= 0 {
		cols = len(t.grid)
	}
	if cols <= 0 {
		cols = 1
	}

	t.ensureGridLength(cols)
	if cols > t.gridColumns {
		t.gridColumns = cols
	}

	row := &TableRow{
		table: t,
		cells: make([]*TableCell, cols),
	}
	for i := 0; i < cols; i++ {
		width := 1440
		if i < len(t.grid) && t.grid[i] > 0 {
			width = t.grid[i]
		}
		cell := &TableCell{
			row:        row,
			paragraphs: []*Paragraph{NewParagraph()},
			width:      width,
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
		if len(t.borders) == 0 {
			t.bordersDefined = false
		}
		return
	}
	copy := border
	t.borders[side] = &copy
	t.bordersDefined = true
}

// SetWidth sets the table width in twentieths of a point (0 for auto width)
func (t *Table) SetWidth(width int) {
	if width <= 0 {
		t.width = 0
		t.widthType = "auto"
		return
	}
	t.width = width
	t.widthType = "dxa"
}

// Width returns the table width in twentieths of a point (0 means auto)
func (t *Table) Width() int {
	return t.width
}

// SetWidthWithType sets the table width value with an explicit width type (e.g. dxa, pct, auto).
func (t *Table) SetWidthWithType(width int, typ string) {
	t.width = width
	if typ == "" {
		if width <= 0 {
			t.widthType = "auto"
			t.width = 0
			return
		}
		t.widthType = "dxa"
		return
	}
	t.widthType = typ
}

// WidthType returns the stored table width type string.
func (t *Table) WidthType() string {
	if t.widthType == "" {
		if t.width <= 0 {
			return "auto"
		}
		return "dxa"
	}
	return t.widthType
}

// SetIndent sets the table indentation in twentieths of a point with the given type (e.g. "dxa").
func (t *Table) SetIndent(value int, typ string) {
	t.indent = value
	t.indentType = typ
	t.indentSet = true
	if t.indentType == "" {
		t.indentType = "dxa"
	}
}

// Indent returns the table indentation value and type if set.
func (t *Table) Indent() (int, string, bool) {
	if !t.indentSet {
		return 0, "", false
	}
	typ := t.indentType
	if typ == "" {
		typ = "dxa"
	}
	return t.indent, typ, true
}

// ClearIndent removes any table indentation settings.
func (t *Table) ClearIndent() {
	t.indent = 0
	t.indentType = ""
	t.indentSet = false
}

// SetStyle applies a table style by id.
func (t *Table) SetStyle(style string) {
	t.style = style
}

// Style returns the currently assigned table style id.
func (t *Table) Style() string {
	return t.style
}

// ClearStyle removes any table style association.
func (t *Table) ClearStyle() {
	t.style = ""
}

// SetLayout stores the table layout (e.g. "fixed" or "autofit").
func (t *Table) SetLayout(layout string) {
	t.layout = layout
}

// Layout returns the stored table layout value.
func (t *Table) Layout() string {
	return t.layout
}

// ClearLayout removes any explicit table layout setting.
func (t *Table) ClearLayout() {
	t.layout = ""
}

// SetAlignment configures the table justification.
func (t *Table) SetAlignment(alignment TableAlignment) {
	t.alignment = alignment
}

// Alignment returns the currently stored table justification.
func (t *Table) Alignment() TableAlignment {
	return t.alignment
}

// ClearAlignment removes any table justification setting.
func (t *Table) ClearAlignment() {
	t.alignment = ""
}

// SetLook stores the look flags associated with the table.
func (t *Table) SetLook(look TableLook) {
	copy := look
	t.look = &copy
}

// Look returns the table look configuration when present.
func (t *Table) Look() (*TableLook, bool) {
	if t.look == nil {
		return nil, false
	}
	return t.look, true
}

// ClearLook removes any table look information.
func (t *Table) ClearLook() {
	t.look = nil
}

// Border returns the configured border for the table on the specified side.
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
	t.bordersDefined = false
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

	if t.style != "" {
		builder.WriteString(fmt.Sprintf(`<w:tblStyle w:val="%s"/>`, xmlEscapeAttribute(t.style)))
	}

	widthType := t.WidthType()
	widthVal := t.width
	if widthType == "auto" {
		widthVal = 0
	}
	builder.WriteString(fmt.Sprintf(`<w:tblW w:w="%d" w:type="%s"/>`, widthVal, xmlEscapeAttribute(widthType)))

	if t.indentSet {
		typeAttr := t.indentType
		if typeAttr == "" {
			typeAttr = "dxa"
		}
		builder.WriteString(fmt.Sprintf(`<w:tblInd w:w="%d" w:type="%s"/>`, t.indent, typeAttr))
	}

	if t.alignment != "" {
		builder.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, xmlEscapeAttribute(string(t.alignment))))
	}

	if t.bordersDefined || len(t.borders) > 0 {
		builder.WriteString(t.bordersXML())
	}

	if t.layout != "" {
		builder.WriteString(fmt.Sprintf(`<w:tblLayout w:type="%s"/>`, xmlEscapeAttribute(t.layout)))
	}

	if t.look != nil {
		builder.WriteString(tableLookElement(t.look))
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

func (t *Table) bordersXML() string {
	if len(t.borders) == 0 {
		return "<w:tblBorders/>"
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
	size := border.Size
	if size < 0 {
		size = 0
	}
	attrs = append(attrs, fmt.Sprintf(`w:sz="%d"`, size))
	space := border.Space
	if space < 0 {
		space = 0
	}
	attrs = append(attrs, fmt.Sprintf(`w:space="%d"`, space))
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

func tableLookElement(look *TableLook) string {
	if look == nil {
		return ""
	}
	attrs := []string{fmt.Sprintf(`w:val="%s"`, xmlEscapeAttribute(look.Val))}
	attrs = append(attrs, fmt.Sprintf(`w:firstRow="%s"`, boolBinaryAttr(look.FirstRow)))
	attrs = append(attrs, fmt.Sprintf(`w:lastRow="%s"`, boolBinaryAttr(look.LastRow)))
	attrs = append(attrs, fmt.Sprintf(`w:firstColumn="%s"`, boolBinaryAttr(look.FirstColumn)))
	attrs = append(attrs, fmt.Sprintf(`w:lastColumn="%s"`, boolBinaryAttr(look.LastColumn)))
	attrs = append(attrs, fmt.Sprintf(`w:noHBand="%s"`, boolBinaryAttr(look.NoHBand)))
	attrs = append(attrs, fmt.Sprintf(`w:noVBand="%s"`, boolBinaryAttr(look.NoVBand)))
	return fmt.Sprintf(`<w:tblLook %s/>`, strings.Join(attrs, " "))
}

func boolBinaryAttr(value bool) string {
	if value {
		return "1"
	}
	return "0"
}

func (t *Table) tblGridXML() string {
	widths := t.columnWidths()
	if len(widths) == 0 {
		return ""
	}

	var builder strings.Builder
	builder.WriteString("<w:tblGrid>")
	for _, w := range widths {
		if w < 0 {
			w = 0
		}
		builder.WriteString(fmt.Sprintf(`<w:gridCol w:w="%d"/>`, w))
	}
	builder.WriteString("</w:tblGrid>")
	return builder.String()
}

func (t *Table) ensureGridLength(cols int) {
	if cols <= 0 {
		return
	}
	if len(t.grid) >= cols {
		return
	}
	missing := cols - len(t.grid)
	for i := 0; i < missing; i++ {
		t.grid = append(t.grid, 1440)
	}
}

// SetColumnWidths updates the table grid column widths (twentieths of a point).
// Existing rows will have their cell widths adjusted to match the new grid.
func (t *Table) SetColumnWidths(widths ...int) {
	if len(widths) == 0 {
		t.grid = nil
		t.gridColumns = 0
		return
	}

	// Defensive copy to avoid external mutation
	if t.grid == nil || cap(t.grid) < len(widths) {
		t.grid = append([]int(nil), widths...)
	} else {
		t.grid = t.grid[:len(widths)]
		copy(t.grid, widths)
	}

	t.gridColumns = len(widths)

	for _, row := range t.rows {
		col := 0
		for _, cell := range row.cells {
			span := cell.GridSpan()
			if span < 1 {
				span = 1
			}
			if col >= len(widths) {
				break
			}
			total := 0
			for i := 0; i < span && col+i < len(widths); i++ {
				total += widths[col+i]
			}
			if total > 0 {
				cell.width = total
			}
			col += span
		}
	}
}

func (t *Table) columnWidths() []int {
	if len(t.grid) > 0 {
		res := make([]int, len(t.grid))
		copy(res, t.grid)
		return res
	}

	if t.gridColumns <= 0 {
		return nil
	}

	widths := make([]int, t.gridColumns)
	filled := make([]bool, t.gridColumns)

	for _, row := range t.rows {
		col := 0
		for _, cell := range row.cells {
			span := cell.GridSpan()
			if span < 1 {
				span = 1
			}
			total := cell.Width()
			if total <= 0 {
				col += span
				continue
			}

			per := total / span
			remainder := 0
			if span > 0 {
				remainder = total % span
			}

			for i := 0; i < span && col < len(widths); i++ {
				w := per
				if remainder > 0 {
					w++
					remainder--
				}
				if !filled[col] && w > 0 {
					widths[col] = w
					filled[col] = true
				}
				col++
			}
		}

		complete := true
		for _, f := range filled {
			if !f {
				complete = false
				break
			}
		}
		if complete {
			break
		}
	}

	for i := range widths {
		if widths[i] == 0 {
			widths[i] = 1440
		}
	}

	return widths
}

// ToXML converts the table to WordprocessingML XML
func (t *Table) ToXML() string {
	var rowsXML strings.Builder

	for _, row := range t.rows {
		rowsXML.WriteString(row.ToXML())
	}

	return fmt.Sprintf(`<w:tbl>%s%s%s</w:tbl>`, t.tblPropertiesXML(), t.tblGridXML(), rowsXML.String())
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

// GetCell is an alias for Cell() - returns the cell at the specified index
func (tr *TableRow) GetCell(index int) *TableCell {
	return tr.Cell(index)
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

// ClearParagraphs removes all paragraphs from the cell
func (tc *TableCell) ClearParagraphs() {
	tc.paragraphs = nil
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

// SetVerticalAlignment sets the vertical alignment for the cell content.
func (tc *TableCell) SetVerticalAlignment(alignment WDVerticalAlignment) {
	tc.verticalAlign = alignment
}

// VerticalAlignment returns the vertical alignment for the cell content.
func (tc *TableCell) VerticalAlignment() WDVerticalAlignment {
	return tc.verticalAlign
}

// ClearVerticalAlignment resets the vertical alignment to default (top).
func (tc *TableCell) ClearVerticalAlignment() {
	tc.verticalAlign = WDVerticalAlignmentTop
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
	if tc.verticalAlign != WDVerticalAlignmentTop {
		switch tc.verticalAlign {
		case WDVerticalAlignmentCenter:
			builder.WriteString(`<w:vAlign w:val="center"/>`)
		case WDVerticalAlignmentBottom:
			builder.WriteString(`<w:vAlign w:val="bottom"/>`)
		}
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
