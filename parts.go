package docx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"strconv"
	"strings"
)

// DocumentPart represents the main document part of a Word document
type documentElement struct {
	paragraph *Paragraph
	table     *Table
	section   *Section
}

type DocumentPart struct {
	*Part
	pkg            *Package
	paragraphs     []*Paragraph
	tables         []*Table
	sections       []*Section
	bodyElements   []documentElement
	drawingCounter int
	headers        []*Header
	footers        []*Footer
	headerByRelID  map[string]*Header
	footerByRelID  map[string]*Footer
	headerByTarget map[string]*Header
	footerByTarget map[string]*Footer
}

// NewDocumentPart creates a new document part
func NewDocumentPart() *DocumentPart {
	// Create default document XML
	docXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`

	part := &Part{
		URI:         "word/document.xml",
		ContentType: ContentTypeWMLDocumentMain,
		Data:        []byte(docXML),
	}

	dp := &DocumentPart{
		Part: part,
	}
	_ = dp.loadFromXML()
	return dp
}

func (dp *DocumentPart) loadFromXML() error {
	dp.paragraphs = make([]*Paragraph, 0)
	dp.tables = make([]*Table, 0)
	dp.sections = make([]*Section, 0)
	dp.bodyElements = make([]documentElement, 0)
	dp.drawingCounter = 0
	dp.headers = make([]*Header, 0)
	dp.footers = make([]*Footer, 0)
	dp.headerByRelID = make(map[string]*Header)
	dp.footerByRelID = make(map[string]*Footer)
	dp.headerByTarget = make(map[string]*Header)
	dp.footerByTarget = make(map[string]*Footer)

	if dp.Part == nil || len(dp.Part.Data) == 0 {
		return nil
	}

	decoder := xml.NewDecoder(bytes.NewReader(dp.Part.Data))
	decoder.Strict = false

	for {
		tok, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "p":
				paragraph, err := parseParagraph(decoder, t, dp)
				if err != nil {
					return fmt.Errorf("failed to parse paragraph: %w", err)
				}
				dp.paragraphs = append(dp.paragraphs, paragraph)
				dp.bodyElements = append(dp.bodyElements, documentElement{paragraph: paragraph})
			case "tbl":
				table, err := parseTable(decoder, t, dp)
				if err != nil {
					return fmt.Errorf("failed to parse table: %w", err)
				}
				dp.tables = append(dp.tables, table)
				dp.bodyElements = append(dp.bodyElements, documentElement{table: table})
			case "sectPr":
				section, err := parseSectionProperties(decoder, t, dp)
				if err != nil {
					return fmt.Errorf("failed to parse section: %w", err)
				}
				dp.sections = append(dp.sections, section)
				dp.bodyElements = append(dp.bodyElements, documentElement{section: section})
			}
		}
	}

	if len(dp.sections) == 0 {
		section := NewSection(SectionStartContinuous)
		section.setOwner(dp)
		dp.sections = append(dp.sections, section)
	}

	return nil
}

func parseParagraph(decoder *xml.Decoder, start xml.StartElement, dp *DocumentPart) (*Paragraph, error) {
	paragraph := NewParagraph()
	paragraph.owner = dp

	var (
		currentRun      *Run
		textBuffer      strings.Builder
		inText          bool
		hyperlinkURL    string
		hyperlinkAnchor string
	)

	applyHyperlinkContext := func(run *Run) {
		if run == nil {
			return
		}
		run.owner = dp
		if hyperlinkURL != "" {
			run.SetHyperlink(hyperlinkURL)
		} else if hyperlinkAnchor != "" {
			run.SetHyperlinkAnchor(hyperlinkAnchor)
		}
	}

	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pPr":
				// handle paragraph properties including potential sectPr nested inside pPr
			case "rPr":
				// run properties container
			case "pStyle":
				if style := attrValue(t.Attr, "val"); style != "" {
					paragraph.SetStyle(style)
				}
			case "jc":
				if align := attrValue(t.Attr, "val"); align != "" {
					paragraph.SetAlignment(mapParagraphAlignment(align))
				}
			case "sectPr":
				// Paragraph-level section break; parse it and attach to this paragraph
				sect, err := parseSectionProperties(decoder, t, dp)
				if err != nil {
					return nil, err
				}
				paragraph.section = sect
				continue
			case "pBdr":
				borders, err := parseParagraphBorders(decoder, t)
				if err != nil {
					return nil, err
				}
				for side, border := range borders {
					if border != nil {
						paragraph.SetBorder(side, *border)
					}
				}
			case "spacing":
				if currentRun != nil {
					if val := attrValue(t.Attr, "val"); val != "" {
						if v, err := strconv.Atoi(val); err == nil {
							currentRun.SetCharacterSpacing(v)
						}
					}
					if err := skipElement(decoder, t); err != nil {
						return nil, err
					}
					break
				}
				if val := attrValue(t.Attr, "before"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.spacingBefore = v
						paragraph.spacingBeforeSet = true
					}
				}
				if val := attrValue(t.Attr, "after"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.spacingAfter = v
						paragraph.spacingAfterSet = true
					}
				}
				if val := attrValue(t.Attr, "line"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.spacingLine = v
						paragraph.spacingLineSet = true
					}
				}
				if val := attrValue(t.Attr, "lineRule"); val != "" {
					paragraph.spacingLineRule = val
					paragraph.spacingLineRuleSet = true
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "ind":
				if val := attrValue(t.Attr, "left"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.indentLeft = v
						paragraph.indentLeftSet = true
					}
				}
				if val := attrValue(t.Attr, "right"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.indentRight = v
						paragraph.indentRightSet = true
					}
				}
				if val := attrValue(t.Attr, "firstLine"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.indentFirstLine = v
						paragraph.indentFirstLineSet = true
					}
				}
				if val := attrValue(t.Attr, "hanging"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.indentHanging = v
						paragraph.indentHangingSet = true
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "numPr":
				if err := parseParagraphNumbering(decoder, paragraph); err != nil {
					return nil, err
				}
			case "keepNext":
				paragraph.keepWithNext = parseOnOff(t.Attr)
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "keepLines":
				paragraph.keepLines = parseOnOff(t.Attr)
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "pageBreakBefore":
				paragraph.pageBreakBefore = parseOnOff(t.Attr)
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "widowControl":
				paragraph.widowControl = parseOnOff(t.Attr)
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "tabs":
				stops, err := parseParagraphTabs(decoder, t)
				if err != nil {
					return nil, err
				}
				for _, stop := range stops {
					paragraph.AddTabStop(stop.Position, stop.Alignment, stop.Leader)
				}
			case "hyperlink":
				hyperlinkURL = ""
				hyperlinkAnchor = attrValue(t.Attr, "anchor")
				if relID := attrValue(t.Attr, "id"); relID != "" && dp != nil {
					if target, mode, ok := dp.relationshipTarget(relID); ok {
						if strings.EqualFold(mode, "External") {
							hyperlinkURL = target
						} else if hyperlinkAnchor == "" {
							hyperlinkAnchor = target
						}
					}
				}
				// Continue parsing child runs within the hyperlink
			case "r":
				currentRun = NewRun("")
				applyHyperlinkContext(currentRun)
			case "t":
				textBuffer.Reset()
				inText = true
			case "b":
				if currentRun != nil {
					currentRun.SetBold(true)
				}
			case "i":
				if currentRun != nil {
					currentRun.SetItalic(true)
				}
			case "strike":
				if currentRun != nil {
					currentRun.SetStrikethrough(true)
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "dstrike":
				if currentRun != nil {
					currentRun.SetDoubleStrikethrough(true)
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "smallCaps":
				if currentRun != nil {
					currentRun.SetSmallCaps(true)
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "caps":
				if currentRun != nil {
					currentRun.SetAllCaps(true)
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "shadow":
				if currentRun != nil {
					currentRun.SetShadow(true)
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "outline":
				if currentRun != nil {
					currentRun.SetOutline(true)
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "emboss":
				if currentRun != nil {
					currentRun.SetEmboss(true)
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "imprint":
				if currentRun != nil {
					currentRun.SetImprint(true)
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "u":
				if currentRun != nil {
					underline := attrValue(t.Attr, "val")
					if underline == "" {
						underline = string(WDUnderlineSingle)
					}
					currentRun.SetUnderline(WDUnderline(underline))
				}
			case "color":
				if currentRun != nil {
					if val := attrValue(t.Attr, "val"); val != "" {
						currentRun.SetColor(val)
					}
				}
			case "rFonts":
				if currentRun != nil {
					font := attrValue(t.Attr, "ascii")
					if font == "" {
						font = attrValue(t.Attr, "hAnsi")
					}
					if font != "" {
						currentRun.SetFont(font)
					}
				}
			case "sz":
				if currentRun != nil {
					if val := attrValue(t.Attr, "val"); val != "" {
						if sz, err := strconv.Atoi(val); err == nil {
							currentRun.setSizeRaw(sz)
						}
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "szCs":
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "highlight":
				if currentRun != nil {
					if val := attrValue(t.Attr, "val"); val != "" {
						currentRun.SetHighlight(WDColorIndex(val))
					}
				}
			case "shd":
				if currentRun == nil {
					paragraph.SetShading(attrValue(t.Attr, "val"), attrValue(t.Attr, "fill"), attrValue(t.Attr, "color"))
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
				continue
			case "kern":
				if currentRun != nil {
					if val := attrValue(t.Attr, "val"); val != "" {
						if v, err := strconv.Atoi(val); err == nil {
							currentRun.SetKerning(v)
						}
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "position":
				if currentRun != nil {
					if val := attrValue(t.Attr, "val"); val != "" {
						if v, err := strconv.Atoi(val); err == nil {
							currentRun.SetBaselineShift(v)
						}
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "br":
				if currentRun == nil {
					currentRun = NewRun("")
					applyHyperlinkContext(currentRun)
				}
				currentRun.AddBreak(mapBreakType(attrValue(t.Attr, "type")))
			case "drawing":
				if currentRun == nil {
					currentRun = NewRun("")
					applyHyperlinkContext(currentRun)
				}
				picture, err := parseDrawing(decoder, t, dp)
				if err != nil {
					return nil, err
				}
				if picture != nil {
					currentRun.picture = picture
				}
			default:
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			}
		case xml.CharData:
			if inText && currentRun != nil {
				textBuffer.Write([]byte(t))
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "t":
				if currentRun != nil {
					existing := currentRun.Text()
					currentRun.SetText(existing + textBuffer.String())
				}
				inText = false
			case "r":
				if currentRun != nil {
					paragraph.runs = append(paragraph.runs, currentRun)
				}
				currentRun = nil
			case "hyperlink":
				hyperlinkURL = ""
				hyperlinkAnchor = ""
			case "p":
				return paragraph, nil
			}
		}
	}
}

func parseDrawing(decoder *xml.Decoder, start xml.StartElement, dp *DocumentPart) (*Picture, error) {
	picture := &Picture{docPart: dp}
	depth := 1

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "extent":
				if val := attrValue(t.Attr, "cx"); val != "" {
					if cx, err := strconv.ParseInt(val, 10, 64); err == nil {
						picture.widthEMU = cx
					}
				}
				if val := attrValue(t.Attr, "cy"); val != "" {
					if cy, err := strconv.ParseInt(val, 10, 64); err == nil {
						picture.heightEMU = cy
					}
				}
			case "docPr":
				if val := attrValue(t.Attr, "id"); val != "" {
					if id, err := strconv.Atoi(val); err == nil {
						picture.docPrID = id
					}
				}
				if name := attrValue(t.Attr, "name"); name != "" {
					picture.name = name
				}
				if descr := attrValue(t.Attr, "descr"); descr != "" {
					picture.description = descr
				}
			case "blip":
				if relID := attrValue(t.Attr, "embed"); relID != "" {
					picture.relID = relID
				}
			}
		case xml.EndElement:
			depth--
			if depth == 0 {
				break
			}
		}
	}

	if picture.relID != "" && dp != nil {
		if target, _, ok := dp.relationshipTarget(picture.relID); ok {
			picture.target = target
		}
	}

	if dp != nil && picture.docPrID > dp.drawingCounter {
		dp.drawingCounter = picture.docPrID
	}

	return picture, nil
}

func parseParagraphBorders(decoder *xml.Decoder, start xml.StartElement) (map[ParagraphBorderSide]*ParagraphBorder, error) {
	borders := make(map[ParagraphBorderSide]*ParagraphBorder)

	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			side := ParagraphBorderSide(t.Name.Local)
			if side != ParagraphBorderTop && side != ParagraphBorderLeft && side != ParagraphBorderBottom && side != ParagraphBorderRight && side != ParagraphBorderBetween && side != ParagraphBorderBar {
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
				continue
			}
			border := &ParagraphBorder{
				Style: attrValue(t.Attr, "val"),
				Color: attrValue(t.Attr, "color"),
			}
			if border.Style == "" {
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
				continue
			}
			if sz := attrValue(t.Attr, "sz"); sz != "" {
				if v, err := strconv.Atoi(sz); err == nil {
					border.Size = v
				}
			}
			if space := attrValue(t.Attr, "space"); space != "" {
				if v, err := strconv.Atoi(space); err == nil {
					border.Space = v
				}
			}
			if shadow := attrValue(t.Attr, "shadow"); shadow != "" {
				border.Shadow = strings.EqualFold(shadow, "1") || strings.EqualFold(shadow, "true")
			}
			borders[side] = border
			if err := skipElement(decoder, t); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return borders, nil
			}
		}
	}
}

func parseParagraphTabs(decoder *xml.Decoder, start xml.StartElement) ([]TabStop, error) {
	stops := make([]TabStop, 0)

	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tab":
				stop := TabStop{
					Alignment: mapTabAlignment(attrValue(t.Attr, "val")),
					Leader:    mapTabLeader(attrValue(t.Attr, "leader")),
				}
				if posStr := attrValue(t.Attr, "pos"); posStr != "" {
					if pos, err := strconv.Atoi(posStr); err == nil {
						stop.Position = pos
					}
				}
				stops = append(stops, stop)
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			default:
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return stops, nil
			}
		}
	}
}

func parseParagraphNumbering(decoder *xml.Decoder, paragraph *Paragraph) error {
	var numIDSet, levelSet bool
	var numID, level int

	for {
		tok, err := decoder.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "ilvl":
				if val := attrValue(t.Attr, "val"); val != "" {
					if lvl, err := strconv.Atoi(val); err == nil {
						level = lvl
						levelSet = true
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			case "numId":
				if val := attrValue(t.Attr, "val"); val != "" {
					if id, err := strconv.Atoi(val); err == nil {
						numID = id
						numIDSet = true
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			default:
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "numPr" {
				if numIDSet {
					if !levelSet {
						level = 0
					}
					paragraph.SetNumbering(numID, level)
				}
				return nil
			}
		}
	}
}

func parseTable(decoder *xml.Decoder, start xml.StartElement, dp *DocumentPart) (*Table, error) {
	table := &Table{rows: make([]*TableRow, 0), owner: dp, borders: make(map[TableBorderSide]*TableBorder)}

	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tr":
				row, err := parseTableRow(decoder, t, table, dp)
				if err != nil {
					return nil, err
				}
				table.rows = append(table.rows, row)
				if len(row.cells) > table.gridColumns {
					table.gridColumns = len(row.cells)
				}
			case "tblPr":
				if err := parseTableProperties(decoder, t, table); err != nil {
					return nil, err
				}
			case "tblGrid":
				widths, err := parseTableGrid(decoder, t)
				if err != nil {
					return nil, err
				}
				if len(widths) > 0 {
					table.grid = append([]int(nil), widths...)
					if len(widths) > table.gridColumns {
						table.gridColumns = len(widths)
					}
				}
			default:
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				if table.gridColumns == 0 && len(table.rows) > 0 {
					table.gridColumns = len(table.rows[0].cells)
				}
				table.setOwner(dp)
				return table, nil
			}
		}
	}
}

func parseTableRow(decoder *xml.Decoder, start xml.StartElement, table *Table, dp *DocumentPart) (*TableRow, error) {
	row := &TableRow{table: table, cells: make([]*TableCell, 0)}

	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tc":
				cell, err := parseTableCell(decoder, t, row, dp)
				if err != nil {
					return nil, err
				}
				row.cells = append(row.cells, cell)
			case "trPr":
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			default:
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return row, nil
			}
		}
	}
}

func parseTableCell(decoder *xml.Decoder, start xml.StartElement, row *TableRow, dp *DocumentPart) (*TableCell, error) {
	cell := &TableCell{
		row:        row,
		paragraphs: make([]*Paragraph, 0),
		width:      1440,
		borders:    make(map[TableBorderSide]*TableBorder),
	}

	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tcPr":
				if err := parseTableCellProperties(decoder, t, cell); err != nil {
					return nil, err
				}
			case "p":
				paragraph, err := parseParagraph(decoder, t, dp)
				if err != nil {
					return nil, err
				}
				cell.paragraphs = append(cell.paragraphs, paragraph)
			case "tbl":
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			default:
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				if len(cell.paragraphs) == 0 {
					paragraph := NewParagraph()
					paragraph.owner = dp
					cell.paragraphs = []*Paragraph{paragraph}
				}
				return cell, nil
			}
		}
	}
}

func parseTableCellProperties(decoder *xml.Decoder, start xml.StartElement, cell *TableCell) error {
	for {
		tok, err := decoder.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tcW":
				if val := attrValue(t.Attr, "w"); val != "" {
					if w, err := strconv.Atoi(val); err == nil {
						cell.width = w
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			case "gridSpan":
				if val := attrValue(t.Attr, "val"); val != "" {
					if span, err := strconv.Atoi(val); err == nil {
						cell.gridSpan = span
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			case "vMerge":
				state := attrValue(t.Attr, "val")
				switch state {
				case "restart":
					cell.verticalMerge = TableVerticalMergeRestart
				case "continue":
					cell.verticalMerge = TableVerticalMergeContinue
				default:
					if state == "" {
						cell.verticalMerge = TableVerticalMergeContinue
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			case "vAlign":
				val := attrValue(t.Attr, "val")
				switch val {
				case "center":
					cell.verticalAlign = WDVerticalAlignmentCenter
				case "bottom":
					cell.verticalAlign = WDVerticalAlignmentBottom
				default:
					cell.verticalAlign = WDVerticalAlignmentTop
				}
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			case "tcBorders":
				borders, err := parseTableBorders(decoder, t)
				if err != nil {
					return err
				}
				for side, border := range borders {
					if border != nil {
						cell.SetBorder(side, *border)
					}
				}
			case "shd":
				cell.SetShading(attrValue(t.Attr, "val"), attrValue(t.Attr, "fill"), attrValue(t.Attr, "color"))
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			default:
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return nil
			}
		}
	}
}

func parseSectionProperties(decoder *xml.Decoder, start xml.StartElement, dp *DocumentPart) (*Section, error) {
	section := NewSection(SectionStartContinuous)
	section.setOwner(dp)

	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "type":
				if val := attrValue(t.Attr, "val"); val != "" {
					section.startType = SectionStartType(val)
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "pgSz":
				if val := attrValue(t.Attr, "w"); val != "" {
					if w, err := strconv.Atoi(val); err == nil {
						section.pageWidth = w
					}
				}
				if val := attrValue(t.Attr, "h"); val != "" {
					if h, err := strconv.Atoi(val); err == nil {
						section.pageHeight = h
					}
				}
				if val := attrValue(t.Attr, "orient"); val != "" {
					section.orientation = val
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "pgMar":
				if val := attrValue(t.Attr, "top"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						section.marginTop = v
					}
				}
				if val := attrValue(t.Attr, "right"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						section.marginRight = v
					}
				}
				if val := attrValue(t.Attr, "bottom"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						section.marginBottom = v
					}
				}
				if val := attrValue(t.Attr, "left"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						section.marginLeft = v
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "headerReference":
				typeVal := HeaderType(attrValue(t.Attr, "type"))
				if typeVal == "" {
					typeVal = HeaderTypeDefault
				}
				relID := attrValue(t.Attr, "id")
				if relID != "" && dp != nil {
					header, err := dp.headerFromRelationship(relID)
					if err != nil {
						return nil, err
					}
					if header != nil {
						section.headerRefs[typeVal] = &headerReference{typeValue: typeVal, relID: relID, header: header}
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "footerReference":
				typeVal := FooterType(attrValue(t.Attr, "type"))
				if typeVal == "" {
					typeVal = FooterTypeDefault
				}
				relID := attrValue(t.Attr, "id")
				if relID != "" && dp != nil {
					footer, err := dp.footerFromRelationship(relID)
					if err != nil {
						return nil, err
					}
					if footer != nil {
						section.footerRefs[typeVal] = &footerReference{typeValue: typeVal, relID: relID, footer: footer}
					}
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			default:
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return section, nil
			}
		}
	}
}

func (dp *DocumentPart) headerFromRelationship(relID string) (*Header, error) {
	if header, ok := dp.headerByRelID[relID]; ok {
		return header, nil
	}
	if dp.pkg == nil {
		return nil, fmt.Errorf("document part is not associated with a package")
	}
	target, _, ok := dp.relationshipTarget(relID)
	if !ok {
		return nil, fmt.Errorf("relationship %s not found", relID)
	}
	fullPath := resolveRelationshipTarget(dp.Part.URI, target)
	if header, ok := dp.headerByTarget[fullPath]; ok {
		dp.headerByRelID[relID] = header
		return header, nil
	}
	part, exists := dp.pkg.parts[fullPath]
	if !exists {
		return nil, fmt.Errorf("header part %s not found", fullPath)
	}
	header := newHeader(dp, part)
	if err := header.loadFromXML(); err != nil {
		return nil, fmt.Errorf("failed to load header %s: %w", fullPath, err)
	}
	dp.headers = append(dp.headers, header)
	dp.headerByTarget[fullPath] = header
	dp.headerByRelID[relID] = header
	return header, nil
}

func (dp *DocumentPart) footerFromRelationship(relID string) (*Footer, error) {
	if footer, ok := dp.footerByRelID[relID]; ok {
		return footer, nil
	}
	if dp.pkg == nil {
		return nil, fmt.Errorf("document part is not associated with a package")
	}
	target, _, ok := dp.relationshipTarget(relID)
	if !ok {
		return nil, fmt.Errorf("relationship %s not found", relID)
	}
	fullPath := resolveRelationshipTarget(dp.Part.URI, target)
	if footer, ok := dp.footerByTarget[fullPath]; ok {
		dp.footerByRelID[relID] = footer
		return footer, nil
	}
	part, exists := dp.pkg.parts[fullPath]
	if !exists {
		return nil, fmt.Errorf("footer part %s not found", fullPath)
	}
	footer := newFooter(dp, part)
	if err := footer.loadFromXML(); err != nil {
		return nil, fmt.Errorf("failed to load footer %s: %w", fullPath, err)
	}
	dp.footers = append(dp.footers, footer)
	dp.footerByTarget[fullPath] = footer
	dp.footerByRelID[relID] = footer
	return footer, nil
}

func (dp *DocumentPart) createHeaderPart() (*Header, string, error) {
	if dp == nil || dp.pkg == nil {
		return nil, "", fmt.Errorf("document part is not associated with a package")
	}
	part := dp.pkg.newHeaderPart()
	header := newHeader(dp, part)
	dp.headers = append(dp.headers, header)
	dp.headerByTarget[part.URI] = header
	target := path.Base(part.URI)
	relID := dp.pkg.ensureRelationship(dp.Part.URI, RelTypeHeader, target)
	dp.headerByRelID[relID] = header
	return header, relID, nil
}

func (dp *DocumentPart) createFooterPart() (*Footer, string, error) {
	if dp == nil || dp.pkg == nil {
		return nil, "", fmt.Errorf("document part is not associated with a package")
	}
	part := dp.pkg.newFooterPart()
	footer := newFooter(dp, part)
	dp.footers = append(dp.footers, footer)
	dp.footerByTarget[part.URI] = footer
	target := path.Base(part.URI)
	relID := dp.pkg.ensureRelationship(dp.Part.URI, RelTypeFooter, target)
	dp.footerByRelID[relID] = footer
	return footer, relID, nil
}

func parseTableProperties(decoder *xml.Decoder, start xml.StartElement, table *Table) error {
	for {
		tok, err := decoder.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblBorders":
				borders, err := parseTableBorders(decoder, t)
				if err != nil {
					return err
				}
				for side, border := range borders {
					if border != nil {
						table.SetBorder(side, *border)
					}
				}
			case "tblCellMar":
				margins, err := parseTableCellMargins(decoder, t)
				if err != nil {
					return err
				}
				if margins != nil {
					table.cellMargins = margins
				}
			case "shd":
				table.SetShading(attrValue(t.Attr, "val"), attrValue(t.Attr, "fill"), attrValue(t.Attr, "color"))
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			case "tblW":
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			default:
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return nil
			}
		}
	}
}

func parseTableBorders(decoder *xml.Decoder, start xml.StartElement) (map[TableBorderSide]*TableBorder, error) {
	borders := make(map[TableBorderSide]*TableBorder)

	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			border := parseBorderAttributes(t.Attr)
			if border.Style != "" {
				side := TableBorderSide(t.Name.Local)
				copy := border
				borders[side] = &copy
			}
			if err := skipElement(decoder, t); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return borders, nil
			}
		}
	}
}

func parseBorderAttributes(attrs []xml.Attr) TableBorder {
	border := TableBorder{}
	for _, attr := range attrs {
		if attr.Name.Local == "val" {
			border.Style = attr.Value
		} else if attr.Name.Local == "color" {
			border.Color = attr.Value
		} else if attr.Name.Local == "sz" {
			if v, err := strconv.Atoi(attr.Value); err == nil {
				border.Size = v
			}
		} else if attr.Name.Local == "space" {
			if v, err := strconv.Atoi(attr.Value); err == nil {
				border.Space = v
			}
		}
	}
	return border
}

func parseTableCellMargins(decoder *xml.Decoder, start xml.StartElement) (*TableCellMargins, error) {
	margins := &TableCellMargins{}

	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			val := attrValue(t.Attr, "w")
			if val != "" {
				if v, err := strconv.Atoi(val); err == nil {
					switch t.Name.Local {
					case "top":
						margins.Top = intPtr(v)
					case "left":
						margins.Left = intPtr(v)
					case "bottom":
						margins.Bottom = intPtr(v)
					case "right":
						margins.Right = intPtr(v)
					}
				}
			}
			if err := skipElement(decoder, t); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return margins, nil
			}
		}
	}
}

func parseTableGrid(decoder *xml.Decoder, start xml.StartElement) ([]int, error) {
	widths := make([]int, 0)
	for {
		tok, err := decoder.Token()
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "gridCol" {
				width := 0
				if val := attrValue(t.Attr, "w"); val != "" {
					if w, err := strconv.Atoi(val); err == nil {
						width = w
					}
				}
				widths = append(widths, width)
			}
			if err := skipElement(decoder, t); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return widths, nil
			}
		}
	}
}

func skipElement(decoder *xml.Decoder, start xml.StartElement) error {
	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			return err
		}

		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return nil
}

func attrValue(attrs []xml.Attr, key string) string {
	for _, attr := range attrs {
		if attr.Name.Local == key {
			return attr.Value
		}
	}
	return ""
}

func parseOnOff(attrs []xml.Attr) *bool {
	val := strings.ToLower(attrValue(attrs, "val"))
	switch val {
	case "", "1", "true", "on":
		return boolPtr(true)
	case "0", "false", "off":
		return boolPtr(false)
	default:
		return boolPtr(true)
	}
}

func mapParagraphAlignment(val string) WDAlignParagraph {
	switch strings.ToLower(val) {
	case "center":
		return WDAlignParagraphCenter
	case "right":
		return WDAlignParagraphRight
	case "both", "justify":
		return WDAlignParagraphJustify
	case "distribute":
		return WDAlignParagraphDistribute
	default:
		return WDAlignParagraphLeft
	}
}

func mapTabAlignment(val string) WDTabAlignment {
	switch strings.ToLower(val) {
	case "center":
		return WDTabAlignmentCenter
	case "right":
		return WDTabAlignmentRight
	case "decimal":
		return WDTabAlignmentDecimal
	case "bar":
		return WDTabAlignmentBar
	default:
		return WDTabAlignmentLeft
	}
}

func mapTabLeader(val string) WDTabLeader {
	switch strings.ToLower(val) {
	case "dot":
		return WDTabLeaderDot
	case "hyphen":
		return WDTabLeaderHyphen
	case "underscore":
		return WDTabLeaderUnderscore
	case "heavy":
		return WDTabLeaderHeavy
	case "middledot":
		return WDTabLeaderMiddleDot
	default:
		return WDTabLeaderNone
	}
}

func mapBreakType(val string) BreakType {
	switch strings.ToLower(val) {
	case "page":
		return BreakTypePage
	case "column":
		return BreakTypeColumn
	default:
		return BreakTypeText
	}
}

// ContentType returns the content type of this part
func (dp *DocumentPart) ContentType() string {
	return dp.Part.ContentType
}

// AddParagraph adds a new paragraph to the document
func (dp *DocumentPart) AddParagraph(text ...string) *Paragraph {
	paragraph := NewParagraph()
	paragraph.owner = dp

	// Add text if provided
	for _, t := range text {
		paragraph.AddRun(t)
	}

	dp.paragraphs = append(dp.paragraphs, paragraph)
	dp.bodyElements = append(dp.bodyElements, documentElement{paragraph: paragraph})

	// Update the XML data
	dp.updateXMLData()

	return paragraph
}

// AddTable adds a new table to the document
func (dp *DocumentPart) AddTable(rows, cols int) *Table {
	table := NewTable(rows, cols)
	table.setOwner(dp)
	dp.tables = append(dp.tables, table)
	dp.bodyElements = append(dp.bodyElements, documentElement{table: table})

	// Update the XML data
	dp.updateXMLData()

	return table
}

// AddSection adds a new section to the document
func (dp *DocumentPart) AddSection(startType SectionStartType) *Section {
	section := NewSection(startType)
	section.setOwner(dp)
	dp.sections = append(dp.sections, section)

	// Update the XML data
	dp.updateXMLData()

	return section
}

// Paragraphs returns all paragraphs in the document
func (dp *DocumentPart) Paragraphs() []*Paragraph {
	return dp.paragraphs
}

// Tables returns all tables in the document
func (dp *DocumentPart) Tables() []*Table {
	return dp.tables
}

// Sections returns all sections in the document
func (dp *DocumentPart) Sections() []*Section {
	return dp.sections
}

// InsertTableAfterParagraph inserts a table immediately after the specified paragraph
func (dp *DocumentPart) InsertTableAfterParagraph(paragraph *Paragraph, rows, cols int) (*Table, error) {
	if paragraph == nil {
		return nil, fmt.Errorf("paragraph cannot be nil")
	}

	// Find the index of the paragraph in bodyElements
	paragraphIndex := -1
	for i, elem := range dp.bodyElements {
		if elem.paragraph == paragraph {
			paragraphIndex = i
			break
		}
	}

	if paragraphIndex == -1 {
		return nil, fmt.Errorf("paragraph not found in document")
	}

	// Create new table
	table := NewTable(rows, cols)
	table.setOwner(dp)

	// Insert the table after the paragraph in bodyElements
	dp.bodyElements = append(dp.bodyElements[:paragraphIndex+1],
		append([]documentElement{{table: table}}, dp.bodyElements[paragraphIndex+1:]...)...)

	// Add to tables list
	dp.tables = append(dp.tables, table)

	// Update the XML data
	dp.updateXMLData()

	return table, nil
}

// RemoveParagraph removes the specified paragraph from the document
func (dp *DocumentPart) RemoveParagraph(paragraph *Paragraph) error {
	if paragraph == nil {
		return fmt.Errorf("paragraph cannot be nil")
	}

	// Find and remove from paragraphs slice
	paragraphIndex := -1
	for i, p := range dp.paragraphs {
		if p == paragraph {
			paragraphIndex = i
			break
		}
	}

	if paragraphIndex == -1 {
		return fmt.Errorf("paragraph not found in document")
	}

	// Remove from paragraphs slice
	dp.paragraphs = append(dp.paragraphs[:paragraphIndex], dp.paragraphs[paragraphIndex+1:]...)

	// Find and remove from bodyElements
	for i, elem := range dp.bodyElements {
		if elem.paragraph == paragraph {
			dp.bodyElements = append(dp.bodyElements[:i], dp.bodyElements[i+1:]...)
			break
		}
	}

	// Update the XML data
	dp.updateXMLData()

	return nil
}

// RemoveTable removes the specified table from the document
func (dp *DocumentPart) RemoveTable(table *Table) error {
	if table == nil {
		return fmt.Errorf("table cannot be nil")
	}

	// Find and remove from tables slice
	tableIndex := -1
	for i, t := range dp.tables {
		if t == table {
			tableIndex = i
			break
		}
	}

	if tableIndex == -1 {
		return fmt.Errorf("table not found in document")
	}

	// Remove from tables slice
	dp.tables = append(dp.tables[:tableIndex], dp.tables[tableIndex+1:]...)

	// Find and remove from bodyElements
	for i, elem := range dp.bodyElements {
		if elem.table == table {
			dp.bodyElements = append(dp.bodyElements[:i], dp.bodyElements[i+1:]...)
			break
		}
	}

	// Update the XML data
	dp.updateXMLData()

	return nil
}

// RemoveSection removes the specified section from the document
func (dp *DocumentPart) RemoveSection(section *Section) error {
	if section == nil {
		return fmt.Errorf("section cannot be nil")
	}

	// Find and remove from sections slice
	sectionIndex := -1
	for i, s := range dp.sections {
		if s == section {
			sectionIndex = i
			break
		}
	}

	if sectionIndex == -1 {
		return fmt.Errorf("section not found in document")
	}

	// Remove from sections slice
	dp.sections = append(dp.sections[:sectionIndex], dp.sections[sectionIndex+1:]...)

	// Find and remove from bodyElements
	for i, elem := range dp.bodyElements {
		if elem.section == section {
			dp.bodyElements = append(dp.bodyElements[:i], dp.bodyElements[i+1:]...)
			break
		}
	}

	// Update the XML data
	dp.updateXMLData()

	return nil
}

func (dp *DocumentPart) updateXMLData() {
	var bodyContent strings.Builder

	hasSectionMarkers := false
	for _, element := range dp.bodyElements {
		if element.paragraph != nil {
			bodyContent.WriteString(element.paragraph.ToXML())
		} else if element.table != nil {
			bodyContent.WriteString(element.table.ToXML())
		} else if element.section != nil {
			bodyContent.WriteString(element.section.ToXML())
			hasSectionMarkers = true
		}
	}

	// Agar body ichida sektsiya belgilanmagan bo'lsa, oxirida kamida bitta sectPr yozamiz
	if !hasSectionMarkers {
		if len(dp.sections) > 0 {
			bodyContent.WriteString(dp.sections[len(dp.sections)-1].ToXML())
		} else {
			bodyContent.WriteString(NewSection(SectionStartContinuous).ToXML())
		}
	}

	docXML := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    %s
  </w:body>
</w:document>`, bodyContent.String())

	dp.Part.Data = []byte(docXML)
}

func (dp *DocumentPart) ensureHyperlinkRelationship(url string) string {
	if dp == nil || dp.pkg == nil {
		return ""
	}
	return dp.pkg.ensureRelationshipWithMode(dp.Part.URI, RelTypeHyperlink, url, "External")
}

func (dp *DocumentPart) relationshipTarget(relID string) (string, string, bool) {
	if dp == nil || dp.pkg == nil {
		return "", "", false
	}
	rels := dp.pkg.relations[dp.Part.URI]
	for _, rel := range rels {
		if rel.ID == relID {
			return rel.Target, rel.TargetMode, true
		}
	}
	return "", "", false
}

func (dp *DocumentPart) nextDrawingID() int {
	dp.drawingCounter++
	return dp.drawingCounter
}

// StylesPart represents the styles part of a Word document
type StylesPart struct {
	*Part
}

// NewStylesPart creates a new styles part
func NewStylesPart() *StylesPart {
	stylesXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri" w:cs="Times New Roman"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="276"/>
</w:styles>`

	part := &Part{
		URI:         "word/styles.xml",
		ContentType: ContentTypeWMLStyles,
		Data:        []byte(stylesXML),
	}

	return &StylesPart{Part: part}
}

// SettingsPart represents the settings part of a Word document
type SettingsPart struct {
	*Part
}

// NewSettingsPart creates a new settings part
func NewSettingsPart() *SettingsPart {
	settingsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="708"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>`

	part := &Part{
		URI:         "word/settings.xml",
		ContentType: ContentTypeWMLSettings,
		Data:        []byte(settingsXML),
	}

	return &SettingsPart{Part: part}
}
