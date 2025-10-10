package docx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// DocumentPart represents the main document part of a Word document
type documentElement struct {
	paragraph *Paragraph
	table     *Table
}

type DocumentPart struct {
	*Part
	pkg            *Package
	paragraphs     []*Paragraph
	tables         []*Table
	sections       []*Section
	bodyElements   []documentElement
	drawingCounter int
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
				dp.sections = append(dp.sections, NewSection(SectionStartContinuous))
				if err := skipElement(decoder, t); err != nil {
					return fmt.Errorf("failed to skip sectPr: %w", err)
				}
			}
		}
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
			case "pPr", "rPr":
				// containers
			case "pStyle":
				if style := attrValue(t.Attr, "val"); style != "" {
					paragraph.SetStyle(style)
				}
			case "jc":
				if align := attrValue(t.Attr, "val"); align != "" {
					paragraph.SetAlignment(mapParagraphAlignment(align))
				}
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
					}
				}
				if val := attrValue(t.Attr, "after"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.spacingAfter = v
					}
				}
				if val := attrValue(t.Attr, "line"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.spacingLine = v
					}
				}
				if val := attrValue(t.Attr, "lineRule"); val != "" {
					paragraph.spacingLineRule = val
				}
				if err := skipElement(decoder, t); err != nil {
					return nil, err
				}
			case "ind":
				if val := attrValue(t.Attr, "left"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.indentLeft = v
					}
				}
				if val := attrValue(t.Attr, "right"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.indentRight = v
					}
				}
				if val := attrValue(t.Attr, "firstLine"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.indentFirstLine = v
					}
				}
				if val := attrValue(t.Attr, "hanging"); val != "" {
					if v, err := strconv.Atoi(val); err == nil {
						paragraph.indentHanging = v
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
	table := &Table{rows: make([]*TableRow, 0), owner: dp}

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
			case "tblPr", "tblGrid":
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

func (dp *DocumentPart) updateXMLData() {
	var bodyContent strings.Builder

	for _, element := range dp.bodyElements {
		if element.paragraph != nil {
			bodyContent.WriteString(element.paragraph.ToXML())
		} else if element.table != nil {
			bodyContent.WriteString(element.table.ToXML())
		}
	}

	if len(dp.sections) > 0 {
		for _, section := range dp.sections {
			bodyContent.WriteString(section.ToXML())
		}
	} else {
		bodyContent.WriteString(NewSection(SectionStartContinuous).ToXML())
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
