package docx

const (
	defaultDecimalNumID = 1
	defaultBulletNumID  = 2
)

const defaultNumberingXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>`

// NumberingPart represents the numbering definitions part of a document
type NumberingPart struct {
	*Part
}

// NewNumberingPart creates a numbering part with default numbering definitions
func NewNumberingPart() *NumberingPart {
	part := &Part{
		URI:         "word/numbering.xml",
		ContentType: ContentTypeWMLNumbering,
		Data:        []byte(defaultNumberingXML),
	}

	return &NumberingPart{Part: part}
}

// Numbering provides helpers to access numbering information within a document
type Numbering struct {
	pkg  *Package
	part *Part
}

// NewNumbering creates a numbering helper for the given package
func NewNumbering(pkg *Package) *Numbering {
	n := &Numbering{pkg: pkg}
	if part, exists := pkg.parts["word/numbering.xml"]; exists {
		n.part = part
	}
	return n
}

func (n *Numbering) ensurePart() {
	if n.part != nil {
		return
	}

	numberingPart := NewNumberingPart()
	n.pkg.parts["word/numbering.xml"] = numberingPart.Part
	n.pkg.contentTypes["/word/numbering.xml"] = ContentTypeWMLNumbering
	n.pkg.ensureRelationship("word/document.xml", RelTypeNumbering, "numbering.xml")
	n.part = numberingPart.Part
}

func (n *Numbering) ensureDefault() {
	n.ensurePart()
	if len(n.part.Data) == 0 {
		n.part.Data = []byte(defaultNumberingXML)
	}
}

// DecimalListID returns the numbering ID for the default decimal list, creating it if necessary
func (n *Numbering) DecimalListID() int {
	n.ensureDefault()
	return defaultDecimalNumID
}

// BulletedListID returns the numbering ID for the default bullet list, creating it if necessary
func (n *Numbering) BulletedListID() int {
	n.ensureDefault()
	return defaultBulletNumID
}

// Part returns the underlying numbering part, ensuring it exists
func (n *Numbering) Part() *Part {
	n.ensureDefault()
	return n.part
}
