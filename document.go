package docx

// Package docx provides functionality for creating and manipulating Microsoft Word documents
// in the OpenXML (DOCX) format. This is a Go port of the popular python-docx library.
//
// The main entry point is the Document type which represents a Word document.
// You can create a new document or open an existing one:
//
//	// Create a new document
//	doc := docx.NewDocument()
//
//	// Add a paragraph
//	paragraph := doc.AddParagraph("Hello, World!")
//
//	// Save the document
//	err := doc.SaveAs("example.docx")
//	if err != nil {
//		log.Fatal(err)
//	}
//
// The library supports:
//   - Creating and editing paragraphs
//   - Adding and formatting text runs
//   - Creating tables
//   - Adding headers and footers
//   - Setting document properties
//   - Working with styles
//   - Adding comments
//   - Inserting images

import (
	"fmt"
)

// Document represents a Word document and provides methods to manipulate its content
type Document struct {
	pkg       *Package
	docPart   *DocumentPart
	comments  *Comments
	settings  *Settings
	styles    *Styles
	numbering *Numbering
}

// NewDocument creates a new empty Word document
func NewDocument() *Document {
	pkg := NewPackage()
	docPart := pkg.MainDocumentPart()

	return &Document{
		pkg:       pkg,
		docPart:   docPart,
		comments:  NewComments(),
		settings:  NewSettings(),
		styles:    NewStyles(),
		numbering: NewNumbering(pkg),
	}
}

// OpenDocument opens an existing Word document from a file path
func OpenDocument(path string) (*Document, error) {
	pkg, err := OpenPackage(path)
	if err != nil {
		return nil, fmt.Errorf("failed to open package: %w", err)
	}

	docPart := pkg.MainDocumentPart()
	if docPart.ContentType() != ContentTypeWMLDocumentMain {
		return nil, fmt.Errorf("file '%s' is not a Word file, content type is '%s'",
			path, docPart.ContentType())
	}

	return &Document{
		pkg:       pkg,
		docPart:   docPart,
		comments:  NewComments(),
		settings:  NewSettings(),
		styles:    NewStyles(),
		numbering: NewNumbering(pkg),
	}, nil
}

// Get XML content of the document as string.
// Returns an error if the document has no main document part.
func (d *Document) GetXML() (string, error) {
	if d.docPart == nil {
		return "", fmt.Errorf("document has no main document part")
	}
	return string(d.docPart.Data), nil
}

// AddParagraph adds a new paragraph to the end of the document and returns it
func (d *Document) AddParagraph(text ...string) *Paragraph {
	return d.docPart.AddParagraph(text...)
}

// AddPicture adds a new paragraph containing the specified image. Width and height are specified in EMUs.
// Passing zero for either dimension will keep the aspect ratio using the source image dimensions.
func (d *Document) AddPicture(path string, widthEMU, heightEMU int64) (*Paragraph, *Picture, error) {
	if d.docPart == nil {
		return nil, nil, fmt.Errorf("document has no main document part")
	}
	picture, err := d.docPart.addPictureFromFile(path, widthEMU, heightEMU)
	if err != nil {
		return nil, nil, err
	}
	paragraph := NewParagraph()
	paragraph.owner = d.docPart
	run := NewRun("")
	run.owner = d.docPart
	run.picture = picture
	paragraph.runs = append(paragraph.runs, run)

	d.docPart.paragraphs = append(d.docPart.paragraphs, paragraph)
	d.docPart.bodyElements = append(d.docPart.bodyElements, documentElement{paragraph: paragraph})
	d.docPart.updateXMLData()

	return paragraph, picture, nil
}

// AddHeading adds a heading paragraph with the specified text and level.
// Level 0 creates a Title style, levels 1-9 create Heading styles.
// Returns an error if level is outside the valid range [0-9].
func (d *Document) AddHeading(text string, level int) (*Paragraph, error) {
	if level < 0 || level > 9 {
		return nil, fmt.Errorf("level must be in range 0-9, got %d", level)
	}

	var style string
	if level == 0 {
		style = "Title"
	} else {
		style = fmt.Sprintf("Heading %d", level)
	}

	paragraph := d.AddParagraph(text)
	paragraph.SetStyle(style)
	return paragraph, nil
}

// AddTable adds a new table with the specified number of rows and columns
func (d *Document) AddTable(rows, cols int) *Table {
	return d.docPart.AddTable(rows, cols)
}

// AddPageBreak adds a page break to the document
func (d *Document) AddPageBreak() {
	paragraph := d.AddParagraph()
	run := paragraph.AddRun("")
	run.AddBreak(BreakTypePage)
}

// AddNumberedParagraph adds a paragraph with default decimal numbering at the specified level.
// Level must be non-negative (negative values will be treated as 0).
// Returns the created paragraph.
func (d *Document) AddNumberedParagraph(text string, level int) *Paragraph {
	if level < 0 {
		level = 0
	}
	numID := d.numbering.DecimalListID()
	paragraph := d.docPart.AddParagraph(text)
	paragraph.SetNumbering(numID, level)
	d.docPart.updateXMLData()
	return paragraph
}

// AddBulletedParagraph adds a paragraph with default bullet numbering at the specified level.
// Level must be non-negative (negative values will be treated as 0).
// Returns the created paragraph.
func (d *Document) AddBulletedParagraph(text string, level int) *Paragraph {
	if level < 0 {
		level = 0
	}
	numID := d.numbering.BulletedListID()
	paragraph := d.docPart.AddParagraph(text)
	paragraph.SetNumbering(numID, level)
	d.docPart.updateXMLData()
	return paragraph
}

// AddSection adds a new section to the document
func (d *Document) AddSection(startType SectionStartType) *Section {
	return d.docPart.AddSection(startType)
}

// Paragraphs returns all paragraphs in the document
func (d *Document) Paragraphs() []*Paragraph {
	return d.docPart.Paragraphs()
}

// Tables returns all tables in the document
func (d *Document) Tables() []*Table {
	return d.docPart.Tables()
}

// Sections returns all sections in the document
func (d *Document) Sections() []*Section {
	return d.docPart.Sections()
}

// Header returns the default header for the first section, creating both if necessary.
func (d *Document) Header() (*Header, error) {
	return d.HeaderOfType(HeaderTypeDefault)
}

// HeaderOfType returns the header of the specified type for the first section.
func (d *Document) HeaderOfType(headerType HeaderType) (*Header, error) {
	section, err := d.firstOrNewSection()
	if err != nil {
		return nil, err
	}
	return section.HeaderOfType(headerType)
}

// Footer returns the default footer for the first section, creating both if necessary.
func (d *Document) Footer() (*Footer, error) {
	return d.FooterOfType(FooterTypeDefault)
}

// FooterOfType returns the footer of the specified type for the first section.
func (d *Document) FooterOfType(footerType FooterType) (*Footer, error) {
	section, err := d.firstOrNewSection()
	if err != nil {
		return nil, err
	}
	return section.FooterOfType(footerType)
}

func (d *Document) firstOrNewSection() (*Section, error) {
	if d.docPart == nil {
		return nil, fmt.Errorf("document has no main document part")
	}
	sections := d.docPart.Sections()
	if len(sections) == 0 {
		section := d.docPart.AddSection(SectionStartContinuous)
		sections = d.docPart.Sections()
		if len(sections) == 0 {
			return nil, fmt.Errorf("failed to create section")
		}
		return section, nil
	}
	return sections[0], nil
}

// Numbering returns the document's numbering helper
func (d *Document) Numbering() *Numbering {
	return d.numbering
}

// CoreProperties returns the document's core properties (metadata)
func (d *Document) CoreProperties() *CoreProperties {
	return d.pkg.CoreProperties()
}

// Comments returns the document's comments collection
func (d *Document) Comments() *Comments {
	return d.comments
}

// Settings returns the document's settings
func (d *Document) Settings() *Settings {
	return d.settings
}

// Styles returns the document's styles collection
func (d *Document) Styles() *Styles {
	return d.styles
}

// SaveAs saves the document to the specified file path
func (d *Document) SaveAs(path string) error {
	if d != nil && d.docPart != nil {
		d.docPart.updateXMLData()
	}
	return d.pkg.SaveAs(path)
}

// Save saves the document to its original location (if opened from file)
func (d *Document) Save() error {
	if d.docPart != nil {
		d.docPart.updateXMLData()
	}
	return d.pkg.Save()
}

// Close closes the document and releases any resources
func (d *Document) Close() error {
	return d.pkg.Close()
}
