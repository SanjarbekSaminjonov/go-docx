package docx

import "fmt"

// Section represents a section in a Word document
type Section struct {
	startType    SectionStartType
	pageWidth    int // in twentieths of a point
	pageHeight   int // in twentieths of a point
	marginTop    int // in twentieths of a point
	marginRight  int
	marginBottom int
	marginLeft   int
}

// NewSection creates a new section with the specified start type
func NewSection(startType SectionStartType) *Section {
	return &Section{
		startType:    startType,
		pageWidth:    11906, // 8.5 inches
		pageHeight:   16838, // 11.69 inches
		marginTop:    1440,  // 1 inch
		marginRight:  1440,  // 1 inch
		marginBottom: 1440,  // 1 inch
		marginLeft:   1440,  // 1 inch
	}
}

// SetPageSize sets the page size in twentieths of a point
func (s *Section) SetPageSize(width, height int) {
	s.pageWidth = width
	s.pageHeight = height
}

// SetMargins sets the margins in twentieths of a point
func (s *Section) SetMargins(top, right, bottom, left int) {
	s.marginTop = top
	s.marginRight = right
	s.marginBottom = bottom
	s.marginLeft = left
}

// SetStartType sets how this section starts
func (s *Section) SetStartType(startType SectionStartType) {
	s.startType = startType
}

// ToXML converts the section to WordprocessingML XML
func (s *Section) ToXML() string {
	var typeXML string
	if s.startType != SectionStartContinuous {
		typeXML = fmt.Sprintf(`<w:type w:val="%s"/>`, s.startType)
	}

	return fmt.Sprintf(`<w:sectPr>
  %s
  <w:pgSz w:w="%d" w:h="%d"/>
  <w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d"/>
</w:sectPr>`, typeXML, s.pageWidth, s.pageHeight, s.marginTop, s.marginRight, s.marginBottom, s.marginLeft)
}

// Comments represents a collection of comments in a document
type Comments struct {
	comments []*Comment
	nextID   int
}

// Comment represents a single comment
type Comment struct {
	ID       int
	Author   string
	Initials string
	Text     string
}

// NewComments creates a new comments collection
func NewComments() *Comments {
	return &Comments{
		comments: make([]*Comment, 0),
		nextID:   1,
	}
}

// AddComment adds a new comment
func (c *Comments) AddComment(text, author, initials string) *Comment {
	comment := &Comment{
		ID:       c.nextID,
		Author:   author,
		Initials: initials,
		Text:     text,
	}

	c.comments = append(c.comments, comment)
	c.nextID++

	return comment
}

// Settings represents document settings
type Settings struct {
	defaultTabStop int
	zoom           int
}

// NewSettings creates new document settings
func NewSettings() *Settings {
	return &Settings{
		defaultTabStop: 708, // 0.5 inch
		zoom:           100,
	}
}

// SetDefaultTabStop sets the default tab stop in twentieths of a point
func (s *Settings) SetDefaultTabStop(tabStop int) {
	s.defaultTabStop = tabStop
}

// SetZoom sets the zoom percentage
func (s *Settings) SetZoom(zoom int) {
	s.zoom = zoom
}

// Styles represents a collection of document styles
type Styles struct {
	styles []*Style
}

// Style represents a document style
type Style struct {
	ID   string
	Name string
	Type string
}

// NewStyles creates a new styles collection
func NewStyles() *Styles {
	return &Styles{
		styles: make([]*Style, 0),
	}
}

// AddStyle adds a new style
func (s *Styles) AddStyle(id, name, styleType string) *Style {
	style := &Style{
		ID:   id,
		Name: name,
		Type: styleType,
	}

	s.styles = append(s.styles, style)
	return style
}
