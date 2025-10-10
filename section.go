package docx

import (
	"fmt"
	"sort"
	"strings"
)

// Section represents a section in a Word document
type Section struct {
	startType    SectionStartType
	pageWidth    int // in twentieths of a point
	pageHeight   int // in twentieths of a point
	marginTop    int // in twentieths of a point
	marginRight  int
	marginBottom int
	marginLeft   int
	owner        *DocumentPart
	headerRefs   map[HeaderType]*headerReference
	footerRefs   map[FooterType]*footerReference
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
		headerRefs:   make(map[HeaderType]*headerReference),
		footerRefs:   make(map[FooterType]*footerReference),
	}
}

func (s *Section) setOwner(owner *DocumentPart) {
	s.owner = owner
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

// Header returns the default header for the section, creating it if necessary.
func (s *Section) Header() (*Header, error) {
	return s.headerOfType(HeaderTypeDefault)
}

// Footer returns the default footer for the section, creating it if necessary.
func (s *Section) Footer() (*Footer, error) {
	return s.footerOfType(FooterTypeDefault)
}

// HeaderOfType returns the header for the specified header type, creating it if necessary.
func (s *Section) HeaderOfType(headerType HeaderType) (*Header, error) {
	return s.headerOfType(headerType)
}

// FooterOfType returns the footer for the specified footer type, creating it if necessary.
func (s *Section) FooterOfType(footerType FooterType) (*Footer, error) {
	return s.footerOfType(footerType)
}

func (s *Section) headerOfType(headerType HeaderType) (*Header, error) {
	if s == nil {
		return nil, fmt.Errorf("section is nil")
	}
	if ref, ok := s.headerRefs[headerType]; ok && ref != nil {
		return ref.header, nil
	}
	if s.owner == nil {
		return nil, fmt.Errorf("section has no owner document part")
	}
	header, relID, err := s.owner.createHeaderPart()
	if err != nil {
		return nil, err
	}
	if s.headerRefs == nil {
		s.headerRefs = make(map[HeaderType]*headerReference)
	}
	s.headerRefs[headerType] = &headerReference{typeValue: headerType, relID: relID, header: header}
	if s.owner != nil {
		s.owner.updateXMLData()
	}
	return header, nil
}

func (s *Section) footerOfType(footerType FooterType) (*Footer, error) {
	if s == nil {
		return nil, fmt.Errorf("section is nil")
	}
	if ref, ok := s.footerRefs[footerType]; ok && ref != nil {
		return ref.footer, nil
	}
	if s.owner == nil {
		return nil, fmt.Errorf("section has no owner document part")
	}
	footer, relID, err := s.owner.createFooterPart()
	if err != nil {
		return nil, err
	}
	if s.footerRefs == nil {
		s.footerRefs = make(map[FooterType]*footerReference)
	}
	s.footerRefs[footerType] = &footerReference{typeValue: footerType, relID: relID, footer: footer}
	if s.owner != nil {
		s.owner.updateXMLData()
	}
	return footer, nil
}

func (s *Section) headerReferenceElements() []string {
	if len(s.headerRefs) == 0 {
		return nil
	}
	var elements []string
	order := []HeaderType{HeaderTypeDefault, HeaderTypeFirst, HeaderTypeEven}
	for _, key := range order {
		if ref, ok := s.headerRefs[key]; ok && ref != nil && ref.header != nil && ref.relID != "" {
			elements = append(elements, fmt.Sprintf(`<w:headerReference w:type="%s" r:id="%s"/>`, ref.typeValue, ref.relID))
		}
	}
	otherKeys := make([]string, 0)
	for key := range s.headerRefs {
		switch key {
		case HeaderTypeDefault, HeaderTypeFirst, HeaderTypeEven:
			// already handled
		default:
			otherKeys = append(otherKeys, string(key))
		}
	}
	if len(otherKeys) > 0 {
		sort.Strings(otherKeys)
		for _, key := range otherKeys {
			if ref := s.headerRefs[HeaderType(key)]; ref != nil && ref.header != nil && ref.relID != "" {
				elements = append(elements, fmt.Sprintf(`<w:headerReference w:type="%s" r:id="%s"/>`, ref.typeValue, ref.relID))
			}
		}
	}
	return elements
}

func (s *Section) footerReferenceElements() []string {
	if len(s.footerRefs) == 0 {
		return nil
	}
	var elements []string
	order := []FooterType{FooterTypeDefault, FooterTypeFirst, FooterTypeEven}
	for _, key := range order {
		if ref, ok := s.footerRefs[key]; ok && ref != nil && ref.footer != nil && ref.relID != "" {
			elements = append(elements, fmt.Sprintf(`<w:footerReference w:type="%s" r:id="%s"/>`, ref.typeValue, ref.relID))
		}
	}
	otherKeys := make([]string, 0)
	for key := range s.footerRefs {
		switch key {
		case FooterTypeDefault, FooterTypeFirst, FooterTypeEven:
			// already handled
		default:
			otherKeys = append(otherKeys, string(key))
		}
	}
	if len(otherKeys) > 0 {
		sort.Strings(otherKeys)
		for _, key := range otherKeys {
			if ref := s.footerRefs[FooterType(key)]; ref != nil && ref.footer != nil && ref.relID != "" {
				elements = append(elements, fmt.Sprintf(`<w:footerReference w:type="%s" r:id="%s"/>`, ref.typeValue, ref.relID))
			}
		}
	}
	return elements
}

// ToXML converts the section to WordprocessingML XML
func (s *Section) ToXML() string {
	var elements []string
	if s.startType != SectionStartContinuous {
		elements = append(elements, fmt.Sprintf(`<w:type w:val="%s"/>`, s.startType))
	}
	if headerElems := s.headerReferenceElements(); len(headerElems) > 0 {
		elements = append(elements, headerElems...)
	}
	if footerElems := s.footerReferenceElements(); len(footerElems) > 0 {
		elements = append(elements, footerElems...)
	}
	elements = append(elements, fmt.Sprintf(`<w:pgSz w:w="%d" w:h="%d"/>`, s.pageWidth, s.pageHeight))
	elements = append(elements, fmt.Sprintf(`<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d"/>`, s.marginTop, s.marginRight, s.marginBottom, s.marginLeft))

	return fmt.Sprintf(`<w:sectPr>
  %s
</w:sectPr>`, strings.Join(elements, "\n  "))
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
