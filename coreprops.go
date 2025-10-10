package docx

import (
	"encoding/xml"
	"time"
)

// CoreProperties represents the core properties (metadata) of a document
type CoreProperties struct {
	Title       string
	Subject     string
	Creator     string
	Keywords    string
	Description string
	Category    string
	Created     time.Time
	Modified    time.Time
	Revision    string
}

// NewCoreProperties creates a new CoreProperties instance with default values
func NewCoreProperties() *CoreProperties {
	now := time.Now()
	return &CoreProperties{
		Created:  now,
		Modified: now,
		Revision: "1",
	}
}

// SetTitle sets the document title
func (cp *CoreProperties) SetTitle(title string) {
	cp.Title = title
}

// SetSubject sets the document subject
func (cp *CoreProperties) SetSubject(subject string) {
	cp.Subject = subject
}

// SetCreator sets the document creator/author
func (cp *CoreProperties) SetCreator(creator string) {
	cp.Creator = creator
}

// SetKeywords sets the document keywords
func (cp *CoreProperties) SetKeywords(keywords string) {
	cp.Keywords = keywords
}

// SetDescription sets the document description
func (cp *CoreProperties) SetDescription(description string) {
	cp.Description = description
}

// SetCategory sets the document category
func (cp *CoreProperties) SetCategory(category string) {
	cp.Category = category
}

// ToXML converts the core properties to XML format
func (cp *CoreProperties) ToXML() ([]byte, error) {
	type CorePropsXML struct {
		XMLName      xml.Name `xml:"cp:coreProperties"`
		Xmlns        string   `xml:"xmlns:cp,attr"`
		XmlnsDC      string   `xml:"xmlns:dc,attr"`
		XmlnsDCTerms string   `xml:"xmlns:dcterms,attr"`
		XmlnsXsi     string   `xml:"xmlns:xsi,attr"`

		Title       string    `xml:"dc:title,omitempty"`
		Subject     string    `xml:"dc:subject,omitempty"`
		Creator     string    `xml:"dc:creator,omitempty"`
		Keywords    string    `xml:"cp:keywords,omitempty"`
		Description string    `xml:"dc:description,omitempty"`
		Category    string    `xml:"cp:category,omitempty"`
		Created     time.Time `xml:"dcterms:created,omitempty"`
		Modified    time.Time `xml:"dcterms:modified,omitempty"`
		Revision    string    `xml:"cp:revision,omitempty"`
	}

	props := CorePropsXML{
		Xmlns:        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
		XmlnsDC:      "http://purl.org/dc/elements/1.1/",
		XmlnsDCTerms: "http://purl.org/dc/terms/",
		XmlnsXsi:     "http://www.w3.org/2001/XMLSchema-instance",
		Title:        cp.Title,
		Subject:      cp.Subject,
		Creator:      cp.Creator,
		Keywords:     cp.Keywords,
		Description:  cp.Description,
		Category:     cp.Category,
		Created:      cp.Created,
		Modified:     cp.Modified,
		Revision:     cp.Revision,
	}

	return xml.MarshalIndent(props, "", "  ")
}
