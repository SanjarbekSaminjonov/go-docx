package docx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"strings"
)

// HeaderType identifies the type of header applied to a section.
type HeaderType string

const (
	HeaderTypeDefault HeaderType = "default"
	HeaderTypeFirst   HeaderType = "first"
	HeaderTypeEven    HeaderType = "even"
)

// FooterType identifies the type of footer applied to a section.
type FooterType string

const (
	FooterTypeDefault FooterType = "default"
	FooterTypeFirst   FooterType = "first"
	FooterTypeEven    FooterType = "even"
)

type headerReference struct {
	typeValue HeaderType
	relID     string
	header    *Header
}

type footerReference struct {
	typeValue FooterType
	relID     string
	footer    *Footer
}

// Header represents a header part in the document.
type Header struct {
	part         *Part
	owner        *DocumentPart
	paragraphs   []*Paragraph
	tables       []*Table
	bodyElements []documentElement
}

// Footer represents a footer part in the document.
type Footer struct {
	part         *Part
	owner        *DocumentPart
	paragraphs   []*Paragraph
	tables       []*Table
	bodyElements []documentElement
}

func newHeader(owner *DocumentPart, part *Part) *Header {
	h := &Header{
		part:         part,
		owner:        owner,
		paragraphs:   make([]*Paragraph, 0),
		tables:       make([]*Table, 0),
		bodyElements: make([]documentElement, 0),
	}
	if h.part != nil && len(h.part.Data) > 0 {
		_ = h.loadFromXML()
	} else {
		h.updateXMLData()
	}
	return h
}

func newFooter(owner *DocumentPart, part *Part) *Footer {
	f := &Footer{
		part:         part,
		owner:        owner,
		paragraphs:   make([]*Paragraph, 0),
		tables:       make([]*Table, 0),
		bodyElements: make([]documentElement, 0),
	}
	if f.part != nil && len(f.part.Data) > 0 {
		_ = f.loadFromXML()
	} else {
		f.updateXMLData()
	}
	return f
}

func (h *Header) AddParagraph(text ...string) *Paragraph {
	paragraph := NewParagraph()
	if h.owner != nil {
		paragraph.owner = h.owner
	}
	for _, t := range text {
		paragraph.AddRun(t)
	}
	h.paragraphs = append(h.paragraphs, paragraph)
	h.bodyElements = append(h.bodyElements, documentElement{paragraph: paragraph})
	h.updateXMLData()
	return paragraph
}

func (h *Header) AddTable(rows, cols int) *Table {
	table := NewTable(rows, cols)
	if h.owner != nil {
		table.setOwner(h.owner)
	}
	h.tables = append(h.tables, table)
	h.bodyElements = append(h.bodyElements, documentElement{table: table})
	h.updateXMLData()
	return table
}

func (h *Header) Paragraphs() []*Paragraph {
	return h.paragraphs
}

func (h *Header) Tables() []*Table {
	return h.tables
}

func (h *Header) updateXMLData() {
	if h.part == nil {
		return
	}
	var content strings.Builder
	for _, element := range h.bodyElements {
		if element.paragraph != nil {
			content.WriteString(element.paragraph.ToXML())
		} else if element.table != nil {
			content.WriteString(element.table.ToXML())
		}
	}
	h.part.Data = []byte(fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
%s
</w:hdr>`, content.String()))
}

func (h *Header) loadFromXML() error {
	if h.part == nil || len(h.part.Data) == 0 {
		return nil
	}
	h.paragraphs = make([]*Paragraph, 0)
	h.tables = make([]*Table, 0)
	h.bodyElements = make([]documentElement, 0)

	decoder := xml.NewDecoder(bytes.NewReader(h.part.Data))
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
			case "hdr":
				// consume children
			case "p":
				paragraph, err := parseParagraph(decoder, t, h.owner)
				if err != nil {
					return err
				}
				h.paragraphs = append(h.paragraphs, paragraph)
				h.bodyElements = append(h.bodyElements, documentElement{paragraph: paragraph})
			case "tbl":
				table, err := parseTable(decoder, t, h.owner)
				if err != nil {
					return err
				}
				h.tables = append(h.tables, table)
				h.bodyElements = append(h.bodyElements, documentElement{table: table})
			default:
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

func (f *Footer) AddParagraph(text ...string) *Paragraph {
	paragraph := NewParagraph()
	if f.owner != nil {
		paragraph.owner = f.owner
	}
	for _, t := range text {
		paragraph.AddRun(t)
	}
	f.paragraphs = append(f.paragraphs, paragraph)
	f.bodyElements = append(f.bodyElements, documentElement{paragraph: paragraph})
	f.updateXMLData()
	return paragraph
}

func (f *Footer) AddTable(rows, cols int) *Table {
	table := NewTable(rows, cols)
	if f.owner != nil {
		table.setOwner(f.owner)
	}
	f.tables = append(f.tables, table)
	f.bodyElements = append(f.bodyElements, documentElement{table: table})
	f.updateXMLData()
	return table
}

func (f *Footer) Paragraphs() []*Paragraph {
	return f.paragraphs
}

func (f *Footer) Tables() []*Table {
	return f.tables
}

func (f *Footer) updateXMLData() {
	if f.part == nil {
		return
	}
	var content strings.Builder
	for _, element := range f.bodyElements {
		if element.paragraph != nil {
			content.WriteString(element.paragraph.ToXML())
		} else if element.table != nil {
			content.WriteString(element.table.ToXML())
		}
	}
	f.part.Data = []byte(fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
%s
</w:ftr>`, content.String()))
}

func (f *Footer) loadFromXML() error {
	if f.part == nil || len(f.part.Data) == 0 {
		return nil
	}
	f.paragraphs = make([]*Paragraph, 0)
	f.tables = make([]*Table, 0)
	f.bodyElements = make([]documentElement, 0)

	decoder := xml.NewDecoder(bytes.NewReader(f.part.Data))
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
			case "ftr":
				// consume children
			case "p":
				paragraph, err := parseParagraph(decoder, t, f.owner)
				if err != nil {
					return err
				}
				f.paragraphs = append(f.paragraphs, paragraph)
				f.bodyElements = append(f.bodyElements, documentElement{paragraph: paragraph})
			case "tbl":
				table, err := parseTable(decoder, t, f.owner)
				if err != nil {
					return err
				}
				f.tables = append(f.tables, table)
				f.bodyElements = append(f.bodyElements, documentElement{table: table})
			default:
				if err := skipElement(decoder, t); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

func resolveRelationshipTarget(baseURI, target string) string {
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}
	baseDir := path.Dir(baseURI)
	if baseDir == "." || baseDir == "" {
		return target
	}
	return path.Clean(path.Join(baseDir, target))
}
