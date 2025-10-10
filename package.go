package docx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"sort"
	"strconv"
	"strings"
)

// Package represents an OpenXML package (DOCX file)
type Package struct {
	zipReader           *zip.ReadCloser
	parts               map[string]*Part
	relations           map[string][]*Relationship
	filePath            string
	coreProps           *CoreProperties
	contentTypes        map[string]string
	defaultContentTypes map[string]string
}

// Part represents a part within the OpenXML package
type Part struct {
	URI         string
	ContentType string
	Data        []byte
	Relations   []*Relationship
}

// Relationship represents a relationship between parts
type Relationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

// NewPackage creates a new empty package
func NewPackage() *Package {
	pkg := &Package{
		parts:               make(map[string]*Part),
		relations:           make(map[string][]*Relationship),
		coreProps:           NewCoreProperties(),
		contentTypes:        make(map[string]string),
		defaultContentTypes: make(map[string]string),
	}

	// Add default parts
	pkg.addDefaultParts()
	return pkg
}

// OpenPackage opens an existing package from file
func OpenPackage(filePath string) (*Package, error) {
	zipReader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open zip file: %w", err)
	}

	pkg := &Package{
		zipReader:           zipReader,
		parts:               make(map[string]*Part),
		relations:           make(map[string][]*Relationship),
		filePath:            filePath,
		coreProps:           NewCoreProperties(),
		contentTypes:        make(map[string]string),
		defaultContentTypes: make(map[string]string),
	}

	// Load all parts from the zip file
	err = pkg.loadParts()
	if err != nil {
		zipReader.Close()
		return nil, fmt.Errorf("failed to load parts: %w", err)
	}

	return pkg, nil
}

// MainDocumentPart returns the main document part
func (p *Package) MainDocumentPart() *DocumentPart {
	// Find the main document part through relationships
	rels := p.relations[""]
	for _, rel := range rels {
		if rel.Type == RelTypeOfficeDocument {
			if part, exists := p.parts[rel.Target]; exists {
				docPart := &DocumentPart{
					Part: part,
					pkg:  p,
				}
				_ = docPart.loadFromXML()
				return docPart
			}
		}
	}

	// If not found, create a new one
	docPart := NewDocumentPart()
	p.parts["word/document.xml"] = docPart.Part
	p.contentTypes["/word/document.xml"] = ContentTypeWMLDocumentMain
	docPart.pkg = p

	// Add relationship
	rel := &Relationship{
		ID:     "rId1",
		Type:   RelTypeOfficeDocument,
		Target: "word/document.xml",
	}
	p.relations[""] = append(p.relations[""], rel)

	return &DocumentPart{
		Part: docPart.Part,
		pkg:  p,
	}
}

// CoreProperties returns the core properties part
func (p *Package) CoreProperties() *CoreProperties {
	return p.coreProps
}

// SaveAs saves the package to a new file
func (p *Package) SaveAs(filePath string) error {
	file, err := os.Create(filePath)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer file.Close()

	zipWriter := zip.NewWriter(file)
	defer zipWriter.Close()

	// Write all parts to the zip file
	for uri, part := range p.parts {
		w, err := zipWriter.Create(uri)
		if err != nil {
			return fmt.Errorf("failed to create zip entry %s: %w", uri, err)
		}

		_, err = w.Write(part.Data)
		if err != nil {
			return fmt.Errorf("failed to write part data %s: %w", uri, err)
		}
	}

	// Write relationships
	for baseURI, rels := range p.relations {
		relsURI := p.relationshipsURI(baseURI)
		w, err := zipWriter.Create(relsURI)
		if err != nil {
			return fmt.Errorf("failed to create relationships entry %s: %w", relsURI, err)
		}

		relsXML, err := p.serializeRelationships(rels)
		if err != nil {
			return fmt.Errorf("failed to serialize relationships: %w", err)
		}

		_, err = w.Write(relsXML)
		if err != nil {
			return fmt.Errorf("failed to write relationships %s: %w", relsURI, err)
		}
	}

	// Write content types
	err = p.writeContentTypes(zipWriter)
	if err != nil {
		return fmt.Errorf("failed to write content types: %w", err)
	}

	p.filePath = filePath
	return nil
}

func (p *Package) parseContentTypes(data []byte) error {
	type contentTypesXML struct {
		Defaults []struct {
			Extension   string `xml:"Extension,attr"`
			ContentType string `xml:"ContentType,attr"`
		} `xml:"Default"`
		Overrides []struct {
			PartName    string `xml:"PartName,attr"`
			ContentType string `xml:"ContentType,attr"`
		} `xml:"Override"`
	}

	var ct contentTypesXML
	if err := xml.Unmarshal(data, &ct); err != nil {
		return err
	}

	for _, def := range ct.Defaults {
		p.defaultContentTypes[def.Extension] = def.ContentType
	}

	for _, ov := range ct.Overrides {
		p.contentTypes[ov.PartName] = ov.ContentType
	}

	return nil
}

func relationshipsBaseURI(relPath string) string {
	if relPath == "_rels/.rels" {
		return ""
	}

	dir := path.Dir(relPath)
	file := path.Base(relPath)
	file = strings.TrimSuffix(file, ".rels")
	dir = strings.TrimSuffix(dir, "/_rels")
	if dir == "." || dir == "" {
		return file
	}
	return path.Join(dir, file)
}

func parseRelationships(data []byte) ([]*Relationship, error) {
	type relationshipsXML struct {
		Relationships []*Relationship `xml:"Relationship"`
	}

	var rels relationshipsXML
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil, err
	}

	return rels.Relationships, nil
}

func (p *Package) lookupContentType(partName string) string {
	if !strings.HasPrefix(partName, "/") {
		partName = "/" + partName
	}

	if ct, ok := p.contentTypes[partName]; ok {
		return ct
	}

	// Use file extension defaults
	ext := path.Ext(partName)
	if ext != "" {
		ext = strings.TrimPrefix(ext, ".")
		if ct, ok := p.defaultContentTypes[ext]; ok {
			return ct
		}
	}

	return ""
}

// Save saves the package to its original location
func (p *Package) Save() error {
	if p.filePath == "" {
		return fmt.Errorf("no file path set, use SaveAs instead")
	}
	return p.SaveAs(p.filePath)
}

// Close closes the package and releases resources
func (p *Package) Close() error {
	if p.zipReader != nil {
		return p.zipReader.Close()
	}
	return nil
}

// addDefaultParts adds the default parts required for a minimal DOCX file
func (p *Package) addDefaultParts() {
	// Add main document part
	docPart := NewDocumentPart()
	p.parts["word/document.xml"] = docPart.Part
	p.contentTypes["/word/document.xml"] = ContentTypeWMLDocumentMain

	// Add styles part
	stylesPart := NewStylesPart()
	p.parts["word/styles.xml"] = stylesPart.Part
	p.contentTypes["/word/styles.xml"] = ContentTypeWMLStyles

	// Add settings part
	settingsPart := NewSettingsPart()
	p.parts["word/settings.xml"] = settingsPart.Part
	p.contentTypes["/word/settings.xml"] = ContentTypeWMLSettings

	// Add numbering part
	numberingPart := NewNumberingPart()
	p.parts["word/numbering.xml"] = numberingPart.Part
	p.contentTypes["/word/numbering.xml"] = ContentTypeWMLNumbering

	// Populate default content types
	p.defaultContentTypes["rels"] = ContentTypeRels
	p.defaultContentTypes["xml"] = "application/xml"

	// Add relationships
	p.relations[""] = []*Relationship{
		{
			ID:     "rId1",
			Type:   RelTypeOfficeDocument,
			Target: "word/document.xml",
		},
	}

	p.ensureRelationship("word/document.xml", RelTypeStyles, "styles.xml")
	p.ensureRelationship("word/document.xml", RelTypeSettings, "settings.xml")
	p.ensureRelationship("word/document.xml", RelTypeNumbering, "numbering.xml")
}

// loadParts loads all parts from the zip file
func (p *Package) loadParts() error {
	// First, parse content types so they are available for subsequent parts
	for _, file := range p.zipReader.File {
		if file.Name != "[Content_Types].xml" {
			continue
		}

		rc, err := file.Open()
		if err != nil {
			return fmt.Errorf("failed to open file %s: %w", file.Name, err)
		}

		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return fmt.Errorf("failed to read file %s: %w", file.Name, err)
		}

		if err := p.parseContentTypes(data); err != nil {
			return fmt.Errorf("failed to parse content types: %w", err)
		}

		break
	}

	for _, file := range p.zipReader.File {
		// Skip directories
		if strings.HasSuffix(file.Name, "/") {
			continue
		}

		if file.Name == "[Content_Types].xml" {
			continue
		}

		rc, err := file.Open()
		if err != nil {
			return fmt.Errorf("failed to open file %s: %w", file.Name, err)
		}

		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return fmt.Errorf("failed to read file %s: %w", file.Name, err)
		}

		if strings.HasSuffix(file.Name, ".rels") {
			baseURI := relationshipsBaseURI(file.Name)
			rels, err := parseRelationships(data)
			if err != nil {
				return fmt.Errorf("failed to parse relationships for %s: %w", file.Name, err)
			}
			p.relations[baseURI] = rels
			continue
		}

		part := &Part{
			URI:         file.Name,
			Data:        data,
			ContentType: p.lookupContentType(file.Name),
		}
		p.parts[file.Name] = part
	}

	return nil
}

// relationshipsURI returns the relationships URI for a given base URI
func (p *Package) relationshipsURI(baseURI string) string {
	if baseURI == "" {
		return "_rels/.rels"
	}

	dir := path.Dir(baseURI)
	base := path.Base(baseURI)
	return path.Join(dir, "_rels", base+".rels")
}

// serializeRelationships serializes relationships to XML
func (p *Package) serializeRelationships(rels []*Relationship) ([]byte, error) {
	type Relationships struct {
		XMLName       xml.Name        `xml:"Relationships"`
		Xmlns         string          `xml:"xmlns,attr"`
		Relationships []*Relationship `xml:"Relationship"`
	}

	relationships := &Relationships{
		Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: rels,
	}

	return xml.MarshalIndent(relationships, "", "  ")
}

// writeContentTypes writes the [Content_Types].xml file
func (p *Package) writeContentTypes(zipWriter *zip.Writer) error {
	type Default struct {
		Extension   string `xml:"Extension,attr"`
		ContentType string `xml:"ContentType,attr"`
	}

	type Override struct {
		PartName    string `xml:"PartName,attr"`
		ContentType string `xml:"ContentType,attr"`
	}

	type Types struct {
		XMLName   xml.Name   `xml:"Types"`
		Xmlns     string     `xml:"xmlns,attr"`
		Defaults  []Default  `xml:"Default"`
		Overrides []Override `xml:"Override"`
	}

	types := &Types{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
	}

	defaultKeys := make([]string, 0, len(p.defaultContentTypes))
	for ext := range p.defaultContentTypes {
		defaultKeys = append(defaultKeys, ext)
	}
	if len(defaultKeys) == 0 {
		defaultKeys = append(defaultKeys, "rels", "xml")
		p.defaultContentTypes["rels"] = ContentTypeRels
		p.defaultContentTypes["xml"] = "application/xml"
	} else {
		sort.Strings(defaultKeys)
	}

	for _, ext := range defaultKeys {
		types.Defaults = append(types.Defaults, Default{
			Extension:   ext,
			ContentType: p.defaultContentTypes[ext],
		})
	}

	overrideKeys := make([]string, 0, len(p.contentTypes))
	for partName := range p.contentTypes {
		overrideKeys = append(overrideKeys, partName)
	}
	sort.Strings(overrideKeys)
	for _, partName := range overrideKeys {
		types.Overrides = append(types.Overrides, Override{
			PartName:    partName,
			ContentType: p.contentTypes[partName],
		})
	}

	w, err := zipWriter.Create("[Content_Types].xml")
	if err != nil {
		return err
	}

	data, err := xml.MarshalIndent(types, "", "  ")
	if err != nil {
		return err
	}

	_, err = w.Write(data)
	return err
}

func (p *Package) ensureRelationship(baseURI, relType, target string) string {
	return p.ensureRelationshipWithMode(baseURI, relType, target, "")
}

func (p *Package) ensureRelationshipWithMode(baseURI, relType, target, targetMode string) string {
	rels := p.relations[baseURI]
	for _, rel := range rels {
		if rel.Type == relType && rel.Target == target {
			existingMode := rel.TargetMode
			if existingMode == targetMode || (existingMode == "" && targetMode == "") {
				return rel.ID
			}
		}
	}

	id := p.nextRelationshipID(baseURI)
	rel := &Relationship{
		ID:     id,
		Type:   relType,
		Target: target,
	}
	if targetMode != "" {
		rel.TargetMode = targetMode
	}
	p.relations[baseURI] = append(rels, rel)
	return id
}

func (p *Package) nextRelationshipID(baseURI string) string {
	rels := p.relations[baseURI]
	maxID := 0
	for _, rel := range rels {
		if strings.HasPrefix(rel.ID, "rId") {
			if n, err := strconv.Atoi(strings.TrimPrefix(rel.ID, "rId")); err == nil {
				if n > maxID {
					maxID = n
				}
			}
		}
	}
	return fmt.Sprintf("rId%d", maxID+1)
}
