package docx

import (
	"bytes"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"math"
	"os"
	"path"
	"path/filepath"
	"strings"
)

const (
	// EMUs per common measurement unit
	EMUsPerInch  = 914400
	EMUsPerCm    = 360000
	EMUsPerPoint = 12700

	defaultImageDPI = 96
)

var imageContentTypes = map[string]string{
	".bmp":  "image/bmp",
	".dib":  "image/bmp",
	".emf":  "image/x-emf",
	".gif":  "image/gif",
	".jpg":  "image/jpeg",
	".jpeg": "image/jpeg",
	".jfif": "image/jpeg",
	".png":  "image/png",
	".tif":  "image/tiff",
	".tiff": "image/tiff",
	".wmf":  "image/x-wmf",
}

// Picture represents an inline picture embedded in a run.
type Picture struct {
	docPart     *DocumentPart
	relID       string
	target      string
	widthEMU    int64
	heightEMU   int64
	docPrID     int
	name        string
	description string
}

// WidthEMU returns the picture width in English Metric Units (EMUs).
func (p *Picture) WidthEMU() int64 {
	return p.widthEMU
}

// HeightEMU returns the picture height in English Metric Units (EMUs).
func (p *Picture) HeightEMU() int64 {
	return p.heightEMU
}

// RelationshipID returns the relationship ID referencing the image part.
func (p *Picture) RelationshipID() string {
	return p.relID
}

// Target returns the relationship target (typically media/imageX.ext).
func (p *Picture) Target() string {
	return p.target
}

// Name returns the docPr name for this picture.
func (p *Picture) Name() string {
	return p.name
}

// Description returns the docPr description for this picture.
func (p *Picture) Description() string {
	return p.description
}

// ImageData returns the raw bytes of the embedded image.
func (p *Picture) ImageData() ([]byte, error) {
	if p == nil || p.docPart == nil || p.docPart.pkg == nil {
		return nil, fmt.Errorf("picture is detached from document")
	}
	target := strings.TrimPrefix(p.target, "/")
	for strings.HasPrefix(target, "../") {
		target = strings.TrimPrefix(target, "../")
	}
	uri := target
	if !strings.HasPrefix(uri, "word/") {
		uri = path.Join("word", uri)
	}
	part, ok := p.docPart.pkg.parts[uri]
	if !ok {
		return nil, fmt.Errorf("image part %s not found", uri)
	}
	return part.Data, nil
}

func (p *Picture) toXML() string {
	if p == nil {
		return ""
	}
	name := p.name
	if name == "" {
		name = fmt.Sprintf("Picture %d", p.docPrID)
	}
	descr := p.description
	var builder strings.Builder
	builder.WriteString(`<w:drawing>`)
	builder.WriteString(`<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" distT="0" distB="0" distL="0" distR="0">`)
	builder.WriteString(fmt.Sprintf(`<wp:extent cx="%d" cy="%d"/>`, p.widthEMU, p.heightEMU))
	builder.WriteString(fmt.Sprintf(`<wp:docPr id="%d" name="%s" descr="%s"/>`, p.docPrID, escapeXML(name), escapeXML(descr)))
	builder.WriteString(`<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`)
	builder.WriteString(`<a:graphic>`)
	builder.WriteString(`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	builder.WriteString(`<pic:pic>`)
	builder.WriteString(`<pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>`)
	builder.WriteString(`<pic:blipFill><a:blip r:embed="` + escapeXML(p.relID) + `"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`)
	builder.WriteString(`<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="`)
	builder.WriteString(fmt.Sprintf("%d", p.widthEMU))
	builder.WriteString(`" cy="`)
	builder.WriteString(fmt.Sprintf("%d", p.heightEMU))
	builder.WriteString(`"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`)
	builder.WriteString(`</pic:pic>`)
	builder.WriteString(`</a:graphicData>`)
	builder.WriteString(`</a:graphic>`)
	builder.WriteString(`</wp:inline>`)
	builder.WriteString(`</w:drawing>`)
	return builder.String()
}

func escapeXML(value string) string {
	replacer := strings.NewReplacer(
		"&", "&amp;",
		"<", "&lt;",
		">", "&gt;",
		"\"", "&quot;",
		"'", "&apos;",
	)
	return replacer.Replace(value)
}

func decodeImageDimensionsEMU(data []byte) (int64, int64, error) {
	cfg, _, err := image.DecodeConfig(bytes.NewReader(data))
	if err != nil {
		return 0, 0, err
	}
	if cfg.Width <= 0 || cfg.Height <= 0 {
		return 0, 0, fmt.Errorf("invalid image dimensions")
	}
	emusPerPixel := EMUsPerInch / defaultImageDPI
	if emusPerPixel <= 0 {
		emusPerPixel = 9525
	}
	widthEMU := int64(cfg.Width) * int64(emusPerPixel)
	heightEMU := int64(cfg.Height) * int64(emusPerPixel)
	return widthEMU, heightEMU, nil
}

func scaleEMU(value, numerator, denominator int64) int64 {
	if value <= 0 || numerator <= 0 || denominator <= 0 {
		return value
	}
	return (value*numerator + denominator/2) / denominator
}

// InchesToEMU converts a measurement in inches to EMUs.
func InchesToEMU(inches float64) int64 {
	return int64(math.Round(inches * float64(EMUsPerInch)))
}

// CentimetersToEMU converts centimeters to EMUs.
func CentimetersToEMU(cm float64) int64 {
	return int64(math.Round(cm * float64(EMUsPerCm)))
}

// PointsToEMU converts points to EMUs.
func PointsToEMU(points float64) int64 {
	return int64(math.Round(points * float64(EMUsPerPoint)))
}

func (dp *DocumentPart) addPictureFromFile(path string, widthEMU, heightEMU int64) (*Picture, error) {
	if dp == nil || dp.pkg == nil {
		return nil, fmt.Errorf("paragraph is not attached to a document package")
	}

	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("failed to read image %s: %w", path, err)
	}

	ext := strings.ToLower(filepath.Ext(path))
	contentType, ok := imageContentTypes[ext]
	if !ok {
		return nil, fmt.Errorf("unsupported image format: %s", ext)
	}

	var (
		defaultWidthEMU  int64
		defaultHeightEMU int64
		dimErr           error
	)

	if widthEMU <= 0 || heightEMU <= 0 {
		defaultWidthEMU, defaultHeightEMU, dimErr = decodeImageDimensionsEMU(data)
	}

	switch {
	case widthEMU <= 0 && heightEMU <= 0:
		if dimErr != nil {
			return nil, fmt.Errorf("picture width and height must be specified: %w", dimErr)
		}
		widthEMU = defaultWidthEMU
		heightEMU = defaultHeightEMU
	case widthEMU <= 0:
		if dimErr != nil || defaultWidthEMU == 0 || defaultHeightEMU == 0 {
			return nil, fmt.Errorf("unable to determine picture width automatically; specify both width and height")
		}
		widthEMU = scaleEMU(heightEMU, defaultWidthEMU, defaultHeightEMU)
	case heightEMU <= 0:
		if dimErr != nil || defaultWidthEMU == 0 || defaultHeightEMU == 0 {
			return nil, fmt.Errorf("unable to determine picture height automatically; specify both width and height")
		}
		heightEMU = scaleEMU(widthEMU, defaultHeightEMU, defaultWidthEMU)
	}

	if widthEMU <= 0 || heightEMU <= 0 {
		return nil, fmt.Errorf("picture width and height must be positive EMU values")
	}

	partURI, err := dp.pkg.addImagePart(data, ext, contentType)
	if err != nil {
		return nil, err
	}
	target := strings.TrimPrefix(partURI, "word/")
	relID := dp.pkg.ensureRelationship(dp.Part.URI, RelTypeImage, target)
	docPrID := dp.nextDrawingID()

	base := strings.TrimSuffix(filepath.Base(path), ext)
	picture := &Picture{
		docPart:     dp,
		relID:       relID,
		target:      target,
		widthEMU:    widthEMU,
		heightEMU:   heightEMU,
		docPrID:     docPrID,
		name:        fmt.Sprintf("Picture %d", docPrID),
		description: base,
	}
	return picture, nil
}
