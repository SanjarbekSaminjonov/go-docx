package docx

import (
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestDocumentCreation(t *testing.T) {
	doc := NewDocument()
	if doc == nil {
		t.Fatal("NewDocument returned nil")
	}
}

func TestAddParagraph(t *testing.T) {
	doc := NewDocument()
	p := doc.AddParagraph("Test paragraph")

	if p == nil {
		t.Fatal("AddParagraph returned nil")
	}

	if p.Text() != "Test paragraph" {
		t.Errorf("Expected 'Test paragraph', got '%s'", p.Text())
	}
}

func TestAddTable(t *testing.T) {
	doc := NewDocument()
	table := doc.AddTable(2, 3)

	if table == nil {
		t.Fatal("AddTable returned nil")
	}

	if len(table.Rows()) != 2 {
		t.Errorf("Expected 2 rows, got %d", len(table.Rows()))
	}

	if len(table.Row(0).Cells()) != 3 {
		t.Errorf("Expected 3 cells, got %d", len(table.Row(0).Cells()))
	}
}

func TestOpenDocumentParsesParagraphs(t *testing.T) {
	doc := NewDocument()
	doc.AddParagraph("First paragraph")
	doc.AddParagraph("Second paragraph")

	outputPath := filepath.Join(t.TempDir(), "paragraphs.docx")
	if err := doc.SaveAs(outputPath); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(outputPath)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	paragraphs := reopened.Paragraphs()
	if len(paragraphs) != 2 {
		t.Fatalf("expected 2 paragraphs, got %d", len(paragraphs))
	}

	if paragraphs[0].Text() != "First paragraph" {
		t.Errorf("expected first paragraph text to be 'First paragraph', got %q", paragraphs[0].Text())
	}

	if paragraphs[1].Text() != "Second paragraph" {
		t.Errorf("expected second paragraph text to be 'Second paragraph', got %q", paragraphs[1].Text())
	}
}

func TestOpenDocumentPreservesRunFormatting(t *testing.T) {
	doc := NewDocument()
	paragraph := doc.AddParagraph()
	run := paragraph.AddRun("Formatted text")
	run.SetBold(true)
	run.SetItalic(true)
	run.SetUnderline(WDUnderlineDouble)
	run.SetSize(16)
	run.SetColor("FF0000")
	run.SetFont("Arial")
	run.SetHighlight(WDColorIndexYellow)
	run.SetStrikethrough(true)
	run.SetDoubleStrikethrough(true)
	run.SetSmallCaps(true)
	run.SetAllCaps(true)
	run.SetShadow(true)
	run.SetOutline(true)
	run.SetEmboss(true)
	run.SetImprint(true)
	run.AddBreak(BreakTypePage)
	run.SetCharacterSpacing(40)
	run.SetKerning(24)
	run.SetBaselineShift(12)

	outputPath := filepath.Join(t.TempDir(), "formatted.docx")
	if err := doc.SaveAs(outputPath); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(outputPath)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	paragraphs := reopened.Paragraphs()
	if len(paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(paragraphs))
	}

	runs := paragraphs[0].Runs()
	if len(runs) != 1 {
		t.Fatalf("expected 1 run, got %d", len(runs))
	}

	reopenedRun := runs[0]
	if !reopenedRun.IsBold() {
		t.Errorf("expected run to be bold")
	}

	if !reopenedRun.IsItalic() {
		t.Errorf("expected run to be italic")
	}

	if reopenedRun.Underline() != WDUnderlineDouble {
		t.Errorf("expected underline %q, got %q", WDUnderlineDouble, reopenedRun.Underline())
	}

	if reopenedRun.Size() != 16 {
		t.Errorf("expected size 16pt, got %d", reopenedRun.Size())
	}

	if reopenedRun.Color() != "FF0000" {
		t.Errorf("expected color FF0000, got %s", reopenedRun.Color())
	}

	if reopenedRun.Font() != "Arial" {
		t.Errorf("expected font Arial, got %s", reopenedRun.Font())
	}

	if reopenedRun.Highlight() != WDColorIndexYellow {
		t.Errorf("expected highlight %q, got %q", WDColorIndexYellow, reopenedRun.Highlight())
	}

	if !reopenedRun.IsStrikethrough() {
		t.Errorf("expected run to be strikethrough")
	}

	if !reopenedRun.IsDoubleStrikethrough() {
		t.Errorf("expected run to be double strikethrough")
	}

	if !reopenedRun.IsSmallCaps() {
		t.Errorf("expected run to have small caps")
	}

	if !reopenedRun.IsAllCaps() {
		t.Errorf("expected run to have all caps")
	}

	if !reopenedRun.HasShadow() {
		t.Errorf("expected run to have shadow effect")
	}

	if !reopenedRun.HasOutline() {
		t.Errorf("expected run to have outline effect")
	}

	if !reopenedRun.IsEmbossed() {
		t.Errorf("expected run to be embossed")
	}

	if !reopenedRun.IsImprinted() {
		t.Errorf("expected run to be imprinted")
	}

	if !reopenedRun.HasBreak() {
		t.Errorf("expected run to contain a break")
	}

	if reopenedRun.BreakType() != BreakTypePage {
		t.Errorf("expected break type %q, got %q", BreakTypePage, reopenedRun.BreakType())
	}

	if spacing, ok := reopenedRun.CharacterSpacing(); !ok || spacing != 40 {
		t.Errorf("expected character spacing 40, got %d (ok=%v)", spacing, ok)
	}

	if kern, ok := reopenedRun.Kerning(); !ok || kern != 24 {
		t.Errorf("expected kerning 24, got %d (ok=%v)", kern, ok)
	}

	if shift, ok := reopenedRun.BaselineShift(); !ok || shift != 12 {
		t.Errorf("expected baseline shift 12, got %d (ok=%v)", shift, ok)
	}
}

func TestAddNumberedParagraphs(t *testing.T) {
	doc := NewDocument()
	doc.AddNumberedParagraph("First", 0)
	doc.AddNumberedParagraph("Second", 0)
	doc.AddBulletedParagraph("Bullet", 0)

	outputPath := filepath.Join(t.TempDir(), "numbered.docx")
	if err := doc.SaveAs(outputPath); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(outputPath)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	paragraphs := reopened.Paragraphs()
	if len(paragraphs) != 3 {
		t.Fatalf("expected 3 paragraphs, got %d", len(paragraphs))
	}

	if numID, level, ok := paragraphs[0].Numbering(); !ok || numID != 1 || level != 0 {
		t.Fatalf("expected first paragraph numbering (id=1, level=0), got (ok=%v, id=%d, level=%d)", ok, numID, level)
	}

	if numID, level, ok := paragraphs[1].Numbering(); !ok || numID != 1 || level != 0 {
		t.Fatalf("expected second paragraph numbering (id=1, level=0), got (ok=%v, id=%d, level=%d)", ok, numID, level)
	}

	if numID, level, ok := paragraphs[2].Numbering(); !ok || numID != 2 || level != 0 {
		t.Fatalf("expected third paragraph numbering (id=2, level=0), got (ok=%v, id=%d, level=%d)", ok, numID, level)
	}

	numberingPart := reopened.Numbering().Part()
	if numberingPart == nil {
		t.Fatal("expected numbering part to be present")
	}

	data := string(numberingPart.Data)
	if !strings.Contains(data, "w:numId=\"1\"") {
		t.Fatal("expected numbering part to contain numId=1 definition")
	}
	if !strings.Contains(data, "w:numId=\"2\"") {
		t.Fatal("expected numbering part to contain numId=2 definition")
	}
}

func TestInlinePictureRoundTrip(t *testing.T) {
	imgPath := filepath.Join(t.TempDir(), "sample.png")
	createTestImage(t, imgPath, 4, 3)

	doc := NewDocument()
	paragraph := doc.AddParagraph("Logo ")
	run := paragraph.AddRun("")
	pic, err := run.AddPicture(imgPath, 0, 0)
	if err != nil {
		t.Fatalf("AddPicture on run failed: %v", err)
	}
	if pic == nil {
		t.Fatalf("expected picture instance")
	}
	if !run.HasPicture() {
		t.Fatalf("expected run to report picture")
	}
	if pic.RelationshipID() == "" {
		t.Fatalf("expected relationship ID to be assigned")
	}
	if pic.WidthEMU() == 0 || pic.HeightEMU() == 0 {
		t.Fatalf("expected non-zero image dimensions")
	}
	data, err := pic.ImageData()
	if err != nil {
		t.Fatalf("ImageData failed: %v", err)
	}
	if len(data) == 0 {
		t.Fatalf("expected image bytes from ImageData")
	}

	_, docPic, err := doc.AddPicture(imgPath, 0, 0)
	if err != nil {
		t.Fatalf("Document.AddPicture failed: %v", err)
	}
	if docPic == nil {
		t.Fatalf("expected document-level picture instance")
	}

	doc.docPart.updateXMLData()
	if xml := string(doc.docPart.Part.Data); !strings.Contains(xml, "<w:drawing>") {
		t.Fatalf("document XML missing drawing element")
	}

	output := filepath.Join(t.TempDir(), "inline-picture.docx")
	if err := doc.SaveAs(output); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(output)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	var pictureRuns int
	for _, para := range reopened.Paragraphs() {
		for _, r := range para.Runs() {
			if r.Picture() != nil {
				pictureRuns++
			}
		}
	}
	if pictureRuns < 2 {
		t.Fatalf("expected at least two picture runs after reopen, got %d", pictureRuns)
	}

	foundData := false
	for _, para := range reopened.Paragraphs() {
		for _, r := range para.Runs() {
			if reopenedPic := r.Picture(); reopenedPic != nil {
				bytes, err := reopenedPic.ImageData()
				if err != nil {
					t.Fatalf("reopened ImageData failed: %v", err)
				}
				if len(bytes) == 0 {
					t.Fatalf("expected non-empty bytes for reopened picture")
				}
				foundData = true
				break
			}
		}
		if foundData {
			break
		}
	}
	if !foundData {
		t.Fatalf("expected to read image data from reopened document")
	}
}

func TestParagraphSpacingAndIndentation(t *testing.T) {
	doc := NewDocument()
	paragraph := doc.AddParagraph("Spacing test")
	paragraph.SetSpacing(240, 120, 360, "auto")
	paragraph.SetIndentation(720, 360, 0, 0)

	outputPath := filepath.Join(t.TempDir(), "spacing.docx")
	if err := doc.SaveAs(outputPath); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(outputPath)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	paragraphs := reopened.Paragraphs()
	if len(paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(paragraphs))
	}

	before, after, line, rule := paragraphs[0].Spacing()
	if before != 240 || after != 120 || line != 360 || rule != "auto" {
		t.Fatalf("unexpected spacing values: before=%d after=%d line=%d rule=%s", before, after, line, rule)
	}

	left, right, first, hanging := paragraphs[0].Indentation()
	if left != 720 || right != 360 || first != 0 || hanging != 0 {
		t.Fatalf("unexpected indentation values: left=%d right=%d first=%d hanging=%d", left, right, first, hanging)
	}
}

func TestParagraphKeepSettingsRoundTrip(t *testing.T) {
	doc := NewDocument()
	paragraph := doc.AddParagraph("Keep options")
	paragraph.SetKeepWithNext(true)
	paragraph.SetKeepLines(true)
	paragraph.SetPageBreakBefore(true)
	paragraph.SetWidowControl(false)

	outputPath := filepath.Join(t.TempDir(), "keep-settings.docx")
	if err := doc.SaveAs(outputPath); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(outputPath)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	paragraphs := reopened.Paragraphs()
	if len(paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(paragraphs))
	}

	reopenedParagraph := paragraphs[0]
	if !reopenedParagraph.KeepWithNext() {
		t.Fatalf("expected keep-with-next to be true")
	}
	if !reopenedParagraph.KeepLines() {
		t.Fatalf("expected keep-lines to be true")
	}
	if !reopenedParagraph.PageBreakBefore() {
		t.Fatalf("expected page-break-before to be true")
	}
	if reopenedParagraph.WidowControl() {
		t.Fatalf("expected widow control to be false")
	}
}

func TestParagraphBordersAndShadingRoundTrip(t *testing.T) {
	doc := NewDocument()
	paragraph := doc.AddParagraph("Bordered paragraph")
	paragraph.SetBorder(ParagraphBorderTop, ParagraphBorder{
		Style:  "single",
		Color:  "FF0000",
		Size:   12,
		Space:  80,
		Shadow: true,
	})
	paragraph.SetBorder(ParagraphBorderBottom, ParagraphBorder{
		Style: "double",
		Color: "00FF00",
		Size:  8,
	})
	paragraph.SetShading("solid", "FFFFAA", "000000")

	outputPath := filepath.Join(t.TempDir(), "paragraph-borders.docx")
	if err := doc.SaveAs(outputPath); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(outputPath)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	paras := reopened.Paragraphs()
	if len(paras) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(paras))
	}

	reopenedParagraph := paras[0]
	top, ok := reopenedParagraph.Border(ParagraphBorderTop)
	if !ok {
		t.Fatalf("expected top border to be present")
	}
	if top.Style != "single" {
		t.Fatalf("expected top border style 'single', got %q", top.Style)
	}
	if top.Color != "FF0000" {
		t.Fatalf("expected top border color FF0000, got %q", top.Color)
	}
	if top.Size != 12 {
		t.Fatalf("expected top border size 12, got %d", top.Size)
	}
	if top.Space != 80 {
		t.Fatalf("expected top border space 80, got %d", top.Space)
	}
	if !top.Shadow {
		t.Fatalf("expected top border shadow to be true")
	}

	bottom, ok := reopenedParagraph.Border(ParagraphBorderBottom)
	if !ok {
		t.Fatalf("expected bottom border to be present")
	}
	if bottom.Style != "double" {
		t.Fatalf("expected bottom border style 'double', got %q", bottom.Style)
	}
	if bottom.Color != "00FF00" {
		t.Fatalf("expected bottom border color 00FF00, got %q", bottom.Color)
	}
	if bottom.Size != 8 {
		t.Fatalf("expected bottom border size 8, got %d", bottom.Size)
	}

	shading, ok := reopenedParagraph.Shading()
	if !ok {
		t.Fatalf("expected shading to be present")
	}
	if shading.Pattern != "solid" {
		t.Fatalf("expected shading pattern 'solid', got %q", shading.Pattern)
	}
	if shading.Fill != "FFFFAA" {
		t.Fatalf("expected shading fill FFFF-AA, got %q", shading.Fill)
	}
	if shading.Color != "000000" {
		t.Fatalf("expected shading color 000000, got %q", shading.Color)
	}

	paragraphXML := string(reopened.docPart.Part.Data)
	if !strings.Contains(paragraphXML, "<w:pBdr>") {
		t.Fatalf("expected paragraph XML to contain border definition")
	}
	if !strings.Contains(paragraphXML, "<w:shd") {
		t.Fatalf("expected paragraph XML to contain shading definition")
	}
}

func TestParagraphHyperlinkRoundTrip(t *testing.T) {
	doc := NewDocument()
	paragraph := doc.AddParagraph()
	run := paragraph.AddHyperlink("Example", "https://example.com")
	run.SetColor("0000FF")
	run.SetUnderline(WDUnderlineSingle)

	outputPath := filepath.Join(t.TempDir(), "hyperlink.docx")
	if err := doc.SaveAs(outputPath); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(outputPath)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	paragraphs := reopened.Paragraphs()
	if len(paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(paragraphs))
	}

	runs := paragraphs[0].Runs()
	if len(runs) != 1 {
		t.Fatalf("expected 1 run, got %d", len(runs))
	}

	reopenedRun := runs[0]
	if !reopenedRun.HasHyperlink() {
		t.Fatalf("expected run to be a hyperlink")
	}
	if reopenedRun.HyperlinkURL() != "https://example.com" {
		t.Fatalf("expected hyperlink URL 'https://example.com', got %q", reopenedRun.HyperlinkURL())
	}
	if reopenedRun.Text() != "Example" {
		t.Fatalf("expected hyperlink text 'Example', got %q", reopenedRun.Text())
	}
	if reopenedRun.Underline() != WDUnderlineSingle {
		t.Fatalf("expected underline %q, got %q", WDUnderlineSingle, reopenedRun.Underline())
	}
}

func TestParagraphTabStopsRoundTrip(t *testing.T) {
	doc := NewDocument()
	paragraph := doc.AddParagraph("Tabs")
	paragraph.AddTabStop(720, WDTabAlignmentCenter, WDTabLeaderDot)
	paragraph.AddTabStop(1440, WDTabAlignmentRight, WDTabLeaderNone)

	outputPath := filepath.Join(t.TempDir(), "tabstops.docx")
	if err := doc.SaveAs(outputPath); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(outputPath)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	paragraphs := reopened.Paragraphs()
	if len(paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(paragraphs))
	}

	stops := paragraphs[0].TabStops()
	if len(stops) != 2 {
		t.Fatalf("expected 2 tab stops, got %d", len(stops))
	}

	if stops[0].Position != 720 {
		t.Fatalf("expected first tab stop position 720, got %d", stops[0].Position)
	}
	if stops[0].Alignment != WDTabAlignmentCenter {
		t.Fatalf("expected first tab stop alignment %q, got %q", WDTabAlignmentCenter, stops[0].Alignment)
	}
	if stops[0].Leader != WDTabLeaderDot {
		t.Fatalf("expected first tab stop leader %q, got %q", WDTabLeaderDot, stops[0].Leader)
	}

	if stops[1].Position != 1440 {
		t.Fatalf("expected second tab stop position 1440, got %d", stops[1].Position)
	}
	if stops[1].Alignment != WDTabAlignmentRight {
		t.Fatalf("expected second tab stop alignment %q, got %q", WDTabAlignmentRight, stops[1].Alignment)
	}
	if stops[1].Leader != WDTabLeaderNone {
		t.Fatalf("expected second tab stop leader %q, got %q", WDTabLeaderNone, stops[1].Leader)
	}
}

func TestOpenDocumentParsesTables(t *testing.T) {
	doc := NewDocument()
	table := doc.AddTable(2, 2)

	table.Row(0).Cell(0).SetText("A1")
	table.Row(0).Cell(1).SetText("A2")
	table.Row(1).Cell(0).SetText("B1")
	table.Row(1).Cell(1).SetText("B2")
	table.Row(1).Cell(1).SetWidth(2400)

	outputPath := filepath.Join(t.TempDir(), "table.docx")
	if err := doc.SaveAs(outputPath); err != nil {
		t.Fatalf("SaveAs failed: %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	reopened, err := OpenDocument(outputPath)
	if err != nil {
		t.Fatalf("OpenDocument failed: %v", err)
	}
	defer reopened.Close()

	tables := reopened.Tables()
	if len(tables) != 1 {
		t.Fatalf("expected 1 table, got %d", len(tables))
	}

	reopenedTable := tables[0]
	if len(reopenedTable.Rows()) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(reopenedTable.Rows()))
	}

	if len(reopenedTable.Row(0).Cells()) != 2 {
		t.Fatalf("expected 2 cells in first row, got %d", len(reopenedTable.Row(0).Cells()))
	}

	if reopenedTable.Row(0).Cell(0).Text() != "A1" {
		t.Errorf("expected cell (0,0) text to be 'A1', got %q", reopenedTable.Row(0).Cell(0).Text())
	}

	if reopenedTable.Row(1).Cell(1).Text() != "B2" {
		t.Errorf("expected cell (1,1) text to be 'B2', got %q", reopenedTable.Row(1).Cell(1).Text())
	}

	if reopenedTable.Row(1).Cell(1).Width() != 2400 {
		t.Errorf("expected cell (1,1) width to be 2400, got %d", reopenedTable.Row(1).Cell(1).Width())
	}
}

func createTestImage(t *testing.T, path string, width, height int) {
	t.Helper()
	img := image.NewRGBA(image.Rect(0, 0, width, height))
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			img.Set(x, y, color.RGBA{R: uint8(50 * (x + 1)), G: uint8(40 * (y + 1)), B: 200, A: 255})
		}
	}
	file, err := os.Create(path)
	if err != nil {
		t.Fatalf("failed to create test image: %v", err)
	}
	defer file.Close()
	if err := png.Encode(file, img); err != nil {
		t.Fatalf("failed to encode test image: %v", err)
	}
}
