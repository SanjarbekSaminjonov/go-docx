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

func TestHeaderFooterRoundTrip(t *testing.T) {
	doc := NewDocument()
	sections := doc.Sections()
	if len(sections) == 0 {
		t.Fatalf("expected at least one section")
	}
	header, err := sections[0].Header()
	if err != nil {
		t.Fatalf("Header() failed: %v", err)
	}
	footer, err := sections[0].Footer()
	if err != nil {
		t.Fatalf("Footer() failed: %v", err)
	}
	header.AddParagraph("Primary header text")
	footer.AddParagraph("Primary footer text")
	doc.docPart.updateXMLData()
	mainXML := string(doc.docPart.Part.Data)
	if !strings.Contains(mainXML, "<w:headerReference") {
		t.Fatalf("expected document XML to contain header reference")
	}
	if !strings.Contains(mainXML, "<w:footerReference") {
		t.Fatalf("expected document XML to contain footer reference")
	}

	outputPath := filepath.Join(t.TempDir(), "header-footer.docx")
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

	reopenedSections := reopened.Sections()
	if len(reopenedSections) == 0 {
		t.Fatalf("expected reopened document to have sections")
	}
	reopenedHeader, err := reopenedSections[0].Header()
	if err != nil {
		t.Fatalf("Header() on reopened doc failed: %v", err)
	}
	reopenedFooter, err := reopenedSections[0].Footer()
	if err != nil {
		t.Fatalf("Footer() on reopened doc failed: %v", err)
	}

	headerParas := reopenedHeader.Paragraphs()
	if len(headerParas) != 1 {
		t.Fatalf("expected 1 header paragraph, got %d", len(headerParas))
	}
	if headerParas[0].Text() != "Primary header text" {
		t.Fatalf("unexpected header text: %q", headerParas[0].Text())
	}
	footerParas := reopenedFooter.Paragraphs()
	if len(footerParas) != 1 {
		t.Fatalf("expected 1 footer paragraph, got %d", len(footerParas))
	}
	if footerParas[0].Text() != "Primary footer text" {
		t.Fatalf("unexpected footer text: %q", footerParas[0].Text())
	}

	headerXML := string(reopenedHeader.part.Data)
	if !strings.Contains(headerXML, "Primary header text") {
		t.Fatalf("expected header XML to contain header text")
	}
	footerXML := string(reopenedFooter.part.Data)
	if !strings.Contains(footerXML, "Primary footer text") {
		t.Fatalf("expected footer XML to contain footer text")
	}
}

func TestTableFormattingRoundTrip(t *testing.T) {
	doc := NewDocument()
	table := doc.AddTable(2, 2)

	table.SetBorder(TableBorderTop, TableBorder{Style: "single", Color: "FF0000", Size: 12, Space: 40})
	table.SetBorder(TableBorderBottom, TableBorder{Style: "double", Color: "00FF00", Size: 8})
	table.SetShading("solid", "CCCCCC", "000000")
	table.SetCellMargins(120, 240, 360, 480)

	cell := table.Row(0).Cell(0)
	cell.SetText("merged")
	cell.SetShading("solid", "FFFFAA", "000000")
	cell.SetBorder(TableBorderLeft, TableBorder{Style: "single", Color: "0000FF", Size: 6})

	if err := table.MergeCellsHorizontally(0, 0, 1); err != nil {
		t.Fatalf("MergeCellsHorizontally failed: %v", err)
	}
	if err := table.MergeCellsVertically(0, 0, 1); err != nil {
		t.Fatalf("MergeCellsVertically failed: %v", err)
	}

	outputPath := filepath.Join(t.TempDir(), "table-formatting.docx")
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

	top, ok := reopenedTable.Border(TableBorderTop)
	if !ok || top.Style != "single" || top.Color != "FF0000" || top.Size != 12 || top.Space != 40 {
		t.Fatalf("expected top border to match, got %+v", top)
	}
	bottom, ok := reopenedTable.Border(TableBorderBottom)
	if !ok || bottom.Style != "double" || bottom.Color != "00FF00" || bottom.Size != 8 {
		t.Fatalf("expected bottom border to match, got %+v", bottom)
	}

	shading, ok := reopenedTable.Shading()
	if !ok || shading.Pattern != "solid" || shading.Fill != "CCCCCC" || shading.Color != "000000" {
		t.Fatalf("expected table shading to match, got %+v", shading)
	}

	margins, ok := reopenedTable.CellMargins()
	if !ok || margins.Top == nil || *margins.Top != 120 || margins.Left == nil || *margins.Left != 240 || margins.Bottom == nil || *margins.Bottom != 360 || margins.Right == nil || *margins.Right != 480 {
		t.Fatalf("expected cell margins to match, got %+v", margins)
	}

	reopenedCell := reopenedTable.Row(0).Cell(0)
	cellShading, ok := reopenedCell.Shading()
	if !ok || cellShading.Pattern != "solid" || cellShading.Fill != "FFFFAA" || cellShading.Color != "000000" {
		t.Fatalf("expected cell shading to match, got %+v", cellShading)
	}
	left, ok := reopenedCell.Border(TableBorderLeft)
	if !ok || left.Style != "single" || left.Color != "0000FF" || left.Size != 6 {
		t.Fatalf("expected cell left border to match, got %+v", left)
	}

	if reopenedCell.GridSpan() != 2 {
		t.Fatalf("expected merged cell grid span 2, got %d", reopenedCell.GridSpan())
	}
	if reopenedCell.VerticalMerge() != TableVerticalMergeRestart {
		t.Fatalf("expected vertical merge restart, got %q", reopenedCell.VerticalMerge())
	}

	row2Cell := reopenedTable.Row(1).Cell(0)
	if row2Cell.VerticalMerge() != TableVerticalMergeContinue {
		t.Fatalf("expected vertical merge continue on second row, got %q", row2Cell.VerticalMerge())
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

func TestGetXML(t *testing.T) {
	// Test GetXML with a new document
	t.Run("NewDocument", func(t *testing.T) {
		doc := NewDocument()

		xmlContent, err := doc.GetXML()
		if err != nil {
			t.Fatalf("GetXML() failed: %v", err)
		}

		if xmlContent == "" {
			t.Fatal("GetXML() returned empty string")
		}

		// Check for basic XML structure
		if !strings.Contains(xmlContent, "<w:document") {
			t.Error("XML content should contain <w:document element")
		}

		if !strings.Contains(xmlContent, "<w:body>") {
			t.Error("XML content should contain <w:body> element")
		}
	})

	// Test GetXML with content
	t.Run("WithContent", func(t *testing.T) {
		doc := NewDocument()

		// Add some content
		doc.AddParagraph("Test paragraph")
		_, err := doc.AddHeading("Test Heading", 1)
		if err != nil {
			t.Fatalf("AddHeading() failed: %v", err)
		}

		xmlContent, err := doc.GetXML()
		if err != nil {
			t.Fatalf("GetXML() failed: %v", err)
		}

		// Check that content is reflected in XML
		if !strings.Contains(xmlContent, "Test paragraph") {
			t.Error("XML content should contain 'Test paragraph'")
		}

		if !strings.Contains(xmlContent, "Test Heading") {
			t.Error("XML content should contain 'Test Heading'")
		}

		// Check for paragraph structure
		if !strings.Contains(xmlContent, "<w:p>") {
			t.Error("XML content should contain paragraph elements")
		}

		if !strings.Contains(xmlContent, "<w:r>") {
			t.Error("XML content should contain run elements")
		}

		if !strings.Contains(xmlContent, "<w:t>") {
			t.Error("XML content should contain text elements")
		}
	})

	// Test GetXML with complex content
	t.Run("WithComplexContent", func(t *testing.T) {
		doc := NewDocument()

		// Add various types of content
		p := doc.AddParagraph()
		p.AddRun("Bold text").SetBold(true)
		p.AddRun(" and ").SetBold(false)
		p.AddRun("italic text").SetItalic(true)

		// Add a table (just test structure, not content for now)
		table := doc.AddTable(2, 2)
		table.Row(0).Cell(0).SetText("Cell 1")
		table.Row(0).Cell(1).SetText("Cell 2")

		xmlContent, err := doc.GetXML()
		if err != nil {
			t.Fatalf("GetXML() failed: %v", err)
		}

		// Check for table structure
		if !strings.Contains(xmlContent, "<w:tbl>") {
			t.Error("XML content should contain table elements")
		}

		if !strings.Contains(xmlContent, "<w:tr>") {
			t.Error("XML content should contain table row elements")
		}

		if !strings.Contains(xmlContent, "<w:tc>") {
			t.Error("XML content should contain table cell elements")
		}

		// Check for formatting
		if !strings.Contains(xmlContent, "<w:b/>") {
			t.Error("XML content should contain bold formatting")
		}

		if !strings.Contains(xmlContent, "<w:i/>") {
			t.Error("XML content should contain italic formatting")
		}

		// Check for text content in runs
		if !strings.Contains(xmlContent, "Bold text") {
			t.Error("XML content should contain 'Bold text'")
		}

		if !strings.Contains(xmlContent, "italic text") {
			t.Error("XML content should contain 'italic text'")
		}
	})

	// Test GetXML after opening an existing document
	t.Run("OpenedDocument", func(t *testing.T) {
		// Create and save a document first
		tempFile := filepath.Join(t.TempDir(), "test_getxml.docx")

		doc := NewDocument()
		doc.AddParagraph("Original content")
		if err := doc.SaveAs(tempFile); err != nil {
			t.Fatalf("Failed to save document: %v", err)
		}
		doc.Close()

		// Open the document and test GetXML
		reopened, err := OpenDocument(tempFile)
		if err != nil {
			t.Fatalf("Failed to open document: %v", err)
		}
		defer reopened.Close()

		xmlContent, err := reopened.GetXML()
		if err != nil {
			t.Fatalf("GetXML() failed on opened document: %v", err)
		}

		if !strings.Contains(xmlContent, "Original content") {
			t.Error("XML content should contain original content from saved document")
		}
	})

	// Test GetXML error case (nil docPart)
	t.Run("ErrorCase", func(t *testing.T) {
		doc := &Document{} // Document with nil docPart

		_, err := doc.GetXML()
		if err == nil {
			t.Error("GetXML() should return error when docPart is nil")
		}

		expectedError := "document has no main document part"
		if !strings.Contains(err.Error(), expectedError) {
			t.Errorf("Expected error to contain '%s', got: %v", expectedError, err)
		}
	})
}

func TestInsertTableAfterParagraph(t *testing.T) {
	doc := NewDocument()

	// Add some paragraphs
	_ = doc.AddParagraph("First paragraph")
	p2 := doc.AddParagraph("Second paragraph")
	_ = doc.AddParagraph("Third paragraph")

	// Insert table after second paragraph
	table, err := doc.InsertTableAfterParagraph(p2, 2, 3)
	if err != nil {
		t.Fatalf("InsertTableAfterParagraph() failed: %v", err)
	}

	if table == nil {
		t.Fatal("InsertTableAfterParagraph() returned nil table")
	}

	// Verify table structure
	if len(table.Rows()) != 2 {
		t.Errorf("Expected 2 rows, got %d", len(table.Rows()))
	}

	if len(table.Row(0).Cells()) != 3 {
		t.Errorf("Expected 3 cells, got %d", len(table.Row(0).Cells()))
	}

	// Verify order of elements
	bodyElements := doc.docPart.bodyElements

	// Find paragraphs and table in bodyElements (ignoring sections)
	var foundElements []string
	for _, elem := range bodyElements {
		if elem.paragraph != nil {
			foundElements = append(foundElements, "paragraph")
		} else if elem.table != nil {
			foundElements = append(foundElements, "table")
		}
	}

	// Expected order: paragraph, paragraph, table, paragraph
	expectedOrder := []string{"paragraph", "paragraph", "table", "paragraph"}

	if len(foundElements) != len(expectedOrder) {
		t.Fatalf("Expected %d elements (paragraphs+tables), got %d", len(expectedOrder), len(foundElements))
	}

	for i, expected := range expectedOrder {
		if foundElements[i] != expected {
			t.Errorf("Element at position %d: expected %s, got %s", i, expected, foundElements[i])
		}
	}

	// Test error case: nil paragraph
	_, err = doc.InsertTableAfterParagraph(nil, 2, 2)
	if err == nil {
		t.Error("InsertTableAfterParagraph() should return error for nil paragraph")
	}

	// Test error case: paragraph not in document
	otherDoc := NewDocument()
	otherP := otherDoc.AddParagraph("Other paragraph")
	_, err = doc.InsertTableAfterParagraph(otherP, 2, 2)
	if err == nil {
		t.Error("InsertTableAfterParagraph() should return error for paragraph not in document")
	}

	// Test round trip
	tempFile := filepath.Join(t.TempDir(), "test_insert_table.docx")
	if err := doc.SaveAs(tempFile); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}
	doc.Close()

	reopened, err := OpenDocument(tempFile)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer reopened.Close()

	// Verify structure after reopening
	if len(reopened.Paragraphs()) != 3 {
		t.Errorf("Expected 3 paragraphs after reopening, got %d", len(reopened.Paragraphs()))
	}

	if len(reopened.Tables()) != 1 {
		t.Errorf("Expected 1 table after reopening, got %d", len(reopened.Tables()))
	}

	// Verify order is preserved
	if reopened.Paragraphs()[0].Text() != "First paragraph" {
		t.Errorf("First paragraph text mismatch: got %q", reopened.Paragraphs()[0].Text())
	}

	if reopened.Paragraphs()[2].Text() != "Third paragraph" {
		t.Errorf("Third paragraph text mismatch: got %q", reopened.Paragraphs()[2].Text())
	}
}

func TestRemoveParagraph(t *testing.T) {
	doc := NewDocument()

	// Add some paragraphs
	p1 := doc.AddParagraph("First paragraph")
	p2 := doc.AddParagraph("Second paragraph")
	p3 := doc.AddParagraph("Third paragraph")

	// Verify initial count
	if len(doc.Paragraphs()) != 3 {
		t.Fatalf("Expected 3 paragraphs, got %d", len(doc.Paragraphs()))
	}

	// Remove middle paragraph
	err := doc.RemoveParagraph(p2)
	if err != nil {
		t.Fatalf("RemoveParagraph() failed: %v", err)
	}

	// Verify count after removal
	if len(doc.Paragraphs()) != 2 {
		t.Errorf("Expected 2 paragraphs after removal, got %d", len(doc.Paragraphs()))
	}

	// Verify remaining paragraphs
	if doc.Paragraphs()[0] != p1 {
		t.Error("First paragraph should still be p1")
	}

	if doc.Paragraphs()[1] != p3 {
		t.Error("Second paragraph should now be p3")
	}

	// Test error case: nil paragraph
	err = doc.RemoveParagraph(nil)
	if err == nil {
		t.Error("RemoveParagraph() should return error for nil paragraph")
	}

	// Test error case: paragraph already removed
	err = doc.RemoveParagraph(p2)
	if err == nil {
		t.Error("RemoveParagraph() should return error for already removed paragraph")
	}

	// Test error case: paragraph not in document
	otherDoc := NewDocument()
	otherP := otherDoc.AddParagraph("Other paragraph")
	err = doc.RemoveParagraph(otherP)
	if err == nil {
		t.Error("RemoveParagraph() should return error for paragraph not in document")
	}

	// Test round trip
	tempFile := filepath.Join(t.TempDir(), "test_remove_paragraph.docx")
	if err := doc.SaveAs(tempFile); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}
	doc.Close()

	reopened, err := OpenDocument(tempFile)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer reopened.Close()

	// Verify structure after reopening
	if len(reopened.Paragraphs()) != 2 {
		t.Errorf("Expected 2 paragraphs after reopening, got %d", len(reopened.Paragraphs()))
	}

	if reopened.Paragraphs()[0].Text() != "First paragraph" {
		t.Errorf("First paragraph text mismatch: got %q", reopened.Paragraphs()[0].Text())
	}

	if reopened.Paragraphs()[1].Text() != "Third paragraph" {
		t.Errorf("Second paragraph text mismatch: got %q", reopened.Paragraphs()[1].Text())
	}
}

func TestRemoveTable(t *testing.T) {
	doc := NewDocument()

	// Add paragraphs and tables
	doc.AddParagraph("First paragraph")
	table1 := doc.AddTable(2, 2)
	table1.Row(0).Cell(0).SetText("Table 1")
	doc.AddParagraph("Second paragraph")
	table2 := doc.AddTable(3, 3)
	table2.Row(0).Cell(0).SetText("Table 2")
	doc.AddParagraph("Third paragraph")

	// Verify initial count
	if len(doc.Tables()) != 2 {
		t.Fatalf("Expected 2 tables, got %d", len(doc.Tables()))
	}

	// Remove first table
	err := doc.RemoveTable(table1)
	if err != nil {
		t.Fatalf("RemoveTable() failed: %v", err)
	}

	// Verify count after removal
	if len(doc.Tables()) != 1 {
		t.Errorf("Expected 1 table after removal, got %d", len(doc.Tables()))
	}

	// Verify remaining table
	if doc.Tables()[0] != table2 {
		t.Error("Remaining table should be table2")
	}

	// Test error case: nil table
	err = doc.RemoveTable(nil)
	if err == nil {
		t.Error("RemoveTable() should return error for nil table")
	}

	// Test error case: table already removed
	err = doc.RemoveTable(table1)
	if err == nil {
		t.Error("RemoveTable() should return error for already removed table")
	}

	// Test round trip
	tempFile := filepath.Join(t.TempDir(), "test_remove_table.docx")
	if err := doc.SaveAs(tempFile); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}
	doc.Close()

	reopened, err := OpenDocument(tempFile)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer reopened.Close()

	// Verify structure after reopening
	if len(reopened.Tables()) != 1 {
		t.Errorf("Expected 1 table after reopening, got %d", len(reopened.Tables()))
	}

	if len(reopened.Paragraphs()) != 3 {
		t.Errorf("Expected 3 paragraphs after reopening, got %d", len(reopened.Paragraphs()))
	}
}

func TestRemoveSection(t *testing.T) {
	doc := NewDocument()

	// Add content with sections
	doc.AddParagraph("First paragraph")
	section1 := doc.AddSection(SectionStartNewPage)
	section1.SetPageSize(11906, 16838)

	doc.AddParagraph("Second paragraph")
	section2 := doc.AddSection(SectionStartContinuous)
	section2.SetPageSize(16838, 11906) // Landscape

	doc.AddParagraph("Third paragraph")

	// Verify initial count
	// Note: NewDocument() creates a default section, so we have 3 sections total
	initialSectionCount := len(doc.Sections())
	if initialSectionCount < 2 {
		t.Fatalf("Expected at least 2 sections, got %d", initialSectionCount)
	}

	// Remove first section
	err := doc.RemoveSection(section1)
	if err != nil {
		t.Fatalf("RemoveSection() failed: %v", err)
	}

	// Verify count after removal (should be one less than initial)
	if len(doc.Sections()) != initialSectionCount-1 {
		t.Errorf("Expected %d sections after removal, got %d", initialSectionCount-1, len(doc.Sections()))
	}

	// Verify section2 still exists in the sections list
	found := false
	for _, s := range doc.Sections() {
		if s == section2 {
			found = true
			break
		}
	}
	if !found {
		t.Error("section2 should still be in the sections list")
	}

	// Test error case: nil section
	err = doc.RemoveSection(nil)
	if err == nil {
		t.Error("RemoveSection() should return error for nil section")
	}

	// Test error case: section already removed
	err = doc.RemoveSection(section1)
	if err == nil {
		t.Error("RemoveSection() should return error for already removed section")
	}
}

func TestGetRowGetCell(t *testing.T) {
	doc := NewDocument()

	// Create a table
	table := doc.AddTable(2, 3)

	// Test GetRow (should be same as Row)
	row1 := table.GetRow(0)
	if row1 == nil {
		t.Fatal("GetRow(0) returned nil")
	}

	row2 := table.Row(0)
	if row1 != row2 {
		t.Error("GetRow() and Row() should return the same reference")
	}

	// Test GetCell (should be same as Cell)
	cell1 := row1.GetCell(0)
	if cell1 == nil {
		t.Fatal("GetCell(0) returned nil")
	}

	cell2 := row1.Cell(0)
	if cell1 != cell2 {
		t.Error("GetCell() and Cell() should return the same reference")
	}

	// Test chaining methods as shown in user's example
	table.GetRow(0).GetCell(1).AddParagraph().AddRun("Test Value").SetBold(true)
	table.GetRow(1).GetCell(0).AddParagraph().AddRun("Another Value").SetItalic(true)

	// Verify content was added (trim whitespace because cells have default empty paragraph)
	cellText := strings.TrimSpace(table.GetRow(0).GetCell(1).Text())
	if cellText != "Test Value" {
		t.Errorf("Expected 'Test Value', got '%s'", cellText)
	}

	// Test out of bounds
	if table.GetRow(10) != nil {
		t.Error("GetRow(10) should return nil for out of bounds")
	}

	if row1.GetCell(10) != nil {
		t.Error("GetCell(10) should return nil for out of bounds")
	}
}

func TestClearRuns(t *testing.T) {
	doc := NewDocument()

	// Create a paragraph with multiple runs
	p := doc.AddParagraph()
	p.AddRun("First run ")
	p.AddRun("Second run ")
	p.AddRun("Third run")

	// Verify initial state
	if len(p.Runs()) != 3 {
		t.Fatalf("Expected 3 runs, got %d", len(p.Runs()))
	}

	if p.Text() != "First run Second run Third run" {
		t.Errorf("Expected 'First run Second run Third run', got '%s'", p.Text())
	}

	// Clear all runs
	p.ClearRuns()

	// Verify runs are cleared
	if len(p.Runs()) != 0 {
		t.Errorf("Expected 0 runs after ClearRuns(), got %d", len(p.Runs()))
	}

	if p.Text() != "" {
		t.Errorf("Expected empty text after ClearRuns(), got '%s'", p.Text())
	}

	// Add new run after clearing
	p.AddRun("New content")

	if len(p.Runs()) != 1 {
		t.Fatalf("Expected 1 run after adding new content, got %d", len(p.Runs()))
	}

	if p.Text() != "New content" {
		t.Errorf("Expected 'New content', got '%s'", p.Text())
	}
}

func TestTemplateReplacement(t *testing.T) {
	doc := NewDocument()

	// Add template content
	doc.AddParagraph("Document Title: ${title}")
	doc.AddParagraph("")
	placeholder := doc.AddParagraph("${signers}")
	doc.AddParagraph("")
	doc.AddParagraph("End of document")

	// Replace ${title}
	for _, p := range doc.Paragraphs() {
		text := p.Text()
		if strings.Contains(text, "${title}") {
			p.ClearRuns()
			p.AddRun(strings.ReplaceAll(text, "${title}", "Important Contract"))
		}
	}

	// Replace ${signers} with table
	table, err := doc.InsertTableAfterParagraph(placeholder, 2, 2)
	if err != nil {
		t.Fatalf("InsertTableAfterParagraph() failed: %v", err)
	}

	// Fill table using GetRow/GetCell
	table.GetRow(0).GetCell(0).AddParagraph().AddRun("Name").SetBold(true)
	table.GetRow(0).GetCell(1).AddParagraph().AddRun("Signature").SetBold(true)
	table.GetRow(1).GetCell(0).AddParagraph().AddRun("John Doe")
	table.GetRow(1).GetCell(1).AddParagraph().AddRun("_________________")

	// Remove placeholder
	if err := doc.RemoveParagraph(placeholder); err != nil {
		t.Fatalf("RemoveParagraph() failed: %v", err)
	}

	// Verify results
	found := false
	for _, p := range doc.Paragraphs() {
		if strings.Contains(p.Text(), "Important Contract") {
			found = true
			break
		}
	}
	if !found {
		t.Error("Title replacement did not work")
	}

	if len(doc.Tables()) != 1 {
		t.Errorf("Expected 1 table, got %d", len(doc.Tables()))
	}

	cellText := strings.TrimSpace(table.GetRow(0).GetCell(0).Text())
	if cellText != "Name" {
		t.Errorf("Expected 'Name' in first cell, got '%s'", cellText)
	}

	// Test round trip
	tempFile := filepath.Join(t.TempDir(), "test_template.docx")
	if err := doc.SaveAs(tempFile); err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}
	doc.Close()

	reopened, err := OpenDocument(tempFile)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}
	defer reopened.Close()

	// Verify after reopening
	if len(reopened.Tables()) != 1 {
		t.Errorf("Expected 1 table after reopening, got %d", len(reopened.Tables()))
	}
}
