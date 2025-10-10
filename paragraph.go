package docx

import (
	"fmt"
	"strings"
)

// Paragraph represents a paragraph in a Word document
type Paragraph struct {
	owner            *DocumentPart
	runs             []*Run
	style            string
	alignment        WDAlignParagraph
	numberingApplied bool
	numberingID      int
	numberingLevel   int
	indentLeft       int
	indentRight      int
	indentFirstLine  int
	indentHanging    int
	spacingBefore    int
	spacingAfter     int
	spacingLine      int
	spacingLineRule  string
	tabStops         []TabStop
	keepWithNext     *bool
	keepLines        *bool
	pageBreakBefore  *bool
	widowControl     *bool
}

// TabStop represents a paragraph tab stop configuration
type TabStop struct {
	Position  int
	Alignment WDTabAlignment
	Leader    WDTabLeader
}

// NewParagraph creates a new paragraph
func NewParagraph() *Paragraph {
	return &Paragraph{
		runs:      make([]*Run, 0),
		tabStops:  make([]TabStop, 0),
		alignment: WDAlignParagraphLeft,
	}
}

// AddRun adds a new run to the paragraph
func (p *Paragraph) AddRun(text string) *Run {
	run := NewRun(text)
	run.owner = p.owner
	p.runs = append(p.runs, run)
	return run
}

// AddPicture creates a new run containing an inline picture
func (p *Paragraph) AddPicture(path string, widthEMU, heightEMU int64) (*Run, *Picture, error) {
	if p.owner == nil {
		return nil, nil, fmt.Errorf("paragraph is not attached to a document")
	}
	run := p.AddRun("")
	picture, err := run.AddPicture(path, widthEMU, heightEMU)
	if err != nil {
		p.runs = p.runs[:len(p.runs)-1]
		return nil, nil, err
	}
	return run, picture, nil
}

// AddHyperlink adds a run with hyperlink formatting
func (p *Paragraph) AddHyperlink(text, url string) *Run {
	run := p.AddRun(text)
	run.SetHyperlink(url)
	return run
}

// SetSpacing configures paragraph spacing (values in twentieths of a point)
func (p *Paragraph) SetSpacing(before, after, line int, lineRule string) {
	p.spacingBefore = before
	p.spacingAfter = after
	p.spacingLine = line
	p.spacingLineRule = lineRule
}

// Spacing returns the spacing configuration
func (p *Paragraph) Spacing() (before, after, line int, lineRule string) {
	return p.spacingBefore, p.spacingAfter, p.spacingLine, p.spacingLineRule
}

// SetIndentation configures paragraph indentation (values in twentieths of a point)
func (p *Paragraph) SetIndentation(left, right, firstLine, hanging int) {
	p.indentLeft = left
	p.indentRight = right
	p.indentFirstLine = firstLine
	p.indentHanging = hanging
}

// Indentation returns the indentation configuration
func (p *Paragraph) Indentation() (left, right, firstLine, hanging int) {
	return p.indentLeft, p.indentRight, p.indentFirstLine, p.indentHanging
}

// SetStyle sets the paragraph style
func (p *Paragraph) SetStyle(style string) {
	p.style = style
}

// Style returns the paragraph style
func (p *Paragraph) Style() string {
	return p.style
}

// SetAlignment sets the paragraph alignment
func (p *Paragraph) SetAlignment(alignment WDAlignParagraph) {
	p.alignment = alignment
}

// Alignment returns the paragraph alignment
func (p *Paragraph) Alignment() WDAlignParagraph {
	return p.alignment
}

// SetNumbering applies numbering to the paragraph using the specified numbering ID and level
func (p *Paragraph) SetNumbering(numID, level int) {
	p.numberingApplied = true
	p.numberingID = numID
	if level < 0 {
		level = 0
	}
	p.numberingLevel = level
}

// ClearNumbering removes numbering from the paragraph
func (p *Paragraph) ClearNumbering() {
	p.numberingApplied = false
}

// HasNumbering reports whether numbering is applied to the paragraph
func (p *Paragraph) HasNumbering() bool {
	return p.numberingApplied
}

// Numbering returns the numbering ID and level for the paragraph
func (p *Paragraph) Numbering() (numID int, level int, ok bool) {
	if !p.numberingApplied {
		return 0, 0, false
	}
	return p.numberingID, p.numberingLevel, true
}

// Runs returns all runs in the paragraph
func (p *Paragraph) Runs() []*Run {
	return p.runs
}

// Text returns the combined text of all runs in the paragraph
func (p *Paragraph) Text() string {
	var text strings.Builder
	for _, run := range p.runs {
		text.WriteString(run.Text())
	}
	return text.String()
}

// Clear removes all runs from the paragraph
func (p *Paragraph) Clear() {
	p.runs = p.runs[:0]
	p.numberingApplied = false
	p.indentLeft = 0
	p.indentRight = 0
	p.indentFirstLine = 0
	p.indentHanging = 0
	p.spacingBefore = 0
	p.spacingAfter = 0
	p.spacingLine = 0
	p.spacingLineRule = ""
	p.tabStops = p.tabStops[:0]
	p.keepWithNext = nil
	p.keepLines = nil
	p.pageBreakBefore = nil
	p.widowControl = nil
}

// ToXML converts the paragraph to WordprocessingML XML
func (p *Paragraph) ToXML() string {
	var runsXML strings.Builder
	for _, run := range p.runs {
		runsXML.WriteString(run.ToXML())
	}

	var pPr string
	if p.style != "" || p.alignment != WDAlignParagraphLeft || p.numberingApplied || p.hasSpacing() || p.hasIndentation() || p.hasTabStops() || p.hasKeepSettings() {
		var pPrContent strings.Builder

		if p.style != "" {
			pPrContent.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, p.style))
		}

		if p.alignment != WDAlignParagraphLeft {
			pPrContent.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, p.alignment))
		}

		if p.numberingApplied {
			pPrContent.WriteString(fmt.Sprintf(`<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>`, p.numberingLevel, p.numberingID))
		}

		if p.hasSpacing() {
			pPrContent.WriteString(p.spacingXML())
		}

		if p.hasIndentation() {
			pPrContent.WriteString(p.indentationXML())
		}

		if p.hasTabStops() {
			pPrContent.WriteString(p.tabsXML())
		}

		if p.hasKeepSettings() {
			pPrContent.WriteString(p.keepSettingsXML())
		}

		pPr = fmt.Sprintf(`<w:pPr>%s</w:pPr>`, pPrContent.String())
	}

	return fmt.Sprintf(`<w:p>%s%s</w:p>`, pPr, runsXML.String())
}

func (p *Paragraph) hasSpacing() bool {
	return p.spacingBefore != 0 || p.spacingAfter != 0 || p.spacingLine != 0 || p.spacingLineRule != ""
}

func (p *Paragraph) spacingXML() string {
	attrs := make([]string, 0, 4)
	if p.spacingBefore != 0 {
		attrs = append(attrs, fmt.Sprintf(`w:before="%d"`, p.spacingBefore))
	}
	if p.spacingAfter != 0 {
		attrs = append(attrs, fmt.Sprintf(`w:after="%d"`, p.spacingAfter))
	}
	if p.spacingLine != 0 {
		attrs = append(attrs, fmt.Sprintf(`w:line="%d"`, p.spacingLine))
	}
	if p.spacingLineRule != "" {
		attrs = append(attrs, fmt.Sprintf(`w:lineRule="%s"`, p.spacingLineRule))
	}

	return fmt.Sprintf(`<w:spacing %s/>`, strings.Join(attrs, " "))
}

func (p *Paragraph) hasIndentation() bool {
	return p.indentLeft != 0 || p.indentRight != 0 || p.indentFirstLine != 0 || p.indentHanging != 0
}

func (p *Paragraph) indentationXML() string {
	attrs := make([]string, 0, 4)
	if p.indentLeft != 0 {
		attrs = append(attrs, fmt.Sprintf(`w:left="%d"`, p.indentLeft))
	}
	if p.indentRight != 0 {
		attrs = append(attrs, fmt.Sprintf(`w:right="%d"`, p.indentRight))
	}
	if p.indentFirstLine != 0 {
		attrs = append(attrs, fmt.Sprintf(`w:firstLine="%d"`, p.indentFirstLine))
	}
	if p.indentHanging != 0 {
		attrs = append(attrs, fmt.Sprintf(`w:hanging="%d"`, p.indentHanging))
	}

	return fmt.Sprintf(`<w:ind %s/>`, strings.Join(attrs, " "))
}

// SetKeepWithNext sets the keep-with-next property (prevents a page break between this and the following paragraph)
func (p *Paragraph) SetKeepWithNext(enabled bool) {
	p.keepWithNext = boolPtr(enabled)
}

// KeepWithNext returns whether the keep-with-next property is enabled
func (p *Paragraph) KeepWithNext() bool {
	if p.keepWithNext == nil {
		return false
	}
	return *p.keepWithNext
}

// ClearKeepWithNext clears the keep-with-next override restoring the default behavior
func (p *Paragraph) ClearKeepWithNext() {
	p.keepWithNext = nil
}

// SetKeepLines sets whether all lines in the paragraph must stay on the same page
func (p *Paragraph) SetKeepLines(enabled bool) {
	p.keepLines = boolPtr(enabled)
}

// KeepLines returns whether the keep lines property is enabled
func (p *Paragraph) KeepLines() bool {
	if p.keepLines == nil {
		return false
	}
	return *p.keepLines
}

// ClearKeepLines clears the keep lines override restoring the default behavior
func (p *Paragraph) ClearKeepLines() {
	p.keepLines = nil
}

// SetPageBreakBefore forces a page break before this paragraph when enabled
func (p *Paragraph) SetPageBreakBefore(enabled bool) {
	p.pageBreakBefore = boolPtr(enabled)
}

// PageBreakBefore reports whether a page break is forced before the paragraph
func (p *Paragraph) PageBreakBefore() bool {
	if p.pageBreakBefore == nil {
		return false
	}
	return *p.pageBreakBefore
}

// ClearPageBreakBefore clears the page-break-before override
func (p *Paragraph) ClearPageBreakBefore() {
	p.pageBreakBefore = nil
}

// SetWidowControl sets widow control (keep minimum lines on a page). Passing false disables the control.
func (p *Paragraph) SetWidowControl(enabled bool) {
	p.widowControl = boolPtr(enabled)
}

// WidowControl returns whether widow control is enabled. If not explicitly set, it defaults to true per Wordprocessing defaults.
func (p *Paragraph) WidowControl() bool {
	if p.widowControl == nil {
		return true
	}
	return *p.widowControl
}

// ClearWidowControl clears the widow control override, reverting to the default
func (p *Paragraph) ClearWidowControl() {
	p.widowControl = nil
}

// AddTabStop adds a tab stop to the paragraph
func (p *Paragraph) AddTabStop(position int, alignment WDTabAlignment, leader WDTabLeader) {
	align := alignment
	if align == "" {
		align = WDTabAlignmentLeft
	}
	lead := leader
	if lead == "" {
		lead = WDTabLeaderNone
	}
	p.tabStops = append(p.tabStops, TabStop{Position: position, Alignment: align, Leader: lead})
}

// SetTabStops replaces the current tab stops with the provided collection
func (p *Paragraph) SetTabStops(stops []TabStop) {
	if len(stops) == 0 {
		p.tabStops = p.tabStops[:0]
		return
	}
	p.tabStops = p.tabStops[:0]
	p.tabStops = append(p.tabStops, stops...)
}

// ClearTabStops removes all tab stops from the paragraph
func (p *Paragraph) ClearTabStops() {
	p.tabStops = p.tabStops[:0]
}

// TabStops returns a copy of the paragraph tab stops
func (p *Paragraph) TabStops() []TabStop {
	stops := make([]TabStop, len(p.tabStops))
	copy(stops, p.tabStops)
	return stops
}

func (p *Paragraph) hasTabStops() bool {
	return len(p.tabStops) > 0
}

func (p *Paragraph) tabsXML() string {
	if len(p.tabStops) == 0 {
		return ""
	}

	var builder strings.Builder
	builder.WriteString("<w:tabs>")
	for _, tab := range p.tabStops {
		alignment := tab.Alignment
		if alignment == "" {
			alignment = WDTabAlignmentLeft
		}
		attrs := []string{
			fmt.Sprintf(`w:val="%s"`, alignment),
			fmt.Sprintf(`w:pos="%d"`, tab.Position),
		}
		if tab.Leader != "" && tab.Leader != WDTabLeaderNone {
			attrs = append(attrs, fmt.Sprintf(`w:leader="%s"`, tab.Leader))
		}
		builder.WriteString(fmt.Sprintf(`<w:tab %s/>`, strings.Join(attrs, " ")))
	}
	builder.WriteString("</w:tabs>")
	return builder.String()
}

func (p *Paragraph) hasKeepSettings() bool {
	return p.keepWithNext != nil || p.keepLines != nil || p.pageBreakBefore != nil || p.widowControl != nil
}

func (p *Paragraph) keepSettingsXML() string {
	var builder strings.Builder
	if p.keepWithNext != nil {
		builder.WriteString(onOffXML("w:keepNext", *p.keepWithNext))
	}
	if p.keepLines != nil {
		builder.WriteString(onOffXML("w:keepLines", *p.keepLines))
	}
	if p.pageBreakBefore != nil {
		builder.WriteString(onOffXML("w:pageBreakBefore", *p.pageBreakBefore))
	}
	if p.widowControl != nil {
		builder.WriteString(onOffXML("w:widowControl", *p.widowControl))
	}
	return builder.String()
}

func onOffXML(tag string, value bool) string {
	if value {
		return fmt.Sprintf(`<%s/>`, tag)
	}
	return fmt.Sprintf(`<%s w:val="0"/>`, tag)
}

func boolPtr(v bool) *bool {
	b := v
	return &b
}

// Run represents a run of text with consistent formatting
type Run struct {
	owner           *DocumentPart
	text            string
	bold            bool
	italic          bool
	underline       WDUnderline
	size            int // font size in half-points
	color           string
	font            string
	highlight       WDColorIndex
	breakType       BreakType // Type of break to add after this run
	hasBreak        bool      // Whether this run has a break
	hyperlinkURL    string
	hyperlinkAnchor string
	strike          bool
	doubleStrike    bool
	smallCaps       bool
	allCaps         bool
	shadow          bool
	outline         bool
	emboss          bool
	imprint         bool
	picture         *Picture
}

// NewRun creates a new run with the specified text
func NewRun(text string) *Run {
	return &Run{
		text:      text,
		underline: WDUnderlineNone,
		size:      22, // 11pt default
		color:     "auto",
		font:      "Calibri",
		highlight: WDColorIndexAuto,
	}
}

// Text returns the text content of the run
func (r *Run) Text() string {
	return r.text
}

// SetText sets the text content of the run
func (r *Run) SetText(text string) {
	r.text = text
}

// SetBold sets the bold formatting
func (r *Run) SetBold(bold bool) {
	r.bold = bold
}

// SetItalic sets the italic formatting
func (r *Run) SetItalic(italic bool) {
	r.italic = italic
}

// SetStrikethrough toggles single strikethrough formatting
func (r *Run) SetStrikethrough(strike bool) {
	r.strike = strike
}

// SetDoubleStrikethrough toggles double strikethrough formatting
func (r *Run) SetDoubleStrikethrough(doubleStrike bool) {
	r.doubleStrike = doubleStrike
}

// SetSmallCaps toggles small caps formatting
func (r *Run) SetSmallCaps(smallCaps bool) {
	r.smallCaps = smallCaps
}

// SetAllCaps toggles all caps formatting
func (r *Run) SetAllCaps(allCaps bool) {
	r.allCaps = allCaps
}

// SetShadow toggles text shadow effect
func (r *Run) SetShadow(shadow bool) {
	r.shadow = shadow
}

// SetOutline toggles outline effect
func (r *Run) SetOutline(outline bool) {
	r.outline = outline
}

// SetEmboss toggles emboss effect
func (r *Run) SetEmboss(emboss bool) {
	r.emboss = emboss
}

// SetImprint toggles imprint (engrave) effect
func (r *Run) SetImprint(imprint bool) {
	r.imprint = imprint
}

// SetUnderline sets the underline formatting
func (r *Run) SetUnderline(underline WDUnderline) {
	r.underline = underline
}

// SetSize sets the font size in points
func (r *Run) SetSize(size int) {
	r.size = size * 2 // Convert to half-points
}

func (r *Run) setSizeRaw(halfPoints int) {
	r.size = halfPoints
}

// SetColor sets the text color
func (r *Run) SetColor(color string) {
	r.color = color
}

// SetFont sets the font family
func (r *Run) SetFont(font string) {
	r.font = font
}

// SetHighlight sets the highlight color
func (r *Run) SetHighlight(highlight WDColorIndex) {
	r.highlight = highlight
}

// SetHyperlink sets an external hyperlink for the run
func (r *Run) SetHyperlink(url string) {
	r.hyperlinkURL = url
	r.hyperlinkAnchor = ""
}

// SetHyperlinkAnchor sets an internal hyperlink anchor for the run
func (r *Run) SetHyperlinkAnchor(anchor string) {
	r.hyperlinkAnchor = anchor
	r.hyperlinkURL = ""
}

// HasHyperlink reports whether the run is a hyperlink
func (r *Run) HasHyperlink() bool {
	return r.hyperlinkURL != "" || r.hyperlinkAnchor != ""
}

// HyperlinkURL returns the hyperlink URL if the run links externally
func (r *Run) HyperlinkURL() string {
	return r.hyperlinkURL
}

// HyperlinkAnchor returns the internal hyperlink anchor if present
func (r *Run) HyperlinkAnchor() string {
	return r.hyperlinkAnchor
}

// HasPicture reports whether the run contains an inline picture
func (r *Run) HasPicture() bool {
	return r.picture != nil
}

// Picture returns the picture embedded in the run, if any
func (r *Run) Picture() *Picture {
	return r.picture
}

// AddPicture embeds an image into the run. Width and height are specified in EMUs.
// Pass zero for either dimension to preserve the image's aspect ratio using the source size.
func (r *Run) AddPicture(path string, widthEMU, heightEMU int64) (*Picture, error) {
	if r.owner == nil {
		return nil, fmt.Errorf("run is not attached to a document")
	}
	picture, err := r.owner.addPictureFromFile(path, widthEMU, heightEMU)
	if err != nil {
		return nil, err
	}
	r.picture = picture
	return picture, nil
}

// AddBreak adds a break to the run
func (r *Run) AddBreak(breakType BreakType) {
	r.breakType = breakType
	r.hasBreak = true
}

// IsBold reports whether the run is bold
func (r *Run) IsBold() bool {
	return r.bold
}

// IsItalic reports whether the run is italic
func (r *Run) IsItalic() bool {
	return r.italic
}

// IsStrikethrough reports whether the run is strikethrough
func (r *Run) IsStrikethrough() bool {
	return r.strike
}

// IsDoubleStrikethrough reports whether the run is double strikethrough
func (r *Run) IsDoubleStrikethrough() bool {
	return r.doubleStrike
}

// IsSmallCaps reports whether the run uses small caps
func (r *Run) IsSmallCaps() bool {
	return r.smallCaps
}

// IsAllCaps reports whether the run uses all caps
func (r *Run) IsAllCaps() bool {
	return r.allCaps
}

// HasShadow reports whether the run has a shadow effect
func (r *Run) HasShadow() bool {
	return r.shadow
}

// HasOutline reports whether the run has an outline effect
func (r *Run) HasOutline() bool {
	return r.outline
}

// IsEmbossed reports whether the run is embossed
func (r *Run) IsEmbossed() bool {
	return r.emboss
}

// IsImprinted reports whether the run is imprinted (engraved)
func (r *Run) IsImprinted() bool {
	return r.imprint
}

// Underline returns the underline style of the run
func (r *Run) Underline() WDUnderline {
	return r.underline
}

// Size returns the font size in points
func (r *Run) Size() int {
	return r.size / 2
}

// Color returns the text color of the run
func (r *Run) Color() string {
	return r.color
}

// Font returns the font family of the run
func (r *Run) Font() string {
	return r.font
}

// Highlight returns the highlight color of the run
func (r *Run) Highlight() WDColorIndex {
	return r.highlight
}

// HasBreak reports whether the run has a break
func (r *Run) HasBreak() bool {
	return r.hasBreak
}

// BreakType returns the break type of the run
func (r *Run) BreakType() BreakType {
	return r.breakType
}

// ToXML converts the run to WordprocessingML XML
func (r *Run) ToXML() string {
	var rPr strings.Builder

	if r.bold {
		rPr.WriteString("<w:b/>")
	}

	if r.italic {
		rPr.WriteString("<w:i/>")
	}

	if r.strike {
		rPr.WriteString("<w:strike/>")
	}

	if r.doubleStrike {
		rPr.WriteString("<w:dstrike/>")
	}

	if r.smallCaps {
		rPr.WriteString("<w:smallCaps/>")
	}

	if r.allCaps {
		rPr.WriteString("<w:caps/>")
	}

	if r.shadow {
		rPr.WriteString("<w:shadow/>")
	}

	if r.outline {
		rPr.WriteString("<w:outline/>")
	}

	if r.emboss {
		rPr.WriteString("<w:emboss/>")
	}

	if r.imprint {
		rPr.WriteString("<w:imprint/>")
	}

	if r.underline != WDUnderlineNone {
		rPr.WriteString(fmt.Sprintf(`<w:u w:val="%s"/>`, r.underline))
	}

	if r.size != 22 {
		rPr.WriteString(fmt.Sprintf(`<w:sz w:val="%d"/>`, r.size))
		rPr.WriteString(fmt.Sprintf(`<w:szCs w:val="%d"/>`, r.size))
	}

	if r.color != "auto" {
		rPr.WriteString(fmt.Sprintf(`<w:color w:val="%s"/>`, r.color))
	}

	if r.font != "Calibri" {
		rPr.WriteString(fmt.Sprintf(`<w:rFonts w:ascii="%s" w:hAnsi="%s"/>`, r.font, r.font))
	}

	if r.highlight != WDColorIndexAuto {
		rPr.WriteString(fmt.Sprintf(`<w:highlight w:val="%s"/>`, r.highlight))
	}

	var rPrXML string
	if rPr.Len() > 0 {
		rPrXML = fmt.Sprintf("<w:rPr>%s</w:rPr>", rPr.String())
	}

	var content strings.Builder

	if r.text != "" {
		escaped := strings.ReplaceAll(r.text, "&", "&amp;")
		escaped = strings.ReplaceAll(escaped, "<", "&lt;")
		escaped = strings.ReplaceAll(escaped, ">", "&gt;")
		content.WriteString(fmt.Sprintf(`<w:t>%s</w:t>`, escaped))
	}

	if r.picture != nil {
		content.WriteString(r.picture.toXML())
	}

	if r.hasBreak {
		switch r.breakType {
		case BreakTypePage:
			content.WriteString(`<w:br w:type="page"/>`)
		case BreakTypeColumn:
			content.WriteString(`<w:br w:type="column"/>`)
		default:
			content.WriteString(`<w:br/>`)
		}
	}

	if content.Len() == 0 {
		content.WriteString("<w:t/>")
	}

	runXML := fmt.Sprintf(`<w:r>%s%s</w:r>`, rPrXML, content.String())

	if r.HasHyperlink() {
		return r.wrapWithHyperlink(runXML)
	}

	return runXML
}

func (r *Run) wrapWithHyperlink(runXML string) string {
	attrs := make([]string, 0, 2)
	if r.hyperlinkURL != "" && r.owner != nil {
		if relID := r.owner.ensureHyperlinkRelationship(r.hyperlinkURL); relID != "" {
			attrs = append(attrs, fmt.Sprintf(`r:id="%s"`, relID))
		}
	}
	if r.hyperlinkAnchor != "" {
		attrs = append(attrs, fmt.Sprintf(`w:anchor="%s"`, r.hyperlinkAnchor))
	}

	attrStr := ""
	if len(attrs) > 0 {
		attrStr = " " + strings.Join(attrs, " ")
	}

	return fmt.Sprintf(`<w:hyperlink%s>%s</w:hyperlink>`, attrStr, runXML)
}
