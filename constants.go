package docx

// ContentType constants for different parts of a Word document
const (
	ContentTypeWMLDocumentMain = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
	ContentTypeWMLStyles       = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
	ContentTypeWMLSettings     = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
	ContentTypeWMLComments     = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
	ContentTypeWMLNumbering    = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
	ContentTypeWMLHeader       = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	ContentTypeWMLFooter       = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	ContentTypeOPCCoreProps    = "application/vnd.openxmlformats-package.core-properties+xml"
	ContentTypeRels            = "application/vnd.openxmlformats-package.relationships+xml"
)

// Relationship Type constants
const (
	RelTypeOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	RelTypeImage          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	RelTypeHyperlink      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	RelTypeStyles         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	RelTypeSettings       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	RelTypeComments       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
	RelTypeNumbering      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
	RelTypeHeader         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	RelTypeFooter         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	RelTypeCoreProps      = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
)

// BreakType represents different types of breaks that can be inserted in a Word document.
type BreakType string

const (
	// BreakTypePage inserts a page break, starting a new page.
	BreakTypePage BreakType = "page"
	// BreakTypeColumn inserts a column break, starting a new column.
	BreakTypeColumn BreakType = "column"
	// BreakTypeText inserts a text wrapping break, continuing on a new line.
	BreakTypeText BreakType = "textWrapping"
)

// SectionStartType represents how a section starts in a Word document.
type SectionStartType string

const (
	// SectionStartContinuous starts the section on the same page.
	SectionStartContinuous SectionStartType = "continuous"
	// SectionStartNewColumn starts the section in a new column.
	SectionStartNewColumn SectionStartType = "nextColumn"
	// SectionStartNewPage starts the section on a new page.
	SectionStartNewPage SectionStartType = "nextPage"
	// SectionStartEvenPage starts the section on the next even page.
	SectionStartEvenPage SectionStartType = "evenPage"
	// SectionStartOddPage starts the section on the next odd page.
	SectionStartOddPage SectionStartType = "oddPage"
)

// WDUnderline represents different underline styles for text formatting.
type WDUnderline string

const (
	// WDUnderlineNone removes any underline formatting.
	WDUnderlineNone WDUnderline = "none"
	// WDUnderlineSingle applies a single underline.
	WDUnderlineSingle WDUnderline = "single"
	// WDUnderlineDouble applies a double underline.
	WDUnderlineDouble WDUnderline = "double"
	// WDUnderlineThick applies a thick underline.
	WDUnderlineThick WDUnderline = "thick"
	// WDUnderlineDotted applies a dotted underline.
	WDUnderlineDotted WDUnderline = "dotted"
	// WDUnderlineDashed applies a dashed underline.
	WDUnderlineDashed WDUnderline = "dash"
)

// WDColorIndex represents predefined color indices for text highlighting in Word documents.
type WDColorIndex string

const (
	WDColorIndexAuto        WDColorIndex = "auto"
	WDColorIndexBlack       WDColorIndex = "black"
	WDColorIndexBlue        WDColorIndex = "blue"
	WDColorIndexBrightGreen WDColorIndex = "brightGreen"
	WDColorIndexDarkBlue    WDColorIndex = "darkBlue"
	WDColorIndexDarkGreen   WDColorIndex = "darkGreen"
	WDColorIndexDarkRed     WDColorIndex = "darkRed"
	WDColorIndexDarkYellow  WDColorIndex = "darkYellow"
	WDColorIndexGray25      WDColorIndex = "gray25"
	WDColorIndexGray50      WDColorIndex = "gray50"
	WDColorIndexGreen       WDColorIndex = "green"
	WDColorIndexPink        WDColorIndex = "pink"
	WDColorIndexRed         WDColorIndex = "red"
	WDColorIndexTurquoise   WDColorIndex = "turquoise"
	WDColorIndexViolet      WDColorIndex = "violet"
	WDColorIndexWhite       WDColorIndex = "white"
	WDColorIndexYellow      WDColorIndex = "yellow"
)

// WDAlignParagraph represents horizontal alignment options for paragraphs.
type WDAlignParagraph string

const (
	WDAlignParagraphLeft       WDAlignParagraph = "left"
	WDAlignParagraphCenter     WDAlignParagraph = "center"
	WDAlignParagraphRight      WDAlignParagraph = "right"
	WDAlignParagraphJustify    WDAlignParagraph = "both"
	WDAlignParagraphDistribute WDAlignParagraph = "distribute"
)

// WDTabAlignment represents alignment options for tab stops in paragraphs.
type WDTabAlignment string

const (
	WDTabAlignmentLeft    WDTabAlignment = "left"
	WDTabAlignmentCenter  WDTabAlignment = "center"
	WDTabAlignmentRight   WDTabAlignment = "right"
	WDTabAlignmentDecimal WDTabAlignment = "decimal"
	WDTabAlignmentBar     WDTabAlignment = "bar"
)

// WDTabLeader represents leader character styles that fill the space before a tab stop.
type WDTabLeader string

const (
	WDTabLeaderNone       WDTabLeader = "none"
	WDTabLeaderDot        WDTabLeader = "dot"
	WDTabLeaderHyphen     WDTabLeader = "hyphen"
	WDTabLeaderUnderscore WDTabLeader = "underscore"
	WDTabLeaderHeavy      WDTabLeader = "heavy"
	WDTabLeaderMiddleDot  WDTabLeader = "middleDot"
)
