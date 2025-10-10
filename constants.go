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

// BreakType represents different types of breaks
type BreakType string

const (
	BreakTypePage   BreakType = "page"
	BreakTypeColumn BreakType = "column"
	BreakTypeText   BreakType = "textWrapping"
)

// SectionStartType represents how a section starts
type SectionStartType string

const (
	SectionStartContinuous SectionStartType = "continuous"
	SectionStartNewColumn  SectionStartType = "nextColumn"
	SectionStartNewPage    SectionStartType = "nextPage"
	SectionStartEvenPage   SectionStartType = "evenPage"
	SectionStartOddPage    SectionStartType = "oddPage"
)

// WDUnderline represents underline types
type WDUnderline string

const (
	WDUnderlineNone   WDUnderline = "none"
	WDUnderlineSingle WDUnderline = "single"
	WDUnderlineDouble WDUnderline = "double"
	WDUnderlineThick  WDUnderline = "thick"
	WDUnderlineDotted WDUnderline = "dotted"
	WDUnderlineDashed WDUnderline = "dash"
)

// WDColorIndex represents color indices
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

// WDAlignParagraph represents paragraph alignment
type WDAlignParagraph string

const (
	WDAlignParagraphLeft       WDAlignParagraph = "left"
	WDAlignParagraphCenter     WDAlignParagraph = "center"
	WDAlignParagraphRight      WDAlignParagraph = "right"
	WDAlignParagraphJustify    WDAlignParagraph = "both"
	WDAlignParagraphDistribute WDAlignParagraph = "distribute"
)

// WDTabAlignment represents tab stop alignment options
type WDTabAlignment string

const (
	WDTabAlignmentLeft    WDTabAlignment = "left"
	WDTabAlignmentCenter  WDTabAlignment = "center"
	WDTabAlignmentRight   WDTabAlignment = "right"
	WDTabAlignmentDecimal WDTabAlignment = "decimal"
	WDTabAlignmentBar     WDTabAlignment = "bar"
)

// WDTabLeader represents the leader characters used for tab stops
type WDTabLeader string

const (
	WDTabLeaderNone       WDTabLeader = "none"
	WDTabLeaderDot        WDTabLeader = "dot"
	WDTabLeaderHyphen     WDTabLeader = "hyphen"
	WDTabLeaderUnderscore WDTabLeader = "underscore"
	WDTabLeaderHeavy      WDTabLeader = "heavy"
	WDTabLeaderMiddleDot  WDTabLeader = "middleDot"
)
