# Go-DOCX - Microsoft Word Document Library for Go

Go-DOCX is a comprehensive Go library for creating and manipulating Microsoft Word (.docx) documents. It is inspired by and provides feature parity (~75%) with the popular python-docx library, offering a type-safe, high-performance alternative for Go developers.

## ✨ Features

### Core Document Features
- ✅ Create new Word documents from scratch
- ✅ Open and modify existing documents
- ✅ Save documents to file or stream
- ✅ Document properties (title, author, subject, keywords, etc.)
- ✅ Multiple sections with different page layouts

### Text and Paragraphs
- ✅ Add paragraphs with rich text formatting
- ✅ Text runs with individual formatting
- ✅ Bold, italic, underline, strikethrough
- ✅ Font family, size, and color
- ✅ Text highlighting with predefined colors
- ✅ Paragraph alignment (left, center, right, justify)
- ✅ Paragraph spacing and indentation
- ✅ Paragraph borders and shading
- ✅ Tab stops and custom tab positions
- ✅ Keep with next, keep together, widow/orphan control

### Headings and Styles
- ✅ Add headings (levels 0-9, including Title)
- ✅ Built-in paragraph styles (Heading 1-9, Normal, etc.)
- ✅ Apply styles to paragraphs
- ✅ Styles part with default definitions

### Tables
- ✅ Create tables with specified rows and columns
- ✅ Add/remove rows dynamically
- ✅ Cell text and content
- ✅ **Advanced table formatting:**
  - Table borders (all sides, customizable style and color)
  - Cell shading/background colors
  - Cell margins (top, bottom, left, right)
  - Horizontal cell merging (merge across columns)
  - Vertical cell merging (merge across rows)
- ✅ Access individual cells, rows, and columns
- ✅ Nested paragraphs in cells

### Images 📷
- ✅ **Insert inline images** (fully functional!)
- ✅ **Supported formats:** PNG, JPEG, GIF, BMP, TIFF
- ✅ **Document-level API:** `doc.AddPicture()`
- ✅ **Run-level API:** `run.AddPicture()`
- ✅ **Automatic aspect ratio** (specify width or height, auto-calculate other)
- ✅ **Custom dimensions** in EMUs (English Metric Units)
- ✅ **Image data access:** retrieve embedded image bytes
- ✅ Proper relationship and part management
- ✅ Round-trip support (read images from existing documents)

### Hyperlinks 🔗
- ✅ **URL hyperlinks** (external links)
- ✅ **Anchor hyperlinks** (internal document bookmarks)
- ✅ **Paragraph-level API:** `paragraph.AddHyperlink()`
- ✅ **Run-level API:** `run.SetHyperlink()`, `run.SetHyperlinkAnchor()`
- ✅ **Hyperlink detection:** `run.HasHyperlink()`
- ✅ Proper relationship management
- ✅ Round-trip support (preserve hyperlinks when editing)

### Lists 📝
- ✅ **Numbered lists** (decimal numbering)
- ✅ **Bulleted lists** (bullet symbols)
- ✅ **Multi-level lists** (up to 9 levels)
- ✅ **Simple API:** `doc.AddNumberedParagraph()`, `doc.AddBulletedParagraph()`
- ✅ **Custom numbering:** `paragraph.SetNumbering(numID, level)`
- ✅ Default numbering definitions included
- ✅ Numbering part with abstract numbering definitions
- ✅ Round-trip support

### Headers and Footers
- ✅ **Section-level headers and footers**
- ✅ **Three types:** default, first page, even page
- ✅ **Add content:** paragraphs, tables, formatted text
- ✅ **Document-level convenience methods**
- ✅ **Section-specific methods**
- ✅ Proper part and relationship management
- ✅ Round-trip support

### Page Layout
- ✅ Page breaks (explicit page breaks)
- ✅ Column breaks
- ✅ Text wrapping breaks
- ✅ Sections with different layouts
- ✅ Page size (width and height)
- ✅ Page orientation (portrait, landscape)
- ✅ Page margins (top, bottom, left, right)

### Advanced Features
- ✅ Comments structure (partial - needs API completion)
- ✅ Settings management
- ✅ Core properties (Dublin Core metadata)
- ✅ Relationship management
- ✅ Content types handling
- ✅ XML part parsing and generation

## 📦 Installation

```bash
go get github.com/SanjarbekSaminjonov/go-docx
```

## 🚀 Quick Start

```go
package main

import (
    "log"
    "github.com/SanjarbekSaminjonov/go-docx"
)

func main() {
    // Create a new document
    doc := docx.NewDocument()
    defer doc.Close()
    
    // Set document properties
    props := doc.CoreProperties()
    props.SetTitle("My Document")
    props.SetCreator("John Doe")
    
    // Add a title
    title, _ := doc.AddHeading("Welcome to Go-DOCX", 0)
    title.SetAlignment(docx.WDAlignParagraphCenter)
    
    // Add a paragraph with formatted text
    p := doc.AddParagraph()
    p.AddRun("This is ").SetBold(false)
    p.AddRun("bold text").SetBold(true)
    p.AddRun(" and this is ").SetBold(false)
    p.AddRun("italic text").SetItalic(true)
    
    // Add a hyperlink
    p2 := doc.AddParagraph("Visit our website: ")
    p2.AddHyperlink("Go-DOCX on GitHub", "https://github.com/SanjarbekSaminjonov/go-docx")
    
    // Add an image
    doc.AddPicture("logo.png", 0, 0) // Auto aspect ratio
    
    // Add a numbered list
    doc.AddNumberedParagraph("First item", 0)
    doc.AddNumberedParagraph("Second item", 0)
    doc.AddNumberedParagraph("Sub-item", 1)
    
    // Add a table
    table := doc.AddTable(3, 3)
    table.Row(0).Cell(0).SetText("Header 1")
    table.Row(0).Cell(1).SetText("Header 2")
    table.Row(0).Cell(2).SetText("Header 3")
    
    // Save the document
    if err := doc.SaveAs("my_document.docx"); err != nil {
        log.Fatal(err)
    }
}
```

## API Reference

### Document

#### Creating Documents

```go
// Create a new document
doc := docx.NewDocument()

// Open an existing document
doc, err := docx.OpenDocument("existing.docx")
```

#### Adding Content

```go
// Add a paragraph
paragraph := doc.AddParagraph("Text content")

// Add a heading (levels 0-9)
heading, err := doc.AddHeading("Chapter Title", 1)

// Add a table
table := doc.AddTable(3, 4) // 3 rows, 4 columns

// Add a page break
doc.AddPageBreak()
```

#### Document Properties

```go
props := doc.CoreProperties()
props.SetTitle("Document Title")
props.SetCreator("Author Name")
props.SetSubject("Document Subject")
```

#### Saving Documents

```go
// Save to a new file
err := doc.SaveAs("output.docx")

// Save to original location (if opened from file)
err := doc.Save()

// Close the document
err := doc.Close()
```

### Paragraphs

```go
paragraph := doc.AddParagraph("Initial text")

// Set alignment
paragraph.SetAlignment(docx.WDAlignParagraphCenter)
paragraph.SetAlignment(docx.WDAlignParagraphRight)
paragraph.SetAlignment(docx.WDAlignParagraphJustify)

// Set style
paragraph.SetStyle("Heading 1")

// Paragraph spacing (in twips)
paragraph.SetSpacingBefore(240) // 240 twips = 1/6 inch
paragraph.SetSpacingAfter(240)
paragraph.SetLineSpacing(360, docx.LineSpacingAuto) // 1.5 line spacing

// Paragraph indentation (in twips)
paragraph.SetIndentation(720, 0, 0) // left, right, firstLine

// Paragraph borders
paragraph.SetBorder(docx.ParagraphBorderTop, docx.ParagraphBorder{
    Style: "single",
    Color: "000000",
    Size:  4,
})

// Paragraph shading
paragraph.SetShading("clear", "D9D9D9", "auto")

// Keep with next paragraph
paragraph.SetKeepWithNext(true)

// Keep lines together
paragraph.SetKeepTogether(true)

// Add runs with different formatting
run1 := paragraph.AddRun("Normal text ")
run2 := paragraph.AddRun("Bold text")
run2.SetBold(true)
```

### Text Formatting (Runs)

```go
run := paragraph.AddRun("Formatted text")

// Basic formatting
run.SetBold(true)
run.SetItalic(true)
run.SetUnderline(docx.WDUnderlineSingle)

// Font and size
run.SetFont("Arial")
run.SetSize(14) // Font size in points

// Color and highlighting
run.SetColor("FF0000") // Red color
run.SetHighlight(docx.WDColorIndexYellow)

// Add breaks
run.AddBreak(docx.BreakTypePage)
run.AddBreak(docx.BreakTypeColumn)
```

### Tables

```go
// Create a table
table := doc.AddTable(3, 4) // 3 rows, 4 columns

// Access cells
cell := table.Row(0).Cell(0)
cell.SetText("Header 1")

// Add content to cells
cell.AddParagraph("Additional content")

// Add new rows
newRow := table.AddRow()

// Table borders
table.SetBorder(docx.TableBorderTop, docx.TableBorder{
    Style: "single",
    Color: "000000",
    Size:  4,
})

// Cell shading
cell.SetShading("clear", "4472C4", "auto") // pattern, fill, color

// Cell margins (in twips - 1/1440 inch)
table.SetCellMargins(100, 100, 100, 100) // top, left, bottom, right

// Merge cells horizontally
table.MergeCellsHorizontally(0, 0, 2) // row, startCol, endCol

// Merge cells vertically
table.MergeCellsVertically(0, 0, 2) // col, startRow, endRow
```

### Images 📷

```go
// Add image to document (creates new paragraph)
paragraph, picture, err := doc.AddPicture("photo.png", 0, 0)
if err != nil {
    log.Fatal(err)
}
// 0, 0 means auto aspect ratio based on image dimensions

// Add image with specific size (in EMUs)
// 914400 EMUs = 1 inch
doc.AddPicture("photo.jpg", 914400, 914400) // 1" x 1"

// Add image to a run
run := paragraph.AddRun()
pic, err := run.AddPicture("logo.png", 0, 0)

// Get image data
imageBytes, err := picture.ImageData()

// Supported formats: PNG, JPEG, GIF, BMP, TIFF
```

### Hyperlinks 🔗

```go
// Simple hyperlink using paragraph method
p := doc.AddParagraph("Visit ")
p.AddHyperlink("Google", "https://www.google.com")

// URL hyperlink using run
run := p.AddRun("GitHub")
run.SetHyperlink("https://github.com")
run.SetColor("0563C1") // Blue color
run.SetUnderline(docx.WDUnderlineSingle)

// Internal anchor/bookmark
run.SetHyperlinkAnchor("section1")

// Check if run has hyperlink
if run.HasHyperlink() {
    // Handle hyperlink
}
```

### Lists (Numbered and Bulleted) 📝

```go
// Numbered list
doc.AddNumberedParagraph("First item", 0)
doc.AddNumberedParagraph("Second item", 0)
doc.AddNumberedParagraph("Third item", 0)

// Multi-level numbered list
doc.AddNumberedParagraph("Level 0 - Item 1", 0)
doc.AddNumberedParagraph("Level 1 - Sub-item 1.1", 1)
doc.AddNumberedParagraph("Level 1 - Sub-item 1.2", 1)
doc.AddNumberedParagraph("Level 2 - Sub-sub-item", 2)
doc.AddNumberedParagraph("Level 0 - Item 2", 0)

// Bulleted list
doc.AddBulletedParagraph("First bullet", 0)
doc.AddBulletedParagraph("Second bullet", 0)
doc.AddBulletedParagraph("Sub-bullet", 1)

// Custom numbering
paragraph := doc.AddParagraph("Custom numbered item")
paragraph.SetNumbering(1, 0) // numID, level
```

### Headers and Footers

```go
// Get or create default section
section := doc.Sections()[0]

// Add default header
header, err := section.Header()
if err != nil {
    log.Fatal(err)
}
headerP := header.AddParagraph("Company Name")
headerP.SetAlignment(docx.WDAlignParagraphCenter)

// Add default footer
footer, err := section.Footer()
if err != nil {
    log.Fatal(err)
}
footerP := footer.AddParagraph("Page ")
footerP.SetAlignment(docx.WDAlignParagraphCenter)

// First page header (different from others)
firstHeader, err := section.HeaderOfType(docx.HeaderTypeFirst)
firstHeader.AddParagraph("First Page Header")

// Even page header
evenHeader, err := section.HeaderOfType(docx.HeaderTypeEven)
evenHeader.AddParagraph("Even Page Header")

// Document-level convenience methods
header, err := doc.Header() // Default header of first section
footer, err := doc.Footer() // Default footer of first section
```

### Sections and Page Layout

```go
// Add a new section
section := doc.AddSection(docx.WDSectionNewPage)

// Set page size (in twips - 1/1440 inch)
section.SetPageSize(12240, 15840) // Letter size: 8.5" x 11"

// Set orientation
section.SetOrientation(docx.WDOrientLandscape)

// Set margins (in twips)
section.SetMargins(1440, 1440, 1440, 1440) // 1 inch margins
```

### Constants

#### Alignment

- `WDAlignParagraphLeft` - Left alignment
- `WDAlignParagraphCenter` - Center alignment
- `WDAlignParagraphRight` - Right alignment
- `WDAlignParagraphJustify` - Justified alignment
- `WDAlignParagraphDistribute` - Distributed alignment

#### Underline Types

- `WDUnderlineNone` - No underline
- `WDUnderlineSingle` - Single underline
- `WDUnderlineDouble` - Double underline
- `WDUnderlineThick` - Thick underline
- `WDUnderlineDotted` - Dotted underline
- `WDUnderlineDashed` - Dashed underline
- `WDUnderlineWave` - Wave underline
- And more...

#### Break Types

- `BreakTypePage` - Page break
- `BreakTypeColumn` - Column break
- `BreakTypeText` - Text wrapping break

#### Color Indices (for highlighting)

- `WDColorIndexAuto` - Automatic color
- `WDColorIndexBlack` - Black
- `WDColorIndexBlue` - Blue
- `WDColorIndexBrightGreen` - Bright green
- `WDColorIndexDarkBlue` - Dark blue
- `WDColorIndexDarkRed` - Dark red
- `WDColorIndexDarkYellow` - Dark yellow
- `WDColorIndexGray25` - 25% gray
- `WDColorIndexGray50` - 50% gray
- `WDColorIndexGreen` - Green
- `WDColorIndexPink` - Pink
- `WDColorIndexRed` - Red
- `WDColorIndexTeal` - Teal
- `WDColorIndexTurquoise` - Turquoise
- `WDColorIndexViolet` - Violet
- `WDColorIndexWhite` - White
- `WDColorIndexYellow` - Yellow

#### Section Start Types

- `WDSectionContinuous` - Continuous section
- `WDSectionNewColumn` - New column
- `WDSectionNewPage` - New page (default)
- `WDSectionEvenPage` - Even page
- `WDSectionOddPage` - Odd page

#### Orientation Types

- `WDOrientPortrait` - Portrait orientation (default)
- `WDOrientLandscape` - Landscape orientation

#### Header/Footer Types

- `HeaderTypeDefault` - Default header (odd pages)
- `HeaderTypeFirst` - First page header
- `HeaderTypeEven` - Even page header
- `FooterTypeDefault` - Default footer (odd pages)
- `FooterTypeFirst` - First page footer
- `FooterTypeEven` - Even page footer

#### Table Border Sides

- `TableBorderTop` - Top border
- `TableBorderLeft` - Left border
- `TableBorderBottom` - Bottom border
- `TableBorderRight` - Right border
- `TableBorderInsideH` - Inside horizontal borders
- `TableBorderInsideV` - Inside vertical borders

#### Units

EMUs (English Metric Units):
- `EMUsPerInch = 914400` - EMUs in one inch
- `EMUsPerCm = 360000` - EMUs in one centimeter
- `EMUsPerPoint = 12700` - EMUs in one point

Twips (Twentieth of a point):
- 1440 twips = 1 inch
- 20 twips = 1 point

## 📚 Examples

### Complete Example

```go
package main

import (
    "log"
    "github.com/SanjarbekSaminjonov/go-docx"
)

func main() {
    // Create document
    doc := docx.NewDocument()
    defer doc.Close()
    
    // Document properties
    props := doc.CoreProperties()
    props.SetTitle("Comprehensive Example")
    props.SetCreator("Go-DOCX")
    props.SetSubject("Feature Demonstration")
    
    // Title
    title, _ := doc.AddHeading("Go-DOCX Feature Showcase", 0)
    title.SetAlignment(docx.WDAlignParagraphCenter)
    
    // Section 1: Text Formatting
    doc.AddHeading("1. Text Formatting", 1)
    p := doc.AddParagraph()
    p.AddRun("Normal text, ")
    p.AddRun("bold text, ").SetBold(true)
    p.AddRun("italic text, ").SetItalic(true)
    p.AddRun("colored text").SetColor("FF0000")
    
    // Section 2: Hyperlinks
    doc.AddHeading("2. Hyperlinks", 1)
    p2 := doc.AddParagraph("Visit ")
    p2.AddHyperlink("our website", "https://example.com")
    
    // Section 3: Lists
    doc.AddHeading("3. Lists", 1)
    doc.AddNumberedParagraph("First item", 0)
    doc.AddNumberedParagraph("Second item", 0)
    doc.AddBulletedParagraph("Bullet point", 0)
    
    // Section 4: Tables
    doc.AddHeading("4. Tables", 1)
    table := doc.AddTable(3, 3)
    table.Row(0).Cell(0).SetText("Name")
    table.Row(0).Cell(1).SetText("Age")
    table.Row(0).Cell(2).SetText("City")
    
    // Set all table borders
    border := docx.TableBorder{Style: "single", Color: "000000", Size: 4}
    table.SetBorder(docx.TableBorderTop, border)
    table.SetBorder(docx.TableBorderBottom, border)
    table.SetBorder(docx.TableBorderLeft, border)
    table.SetBorder(docx.TableBorderRight, border)
    table.SetBorder(docx.TableBorderInsideH, border)
    table.SetBorder(docx.TableBorderInsideV, border)
    
    // Section 5: Images
    doc.AddHeading("5. Images", 1)
    doc.AddPicture("logo.png", 0, 0) // Add your image
    
    // Headers and Footers
    section := doc.Sections()[0]
    header, _ := section.Header()
    header.AddParagraph("Document Header").SetAlignment(docx.WDAlignParagraphCenter)
    
    footer, _ := section.Footer()
    footer.AddParagraph("Page Footer").SetAlignment(docx.WDAlignParagraphCenter)
    
    // Save
    if err := doc.SaveAs("comprehensive_example.docx"); err != nil {
        log.Fatal(err)
    }
}
```


## 📊 Feature Comparison with python-docx

| Feature | python-docx | go-docx | Status |
|---------|-------------|---------|--------|
| **Core Document** |
| Create document | `Document()` | `NewDocument()` | ✅ Full |
| Open document | `Document('file.docx')` | `OpenDocument('file.docx')` | ✅ Full |
| Save document | `doc.save()` | `doc.SaveAs()` | ✅ Full |
| Document properties | `doc.core_properties` | `doc.CoreProperties()` | ✅ Full |
| **Paragraphs & Text** |
| Add paragraph | `doc.add_paragraph()` | `doc.AddParagraph()` | ✅ Full |
| Add heading | `doc.add_heading()` | `doc.AddHeading()` | ✅ Full |
| Text formatting | `run.bold = True` | `run.SetBold(true)` | ✅ Full |
| Paragraph alignment | `p.alignment = WD_ALIGN_*` | `p.SetAlignment()` | ✅ Full |
| Paragraph spacing | `p.paragraph_format` | `p.SetSpacing*()` | ✅ Full |
| Paragraph indentation | `p.paragraph_format` | `p.SetIndentation()` | ✅ Full |
| Paragraph borders | XML manipulation | `p.SetBorder()` | ✅ Full |
| **Tables** |
| Add table | `doc.add_table()` | `doc.AddTable()` | ✅ Full |
| Table borders | Limited | `table.SetBorder()` | ✅ Enhanced |
| Cell shading | XML manipulation | `cell.SetShading()` | ✅ Full |
| Cell margins | XML manipulation | `table.SetCellMargins()` | ✅ Full |
| Merge cells | XML manipulation | `table.MergeCells*()` | ✅ Full |
| **Images** |
| Add picture | `run.add_picture()` | `run.AddPicture()` | ✅ Full |
| Document picture | `doc.add_picture()` | `doc.AddPicture()` | ✅ Full |
| Image formats | PNG, JPEG, GIF, BMP | PNG, JPEG, GIF, BMP, TIFF | ✅ Full |
| Auto aspect ratio | `width=None` or `height=None` | `widthEMU=0` or `heightEMU=0` | ✅ Full |
| **Hyperlinks** |
| Add hyperlink | Paragraph method | `p.AddHyperlink()` | ✅ Full |
| Run hyperlink | XML manipulation | `run.SetHyperlink()` | ✅ Full |
| Internal anchor | Limited | `run.SetHyperlinkAnchor()` | ✅ Full |
| **Lists** |
| Numbered list | Style-based | `doc.AddNumberedParagraph()` | ✅ Full |
| Bulleted list | Style-based | `doc.AddBulletedParagraph()` | ✅ Full |
| Multi-level | `paragraph.style` | Level parameter | ✅ Full |
| Custom numbering | `numbering_part` | `paragraph.SetNumbering()` | ✅ Full |
| **Headers & Footers** |
| Add header | `section.header` | `section.Header()` | ✅ Full |
| Add footer | `section.footer` | `section.Footer()` | ✅ Full |
| First page header | `section.first_page_header` | `section.HeaderOfType()` | ✅ Full |
| Even page header | `section.even_page_header` | `section.HeaderOfType()` | ✅ Full |
| **Sections** |
| Add section | `doc.add_section()` | `doc.AddSection()` | ✅ Full |
| Page size | `section.page_width` | `section.SetPageSize()` | ✅ Full |
| Orientation | `section.orientation` | `section.SetOrientation()` | ✅ Full |
| Margins | `section.left_margin` | `section.SetMargins()` | ✅ Full |
| **Styles** |
| Built-in styles | `doc.styles` | `paragraph.SetStyle()` | ✅ Full |
| Custom styles | `styles.add_style()` | Limited API | ⚠️ Partial |
| **Advanced** |
| Comments | `doc.add_comment()` | Struct exists | ⚠️ Partial |
| Track changes | Yes | No | ❌ Not implemented |
| Charts | Yes | No | ❌ Not implemented |
| SmartArt | Yes | No | ❌ Not implemented |

**Overall Feature Parity: ~75%** 🎯

- ✅ **Full Support:** ~90% of common use cases
- ⚠️ **Partial Support:** ~5% (styles, comments)
- ❌ **Not Supported:** ~5% (charts, SmartArt, mail merge)

## 🚀 Performance & Advantages

### Why Choose Go-DOCX?

1. **Type Safety** - Go's static typing prevents runtime errors
2. **Performance** - Compiled binary, faster than interpreted Python
3. **Deployment** - Single binary, no dependencies to install
4. **Concurrency** - Native goroutines for parallel processing
5. **Memory Efficient** - Better memory management than Python
6. **Production Ready** - Tested and stable for production use

### When to Use Go-DOCX

✅ Server-side document generation  
✅ Microservices architecture  
✅ High-performance scenarios  
✅ Production deployments  
✅ Documents with tables, images, and lists  
✅ Automated report generation  
✅ Go ecosystem integration  

### When to Use python-docx

✅ Charts and SmartArt required  
✅ Mail merge functionality  
✅ Track changes needed  
✅ Python ecosystem integration  
✅ Rapid prototyping  

## 📖 Documentation

For detailed API documentation, see:
- [Feature Parity Analysis](FEATURE_PARITY.md) - Complete comparison with python-docx
- [Discovery Document](DISCOVERY.md) - Detailed feature discovery and analysis
- [API Examples](example/demo/) - Working code examples

## 🧪 Testing

Run the test suite:

```bash
go test ./...
```

Run specific tests:

```bash
go test -v -run TestInlinePicture
go test -v -run TestHyperlink
go test -v -run TestNumbered
```

All tests include round-trip validation (save and reopen).

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup

```bash
# Clone the repository
git clone https://github.com/SanjarbekSaminjonov/go-docx.git
cd go-docx

# Run tests
go test ./...

# Run tests with coverage
go test -cover ./...

# Format code
gofmt -w .
```

### Areas for Contribution

We welcome contributions in these areas:

1. **Custom Styles API** - Complete the styles creation and management
2. **Comments Integration** - Finish the comments API implementation
3. **Charts Support** - Add basic chart creation capabilities
4. **Performance Optimization** - Improve XML parsing/generation
5. **Documentation** - More examples and tutorials
6. **Bug Fixes** - Report and fix any issues you find

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Inspired by the excellent [python-docx](https://github.com/python-openxml/python-docx) library
- Built following the [Office Open XML](http://officeopenxml.com/) specification
- Thanks to all contributors and users

## 📞 Support

- **Issues:** [GitHub Issues](https://github.com/SanjarbekSaminjonov/go-docx/issues)
- **Discussions:** [GitHub Discussions](https://github.com/SanjarbekSaminjonov/go-docx/discussions)

## 🗺️ Roadmap

### Completed Features ✅
- ✅ Core document operations (create, open, save)
- ✅ Paragraphs and text formatting
- ✅ Tables with advanced formatting
- ✅ Headers and footers
- ✅ Images (PNG, JPEG, GIF, BMP, TIFF)
- ✅ Hyperlinks (URL and anchor)
- ✅ Lists (numbered and bulleted)
- ✅ Sections and page layout
- ✅ Built-in styles

### In Progress 🚧
- ⚠️ Custom styles creation API
- ⚠️ Comments integration
- ⚠️ Advanced paragraph formatting options

### Planned Features 📋
- 📋 Charts and graphs
- 📋 SmartArt support
- 📋 Text boxes
- 📋 Content controls
- 📋 Mail merge templates
- 📋 Track changes support

### Future Enhancements 🔮
- 🔮 Performance optimizations
- 🔮 Streaming API for large documents
- 🔮 Template engine integration
- 🔮 PDF conversion support

---

**Made with ❤️ for the Go community**

*Go-DOCX - Production-ready Word document generation for Go*