# Go-DOCX - Microsoft Word Document Library for Go

Go-DOCX is a Go library for creating and manipulating Microsoft Word (.docx) documents. It is inspired by and aims to provide similar functionality to the popular python-docx library.

## Features

- ✅ Create new Word documents
- ✅ Add paragraphs with formatted text
- ✅ Add headings with different levels
- ✅ Create and populate tables
- ✅ Text formatting (bold, italic, underline, color, font)
- ✅ Paragraph alignment
- ✅ Page breaks
- ✅ Document properties (title, author, etc.)
- ✅ Save documents to file

## Installation

```bash
go get github.com/sanjarbek/go-docx
```

## Quick Start

```go
package main

import (
    "log"
    "github.com/sanjarbek/go-docx"
)

func main() {
    // Create a new document
    doc := docx.NewDocument()
    
    // Add a title
    title, err := doc.AddHeading("My Document", 0)
    if err != nil {
        log.Fatal(err)
    }
    title.SetAlignment(docx.WDAlignParagraphCenter)
    
    // Add a paragraph
    paragraph := doc.AddParagraph("This is a sample paragraph.")
    
    // Add formatted text
    p := doc.AddParagraph()
    run := p.AddRun("This text is bold.")
    run.SetBold(true)
    
    // Save the document
    err = doc.SaveAs("my_document.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    doc.Close()
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
table := doc.AddTable(3, 4) // 3 rows, 4 columns

// Access cells
cell := table.Row(0).Cell(0)
cell.SetText("Header 1")

// Add content to cells
cell.AddParagraph("Additional content")

// Add new rows
newRow := table.AddRow()
```

### Constants

#### Alignment

- `WDAlignParagraphLeft`
- `WDAlignParagraphCenter`
- `WDAlignParagraphRight`
- `WDAlignParagraphJustify`

#### Underline Types

- `WDUnderlineNone`
- `WDUnderlineSingle`
- `WDUnderlineDouble`
- `WDUnderlineThick`
- `WDUnderlineDotted`
- `WDUnderlineDashed`

#### Break Types

- `BreakTypePage`
- `BreakTypeColumn`
- `BreakTypeText`

#### Color Indices

- `WDColorIndexAuto`
- `WDColorIndexBlack`
- `WDColorIndexBlue`
- `WDColorIndexRed`
- `WDColorIndexYellow`
- And many more...

## Examples

See the `example/` directory for complete working examples.

## Comparison with python-docx

This library aims to provide similar functionality to python-docx with Go-idiomatic APIs:

| Feature | python-docx | go-docx |
|---------|-------------|---------|
| Create document | `Document()` | `NewDocument()` |
| Add paragraph | `doc.add_paragraph()` | `doc.AddParagraph()` |
| Add heading | `doc.add_heading()` | `doc.AddHeading()` |
| Add table | `doc.add_table()` | `doc.AddTable()` |
| Text formatting | `run.bold = True` | `run.SetBold(true)` |
| Save document | `doc.save()` | `doc.SaveAs()` |

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License.

## Acknowledgments

- Inspired by the excellent [python-docx](https://github.com/python-openxml/python-docx) library
- Built following the OpenXML specification