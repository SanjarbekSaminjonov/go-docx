package main

import (
	"log"

	"github.com/SanjarbekSaminjonov/go-docx"
)

// This example demonstrates basic usage of the go-docx library
// including creating documents, adding content, and saving files.
func main() {
	// Create a new document
	doc := docx.NewDocument()
	defer doc.Close()

	// Set document properties (metadata)
	props := doc.CoreProperties()
	props.SetTitle("Go-DOCX Basic Example")
	props.SetCreator("Go-DOCX Library")
	props.SetSubject("Basic Document Creation")
	props.SetKeywords("golang, docx, example")
	props.SetDescription("A simple example demonstrating go-docx basic features")

	// Add a title heading
	title, err := doc.AddHeading("Go-DOCX Basic Example", 0)
	if err != nil {
		log.Fatalf("Failed to add title: %v", err)
	}
	title.SetAlignment(docx.WDAlignParagraphCenter)

	// Add introduction section
	intro, err := doc.AddHeading("1. Introduction", 1)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}
	intro.SetAlignment(docx.WDAlignParagraphLeft)

	// Add a paragraph with basic text
	p1 := doc.AddParagraph("This is a basic example of using the go-docx library to create Word documents programmatically. ")
	p1.AddRun("The library provides a simple and intuitive API ").SetBold(false)
	p1.AddRun("for creating professional documents.").SetBold(true)

	// Add text formatting examples
	doc.AddHeading("2. Text Formatting", 1)

	p2 := doc.AddParagraph()
	p2.AddRun("Normal text, ")
	p2.AddRun("bold text, ").SetBold(true)
	p2.AddRun("italic text, ").SetItalic(true)
	p2.AddRun("underlined text, ").SetUnderline(docx.WDUnderlineSingle)
	p2.AddRun("colored text.").SetColor("FF0000")

	// Add a paragraph with highlighting
	p3 := doc.AddParagraph()
	p3.AddRun("You can also add ").SetBold(false)
	highlightedRun := p3.AddRun("highlighted text")
	highlightedRun.SetHighlight(docx.WDColorIndexYellow)
	highlightedRun.SetBold(true)
	p3.AddRun(" to emphasize important content.")

	// Add lists section
	doc.AddHeading("3. Lists", 1)

	// Numbered list
	doc.AddParagraph("Here's a numbered list:")
	doc.AddNumberedParagraph("First item", 0)
	doc.AddNumberedParagraph("Second item", 0)
	doc.AddNumberedParagraph("Sub-item 2.1", 1)
	doc.AddNumberedParagraph("Sub-item 2.2", 1)
	doc.AddNumberedParagraph("Third item", 0)

	// Bulleted list
	doc.AddParagraph("And here's a bulleted list:")
	doc.AddBulletedParagraph("First bullet point", 0)
	doc.AddBulletedParagraph("Second bullet point", 0)
	doc.AddBulletedParagraph("Nested bullet point", 1)
	doc.AddBulletedParagraph("Third bullet point", 0)

	// Add tables section
	doc.AddHeading("4. Tables", 1)

	doc.AddParagraph("Tables are useful for organizing data:")

	// Create a simple 4x3 table
	table := doc.AddTable(4, 3)

	// Add header row
	table.Row(0).Cell(0).SetText("Name")
	table.Row(0).Cell(1).SetText("Age")
	table.Row(0).Cell(2).SetText("City")

	// Add data rows
	table.Row(1).Cell(0).SetText("Alice")
	table.Row(1).Cell(1).SetText("28")
	table.Row(1).Cell(2).SetText("New York")

	table.Row(2).Cell(0).SetText("Bob")
	table.Row(2).Cell(1).SetText("34")
	table.Row(2).Cell(2).SetText("London")

	table.Row(3).Cell(0).SetText("Charlie")
	table.Row(3).Cell(1).SetText("25")
	table.Row(3).Cell(2).SetText("Tokyo")

	// Style the table with borders
	border := docx.TableBorder{
		Style: "single",
		Color: "000000",
		Size:  4,
		Space: 0,
	}
	table.SetBorder(docx.TableBorderTop, border)
	table.SetBorder(docx.TableBorderBottom, border)
	table.SetBorder(docx.TableBorderLeft, border)
	table.SetBorder(docx.TableBorderRight, border)
	table.SetBorder(docx.TableBorderInsideH, border)
	table.SetBorder(docx.TableBorderInsideV, border)

	// Add paragraph alignment section
	doc.AddHeading("5. Paragraph Alignment", 1)

	leftP := doc.AddParagraph("This paragraph is left-aligned (default).")
	leftP.SetAlignment(docx.WDAlignParagraphLeft)

	centerP := doc.AddParagraph("This paragraph is center-aligned.")
	centerP.SetAlignment(docx.WDAlignParagraphCenter)

	rightP := doc.AddParagraph("This paragraph is right-aligned.")
	rightP.SetAlignment(docx.WDAlignParagraphRight)

	justifyP := doc.AddParagraph("This paragraph is justified. It will align text on both left and right margins by adjusting spacing between words. This is commonly used in formal documents and books.")
	justifyP.SetAlignment(docx.WDAlignParagraphJustify)

	// Add page break
	doc.AddPageBreak()

	// Add second page content
	doc.AddHeading("6. Additional Features", 1)

	doc.AddParagraph("The go-docx library supports many additional features including:")

	features := []string{
		"Headers and footers for each section",
		"Multiple sections with different page layouts",
		"Page size and orientation control",
		"Page margins customization",
		"Paragraph spacing and indentation",
		"Tab stops and custom tab positions",
		"Paragraph borders and shading",
		"Table cell merging (horizontal and vertical)",
		"Images and pictures (inline)",
		"Hyperlinks (both URL and anchor)",
		"Document properties and metadata",
	}

	for _, feature := range features {
		doc.AddBulletedParagraph(feature, 0)
	}

	// Add conclusion
	doc.AddHeading("7. Conclusion", 1)

	conclusion := doc.AddParagraph()
	conclusion.AddRun("The ").SetBold(false)
	conclusion.AddRun("go-docx").SetBold(true)
	conclusion.AddRun(" library provides a comprehensive API for creating professional Word documents in Go. ")
	conclusion.AddRun("It's easy to use, well-documented, and actively maintained.").SetItalic(true)

	// Save the document
	outputPath := "basic_example.docx"
	if err := doc.SaveAs(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	log.Printf("Document successfully created: %s", outputPath)
}
