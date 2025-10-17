package main

import (
	"errors"
	"fmt"
	"log"

	"github.com/SanjarbekSaminjonov/go-docx"
)

// This example demonstrates proper error handling when using the go-docx library.
// It shows best practices for checking errors and handling edge cases.
func main() {
	if err := createDocumentWithErrorHandling(); err != nil {
		log.Fatalf("Failed to create document: %v", err)
	}
	log.Println("Document created successfully with proper error handling")
}

func createDocumentWithErrorHandling() error {
	// Create a new document
	doc := docx.NewDocument()
	if doc == nil {
		return errors.New("failed to create new document")
	}
	defer func() {
		if err := doc.Close(); err != nil {
			log.Printf("Warning: failed to close document: %v", err)
		}
	}()

	// Set document properties with nil checks
	props := doc.CoreProperties()
	if props != nil {
		props.SetTitle("Error Handling Example")
		props.SetCreator("Go-DOCX")
		props.SetDescription("Demonstrates proper error handling patterns")
	}

	// Add heading with error handling
	title, err := doc.AddHeading("Error Handling Best Practices", 0)
	if err != nil {
		return fmt.Errorf("failed to add title heading: %w", err)
	}
	if title != nil {
		title.SetAlignment(docx.WDAlignParagraphCenter)
	}

	// Validate heading level (this will return an error for invalid levels)
	if _, err := doc.AddHeading("Invalid Level Test", 10); err != nil {
		// Expected error for level > 9
		log.Printf("Expected error for invalid heading level: %v", err)
	}

	// Example 1: Safe table creation
	section1, err := doc.AddHeading("1. Safe Table Creation", 1)
	if err != nil {
		return fmt.Errorf("failed to add section heading: %w", err)
	}
	if section1 != nil {
		section1.SetAlignment(docx.WDAlignParagraphLeft)
	}

	doc.AddParagraph("Always validate table dimensions before use:")

	// Create table with validation
	rows, cols := 3, 3
	if rows <= 0 || cols <= 0 {
		return fmt.Errorf("invalid table dimensions: rows=%d, cols=%d", rows, cols)
	}

	table := doc.AddTable(rows, cols)
	if table == nil {
		return errors.New("failed to create table")
	}

	// Safe cell access with bounds checking
	for i := 0; i < rows; i++ {
		row := table.Row(i)
		if row == nil {
			log.Printf("Warning: row %d is nil", i)
			continue
		}

		for j := 0; j < cols; j++ {
			cell := row.Cell(j)
			if cell == nil {
				log.Printf("Warning: cell [%d,%d] is nil", i, j)
				continue
			}
			cell.SetText(fmt.Sprintf("Cell %d,%d", i, j))
		}
	}

	// Example 2: Safe file operations
	doc.AddHeading("2. File Operations", 1)
	doc.AddParagraph("When working with files, always check for errors:")

	// Example of opening a document (would fail if file doesn't exist)
	testPath := "nonexistent_file.docx"
	if _, err := docx.OpenDocument(testPath); err != nil {
		// Expected error - file doesn't exist
		p := doc.AddParagraph()
		p.AddRun("Opening non-existent file correctly returns error: ").SetBold(false)
		run := p.AddRun(err.Error())
		run.SetItalic(true)
		run.SetColor("FF0000")
	}

	// Example 3: Safe list operations
	doc.AddHeading("3. List Operations", 1)
	doc.AddParagraph("Always handle negative or invalid list levels:")

	// Valid list items
	validLevels := []int{0, 1, 2, 0}
	for i, level := range validLevels {
		text := fmt.Sprintf("Item %d (level %d)", i+1, level)
		doc.AddNumberedParagraph(text, level)
	}

	// Test with negative level (library handles this gracefully)
	negativeLevel := -1
	p := doc.AddParagraph()
	p.AddRun(fmt.Sprintf("Negative level (%d) is automatically corrected to 0:", negativeLevel))
	doc.AddNumberedParagraph("This item had level -1", negativeLevel)

	// Example 4: Nil pointer checks
	doc.AddHeading("4. Nil Pointer Safety", 1)
	doc.AddParagraph("The library includes nil pointer checks in critical methods:")

	exampleP := doc.AddParagraph()
	if exampleP != nil {
		run := exampleP.AddRun("Safe operations")
		if run != nil {
			run.SetBold(true)
			run.SetColor("008000")
		}
	}

	// Example 5: Resource cleanup
	doc.AddHeading("5. Resource Cleanup", 1)
	doc.AddParagraph("Always use defer to ensure proper cleanup:")

	codeExample := doc.AddParagraph()
	codeExample.AddRun("doc := docx.NewDocument()\n").SetFont("Courier New")
	codeExample.AddRun("defer doc.Close() // Ensures cleanup even if errors occur\n").SetFont("Courier New")

	// Example 6: Error wrapping
	doc.AddHeading("6. Error Wrapping", 1)
	doc.AddParagraph("Use fmt.Errorf with %w to maintain error chains:")

	errorExample := doc.AddParagraph()
	run := errorExample.AddRun("return fmt.Errorf(\"failed to process: %w\", err)")
	run.SetFont("Courier New")
	run.SetColor("0000FF")

	doc.AddParagraph("This preserves the error chain and allows errors.Is() and errors.As() to work correctly.")

	// Example 7: Validation
	doc.AddHeading("7. Input Validation", 1)
	doc.AddParagraph("Always validate inputs before processing:")

	validationList := []string{
		"Check for nil pointers",
		"Validate numeric ranges (e.g., heading levels 0-9)",
		"Verify array bounds before access",
		"Ensure file paths are valid",
		"Confirm data is not empty when required",
	}

	for _, item := range validationList {
		doc.AddBulletedParagraph(item, 0)
	}

	// Save with error handling
	outputPath := "error_handling_example.docx"
	if err := doc.SaveAs(outputPath); err != nil {
		return fmt.Errorf("failed to save document to %s: %w", outputPath, err)
	}

	log.Printf("Successfully saved document to: %s", outputPath)
	return nil
}
