# Go-DOCX Examples

This directory contains example programs demonstrating how to use the go-docx library effectively.

## Running Examples

To run any example:

```bash
cd examples
go run basic_example.go
```

Or build and run:

```bash
cd examples
go build basic_example.go
./basic_example
```

## Available Examples

### 1. basic_example.go

**Purpose**: Demonstrates fundamental features of the library

**Features shown**:
- Creating a new document
- Setting document properties (metadata)
- Adding headings at different levels
- Text formatting (bold, italic, underline, colors)
- Text highlighting
- Creating numbered lists
- Creating bulleted lists
- Multi-level lists
- Creating tables
- Table formatting with borders
- Paragraph alignment options
- Page breaks
- Saving documents

**Output**: Creates `basic_example.docx`

**Recommended for**: Beginners wanting to understand the basics

### 2. error_handling_example.go

**Purpose**: Demonstrates proper error handling and defensive programming

**Features shown**:
- Proper error checking patterns
- Nil pointer validation
- Input validation
- Safe table creation with bounds checking
- Handling file operation errors
- Error wrapping with `%w`
- Resource cleanup with defer
- Graceful handling of invalid inputs

**Output**: Creates `error_handling_example.docx`

**Recommended for**: Developers wanting to write robust code

## Best Practices Demonstrated

All examples follow these best practices:

1. **Error Handling**: Check all errors and handle them appropriately
2. **Resource Cleanup**: Use `defer doc.Close()` to ensure cleanup
3. **Input Validation**: Validate inputs before processing
4. **Nil Checks**: Check for nil pointers before dereferencing
5. **Clear Code**: Use descriptive variable names and comments
6. **Error Messages**: Provide helpful error messages with context

## Adding Your Own Examples

When creating examples:

1. Keep them focused on specific features
2. Include clear comments explaining what each section does
3. Handle errors properly
4. Follow Go naming conventions
5. Test the example before submitting

## Example Template

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
    
    // Add content
    doc.AddParagraph("Hello, World!")
    
    // Save with error handling
    if err := doc.SaveAs("output.docx"); err != nil {
        log.Fatalf("Failed to save: %v", err)
    }
    
    log.Println("Document created successfully")
}
```

## More Examples Coming Soon

We're working on adding more examples covering:
- Advanced table features (cell merging, styling)
- Images and pictures
- Headers and footers
- Hyperlinks
- Sections with different layouts
- Advanced text formatting
- Custom styles

## Contributing Examples

Have a great example to share? Please:

1. Ensure it follows the best practices above
2. Test it thoroughly
3. Add it to this README
4. Submit a pull request

Your examples help other developers learn the library!

## Questions?

- Check the main README.md for API reference
- Review the CONTRIBUTING.md for coding standards
- Open an issue if you need help

Happy coding! 🚀
