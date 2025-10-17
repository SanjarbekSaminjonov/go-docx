# Contributing to Go-DOCX

Thank you for your interest in contributing to Go-DOCX! We welcome contributions from the community.

## Getting Started

1. Fork the repository
2. Clone your fork: `git clone https://github.com/YOUR_USERNAME/go-docx.git`
3. Create a new branch: `git checkout -b feature/your-feature-name`
4. Make your changes
5. Run tests: `go test ./...`
6. Commit your changes: `git commit -am 'Add some feature'`
7. Push to the branch: `git push origin feature/your-feature-name`
8. Create a Pull Request

## Development Setup

### Prerequisites

- Go 1.23 or higher
- Git
- (Optional) golangci-lint for code quality checks

### Building

```bash
go build ./...
```

### Testing

```bash
# Run all tests
go test ./...

# Run tests with coverage
go test -cover ./...

# Run tests with verbose output
go test -v ./...

# Run specific tests
go test -v -run TestInlinePicture

# Generate coverage report
go test -coverprofile=coverage.out ./...
go tool cover -html=coverage.out
```

### Code Quality

We use `golangci-lint` to maintain code quality. Install it from [golangci-lint.run](https://golangci-lint.run/usage/install/).

```bash
# Run all linters
golangci-lint run

# Auto-fix issues where possible
golangci-lint run --fix
```

### Code Style

- Follow standard Go formatting: `gofmt -w .` or `go fmt ./...`
- Run `go vet ./...` to catch common errors
- Write clear, descriptive commit messages
- Add godoc comments to all exported types and functions
- Keep functions focused and reasonably short (< 100 lines)
- Handle errors appropriately - don't ignore them
- Add nil checks for pointer receivers where appropriate

## Coding Standards

### Documentation

All exported types, functions, constants, and variables must have godoc comments:

```go
// MyFunction performs a specific operation and returns a result.
// It returns an error if the operation fails.
func MyFunction(param string) (string, error) {
    // implementation
}
```

### Error Handling

- Always check and handle errors
- Use fmt.Errorf with %w for error wrapping to maintain error chains
- Provide meaningful error messages that help users understand what went wrong

```go
if err != nil {
    return fmt.Errorf("failed to process document: %w", err)
}
```

### Input Validation

- Validate function inputs at the beginning
- Check for nil pointers in receiver methods
- Provide clear error messages for invalid inputs

```go
func (d *Document) ProcessData(data []byte) error {
    if d == nil {
        return fmt.Errorf("document is nil")
    }
    if len(data) == 0 {
        return fmt.Errorf("data cannot be empty")
    }
    // ... rest of implementation
}
```

### Testing

- Write tests for new functionality
- Aim for at least 70% code coverage
- Use table-driven tests where appropriate
- Include both positive and negative test cases
- Test edge cases and error conditions

```go
func TestMyFunction(t *testing.T) {
    tests := []struct {
        name    string
        input   string
        want    string
        wantErr bool
    }{
        {"valid input", "test", "result", false},
        {"empty input", "", "", true},
    }
    for _, tt := range tests {
        t.Run(tt.name, func(t *testing.T) {
            got, err := MyFunction(tt.input)
            if (err != nil) != tt.wantErr {
                t.Errorf("error = %v, wantErr %v", err, tt.wantErr)
                return
            }
            if got != tt.want {
                t.Errorf("got %v, want %v", got, tt.want)
            }
        })
    }
}
```

## Areas Where We Need Help

1. **Custom Styles API** - Complete the styles creation and management
2. **Comments Integration** - Finish the comments API implementation
3. **Charts Support** - Add basic chart creation capabilities
4. **Performance Optimization** - Improve XML parsing/generation
5. **Documentation** - More examples and tutorials
6. **Bug Fixes** - Report and fix any issues you find
7. **Test Coverage** - Increase test coverage to 80%+

## Pull Request Guidelines

- Keep pull requests focused on a single feature/fix
- Include tests for new functionality
- Update documentation as needed (README, godoc comments)
- Ensure all tests pass: `go test ./...`
- Ensure code is formatted: `go fmt ./...`
- Run `go vet ./...` and fix any issues
- If available, run `golangci-lint run` and address issues
- Follow the existing code style and patterns
- Keep commits atomic and well-described

## Commit Message Guidelines

Use clear and descriptive commit messages:

```
Add support for document comments

- Implement Comment struct and methods
- Add CommentsPart for managing comments
- Include tests for comment functionality
- Update documentation with comment examples
```

## Code Review Process

1. Maintainers will review your PR within a few days
2. Address any feedback or requested changes
3. Once approved, a maintainer will merge your PR
4. Your contribution will be included in the next release

## Code of Conduct

- Be respectful and inclusive
- Provide constructive feedback
- Focus on what is best for the project
- Welcome newcomers and help them get started
- Assume good intentions

## Questions?

Feel free to:
- Open an issue for questions or concerns
- Start a discussion in GitHub Discussions
- Reach out to maintainers for guidance

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

Thank you for contributing! 🎉
