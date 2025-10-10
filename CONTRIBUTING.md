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

- Go 1.21 or higher
- Git

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

# Run specific tests
go test -v -run TestInlinePicture
```

### Code Style

- Follow standard Go formatting: `gofmt -w .`
- Run `go vet` to catch common errors
- Write clear, descriptive commit messages

## Areas Where We Need Help

1. **Custom Styles API** - Complete the styles creation and management
2. **Comments Integration** - Finish the comments API implementation
3. **Charts Support** - Add basic chart creation capabilities
4. **Performance Optimization** - Improve XML parsing/generation
5. **Documentation** - More examples and tutorials
6. **Bug Fixes** - Report and fix any issues you find

## Pull Request Guidelines

- Keep pull requests focused on a single feature/fix
- Include tests for new functionality
- Update documentation as needed
- Ensure all tests pass
- Follow the existing code style

## Code of Conduct

- Be respectful and inclusive
- Provide constructive feedback
- Focus on what is best for the project

## Questions?

Feel free to open an issue for any questions or concerns.

Thank you for contributing! 🎉
