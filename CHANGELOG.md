# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-10-10

### Added
- Initial release with comprehensive DOCX support
- Core document operations (create, open, save)
- Paragraph and text formatting with runs
- Table creation with advanced formatting:
  - Table borders (all sides, customizable)
  - Cell shading/background colors
  - Cell margins
  - Horizontal and vertical cell merging
- **Images support** (PNG, JPEG, GIF, BMP, TIFF)
  - Document-level API: `doc.AddPicture()`
  - Run-level API: `run.AddPicture()`
  - Auto aspect ratio
  - Custom dimensions in EMUs
- **Hyperlinks support** (URL and anchor)
  - `paragraph.AddHyperlink()`
  - `run.SetHyperlink()`
  - `run.SetHyperlinkAnchor()`
  - `run.HasHyperlink()`
- **Lists support** (numbered and bulleted)
  - `doc.AddNumberedParagraph()`
  - `doc.AddBulletedParagraph()`
  - Multi-level lists (0-8 levels)
  - Custom numbering
- Headers and footers (default, first page, even page)
- Sections with page layout (size, orientation, margins)
- Document properties (title, author, subject, etc.)
- Built-in styles support
- Comprehensive test suite with round-trip validation

### Documentation
- Complete README with API reference
- Feature parity analysis (FEATURE_PARITY.md)
- Discovery document (DISCOVERY.md)
- Working examples in example/demo/

### Performance
- ~75% feature parity with python-docx
- Type-safe API
- High-performance compiled binary

[1.0.0]: https://github.com/SanjarbekSaminjonov/go-docx/releases/tag/v1.0.0
