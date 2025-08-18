# Document Operations MCP Server

[![npm version](https://img.shields.io/npm/v/doc-ops-mcp.svg)](https://www.npmjs.com/package/doc-ops-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://img.shields.io/npm/dm/doc-ops-mcp.svg)](https://www.npmjs.com/package/doc-ops-mcp)

**Language / 语言**: [English](README.md) | [中文](README_zh.md)

> **Document Operations MCP Server** - A universal MCP server for document processing, conversion, and automation. Handle PDF, DOCX, HTML, Markdown, and more through a unified API and toolset.

## Table of Contents

1. [Getting Started](#1-getting-started)
2. [System Architecture](#2-system-architecture)
3. [Optional Integration](#3-optional-integration)
4. [Features](#4-features)
5. [Performance Metrics](#5-performance-metrics)
6. [Open Source Licenses](#6-open-source-licenses)
7. [Future Roadmap](#7-future-roadmap)
8. [Docker Deployment](#8-docker-deployment)
9. [Development Guide](#9-development-guide)
10. [Troubleshooting](#10-troubleshooting)
11. [Contributing](#11-contributing)

## 1. Getting Started

First, add the Document Operations MCP server to your MCP client.

**Standard config** works in most MCP clients:

```json
{
  "mcpServers": {
    "doc-ops-mcp": {
      "command": "npx",
      "args": ["-y", "doc-ops-mcp"]
    }
  }
}
```

<details>
<summary>Claude Desktop</summary>

Follow the MCP install [guide](https://modelcontextprotocol.io/quickstart/user), use the standard config above.

</details>

<details>
<summary>VS Code</summary>

Follow the MCP install [guide](https://code.visualstudio.com/docs/copilot/chat/mcp-servers#_add-an-mcp-server), use the standard config above.

</details>

<details>
<summary>Cursor</summary>

Go to `Cursor Settings` -> `MCP` -> `Add new MCP Server`. Name to your liking, use `command` type with the command `npx -y doc-ops-mcp`.

</details>

<details>
<summary>Other MCP Clients</summary>

For other MCP clients, use the standard config above and refer to your client's documentation for MCP server installation.

</details>

### Configuration

The Document Operations MCP server supports configuration through environment variables. These can be provided in the MCP client configuration as part of the `"env"` object:

```json
{
  "mcpServers": {
    "doc-ops-mcp": {
      "command": "npx",
>>>>>>> main-test
      "args": ["-y", "doc-ops-mcp"],
      "env": {
        "OUTPUT_DIR": "/path/to/your/output/directory",
        "CACHE_DIR": "/path/to/your/cache/directory",
        "WATERMARK_IMAGE": "/path/to/watermark.png",
        "QR_CODE_IMAGE": "/path/to/qrcode.png"
      }
    }
  }
}
```

#### Environment Variables

**Core Directories:**
- **`OUTPUT_DIR`**: Controls where all generated files are saved (default: `~/Documents`)
- **`CACHE_DIR`**: Directory for temporary and cache files (default: `~/.cache/doc-ops-mcp`)

**PDF Enhancement Features:**
- **`WATERMARK_IMAGE`**: Default watermark image path for PDF files
  - Automatically added to all PDF conversions
  - Supported formats: PNG, JPG
  - If not set, default text watermark "doc-ops-mcp" will be used
- **`QR_CODE_IMAGE`**: Default QR code image path for PDF files
  - Added to PDFs only when explicitly requested (`addQrCode=true`)
  - Supported formats: PNG, JPG
  - If not set, QR code functionality will be unavailable

**Output Path Rules:**
1. If `outputPath` is not provided → files saved to `OUTPUT_DIR` with auto-generated names
2. If `outputPath` is relative → resolved relative to `OUTPUT_DIR`
3. If `outputPath` is absolute → used as-is, ignoring `OUTPUT_DIR`

See [OUTPUT_PATH_CONTROL.md](./OUTPUT_PATH_CONTROL.md) for detailed documentation.

## 2. System Architecture

Document Operations MCP Server adopts a pure JavaScript architecture design, providing complete document processing capabilities:

```
┌─────────────────────────────────────────────────────────────┐
│                    MCP Client Layer                         │
│           (Claude Desktop, Cursor, VS Code, etc.)           │
└─────────────────────┬───────────────────────────────────────┘
                      │ JSON-RPC 2.0
┌─────────────────────┴───────────────────────────────────────┐
│                 Doc-Ops-MCP Server                         │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────┐ │
│  │   Tool Router   │  │  Request        │  │  Response   │ │
│  │   & Handler     │  │  Validator      │  │  Formatter  │ │
│  └────────┬────────┘  └────────┬────────┘  └──────┬──────┘ │
│           │                    │                  │        │
│  ┌────────┴────────────────────┴──────────────────┴─────┐ │
│  │                Document Processing Engine             │ │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐   │ │
│  │  │  Document   │  │   Format    │  │   Style     │   │ │
│  │  │   Reader    │  │  Converter  │  │  Processor  │   │ │
│  │  └─────────────┘  └─────────────┘  └─────────────┘   │ │
│  │                                                        │ │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐   │ │
│  │  │    PDF      │  │  Watermark/ │  │ Conversion  │   │ │
│  │  │ Enhancement │  │   QR Code   │  │  Planner    │   │ │
│  │  └─────────────┘  └─────────────┘  └─────────────┘   │ │
└────┴───────────────────────────────────────────────────────┴─┘
                            │
┌───────────────────────────┴─────────────────────────────────┐
│                    Core Dependencies Layer                  │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐          │
│  │   pdf-lib   │  │   mammoth   │  │   marked    │          │
│  │ (PDF Tools) │  │(DOCX Tools) │  │ (Markdown)  │          │
│  └─────────────┘  └─────────────┘  └─────────────┘          │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐          │
│  │   cheerio   │  │   turndown  │  │    docx     │          │
│  │ (HTML Parse)│  │ (HTML to MD)│  │ (DOCX Gen)  │          │
│  └─────────────┘  └─────────────┘  └─────────────┘          │
└─────────────────────────────────────────────────────────────┘
```

### Architecture Overview

**Core Features**:
- Pure JavaScript implementation with no external system dependencies
- Complete document reading, format conversion, and style processing capabilities
- Built-in PDF watermark and QR code addition functionality
- Intelligent conversion planning and path optimization

**Conversion Flow**:
- **Direct Conversion**: Supports most format-to-format conversions directly
- **Multi-step Conversion**: Complex conversions achieved through intermediate formats
- **Style Preservation**: Uses OOXML parser to ensure complete style integrity

## 3. Optional Integration

This server can work with `playwright-mcp` for enhanced PDF conversion capabilities. Please refer to the official `playwright-mcp` documentation for detailed configuration.

### 🔧 PDF Conversion Workflow

This server supports complete PDF conversion functionality:
1. **Document Parsing**: Use OOXML parser to ensure complete style preservation
2. **Format Conversion**: Convert documents to high-quality HTML format
3. **PDF Generation**: Built-in converter or optionally work with `playwright-mcp`
4. **Enhancement Processing**: Automatically add watermarks and QR codes (if configured)

### How It Works

This server uses intelligent conversion architecture:
1. **Smart Planning**: `plan_conversion` analyzes conversion requirements and selects optimal paths
2. **Format Conversion**: Use specialized converters to handle various document formats
3. **Style Preservation**: Ensure style integrity through OOXML parser
4. **Enhancement Processing**: Automatically add watermarks, QR codes and other enhancements
5. **Optional Integration**: Support working with `playwright-mcp` for enhanced capabilities

## 4. Features

### MCP Tools

#### Core Document Tools

| Tool Name | Description | Input Parameters | External Dependencies |
|-----------|-------------|------------------|----------------------|
| `read_document` | Read document content | `filePath`: Document path<br>`extractMetadata`: Extract metadata<br>`preserveFormatting`: Preserve formatting | None |
| `write_document` | Write document content | `content`: Document content<br>`outputPath`: Output file path<br>`encoding`: File encoding | None |
| `convert_document` | Smart document conversion | `inputPath`: Input file path<br>`outputPath`: Output file path<br>`preserveFormatting`: Preserve formatting | None |
| `plan_conversion` | Conversion planner | `sourceFormat`: Source format<br>`targetFormat`: Target format<br>`preserveStyles`: Preserve styles<br>`quality`: Conversion quality | None |

##### **read_document**
Read various document formats including PDF, DOCX, DOC, HTML, MD, and more.

**Parameters:**
- `filePath` (string, required) - Document path to read
- `extractMetadata` (boolean, optional) - Extract document metadata, defaults to `false`
- `preserveFormatting` (boolean, optional) - Preserve formatting (HTML output), defaults to `false`

##### **write_document**
Write content to document files in specified formats.

**Parameters:**
- `content` (string, required) - Content to write
- `outputPath` (string, optional) - Output file path (auto-generated if not provided)
- `encoding` (string, optional) - File encoding, defaults to `utf-8`

##### **convert_document**
Convert documents between formats with enhanced style preservation.

**Parameters:**
- `inputPath` (string, required) - Input file path
- `outputPath` (string, optional) - Output file path (auto-generated if not provided)
- `preserveFormatting` (boolean, optional) - Preserve formatting, defaults to `true`
- `useInternalPlaywright` (boolean, optional) - Use built-in Playwright for PDF conversion, defaults to `false`

##### **convert_docx_to_pdf**
Convert DOCX to PDF with automatic watermark addition (if configured).

**Parameters:**
- `docxPath` (string, required) - DOCX file path
- `outputPath` (string, optional) - Output PDF path (auto-generated if not provided)
- `addQrCode` (boolean, optional) - Whether to add QR code, defaults to `false`
- `preserveFormatting` (boolean, optional) - Preserve original formatting, defaults to `true`
- `chineseFont` (string, optional) - Chinese font, defaults to `Microsoft YaHei`

##### **convert_markdown_to_pdf**
Convert Markdown to PDF with automatic watermark addition (if configured).

**Parameters:**
- `markdownPath` (string, required) - Markdown file path
- `outputPath` (string, optional) - Output PDF path (auto-generated if not provided)
- `theme` (string, optional) - Theme style, defaults to `"github"`
- `includeTableOfContents` (boolean, optional) - Include table of contents, defaults to `false`
- `addQrCode` (boolean, optional) - Whether to add QR code, defaults to `false`

##### **convert_markdown_to_html**
Convert Markdown to HTML.

**Parameters:**
- `markdownPath` (string, required) - Markdown file path
- `outputPath` (string, optional) - Output HTML path (auto-generated if not provided)
- `theme` (string, optional) - Theme style, defaults to `"github"`
- `includeTableOfContents` (boolean, optional) - Include table of contents, defaults to `false`

##### **convert_markdown_to_docx**
Convert Markdown to DOCX.

**Parameters:**
- `markdownPath` (string, required) - Markdown file path
- `outputPath` (string, optional) - Output DOCX path (auto-generated if not provided)

##### **convert_html_to_markdown**
Convert HTML to Markdown.

**Parameters:**
- `htmlPath` (string, required) - HTML file path
- `outputPath` (string, optional) - Output Markdown path (auto-generated if not provided)

##### **plan_conversion**
🎯 Smart Conversion Planner - Analyze conversion requirements and generate optimal conversion plans.

**Parameters:**
- `sourceFormat` (string, required) - Source file format (pdf, docx, html, markdown, md, txt, doc)
- `targetFormat` (string, required) - Target file format (pdf, docx, html, markdown, md, txt, doc)
- `sourceFile` (string, optional) - Source file path (for generating specific conversion parameters)
- `preserveStyles` (boolean, optional) - Whether to preserve style formatting, defaults to `true`
- `includeImages` (boolean, optional) - Whether to include images, defaults to `true`
- `theme` (string, optional) - Conversion theme, defaults to `github`
- `quality` (string, optional) - Conversion quality requirement (fast, balanced, high), defaults to `balanced`

##### **process_pdf_post_conversion**

**Parameters:**
- `playwrightPdfPath` (string, required) - Generated PDF file path
- `targetPath` (string, optional) - Target PDF file path (auto-generated if not provided)
- `addWatermark` (boolean, optional) - Whether to add watermark, defaults to `false`
- `addQrCode` (boolean, optional) - Whether to add QR code, defaults to `false`
- `watermarkImage` (string, optional) - Watermark image path
- `qrCodePath` (string, optional) - QR code image path

#### PDF Enhancement Tools

##### **add_watermark**
🎨 PDF Watermark Addition Tool - Add image or text watermarks to PDF documents.

**Parameters:**
- `pdfPath` (string, required) - PDF file path
- `watermarkImage` (string, optional) - Watermark image path (PNG/JPG)
- `watermarkText` (string, optional) - Watermark text content
- `watermarkImageScale` (number, optional) - Image scale ratio, defaults to `0.25`
- `watermarkImageOpacity` (number, optional) - Image opacity, defaults to `0.6`
- `watermarkImagePosition` (string, optional) - Image position, defaults to `fullscreen`

##### **add_qrcode**
📱 PDF QR Code Addition Tool - Add QR codes to PDF documents.

**Parameters:**
- `pdfPath` (string, required) - PDF file path
- `qrCodePath` (string, optional) - QR code image path
- `qrScale` (number, optional) - QR code scale ratio, defaults to `0.15`
- `qrOpacity` (number, optional) - QR code opacity, defaults to `1.0`
- `qrPosition` (string, optional) - QR code position, defaults to `bottom-center`
- `addText` (boolean, optional) - Whether to add explanatory text, defaults to `true`

## 5. Performance Metrics

### Document Processing Capabilities

| Document Type | Max File Size | Processing Speed | Memory Usage |
|---------------|---------------|------------------|---------------|
| **PDF** | 50MB | 2-5MB/s | ~File size×1.5 |
| **DOCX** | 50MB | 5-10MB/s | ~File size×2 |
| **HTML** | 50MB | 10-20MB/s | ~File size×1.2 |
| **Markdown** | 50MB | 15-30MB/s | ~File size×1.1 |

### Conversion Performance

- **PDF Conversion**: Requires playwright-mcp integration, ~1-3 pages/second
- **DOCX Conversion**: Pure JavaScript processing, ~5-15 pages/second
- **HTML Conversion**: Fastest, ~20-50 pages/second
- **Concurrent Processing**: Supports up to 5 concurrent tasks

### System Resource Requirements

- **Minimum Memory**: 512MB
- **Recommended Memory**: 2GB (for large files)
- **CPU**: Single core sufficient, multi-core improves concurrency
- **Disk Space**: Temporary files require 2-3x original file size

## System Requirements

### System Requirements
- **Node.js** ≥ 18.0.0
- **Zero external system dependencies** - All processing via npm packages
- **Optional Integration**: playwright-mcp for enhanced PDF conversion

### Core Technology Stack
- **pdf-lib** - PDF operations and enhancement
- **mammoth** - DOCX document processing  
- **marked** - Markdown parsing and rendering
- **cheerio** - HTML parsing and manipulation
- **turndown** - HTML to Markdown conversion
- **docx** - DOCX document generation

### Installation
```bash
# Global installation
npm install -g doc-ops-mcp

# Or using pnpm
pnpm add -g doc-ops-mcp

# Or using bun
bun add -g doc-ops-mcp
```

### Architecture Components

- **MCP Server Core**: Handles JSON-RPC 2.0 communication and tool registration
- **Smart Router**: Routes requests to optimal processing modules
- **Conversion Engine**: Contains specialized converters for different document types
- **Style Processor**: Ensures style preservation during format conversion
- **Security Module**: Provides path validation and content security handling

## 6. Open Source Licenses

### Project License
- **This Project**: MIT License
- **Compatibility**: Available for commercial and non-commercial use

### Third-Party Dependencies

| Library | Version | License | Purpose |
|---------|---------|---------|----------|
| **pdf-lib** | ^1.17.1 | MIT | PDF document manipulation |
| **mammoth** | ^1.6.0 | BSD-2-Clause | DOCX parsing and conversion |
| **marked** | ^9.1.6 | MIT | Markdown parsing and rendering |
| **exceljs** | ^4.4.0 | MIT | Excel file processing |
| **jsdom** | ^23.0.1 | MIT | HTML DOM manipulation |
| **turndown** | ^7.1.2 | MIT | HTML to Markdown conversion |

### License Compatibility
- ✅ **Commercial Use**: All dependencies support commercial use
- ✅ **Distribution**: Free to distribute and modify
- ✅ **Patent Protection**: Apache-2.0 provides patent protection
- ⚠️ **Notice**: Original license notices must be retained

## 7. Future Roadmap

### Core Features
- 🔄 **Enhanced Conversion Quality**: Improve style preservation for complex documents
- 📊 **Excel Support**: Complete Excel read/write and conversion functionality
- 🎨 **Template System**: Support for custom document templates
- 🔍 **OCR Integration**: Image text recognition capabilities

### System Improvements
- 🌐 **Multi-language Support**: Internationalization and localization
- 🔐 **Security Enhancements**: Document encryption and access control
- ⚡ **Performance Optimization**: Large file handling and memory optimization
- 🔌 **Plugin System**: Extensible processor architecture

### Version Roadmap
- **v2.0**: Complete Excel support and template system
- **v3.0**: OCR integration and multi-language support
- **v4.0**: Advanced security features and plugin system

## 8. Docker Deployment

### Quick Start

#### Using Pre-built Image

```bash
# Pull the latest image
docker pull docops/doc-ops-mcp:latest

# Run with default configuration
docker run -d \
  --name doc-ops-mcp \
  -p 3000:3000 \
  docops/doc-ops-mcp:latest
```

#### Building from Source

```bash
# Clone the repository
git clone https://github.com/JefferyMunoz/doc-ops-mcp.git
cd doc-ops-mcp

# Build the Docker image
docker build -t doc-ops-mcp .

# Run the container
docker run -d \
  --name doc-ops-mcp \
  -p 3000:3000 \
  -v $(pwd)/documents:/app/documents \
  doc-ops-mcp
```

### Docker Compose Deployment

Create a `docker-compose.yml` file:

```yaml
version: '3.8'

services:
  doc-ops-mcp:
    image: docops/doc-ops-mcp:latest
    container_name: doc-ops-mcp
    ports:
      - "3000:3000"
    volumes:
      - ./documents:/app/documents
      - ./config:/app/config
    environment:
      - NODE_ENV=production
      - PORT=3000
    restart: unless-stopped
    
  # Optional: Add Nginx for reverse proxy
  nginx:
    image: nginx:alpine
    container_name: doc-ops-nginx
    ports:
      - "80:80"
    volumes:
      - ./nginx.conf:/etc/nginx/nginx.conf:ro
    depends_on:
      - doc-ops-mcp
    restart: unless-stopped
```

### Environment Variables

| Variable | Description | Default |
|----------|-------------|----------|
| `PORT` | Server port | `3000` |
| `NODE_ENV` | Environment mode | `production` |
| `LOG_LEVEL` | Logging level | `info` |
| `MAX_FILE_SIZE` | Maximum file size (MB) | `50` |

### Volume Mounts

Mount local directories for persistent storage:

```bash
# Documents directory for file processing
docker run -d \
  --name doc-ops-mcp \
  -p 3000:3000 \
  -v $(pwd)/documents:/app/documents \
  -v $(pwd)/output:/app/output \
  doc-ops-mcp
```

### Docker Configuration Examples

#### Production Deployment

```bash
# Production setup with Docker Swarm
docker swarm init
docker stack deploy -c docker-compose.yml doc-ops

# Scale the service
docker service scale doc-ops_mcp=3
```

### Health Checks

The container includes built-in health checks:

```bash
# Check container health
docker ps

# View health check logs
docker inspect --format='{{.State.Health.Status}}' doc-ops-mcp

# Manual health check
docker exec doc-ops-mcp curl -f http://localhost:3000/health || exit 1
```

## 9. Development Guide

### Local Development

```bash
# Clone the repository
git clone https://github.com/your-org/doc-ops-mcp.git
cd doc-ops-mcp

# Install dependencies
npm install

# Run in development mode
npm run dev

# Build the project
npm run build

# Run tests
npm test
```

### Project Structure

```
src/
├── index.ts          # MCP server entry point
├── tools/            # Tool implementations
│   ├── documentConverter.ts
│   ├── pdfTools.ts
│   └── ...
├── types/            # Type definitions
└── utils/            # Utility functions
```

### Adding New Tools

1. Create a new tool file in `src/tools/`
2. Implement the tool logic
3. Register the tool in `src/index.ts`
4. Add test cases
5. Update documentation

## 10. Troubleshooting

### Common Issues

1. **Port conflicts**: Change the host port in docker-compose.yml
2. **Permission issues**: Ensure volume mounts have correct permissions
3. **Memory issues**: Increase Docker memory allocation

### Debug Mode

```bash
# Run with debug logging
docker run -d \
  --name doc-ops-mcp \
  -p 3000:3000 \
  -e LOG_LEVEL=debug \
  doc-ops-mcp

# View logs
docker logs -f doc-ops-mcp
```

## 11. Contributing

### How to Contribute

1. **Fork the Project**
2. **Create a Feature Branch** (`git checkout -b feature/AmazingFeature`)
3. **Commit Your Changes** (`git commit -m 'Add some AmazingFeature'`)
4. **Push to the Branch** (`git push origin feature/AmazingFeature`)
5. **Open a Pull Request**

#### Intellectual Property License

**By submitting a Pull Request, you agree that all contributions submitted through Pull Requests will be licensed under the MIT License.** This means:

- You grant the project maintainers and users the right to use, modify, and distribute your contributions under the MIT License
- You confirm that you have the right to make these contributions
- You understand that your contributions will become part of the open source project
- You waive any claims to exclusive ownership of the contributed code

If you cannot agree to these terms, please do not submit a Pull Request.

### Code Standards

- Use TypeScript
- Follow ESLint configuration
- Add appropriate tests
- Update relevant documentation

### Reporting Issues

- Use [GitHub Issues](https://github.com/your-org/doc-ops-mcp/issues)
- Provide detailed error information and reproduction steps
- Include system environment information

### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.