# Document Operations MCP Server

[![npm version](https://badge.fury.io/js/doc-ops-mcp.svg)](https://badge.fury.io/js/doc-ops-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://img.shields.io/npm/dm/doc-ops-mcp.svg)](https://www.npmjs.com/package/doc-ops-mcp)

**Language / 语言**: [English](README.md) | [中文](README_zh.md)

> **Document Operations MCP Server** - A universal MCP server for document processing, conversion, and automation. Handle PDF, DOCX, HTML, Markdown, and more through a unified API and toolset.

## Table of Contents

1. [Quick Start](#1-quick-start)
2. [System Architecture](#2-system-architecture)
3. [External Dependencies](#3-external-dependencies)
4. [Features](#4-features)
5. [Performance Metrics](#5-performance-metrics)
6. [Open Source Licenses](#6-open-source-licenses)
7. [Future Roadmap](#7-future-roadmap)
8. [Docker Deployment](#8-docker-deployment)
9. [Development Guide](#9-development-guide)
10. [Troubleshooting](#10-troubleshooting)
11. [Contributing](#11-contributing)

## 1. Quick Start

### Installation
```bash
# Via npm
npm install -g doc-ops-mcp

# Via pnpm
pnpm add -g doc-ops-mcp

# Via bun
bun add -g doc-ops-mcp
```

### Configuration

```json
{
  "mcpServers": {
    "doc-ops-mcp": {
      "command": "npx",
      "args": ["-y", "doc-ops-mcp@latest"],
      "env": {
        "OUTPUT_DIR": "/path/to/your/output/directory",
        "CACHE_DIR": "/path/to/your/cache/directory"
      }
    }
  }
}
```

### Environment Variables

The server supports environment variables for controlling output paths and PDF enhancement features:

#### Core Directories
- **`OUTPUT_DIR`**: Controls where all generated files are saved (default: `~/Documents`)
- **`CACHE_DIR`**: Directory for temporary and cache files (default: `~/.cache/doc-ops-mcp`)

#### PDF Enhancement Features
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

Document Operations MCP Server adopts a hybrid architecture design, combining internal processing with external dependencies:

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
│  │  │    PDF      │  │  Watermark/ │  │    Web      │   │ │
│  │  │ Enhancement │  │   QR Code   │  │  Scraper    │   │ │
│  │  └─────────────┘  └─────────────┘  └─────────────┘   │ │
└────┴───────────────────────────────────────────────────────┴─┘
                            │
┌───────────────────────────┴─────────────────────────────────┐
│                    Core Dependencies Layer                  │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐          │
│  │   pdf-lib   │  │   mammoth   │  │   marked    │          │
│  │ (PDF Tools) │  │(DOCX Tools) │  │ (Markdown)  │          │
│  └─────────────┘  └─────────────┘  └─────────────┘          │
└─────────────────────────────────────────────────────────────┘
                            │
┌───────────────────────────┴─────────────────────────────────┐
│                External Dependencies (PDF Conversion)       │
│  ┌─────────────────────────────────────────────────────┐   │
│  │                playwright-mcp                       │   │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐ │   │
│  │  │  Browser    │  │    HTML     │  │    PDF      │ │   │
│  │  │ Automation  │  │  Rendering  │  │ Generation  │ │   │
│  │  └─────────────┘  └─────────────┘  └─────────────┘ │   │
│  └─────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

### Architecture Overview

**Internal Processing Layer**:
- Document reading, format conversion, style processing
- PDF watermark and QR code addition
- Web content scraping

**External Dependencies Layer**:
- **PDF Conversion**: Relies on `playwright-mcp` for HTML → PDF conversion
- **Conversion Flow**: DOCX/Markdown → HTML → PDF (via playwright-mcp)

**Important Note**: All PDF conversion features require `playwright-mcp` to work properly.

## 3. External Dependencies

### playwright-mcp Dependency

This MCP server's PDF conversion functionality depends on the `playwright-mcp` server:

- **Dependent Tools**: `convert_docx_to_pdf`, `convert_markdown_to_pdf`
- **Important Configuration**: `playwright-mcp` must use `--caps=pdf` parameter to provide `browser_pdf_save` command
- **Conversion Flow**:
  1. Convert source document to HTML format
  2. Use `playwright-mcp`'s `browser_pdf_save` command to render HTML to PDF
  3. Automatically add watermark (if `WATERMARK_IMAGE` is configured)
  4. Optionally add QR code (if `addQrCode=true` and `QR_CODE_IMAGE` is configured)

#### Third-Party Dependency Disclaimer

**Important Notice**: `playwright-mcp` is an independent third-party project. This project (doc-ops-mcp) does not assume responsibility for:

- The availability, security, or behavior of `playwright-mcp`
- Any issues, vulnerabilities, or data loss caused by `playwright-mcp`
- The maintenance, updates, or support of `playwright-mcp`

**Users must:**
- Review and comply with `playwright-mcp`'s own license terms and conditions
- Understand the risks associated with using third-party dependencies
- Ensure `playwright-mcp` meets their security and compliance requirements
- Monitor `playwright-mcp` for updates and security patches independently

Please refer to the official `playwright-mcp` documentation and repository for license information, security advisories, and usage guidelines.

### Configuration Requirements

1. **Install playwright-mcp**:
   ```bash
   # Please refer to playwright-mcp official documentation for installation and configuration
   ```

2. **MCP Client Configuration**:
   Ensure both this server and `playwright-mcp` are configured in your MCP client
   
   **Important**: `playwright-mcp` must use `--caps=pdf` parameter:
   ```json
   {
     "mcpServers": {
       "playwright": {
         "command": "npx",
         "args": ["@playwright/mcp@latest", "--caps=pdf"]
       }
     }
   }
   ```

3. **Environment Variables**:
   - `WATERMARK_IMAGE`: Watermark image path (optional, automatically added during PDF conversion)
   - `QR_CODE_IMAGE`: QR code image path (optional, added when requested)

## PDF Conversion Workflow

When converting documents to PDF, `doc-ops-mcp` follows this workflow:

1. **HTML Generation Stage**: Convert source documents (like Markdown, DOCX) to styled HTML files
2. **PDF Generation Stage**: Call `playwright-mcp`'s `browser_pdf_save` command to convert HTML to PDF
3. **Post-processing Stage**: **Must** call `process_pdf_post_conversion` tool to complete path conversion, watermark and QR code addition

### Important Configuration Requirements

⚠️ **Critical Note**: `playwright-mcp` must be started with `--caps=pdf` parameter to provide `browser_pdf_save` command.

### 🔧 PDF Post-processing Workflow (Important)

**When using `playwright-mcp` to generate PDF, you must follow these steps:**

1. Use `playwright-mcp`'s `browser_pdf_save` command to generate PDF (saved in temporary directory)
2. **Immediately call** `doc-ops-mcp`'s `process_pdf_post_conversion` tool:
   - Automatically move PDF from temporary path to target location
   - Add watermark and QR code based on environment variables or parameters
   - Clean up temporary files

**Example Workflow:**
```
1. playwright-mcp:browser_pdf_save -> Generate temporary PDF: /tmp/playwright-output/file.pdf
2. doc-ops-mcp:process_pdf_post_conversion -> Process and save to: /output/final.pdf
```

⚠️ **Critical Note**: After using `playwright-mcp` to generate PDF, you must immediately call `doc-ops-mcp`'s `process_pdf_post_conversion` tool to complete final processing.

### 🤖 AI Model Workflow Guidance

**To ensure AI models correctly execute PDF conversion post-processing, follow these guidelines:**

#### 1. Tool Call Sequence Constraints
- When PDF generation is needed, tools must be called in this order:
  1. First generate HTML file (using appropriate conversion tools)
  2. Then call `playwright-mcp:browser_pdf_save` to generate PDF
  3. **Immediately** call `doc-ops-mcp:process_pdf_post_conversion` to complete post-processing

#### 2. Key Information in Tool Descriptions
- The `process_pdf_post_conversion` tool description clearly states "⚠️ Important: This is a necessary follow-up step for playwright-mcp's browser_pdf_save command!"
- AI models should recognize this key information and automatically execute the follow-up steps

#### 3. Recommended Prompt Patterns
When users request PDF conversion, use this approach:
```
User request: Convert Markdown to PDF
Execution steps:
1. Use convert_markdown_to_html to generate HTML
2. Use playwright-mcp:browser_pdf_save to generate PDF
3. Use process_pdf_post_conversion to complete final processing
```

#### 4. Automated Workflow
- AI models should treat `playwright-mcp:browser_pdf_save` and `process_pdf_post_conversion` as an atomic operation
- Should not wait for user confirmation between the two steps
- Should automatically pass correct file path parameters

#### 5. Complete Example Workflow

**Scenario: User requests converting `document.md` to PDF**

```json
// Step 1: Convert Markdown to HTML
{
  "tool": "doc-ops-mcp:convert_markdown_to_html",
  "params": {
    "markdownPath": "/path/to/document.md",
    "theme": "github"
  }
}
// Result: Generate /output/document.html

// Step 2: Use playwright-mcp to generate PDF
{
  "tool": "playwright-mcp:browser_pdf_save",
  "params": {
    "url": "file:///output/document.html",
    "path": "/tmp/playwright-output/document.pdf"
  }
}
// Result: Generate temporary PDF /tmp/playwright-output/document.pdf

// Step 3: Immediately call post-processing tool (Required!)
{
  "tool": "doc-ops-mcp:process_pdf_post_conversion",
  "params": {
    "playwrightPdfPath": "/tmp/playwright-output/document.pdf",
    "targetPath": "document.pdf",
    "addWatermark": true,
    "addQrCode": true
  }
}
// Result: Final PDF saved to /output/document.pdf with watermark and QR code
```

**Key Points:**
- No interruption allowed between Step 2 and Step 3
- `playwrightPdfPath` must be the exact path generated in Step 2
- Watermark and QR code will be automatically added if environment variables are set

### How It Works

When performing PDF conversion, this server will:
1. Process the source document and generate HTML
2. Call `playwright-mcp` tools for HTML → PDF conversion
3. Use `process_pdf_post_conversion` to post-process the generated PDF (path movement, watermark, QR code)

## 4. Features

### MCP Tools

#### Core Document Tools

| Tool Name | Description | Input Parameters | External Dependencies |
|-----------|-------------|------------------|----------------------|
| `read_document` | Read document content | `filePath`: Document path<br>`extractMetadata`: Extract metadata<br>`preserveFormatting`: Preserve formatting | None |
| `write_document` | Write document content | `content`: Document content<br>`outputPath`: Output file path<br>`encoding`: File encoding | None |
| `convert_document` | Smart document conversion | `inputPath`: Input file path<br>`outputPath`: Output file path<br>`preserveFormatting`: Preserve formatting<br>`useInternalPlaywright`: Use built-in Playwright | Depends on conversion type |
| `rewrite_document` | Smart document rewriting | `inputPath`: Document path<br>`rewriteRules`: Rewrite rules list<br>`outputPath`: Output path<br>`preserveOriginalFormat`: Keep original format | None |

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

**External Dependency:** Requires `playwright-mcp` for PDF conversion

##### **convert_markdown_to_pdf**
Convert Markdown to PDF with automatic watermark addition (if configured).

**Parameters:**
- `markdownPath` (string, required) - Markdown file path
- `outputPath` (string, optional) - Output PDF path (auto-generated if not provided)
- `theme` (string, optional) - Theme style, defaults to `"github"`
- `includeTableOfContents` (boolean, optional) - Include table of contents, defaults to `false`
- `addQrCode` (boolean, optional) - Whether to add QR code, defaults to `false`

**External Dependency:** Requires `playwright-mcp` for PDF conversion

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
Generate conversion plan by analyzing input file and providing conversion suggestions.

**Parameters:**
- `inputPath` (string, required) - Input file path
- `outputPath` (string, optional) - Output file path

##### **rewrite_document**
📝 Smart Document Rewriting Tool - Perform content rewriting, text replacement, and format adjustment on existing documents.

**Key Features:**
- Support intelligent rewriting for multiple document formats
- Text content replacement (supports regular expressions)
- Format-preserving rewriting and style adjustment
- Document structure reorganization and content optimization
- Batch rewrite rule application

**Supported Rewrite Formats:**
- **Direct Rewriting**: MD, HTML (direct text operations)
- **Conversion Rewriting**: DOCX → HTML → Rewrite → DOCX
- **Smart Rewriting**: Automatically identify format and select optimal rewrite strategy

**Parameters:**
- `inputPath` (string, required) - Path to the document to be rewritten
- `outputPath` (string, optional) - Output file path (overwrites original file if not specified)
- `rewriteRules` (array, required) - List of rewrite rules:
  - `type` (string) - Rewrite type: `"replace"` | `"format"` | `"structure"`
  - `oldText` (string) - Original text to be replaced
  - `newText` (string) - New text after replacement
  - `useRegex` (boolean, optional) - Whether to use regular expressions, defaults to `false`
  - `preserveCase` (boolean, optional) - Whether to preserve case, defaults to `false`
  - `preserveFormatting` (boolean, optional) - Whether to preserve original formatting, defaults to `true`
- `preserveOriginalFormat` (boolean, optional) - Whether to maintain original document format, defaults to `true`
- `backupOriginal` (boolean, optional) - Whether to backup original file, defaults to `false`

**Rewrite Rule Types:**
- **replace**: Text content replacement
  - Supports plain text and regular expression replacement
  - Configurable case sensitivity
- **format**: Format markup adjustment
  - Modify Markdown markup, HTML tags, etc.
  - Adjust document styles and layout
- **structure**: Structure reorganization
  - Reorganize paragraphs and sections
  - Adjust heading levels and content order

**Usage Example:**
```json
{
  "inputPath": "/path/to/document.md",
  "outputPath": "/path/to/rewritten_document.md",
  "rewriteRules": [
    {
      "type": "replace",
      "oldText": "Old Company Name",
      "newText": "New Company Name",
      "preserveCase": true
    },
    {
      "type": "format",
      "oldText": "## (.*)",
      "newText": "### $1",
      "useRegex": true
    }
  ],
  "preserveOriginalFormat": true,
  "backupOriginal": true
}
```

**Important Notes:**
- Rewrite operations are executed in the order of the rules list
- When using regular expressions, ensure expression safety
- Recommend enabling backup for important document rewrites
- For complex format documents (like DOCX), the system will automatically select the optimal conversion path

##### **process_pdf_post_conversion**
🔧 PDF post-processing unified tool - ⚠️ **Important**: This is a necessary follow-up step for playwright-mcp's browser_pdf_save command! When using playwright-mcp to generate PDF, you must immediately call this tool to complete final processing. Features include: 1) Automatically move PDF from playwright temporary path to target location 2) Unified watermark and QR code addition 3) Clean up temporary files. Workflow: playwright-mcp:browser_pdf_save → doc-ops-mcp:process_pdf_post_conversion

**Parameters:**
- `playwrightPdfPath` (string, required) - PDF file path generated by playwright-mcp (usually in temporary directory)
- `targetPath` (string, optional) - Target PDF file path (auto-generated if not provided). If not absolute path, will be resolved relative to OUTPUT_DIR environment variable
- `addWatermark` (boolean, optional) - Whether to add watermark (automatically added if WATERMARK_IMAGE environment variable is set), defaults to `false`
- `addQrCode` (boolean, optional) - Whether to add QR code (automatically added if QR_CODE_IMAGE environment variable is set), defaults to `false`
- `watermarkImage` (string, optional) - Watermark image path (overrides environment variable)
- `watermarkText` (string, optional) - Watermark text content (defaults to "doc-ops-mcp" if no watermark image is provided)
- `watermarkImageScale` (number, optional) - Watermark image scale ratio, defaults to `0.25`
- `watermarkImageOpacity` (number, optional) - Watermark image opacity, defaults to `0.3`
- `watermarkImagePosition` (string, optional) - Watermark image position, options: `["top-left", "top-right", "bottom-left", "bottom-right", "center"]`, defaults to `"top-right"`
- `qrCodePath` (string, optional) - QR code image path (overrides environment variable)
- `qrScale` (number, optional) - QR code scale ratio, defaults to `0.15`
- `qrOpacity` (number, optional) - QR code opacity, defaults to `1.0`
- `qrPosition` (string, optional) - QR code position, options: `["top-left", "top-right", "top-center", "bottom-left", "bottom-right", "bottom-center", "center"]`, defaults to `"bottom-center"`
- `customText` (string, optional) - Custom text below QR code, defaults to `"Scan QR code for more information"`

**External Dependency:** Works with `playwright-mcp` generated PDF files

#### Web Scraping Tools

**⚠️ Legal and Ethical Use Notice**: Web scraping tools should be used in compliance with target websites' Terms of Service and applicable laws and regulations. Users are responsible for:

- Respecting robots.txt files and website scraping policies
- Complying with data protection and privacy laws (GDPR, CCPA, etc.)
- Avoiding excessive requests that may impact website performance
- Obtaining necessary permissions for commercial use of scraped data
- Respecting intellectual property rights and copyright laws

This tool is provided for legitimate research, development, and automation purposes. Misuse of these tools may result in legal consequences.

##### **take_screenshot**
🖼️ Web screenshot tool - Capture webpage or HTML content screenshot using Playwright Chromium.

**Parameters:**
- `urlOrHtml` (string, required) - Webpage URL or HTML content
- `outputPath` (string, required) - Screenshot output path
- `options` (object, optional) - Screenshot options:
  - `width` (number) - Screenshot width
  - `height` (number) - Screenshot height
  - `format` (string) - Image format, options: `["png", "jpeg"]`
  - `quality` (number) - JPEG quality (1-100)
  - `fullPage` (boolean) - Whether to capture the full page

##### **document_preview_screenshot**
📋 Document preview screenshot - Convert DOCX and similar documents to preview screenshot.

**Parameters:**
- `documentPath` (string, required) - Document file path
- `outputPath` (string, required) - Screenshot output path
- `options` (object, optional) - Screenshot options:
  - `width` (number) - Screenshot width
  - `height` (number) - Screenshot height
  - `fullPage` (boolean) - Whether to capture the full page

##### **scrape_web_content**
🕷️ Web content scraping - Use Playwright Chromium to scrape webpage content.

**Parameters:**
- `url` (string, required) - Webpage URL to scrape
- `options` (object, optional) - Scraping options:
  - `waitForSelector` (string) - CSS selector to wait for
  - `timeout` (number) - Timeout in milliseconds
  - `textOnly` (boolean) - Extract only plain text

##### **scrape_structured_data**
📊 Structured data scraping - Scrape structured data from webpages using a CSS selector.

**Parameters:**
- `url` (string, required) - Webpage URL to scrape
- `selector` (string, required) - CSS selector
- `options` (object, optional) - Scraping options:
  - `timeout` (number) - Timeout in milliseconds

### Supported Conversions
| From\To | PDF | DOCX | HTML | Markdown |
|---------|----|----- |------|----------|
| **PDF** | ✅ | ❌ | ❌ | ❌ |
| **DOCX** | ✅ | ✅ | ✅ | ✅ |
| **HTML** | ✅ | ❌ | ✅ | ✅ |
| **Markdown** | ✅ | ✅ | ✅ | ✅ |

## Usage Examples

```
Convert /Users/docs/report.pdf to DOCX
Rewrite company names in /Users/docs/contract.md
Batch replace terminology in /Users/docs/manual.docx
Adjust heading formats in /Users/docs/article.html
```

## 5. Performance Metrics

### Document Processing Capabilities

| Document Type | Max File Size | Processing Speed | Memory Usage |
|---------------|---------------|------------------|---------------|
| **PDF** | 50MB | 2-5MB/s | ~File size×1.5 |
| **DOCX** | 50MB | 5-10MB/s | ~File size×2 |
| **HTML** | 50MB | 10-20MB/s | ~File size×1.2 |
| **Markdown** | 50MB | 15-30MB/s | ~File size×1.1 |

### Conversion Performance

- **PDF Conversion**: Depends on playwright-mcp, ~1-3 pages/second
- **DOCX Conversion**: Pure JavaScript processing, ~5-15 pages/second
- **HTML Conversion**: Fastest, ~20-50 pages/second
- **Concurrent Processing**: Supports up to 5 concurrent tasks

### System Resource Requirements

- **Minimum Memory**: 512MB
- **Recommended Memory**: 2GB (for large files)
- **CPU**: Single core sufficient, multi-core improves concurrency
- **Disk Space**: Temporary files require 2-3x original file size

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
| **playwright** | ^1.40.0 | Apache-2.0 | Browser automation (optional) |
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
- **v3.0**: Enhanced security and performance optimization
- **v4.0**: Plugin system and multi-language support

## Requirements

### Dependencies
- **Node.js** ≥ 18.0.0
- **Zero external tools** - All processing via npm packages
- **Optional**: playwright-mcp for external browser automation

### Pure JavaScript Stack
- **pdf-lib** - PDF manipulation
- **mammoth** - DOCX processing  
- **playwright** - Web automation
- **marked** - Markdown processing
- **exceljs** - Spreadsheet handling
- **puppeteer** - PDF generation from HTML

### Installation
```bash
# Only Node.js required
npm install -g doc-ops-mcp
```

### Component Overview

- **MCP Server Core**: Handles JSON-RPC 2.0 communication and tool registration
- **Tool Router**: Routes requests to appropriate processing modules
- **Processing Engine**: Contains specialized processors for different document types
- **Data Processing Layer**: Pure JavaScript libraries for document manipulation
- **Zero External Dependencies**: All processing done via npm packages

## 8. Docker Deployment

### Quick Start with Docker

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
|----------|-------------|---------|
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

### Troubleshooting

#### Common Issues

1. **Port conflicts**: Change the host port in docker-compose.yml
2. **Permission issues**: Ensure volume mounts have correct permissions
3. **Memory issues**: Increase Docker memory allocation

#### Debug Mode

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

#### Memory Issues
- **Problem**: Out of memory errors with large files
- **Solution**: Increase Node.js memory limit: `node --max-old-space-size=4096`

#### PDF Conversion Fails
- **Problem**: PDF conversion not working
- **Solution**: Ensure `playwright-mcp` is properly configured

#### Permission Errors
- **Problem**: Cannot write to output directory
- **Solution**: Check file permissions and `OUTPUT_DIR` configuration

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