# Document Operations MCP Server

[![npm version](https://badge.fury.io/js/doc-ops-mcp.svg)](https://badge.fury.io/js/doc-ops-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://img.shields.io/npm/dm/doc-ops-mcp.svg)](https://www.npmjs.com/package/doc-ops-mcp)

**Language / 语言**: [English](README.md) | [中文](README_zh.md)

> **Document Operations MCP Server** - 一个用于文档处理、转换和自动化的通用MCP服务器。通过统一的API和工具集处理PDF、DOCX、HTML、Markdown等多种格式。

## 目录

1. [快速开始](#1-快速开始)
2. [系统架构](#2-系统架构)
3. [外部依赖说明](#3-外部依赖说明)
4. [功能特性](#4-功能特性)
5. [性能指标](#5-性能指标)
6. [开源协议](#6-开源协议)
7. [未来规划](#7-未来规划)
8. [Docker部署](#8-docker部署)
9. [开发指南](#9-开发指南)
10. [故障排除](#10-故障排除)
11. [贡献指南](#11-贡献指南)

## 1. 快速开始

### 安装
```bash
# 通过 npm
npm install -g doc-ops-mcp

# 通过 pnpm
pnpm add -g doc-ops-mcp

# 通过 bun
bun add -g doc-ops-mcp
```

### 配置

```json
{
  "mcpServers": {
    "doc-ops-mcp": {
      "command": "npx",
      "args": ["-y", "doc-ops-mcp@latest"],
      "env": {
        "OUTPUT_DIR": "/path/to/your/output/directory",
        "CACHE_DIR": "/path/to/your/cache/directory",
        "WATERMARK_IMAGE": "/path/to/your/watermark.png",
        "QR_CODE_IMAGE": "/path/to/your/qrcode.png"
      }
    }
  }
}
```

### 环境变量

服务器支持环境变量来控制输出路径和PDF增强功能：

#### 核心目录
- **`OUTPUT_DIR`**: 控制所有生成文件的保存位置（默认：`~/Documents`）
- **`CACHE_DIR`**: 临时文件和缓存文件的目录（默认：`~/.cache/doc-ops-mcp`）

#### PDF 增强功能
- **`WATERMARK_IMAGE`**: PDF 文件的默认水印图片路径
  - 自动添加到所有 PDF 转换中
  - 支持格式：PNG、JPG
  - 如果未设置，将使用默认文字水印"doc-ops-mcp"
- **`QR_CODE_IMAGE`**: PDF 文件的默认二维码图片路径
  - 仅在明确要求时添加到 PDF 中（`addQrCode=true`）
  - 支持格式：PNG、JPG
  - 如果未设置，二维码功能将不可用

**输出路径规则：**
1. 如果未提供 `outputPath` → 文件保存到 `OUTPUT_DIR`，使用自动生成的名称
2. 如果 `outputPath` 是相对路径 → 相对于 `OUTPUT_DIR` 解析
3. 如果 `outputPath` 是绝对路径 → 按原样使用，忽略 `OUTPUT_DIR`

详细文档请参见 [OUTPUT_PATH_CONTROL.md](./OUTPUT_PATH_CONTROL.md)。

## 2. 系统架构

Document Operations MCP Server 采用混合架构设计，结合内部处理和外部依赖：

```
┌─────────────────────────────────────────────────────────────┐
│                    MCP 客户端层                             │
│           (Claude Desktop, Cursor, VS Code等)               │
└─────────────────────┬───────────────────────────────────────┘
                      │ JSON-RPC 2.0
┌─────────────────────┴───────────────────────────────────────┐
│                 Doc-Ops-MCP 服务器                         │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────┐ │
│  │   工具路由器    │  │   请求验证器    │  │  响应格式化 │ │
│  │   & 处理器      │  │                 │  │     器      │ │
│  └────────┬────────┘  └────────┬────────┘  └──────┬──────┘ │
│           │                    │                  │        │
│  ┌────────┴────────────────────┴──────────────────┴─────┐ │
│  │                   文档处理引擎                         │ │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐   │ │
│  │  │   文档      │  │   格式      │  │   样式      │   │ │
│  │  │   读取器    │  │   转换器    │  │   处理器    │   │ │
│  │  └─────────────┘  └─────────────┘  └─────────────┘   │ │
│  │                                                        │ │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐   │ │
│  │  │   PDF       │  │   水印/     │  │   网页      │   │ │
│  │  │   增强      │  │   二维码    │  │   抓取      │   │ │
│  │  └─────────────┘  └─────────────┘  └─────────────┘   │ │
└────┴───────────────────────────────────────────────────────┴─┘
                            │
┌───────────────────────────┴─────────────────────────────────┐
│                    核心依赖层                               │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐          │
│  │   pdf-lib   │  │   mammoth   │  │   marked    │          │
│  │   (PDF处理) │  │  (DOCX处理) │  │ (Markdown)  │          │
│  └─────────────┘  └─────────────┘  └─────────────┘          │
└─────────────────────────────────────────────────────────────┘
                            │
┌───────────────────────────┴─────────────────────────────────┐
│                   外部依赖层 (PDF转换)                      │
│  ┌─────────────────────────────────────────────────────┐   │
│  │                playwright-mcp                       │   │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐ │   │
│  │  │  浏览器     │  │   HTML      │  │   PDF       │ │   │
│  │  │  自动化     │  │   渲染      │  │   生成      │ │   │
│  │  └─────────────┘  └─────────────┘  └─────────────┘ │   │
│  └─────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

### 架构说明

**内部处理层**：
- 文档读取、格式转换、样式处理
- PDF 水印和二维码添加
- 网页内容抓取

**外部依赖层**：
- **PDF 转换**：依赖 `playwright-mcp` 进行 HTML → PDF 转换
- **转换流程**：DOCX/Markdown → HTML → PDF (通过 playwright-mcp)

**重要提示**：所有 PDF 转换功能都需要配合 `playwright-mcp` 使用。

## 3. 外部依赖说明

### playwright-mcp 依赖

本 MCP 服务器的 PDF 转换功能依赖于 `playwright-mcp` 服务器：

- **依赖工具**：`convert_docx_to_pdf`、`convert_markdown_to_pdf`
- **重要配置**：`playwright-mcp` 必须使用 `--caps=pdf` 参数才能提供 `browser_pdf_save` 命令
- **转换流程**：
  1. 将源文档转换为 HTML 格式
  2. 通过 `playwright-mcp` 的 `browser_pdf_save` 命令将 HTML 渲染为 PDF
  3. 自动添加水印（如果配置了 `WATERMARK_IMAGE`）
  4. 可选添加二维码（如果设置了 `addQrCode=true` 且配置了 `QR_CODE_IMAGE`）

#### 第三方依赖免责声明

**重要提示**：`playwright-mcp` 是一个独立的第三方项目。本项目（doc-ops-mcp）不对以下方面承担责任：

- `playwright-mcp` 的可用性、安全性或行为
- 由 `playwright-mcp` 引起的任何问题、漏洞或数据丢失
- `playwright-mcp` 的维护、更新或支持

**用户必须：**
- 审查并遵守 `playwright-mcp` 自身的许可条款和条件
- 了解使用第三方依赖的相关风险
- 确保 `playwright-mcp` 符合其安全性和合规性要求
- 独立监控 `playwright-mcp` 的更新和安全补丁

请参考 `playwright-mcp` 的官方文档和代码库，了解许可信息、安全公告和使用指南。

### 配置要求

1. **安装 playwright-mcp**：
   ```bash
   # 请参考 playwright-mcp 的官方文档进行安装和配置
   ```

2. **MCP 客户端配置**：
   确保在 MCP 客户端中同时配置了本服务器和 `playwright-mcp`
   
   **重要**：`playwright-mcp` 必须使用 `--caps=pdf` 参数：
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

3. **环境变量**：
   - `WATERMARK_IMAGE`：水印图片路径（可选，PDF 转换时自动添加）
   - `QR_CODE_IMAGE`：二维码图片路径（可选，需要时添加）

### 🔧 PDF转换工作流程（重要）

**当使用 `playwright-mcp` 生成PDF后，必须按以下步骤操作：**

1. **HTML生成阶段**：将源文档（如Markdown、DOCX）转换为带样式的HTML文件
2. **PDF生成阶段**：调用 `playwright-mcp` 的 `browser_pdf_save` 命令将HTML转换为PDF（会保存在临时目录）
3. **后处理阶段**：**必须**调用 `process_pdf_post_conversion` 工具完成路径转换、水印和二维码添加

**示例工作流程：**
```
1. playwright-mcp:browser_pdf_save -> 生成临时PDF: /tmp/playwright-output/file.pdf
2. doc-ops-mcp:process_pdf_post_conversion -> 处理并保存到: /output/final.pdf
```

⚠️ **关键提示**：使用 `playwright-mcp` 生成PDF后，必须立即调用 `doc-ops-mcp` 的 `process_pdf_post_conversion` 工具来完成最终处理。

### 🤖 大模型工作流程指导

**为确保大模型正确执行PDF转换后处理，请遵循以下指导原则：**

#### 1. 工具调用顺序约束
- 当需要生成PDF时，必须按照以下顺序调用工具：
  1. 首先生成HTML文件（使用相应的转换工具）
  2. 然后调用 `playwright-mcp:browser_pdf_save` 生成PDF
  3. **立即**调用 `doc-ops-mcp:process_pdf_post_conversion` 完成后处理

#### 2. 工具描述中的关键信息
- `process_pdf_post_conversion` 工具的描述明确标注了 "⚠️ 重要：这是playwright-mcp的browser_pdf_save命令的必要后续步骤！"
- 大模型应该识别这个关键信息并自动执行后续步骤

#### 3. 推荐的提示词模式
当用户请求PDF转换时，建议使用以下思路：
```
用户要求：将Markdown转换为PDF
执行步骤：
1. 使用 convert_markdown_to_html 生成HTML
2. 使用 playwright-mcp:browser_pdf_save 生成PDF
3. 使用 process_pdf_post_conversion 完成最终处理
```

#### 4. 自动化工作流程
- 大模型应该将 `playwright-mcp:browser_pdf_save` 和 `process_pdf_post_conversion` 视为一个原子操作
- 不应该在两个步骤之间等待用户确认
- 应该自动传递正确的文件路径参数

#### 5. 完整示例工作流程

**场景：用户要求将 `document.md` 转换为PDF**

```json
// 步骤1：转换Markdown到HTML
{
  "tool": "doc-ops-mcp:convert_markdown_to_html",
  "params": {
    "markdownPath": "/path/to/document.md",
    "theme": "github"
  }
}
// 结果：生成 /output/document.html

// 步骤2：使用playwright-mcp生成PDF
{
  "tool": "playwright-mcp:browser_pdf_save",
  "params": {
    "url": "file:///output/document.html",
    "path": "/tmp/playwright-output/document.pdf"
  }
}
// 结果：生成临时PDF /tmp/playwright-output/document.pdf

// 步骤3：立即调用后处理工具（必须！）
{
  "tool": "doc-ops-mcp:process_pdf_post_conversion",
  "params": {
    "playwrightPdfPath": "/tmp/playwright-output/document.pdf",
    "targetPath": "document.pdf",
    "addWatermark": true,
    "addQrCode": true
  }
}
// 结果：最终PDF保存到 /output/document.pdf，包含水印和二维码
```

**关键点：**
- 步骤2和步骤3之间不能有任何中断
- `playwrightPdfPath` 必须是步骤2生成的确切路径
- 如果设置了环境变量，水印和二维码会自动添加

### 工作原理

当执行 PDF 转换时，本服务器会：
1. 处理源文档并生成 HTML
2. 调用 `playwright-mcp` 的工具进行 HTML → PDF 转换
3. 通过 `process_pdf_post_conversion` 对生成的 PDF 进行后处理（路径移动、水印、二维码）

## 4. 功能特性

### MCP 工具

#### 核心文档工具

| 工具名称 | 功能描述 | 输入参数 | 外部依赖 |
|---------|----------|----------|----------|
| `read_document` | 读取文档内容 | `filePath`: 文档路径<br>`extractMetadata`: 提取元数据<br>`preserveFormatting`: 保留格式 | 无 |
| `write_document` | 写入文档内容 | `content`: 文档内容<br>`outputPath`: 输出文件路径<br>`encoding`: 文件编码 | 无 |
| `convert_document` | 智能文档转换 | `inputPath`: 输入文件路径<br>`outputPath`: 输出文件路径<br>`preserveFormatting`: 保留格式<br>`useInternalPlaywright`: 使用内置Playwright | 根据转换类型 |
| `rewrite_document` | 智能文档改写 | `inputPath`: 文档路径<br>`rewriteRules`: 改写规则列表<br>`outputPath`: 输出路径<br>`preserveOriginalFormat`: 保持原格式 | 无 |

##### **read_document**
读取各种文档格式，包括PDF、DOCX、DOC、HTML、MD等格式。

**参数：**
- `filePath` (string, 必需) - 要读取的文档路径
- `extractMetadata` (boolean, 可选) - 提取文档元数据，默认为`false`
- `preserveFormatting` (boolean, 可选) - 保留格式（HTML输出），默认为`false`

##### **write_document**
将内容写入指定格式的文档文件。

**参数：**
- `content` (string, 必需) - 要写入的内容
- `outputPath` (string, 可选) - 输出文件路径（不指定则自动生成）
- `encoding` (string, 可选) - 文件编码，默认为`utf-8`

##### **convert_document**
在格式间转换文档，支持样式保留增强。

**参数：**
- `inputPath` (string, 必需) - 输入文件路径
- `outputPath` (string, 可选) - 输出文件路径（不指定则自动生成）
- `preserveFormatting` (boolean, 可选) - 保留格式，默认为`true`
- `useInternalPlaywright` (boolean, 可选) - 使用内置Playwright进行PDF转换，默认为`false`

##### **convert_docx_to_pdf**
DOCX转PDF，自动添加水印（如果配置）。

**参数：**
- `docxPath` (string, 必需) - DOCX文件路径
- `outputPath` (string, 可选) - 输出PDF路径（不指定则自动生成）
- `addQrCode` (boolean, 可选) - 是否添加二维码，默认为`false`

**外部依赖：** 需要 `playwright-mcp` 进行PDF转换

##### **convert_markdown_to_pdf**
Markdown转PDF，自动添加水印（如果配置）。

**参数：**
- `markdownPath` (string, 必需) - Markdown文件路径
- `outputPath` (string, 可选) - 输出PDF路径（不指定则自动生成）
- `theme` (string, 可选) - 主题样式，默认为`"github"`
- `includeTableOfContents` (boolean, 可选) - 是否包含目录，默认为`false`
- `addQrCode` (boolean, 可选) - 是否添加二维码，默认为`false`

**外部依赖：** 需要 `playwright-mcp` 进行PDF转换

##### **convert_markdown_to_html**
Markdown转HTML。

**参数：**
- `markdownPath` (string, 必需) - Markdown文件路径
- `outputPath` (string, 可选) - 输出HTML路径（不指定则自动生成）
- `theme` (string, 可选) - 主题样式，默认为`"github"`
- `includeTableOfContents` (boolean, 可选) - 是否包含目录，默认为`false`

##### **convert_markdown_to_docx**
Markdown转DOCX。

**参数：**
- `markdownPath` (string, 必需) - Markdown文件路径
- `outputPath` (string, 可选) - 输出DOCX路径（不指定则自动生成）

##### **convert_html_to_markdown**
HTML转Markdown。

**参数：**
- `htmlPath` (string, 必需) - HTML文件路径
- `outputPath` (string, 可选) - 输出Markdown路径（不指定则自动生成）

##### **plan_conversion**
转换计划生成，分析输入文件并生成转换建议。

**参数：**
- `inputPath` (string, 必需) - 输入文件路径
- `outputPath` (string, 可选) - 输出文件路径

##### **rewrite_document**
📝 智能文档改写工具 - 对现有文档进行内容改写、文本替换和格式调整。

**功能特性：**
- 支持多种文档格式的智能改写
- 文本内容替换（支持正则表达式）
- 格式保持改写和样式调整
- 文档结构重组和内容优化
- 批量改写规则应用

**支持的改写格式：**
- **直接改写**：MD、HTML（直接文本操作）
- **转换改写**：DOCX → HTML → 改写 → DOCX
- **智能改写**：自动识别格式并选择最佳改写策略

**参数：**
- `inputPath` (string, 必需) - 要改写的文档路径
- `outputPath` (string, 可选) - 输出文件路径（不指定则覆盖原文件）
- `rewriteRules` (array, 必需) - 改写规则列表：
  - `type` (string) - 改写类型：`"replace"` | `"format"` | `"structure"`
  - `oldText` (string) - 要替换的原文本
  - `newText` (string) - 替换后的新文本
  - `useRegex` (boolean, 可选) - 是否使用正则表达式，默认为`false`
  - `preserveCase` (boolean, 可选) - 是否保持大小写，默认为`false`
  - `preserveFormatting` (boolean, 可选) - 是否保留原格式，默认为`true`
- `preserveOriginalFormat` (boolean, 可选) - 是否保持原始文档格式，默认为`true`
- `backupOriginal` (boolean, 可选) - 是否备份原文件，默认为`false`

**改写规则类型说明：**
- **replace**：文本内容替换
  - 支持普通文本和正则表达式替换
  - 可配置大小写敏感性
- **format**：格式标记调整
  - 修改Markdown标记、HTML标签等
  - 调整文档样式和排版
- **structure**：结构重组
  - 重新组织段落和章节
  - 调整标题层级和内容顺序

**使用示例：**
```json
{
  "inputPath": "/path/to/document.md",
  "outputPath": "/path/to/rewritten_document.md",
  "rewriteRules": [
    {
      "type": "replace",
      "oldText": "旧公司名称",
      "newText": "新公司名称",
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

**注意事项：**
- 改写操作会按照规则列表的顺序依次执行
- 使用正则表达式时请确保表达式的安全性
- 建议在重要文档改写前启用备份功能
- 对于复杂格式文档（如DOCX），系统会自动选择最佳转换路径

#### 网页抓取工具

**⚠️ 法律和道德使用声明**：网页抓取工具应在遵守目标网站服务条款和相关法律法规的前提下使用。用户有责任：

- 遵守 robots.txt 文件和网站抓取政策
- 遵守数据保护和隐私法律（GDPR、CCPA 等）
- 避免过度请求影响网站性能
- 获得商业使用抓取数据的必要许可
- 尊重知识产权和版权法

此工具仅供合法的研究、开发和自动化目的使用。滥用这些工具可能导致法律后果。

##### **take_screenshot**
🖼️ 网页截图工具 - 使用Playwright Chromium捕获网页或HTML内容截图。

**参数：**
- `urlOrHtml` (string, 必需) - 网页URL或HTML内容
- `outputPath` (string, 必需) - 截图输出路径
- `options` (object, 可选) - 截图选项：
  - `width` (number) - 截图宽度
  - `height` (number) - 截图高度
  - `format` (string) - 图片格式，可选值：`["png", "jpeg"]`
  - `quality` (number) - JPEG质量（1-100）
  - `fullPage` (boolean) - 是否捕获完整页面

##### **document_preview_screenshot**
📋 文档预览截图 - 将DOCX等文档转换为预览截图。

**参数：**
- `documentPath` (string, 必需) - 文档文件路径
- `outputPath` (string, 必需) - 截图输出路径
- `options` (object, 可选) - 截图选项：
  - `width` (number) - 截图宽度
  - `height` (number) - 截图高度
  - `fullPage` (boolean) - 是否捕获完整页面

##### **scrape_web_content**
🕷️ 网页内容抓取 - 使用Playwright Chromium抓取网页内容。

**参数：**
- `url` (string, 必需) - 要抓取的网页URL
- `options` (object, 可选) - 抓取选项：
  - `waitForSelector` (string) - 等待的CSS选择器
  - `timeout` (number) - 超时时间（毫秒）
  - `textOnly` (boolean) - 仅提取纯文本内容

##### **scrape_structured_data**
📊 结构化数据抓取 - 使用CSS选择器从网页抓取结构化数据。

**参数：**
- `url` (string, 必需) - 要抓取的网页URL
- `selector` (string, 必需) - CSS选择器
- `options` (object, 可选) - 抓取选项：
  - `timeout` (number) - 超时时间（毫秒）

### 支持的格式转换
| 从\到 | PDF | DOCX | HTML | Markdown |
|---------|----|----- |------|----------|
| **PDF** | ✅ | ❌ | ❌ | ❌ |
| **DOCX** | ✅ | ✅ | ✅ | ✅ |
| **HTML** | ✅ | ❌ | ✅ | ✅ |
| **Markdown** | ✅ | ✅ | ✅ | ✅ |

## 使用示例

```
将 /Users/docs/report.pdf 转换为 DOCX
改写 /Users/docs/contract.md 中的公司名称
批量替换 /Users/docs/manual.docx 中的术语
调整 /Users/docs/article.html 的标题格式
```

## 5. 性能指标

### 文档处理能力

| 文档类型 | 最大文件大小 | 处理速度 | 内存占用 |
|---------|-------------|----------|----------|
| **PDF** | 50MB | 2-5MB/s | ~文件大小×1.5 |
| **DOCX** | 50MB | 5-10MB/s | ~文件大小×2 |
| **HTML** | 50MB | 10-20MB/s | ~文件大小×1.2 |
| **Markdown** | 50MB | 15-30MB/s | ~文件大小×1.1 |

### 转换性能

- **PDF 转换**：依赖 playwright-mcp，速度约 1-3 页/秒
- **DOCX 转换**：纯 JavaScript 处理，速度约 5-15 页/秒
- **HTML 转换**：最快，速度约 20-50 页/秒
- **并发处理**：支持最多 5 个并发任务

### 系统资源要求

- **最小内存**：512MB
- **推荐内存**：2GB（处理大文件时）
- **CPU**：单核心即可，多核心可提升并发性能
- **磁盘空间**：临时文件约占原文件大小的 2-3 倍

## 系统要求

### 依赖项
- **Node.js** ≥ 18.0.0
- **零外部依赖** - 所有处理通过npm包实现
- **可选**：playwright-mcp用于外部浏览器自动化

### 纯JavaScript技术栈
- **pdf-lib** - PDF操作
- **mammoth** - DOCX处理  
- **playwright** - 网页自动化
- **marked** - Markdown处理
- **exceljs** - 表格处理

### 安装
```bash
# 仅需Node.js
npm install -g doc-ops-mcp
```

### 组件说明

- **MCP服务器核心**: 处理JSON-RPC 2.0通信和工具注册
- **工具路由器**: 将请求路由至相应处理模块
- **处理引擎**: 包含针对不同文档类型的专用处理器
- **数据处理层**: 用于文档操作的纯JavaScript库
- **零外部依赖**: 所有处理通过npm包完成

## 6. 开源协议

### 项目协议
- **本项目**：MIT License
- **兼容性**：可用于商业和非商业项目

### 第三方依赖协议

| 依赖库 | 版本 | 协议 | 用途 |
|--------|------|------|------|
| **pdf-lib** | ^1.17.1 | MIT | PDF 文档操作和处理 |
| **mammoth** | ^1.6.0 | BSD-2-Clause | DOCX 文档解析和转换 |
| **marked** | ^9.1.6 | MIT | Markdown 解析和渲染 |
| **playwright** | ^1.40.0 | Apache-2.0 | 浏览器自动化（可选） |
| **exceljs** | ^4.4.0 | MIT | Excel 文件处理 |
| **jsdom** | ^23.0.1 | MIT | HTML DOM 操作 |
| **turndown** | ^7.1.2 | MIT | HTML 转 Markdown |

### 协议兼容性
- ✅ **商业使用**：所有依赖均支持商业使用
- ✅ **分发**：可自由分发和修改
- ✅ **专利保护**：Apache-2.0 提供专利保护
- ⚠️ **注意事项**：使用时需保留原始协议声明

## 7. 未来规划

### 核心功能
- 🔄 **增强转换质量**：改进复杂文档的样式保留
- 📊 **Excel 支持**：完整的 Excel 读写和转换功能
- 🎨 **模板系统**：支持自定义文档模板
- 🔍 **OCR 集成**：图片文字识别功能

### 系统改进
- 🌐 **多语言支持**：国际化和本地化
- 🔐 **安全增强**：文档加密和权限控制
- ⚡ **性能优化**：大文件处理和内存优化
- 🔌 **插件系统**：可扩展的处理器架构

### 版本路线图
- **v2.0**：完整的 Excel 支持和模板系统
- **v3.0**：OCR 集成和多语言支持
- **v4.0**：高级安全功能和插件系统

## 8. Docker部署

### 快速开始

#### 使用预构建镜像

```bash
# 拉取最新镜像
docker pull docops/doc-ops-mcp:latest

# 使用默认配置运行
docker run -d \
  --name doc-ops-mcp \
  -p 3000:3000 \
  docops/doc-ops-mcp:latest
```

#### 从源码构建

```bash
# 克隆仓库
git clone https://github.com/JefferyMunoz/doc-ops-mcp.git
cd doc-ops-mcp

# 构建Docker镜像
docker build -t doc-ops-mcp .

# 运行容器
docker run -d \
  --name doc-ops-mcp \
  -p 3000:3000 \
  -v $(pwd)/documents:/app/documents \
  doc-ops-mcp
```

### Docker Compose部署

创建 `docker-compose.yml` 文件：

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
    
  # 可选：添加Nginx反向代理
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

### 环境变量

| 变量名 | 描述 | 默认值 |
|----------|-------------|---------|
| `PORT` | 服务器端口 | `3000` |
| `NODE_ENV` | 环境模式 | `production` |
| `LOG_LEVEL` | 日志级别 | `info` |
| `MAX_FILE_SIZE` | 最大文件大小(MB) | `50` |

### 卷挂载

为持久化存储挂载本地目录：

```bash
# 文档目录用于文件处理
docker run -d \
  --name doc-ops-mcp \
  -p 3000:3000 \
  -v $(pwd)/documents:/app/documents \
  -v $(pwd)/output:/app/output \
  doc-ops-mcp
```

### Docker配置示例

#### 生产环境部署

```bash
# 使用Docker Swarm进行生产部署
docker swarm init
docker stack deploy -c docker-compose.yml doc-ops

# 扩展服务
docker service scale doc-ops_mcp=3
```

### 健康检查

容器包含内置健康检查：

```bash
# 检查容器健康状态
docker ps

# 查看健康检查日志
docker inspect --format='{{.State.Health.Status}}' doc-ops-mcp

# 手动健康检查
docker exec doc-ops-mcp curl -f http://localhost:3000/health || exit 1
```

## 9. 开发指南

### 本地开发

```bash
# 克隆项目
git clone https://github.com/your-org/doc-ops-mcp.git
cd doc-ops-mcp

# 安装依赖
npm install

# 开发模式运行
npm run dev

# 构建项目
npm run build

# 运行测试
npm test
```

### 项目结构

```
src/
├── index.ts          # MCP 服务器入口
├── tools/            # 工具实现
│   ├── documentConverter.ts
│   ├── pdfTools.ts
│   └── ...
├── types/            # 类型定义
└── utils/            # 工具函数
```

### 添加新工具

1. 在 `src/tools/` 中创建新的工具文件
2. 实现工具逻辑
3. 在 `src/index.ts` 中注册工具
4. 添加测试用例
5. 更新文档

## 10. 故障排除

#### 常见问题

1. **端口冲突**：在docker-compose.yml中修改主机端口
2. **权限问题**：确保卷挂载具有正确的权限
3. **内存问题**：增加Docker内存分配

#### 调试模式

```bash
# 使用调试日志运行
docker run -d \
  --name doc-ops-mcp \
  -p 3000:3000 \
  -e LOG_LEVEL=debug \
  doc-ops-mcp

# 查看日志
docker logs -f doc-ops-mcp
```

## 11. 贡献指南

### 如何贡献

1. **Fork 项目**
2. **创建功能分支** (`git checkout -b feature/AmazingFeature`)
3. **提交更改** (`git commit -m 'Add some AmazingFeature'`)
4. **推送到分支** (`git push origin feature/AmazingFeature`)
5. **创建 Pull Request**

#### 知识产权授权

**通过提交 Pull Request，您同意所有通过 Pull Request 提交的贡献都将在 MIT 许可证下授权。** 这意味着：

- 您授予项目维护者和用户在 MIT 许可证下使用、修改和分发您的贡献的权利
- 您确认您有权进行这些贡献
- 您理解您的贡献将成为开源项目的一部分
- 您放弃对贡献代码的独占所有权声明

如果您不能同意这些条款，请不要提交 Pull Request。

### 代码规范

- 使用 TypeScript
- 遵循 ESLint 配置
- 添加适当的测试
- 更新相关文档

### 报告问题

- 使用 [GitHub Issues](https://github.com/your-org/doc-ops-mcp/issues)
- 提供详细的错误信息和重现步骤
- 包含系统环境信息

### 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。