# Document Operations MCP Server

[![npm version](https://badge.fury.io/js/doc-ops-mcp.svg)](https://badge.fury.io/js/doc-ops-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Downloads](https://img.shields.io/npm/dm/doc-ops-mcp.svg)](https://www.npmjs.com/package/doc-ops-mcp)

**Language / 语言**: [English](README.md) | [中文](README_zh.md)

> **Document Operations MCP Server** - 一个用于文档处理、转换和自动化的通用MCP服务器。通过统一的API和工具集处理PDF、DOCX、HTML、Markdown等多种格式。

## 目录

1. [快速开始](#1-快速开始)
2. [系统架构](#2-系统架构)
3. [可选集成说明](#3-可选集成说明)
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

### 支持的文档操作

| 格式 | 转换到PDF | 转换到DOCX | 转换到HTML | 转换到Markdown | 内容改写 | 水印/二维码 |
|------|----------|-----------|-----------|--------------|----------|------------|
| **PDF** | ✅ | ❌ | ❌ | ❌ | ❌ | ✅ |
| **DOCX** | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ |
| **HTML** | ✅ | ❌ | ✅ | ✅ | ✅ | ❌ |
| **Markdown** | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ |

**改写功能说明：**
- **内容替换**：支持文本内容的批量替换和正则表达式替换
- **格式调整**：修改文档结构、标题层级和样式格式
- **智能改写**：保持原文档格式的同时进行内容优化

### 使用示例

**格式转换：**
```
将 /Users/docs/report.docx 转换为 PDF
将 /Users/docs/article.md 转换为 HTML
将 /Users/docs/presentation.html 转换为 DOCX
将 /Users/docs/readme.md 转换为 PDF（带主题样式）
```

**文档改写：**
```
改写 /Users/docs/contract.md 中的公司名称
批量替换 /Users/docs/manual.docx 中的术语
调整 /Users/docs/article.html 的标题层级
更新 /Users/docs/policy.md 中的日期和版本号
```

**PDF增强：**
```
为 /Users/docs/document.pdf 添加水印
为 /Users/docs/report.pdf 添加二维码
为 /Users/docs/invoice.pdf 添加公司logo水印
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

Document Operations MCP Server 采用纯 JavaScript 架构设计，提供完整的文档处理能力：

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
│  │  │   PDF       │  │   水印/     │  │   转换      │   │ │
│  │  │   增强      │  │   二维码    │  │   规划器    │   │ │
│  │  └─────────────┘  └─────────────┘  └─────────────┘   │ │
└────┴───────────────────────────────────────────────────────┴─┘
                            │
┌───────────────────────────┴─────────────────────────────────┐
│                    核心依赖层                               │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐          │
│  │   pdf-lib   │  │   mammoth   │  │   marked    │          │
│  │   (PDF处理) │  │  (DOCX处理) │  │ (Markdown)  │          │
│  └─────────────┘  └─────────────┘  └─────────────┘          │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐          │
│  │   cheerio   │  │   turndown  │  │   docx      │          │
│  │ (HTML解析)  │  │ (HTML转MD)  │  │ (DOCX生成)  │          │
│  └─────────────┘  └─────────────┘  └─────────────┘          │
└─────────────────────────────────────────────────────────────┘
```

### 架构说明

**核心特性**：
- 纯 JavaScript 实现，无外部系统依赖
- 完整的文档读取、转换、样式处理能力
- 内置 PDF 水印和二维码添加功能
- 智能转换规划和路径优化

**转换流程**：
- **直接转换**：支持大部分格式间的直接转换
- **多步转换**：复杂转换通过中间格式实现
- **样式保留**：使用 OOXML 解析器确保样式完整性

## 3. 可选集成说明

本服务器可与 `playwright-mcp` 配合使用，获得增强的 PDF 转换能力。详细配置请参考 `playwright-mcp` 官方文档。

### 🔧 PDF转换工作流程

本服务器支持完整的PDF转换功能：
1. **文档解析**：使用OOXML解析器确保样式完整保留
2. **格式转换**：将文档转换为高质量HTML格式
3. **PDF生成**：内置转换器或可选配合 `playwright-mcp` 使用
4. **增强处理**：自动添加水印和二维码（如果配置）

### 工作原理

本服务器采用智能转换架构：
1. **智能规划**：`plan_conversion` 分析转换需求，选择最优路径
2. **格式转换**：使用专用转换器处理各种文档格式
3. **样式保留**：通过 OOXML 解析器确保样式完整性
4. **增强处理**：自动添加水印、二维码等增强功能
5. **可选集成**：支持与 `playwright-mcp` 配合获得增强能力

## 4. 功能特性

### MCP 工具

#### 核心文档工具

| 工具名称 | 功能描述 | 输入参数 | 外部依赖 |
|---------|----------|----------|----------|
| `read_document` | 读取文档内容 | `filePath`: 文档路径<br>`extractMetadata`: 提取元数据<br>`preserveFormatting`: 保留格式 | 无 |
| `write_document` | 写入文档内容 | `content`: 文档内容<br>`outputPath`: 输出文件路径<br>`encoding`: 文件编码 | 无 |
| `convert_document` | 智能文档转换 | `inputPath`: 输入文件路径<br>`outputPath`: 输出文件路径<br>`preserveFormatting`: 保留格式 | 无 |
| `plan_conversion` | 转换规划器 | `sourceFormat`: 源格式<br>`targetFormat`: 目标格式<br>`preserveStyles`: 保留样式<br>`quality`: 转换质量 | 无 |

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
- `preserveFormatting` (boolean, 可选) - 保留原始格式，默认为`true`
- `chineseFont` (string, 可选) - 中文字体，默认为`Microsoft YaHei`

**可选集成：** 可与 `playwright-mcp` 配合获得增强PDF转换

##### **convert_markdown_to_pdf**
Markdown转PDF，自动添加水印（如果配置）。

**参数：**
- `markdownPath` (string, 必需) - Markdown文件路径
- `outputPath` (string, 可选) - 输出PDF路径（不指定则自动生成）
- `theme` (string, 可选) - 主题样式，默认为`"github"`
- `includeTableOfContents` (boolean, 可选) - 是否包含目录，默认为`false`
- `addQrCode` (boolean, 可选) - 是否添加二维码，默认为`false`

**可选集成：** 可与 `playwright-mcp` 配合获得增强PDF转换

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
🎯 智能转换规划器 - 分析转换需求并生成最优转换方案。

**参数：**
- `sourceFormat` (string, 必需) - 源文件格式（pdf, docx, html, markdown, md, txt, doc）
- `targetFormat` (string, 必需) - 目标文件格式（pdf, docx, html, markdown, md, txt, doc）
- `sourceFile` (string, 可选) - 源文件路径（用于生成具体转换参数）
- `preserveStyles` (boolean, 可选) - 是否保留样式格式，默认为`true`
- `includeImages` (boolean, 可选) - 是否包含图片，默认为`true`
- `theme` (string, 可选) - 转换主题，默认为`github`
- `quality` (string, 可选) - 转换质量要求（fast, balanced, high），默认为`balanced`

##### **process_pdf_post_conversion**
🔧 PDF后处理统一工具 - playwright-mcp的browser_pdf_save命令的必要后续步骤。

**参数：**
- `playwrightPdfPath` (string, 必需) - playwright-mcp生成的PDF文件路径
- `targetPath` (string, 可选) - 目标PDF文件路径（不指定则自动生成）
- `addWatermark` (boolean, 可选) - 是否添加水印，默认为`false`
- `addQrCode` (boolean, 可选) - 是否添加二维码，默认为`false`
- `watermarkImage` (string, 可选) - 水印图片路径
- `qrCodePath` (string, 可选) - 二维码图片路径

**功能：**
1. 自动移动PDF从playwright临时路径到目标位置
2. 统一添加水印和二维码
3. 清理临时文件



#### PDF 增强工具

##### **add_watermark**
🎨 PDF水印添加工具 - 为PDF文档添加图片或文字水印。

**参数：**
- `pdfPath` (string, 必需) - PDF文件路径
- `watermarkImage` (string, 可选) - 水印图片路径（PNG/JPG）
- `watermarkText` (string, 可选) - 水印文字内容
- `watermarkImageScale` (number, 可选) - 图片缩放比例，默认为`0.25`
- `watermarkImageOpacity` (number, 可选) - 图片透明度，默认为`0.6`
- `watermarkImagePosition` (string, 可选) - 图片位置，默认为`fullscreen`

##### **add_qrcode**
📱 PDF二维码添加工具 - 为PDF文档添加二维码。

**参数：**
- `pdfPath` (string, 必需) - PDF文件路径
- `qrCodePath` (string, 可选) - 二维码图片路径
- `qrScale` (number, 可选) - 二维码缩放比例，默认为`0.15`
- `qrOpacity` (number, 可选) - 二维码透明度，默认为`1.0`
- `qrPosition` (string, 可选) - 二维码位置，默认为`bottom-center`
- `addText` (boolean, 可选) - 是否添加说明文字，默认为`true`

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

### 系统要求
- **Node.js** ≥ 18.0.0
- **零外部系统依赖** - 所有处理通过npm包实现
- **可选集成**：playwright-mcp用于增强PDF转换

### 核心技术栈
- **pdf-lib** - PDF操作和增强
- **mammoth** - DOCX文档处理  
- **marked** - Markdown解析和渲染
- **cheerio** - HTML解析和操作
- **turndown** - HTML到Markdown转换
- **docx** - DOCX文档生成

### 安装
```bash
# 全局安装
npm install -g doc-ops-mcp

# 或使用 pnpm
pnpm add -g doc-ops-mcp

# 或使用 bun
bun add -g doc-ops-mcp
```

### 架构组件

- **MCP服务器核心**: 处理JSON-RPC 2.0通信和工具注册
- **智能路由器**: 将请求路由至最优处理模块
- **转换引擎**: 包含针对不同文档类型的专用转换器
- **样式处理器**: 确保格式转换中的样式保留
- **安全模块**: 提供路径验证和内容安全处理

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