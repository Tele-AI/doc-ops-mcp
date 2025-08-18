import * as fs from 'fs';
import { promisify } from 'util';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
const { rgb } = require('pdf-lib');

interface DocumentContent {
  title?: string;
  content: string;
  author?: string;
  description?: string;
  metadata?: {
    [key: string]: string;
  };
}

interface ConversionOptions {
  outputPath?: string;
  format: 'md' | 'pdf' | 'docx' | 'html';
  styling?: {
    fontSize?: number;
    fontFamily?: string;
    lineHeight?: number;
    margins?: {
      top?: number;
      bottom?: number;
      left?: number;
      right?: number;
    };
    colors?: {
      primary?: string;
      secondary?: string;
      text?: string;
    };
  };
  // PDF水印配置
  watermark?: {
    enabled?: boolean; // 是否启用水印，PDF格式默认为true
    text?: string; // 水印文字，默认为'doc-ops-mcp'
    imagePath?: string; // 水印图片路径，如果提供则优先使用图片
    opacity?: number; // 透明度，0-1之间，默认0.1
    fontSize?: number; // 文字水印字体大小，默认48
    rotation?: number; // 旋转角度，默认-45度
    spacing?: {
      x?: number; // 水印间距X，默认200
      y?: number; // 水印间距Y，默认150
    };
  };
}

interface DocumentConversionResult {
  success: boolean;
  outputPath?: string;
  error?: string;
  fileSize?: number;
  format: string;
}

export class DocumentConverter {
  private defaultStyling = {
    fontSize: 12,
    fontFamily: 'Arial',
    lineHeight: 1.6,
    margins: {
      top: 72,
      bottom: 72,
      left: 72,
      right: 72,
    },
    colors: {
      primary: '#0066cc',
      secondary: '#666666',
      text: '#333333',
    },
  };

  /**
   * 转换文档到指定格式
   */
  async convertDocument(
    content: DocumentContent,
    options: ConversionOptions
  ): Promise<DocumentConversionResult> {
    try {
      const styling = { ...this.defaultStyling, ...options.styling };
      const outputPath =
        options.outputPath ?? this.generateOutputPath(content.title ?? 'document', options.format);

      // 为PDF格式设置默认水印配置
      if (options.format === 'pdf' && !options.watermark) {
        options.watermark = {
          enabled: true,
          text: 'doc-ops-mcp',
          opacity: 0.1,
          fontSize: 48,
          rotation: -45,
          spacing: {
            x: 200,
            y: 150
          }
        };
      }

      switch (options.format) {
        case 'md':
          return await this.convertToMarkdown(content, outputPath);
        case 'html':
          return await this.convertToHTML(content, outputPath, styling);
        case 'pdf':
          return await this.convertToPDF(content, outputPath, styling, options.watermark);
        case 'docx':
          return await this.convertToDocx(content, outputPath, styling);
        default:
          return {
            success: false,
            error: `不支持的格式: ${options.format}`,
            format: options.format,
          };
      }
    } catch (error: any) {
      return {
        success: false,
        error: `转换失败: ${error.message}`,
        format: options.format,
      };
    }
  }

  /**
   * 转换为Markdown格式
   */
  private async convertToMarkdown(
    content: DocumentContent,
    outputPath: string
  ): Promise<DocumentConversionResult> {
    try {
      let markdown = '';

      // 添加标题
      if (content.title) {
        markdown += `# ${content.title}\n\n`;
      }

      // 添加作者信息
      if (content.author) {
        markdown += `**作者:** ${content.author}\n\n`;
      }

      // 添加描述
      if (content.description) {
        markdown += `**描述:** ${content.description}\n\n`;
      }

      // 添加分隔线
      if (content.title || content.author || content.description) {
        markdown += '---\n\n';
      }

      // 添加主要内容
      markdown += content.content;

      // 保存文件
      const { validateAndSanitizePath } = require('../security/securityConfig');
      const allowedPaths = [process.cwd()];
      const validatedPath = validateAndSanitizePath(outputPath, allowedPaths);
      if (!validatedPath) {
        throw new Error('Invalid output path');
      }
      const writeFile = promisify(fs.writeFile);
      await writeFile(validatedPath, markdown, 'utf-8');
      outputPath = validatedPath;

      const stats = await promisify(fs.stat)(outputPath);

      return {
        success: true,
        outputPath,
        fileSize: stats.size,
        format: 'md',
      };
    } catch (error: any) {
      return {
        success: false,
        error: `Markdown转换失败: ${error.message}`,
        format: 'md',
      };
    }
  }

  /**
   * 转换为HTML格式
   */
  private async convertToHTML(
    content: DocumentContent,
    outputPath: string,
    styling: any
  ): Promise<DocumentConversionResult> {
    try {
      // 将内容转换为HTML
      const htmlContent = this.parseContentToHTML(content.content);

      const html = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${this.escapeHtml(content.title ?? '文档')}</title>
    ${content.author ? `<meta name="author" content="${this.escapeHtml(content.author)}">` : ''}
    ${content.description ? `<meta name="description" content="${this.escapeHtml(content.description)}">` : ''}
    <style>
        body {
            font-family: '${this.escapeHtml(styling.fontFamily)}', 'Microsoft YaHei', 'SimHei', Arial, sans-serif;
            font-size: ${styling.fontSize}px;
            line-height: ${styling.lineHeight};
            color: ${this.escapeHtml(styling.colors.text)};
            max-width: 800px;
            margin: 0 auto;
            padding: ${styling.margins.top / 4}px ${styling.margins.left / 4}px;
            background-color: #fff;
        }
        h1, h2, h3, h4, h5, h6 {
            color: ${this.escapeHtml(styling.colors.primary)};
            margin-top: 30px;
            margin-bottom: 15px;
        }
        h1 {
            font-size: 2.5em;
            text-align: center;
            border-bottom: 3px solid ${this.escapeHtml(styling.colors.primary)};
            padding-bottom: 10px;
        }
        h2 {
            font-size: 1.8em;
            border-left: 4px solid ${this.escapeHtml(styling.colors.primary)};
            padding-left: 15px;
        }
        h3 {
            font-size: 1.4em;
        }
        p {
            margin-bottom: 15px;
            text-align: justify;
            text-indent: 2em;
        }
        .meta-info {
            background-color: #f8f9fa;
            padding: 15px;
            border-left: 4px solid ${this.escapeHtml(styling.colors.secondary)};
            margin-bottom: 30px;
            border-radius: 4px;
        }
        .meta-info p {
            margin: 5px 0;
            text-indent: 0;
        }
        ul, ol {
            margin-bottom: 15px;
            padding-left: 30px;
        }
        li {
            margin-bottom: 5px;
        }
        blockquote {
            border-left: 4px solid ${this.escapeHtml(styling.colors.secondary)};
            margin: 20px 0;
            padding: 10px 20px;
            background-color: #f8f9fa;
            font-style: italic;
        }
        code {
            background-color: #f1f3f4;
            padding: 2px 4px;
            border-radius: 3px;
            font-family: 'Courier New', monospace;
        }
        pre {
            background-color: #f1f3f4;
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
        }
        @media print {
            body {
                margin: 20px;
                max-width: none;
            }
        }
    </style>
</head>
<body>
    ${content.title ? `<h1>${this.escapeHtml(content.title)}</h1>` : ''}
    ${
      content.author || content.description
        ? `
    <div class="meta-info">
        ${content.author ? `<p><strong>作者:</strong> ${this.escapeHtml(content.author)}</p>` : ''}
        ${content.description ? `<p><strong>描述:</strong> ${this.escapeHtml(content.description)}</p>` : ''}
    </div>`
        : ''
    }
    ${this.escapeHtml(htmlContent)}
</body>
</html>
      `;

      const writeFile = promisify(fs.writeFile);
      await writeFile(outputPath, html, 'utf-8');

      const stats = await promisify(fs.stat)(outputPath);

      return {
        success: true,
        outputPath,
        fileSize: stats.size,
        format: 'html',
      };
    } catch (error: any) {
      return {
        success: false,
        error: `HTML转换失败: ${error.message}`,
        format: 'html',
      };
    }
  }

  /**
   * 转换为PDF格式（使用pdf-lib）
   */
  private async convertToPDF(
    content: DocumentContent,
    outputPath: string,
    styling: any,
    watermarkConfig?: any
  ): Promise<DocumentConversionResult> {
    try {
      // 使用pdf-lib直接生成PDF，类似Word转PDF的方式
      const { PDFDocument, rgb, StandardFonts } = await import('pdf-lib');

      const pdfDoc = await PDFDocument.create();
      let currentPage = pdfDoc.addPage();
      const { width, height } = currentPage.getSize();

      // 设置字体
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

      let yPosition = height - styling.margins.top;
      const lineHeight = styling.fontSize * styling.lineHeight;

      // 绘制标题
      if (content.title) {
        currentPage.drawText(content.title, {
          x: styling.margins.left,
          y: yPosition,
          size: styling.fontSize * 1.5,
          font: boldFont,
          color: this.hexToRgb(styling.colors.primary),
        });
        yPosition -= lineHeight * 2;
      }

      // 绘制作者信息
      if (content.author) {
        currentPage.drawText(`作者: ${content.author}`, {
          x: styling.margins.left,
          y: yPosition,
          size: styling.fontSize * 0.9,
          font: font,
          color: this.hexToRgb(styling.colors.secondary),
        });
        yPosition -= lineHeight;
      }

      // 绘制描述
      if (content.description) {
        currentPage.drawText(content.description, {
          x: styling.margins.left,
          y: yPosition,
          size: styling.fontSize * 0.9,
          font: font,
          color: this.hexToRgb(styling.colors.secondary),
        });
        yPosition -= lineHeight * 2;
      }

      // 处理内容 - 简单的文本处理，保持中文字符
      const lines = this.splitContentIntoLines(
        content.content,
        width - styling.margins.left - styling.margins.right,
        font,
        styling.fontSize
      );

      for (const line of lines) {
        if (yPosition < styling.margins.bottom + lineHeight) {
          // 为当前页面添加水印
          if (watermarkConfig?.enabled !== false) {
            await this.addWatermarkToPage(currentPage, watermarkConfig, pdfDoc);
          }
          
          // 添加新页面
          currentPage = pdfDoc.addPage();
          yPosition = currentPage.getSize().height - styling.margins.top;
        }

        // 直接使用原始文本，不进行中文字符替换
        currentPage.drawText(line.text, {
          x: styling.margins.left,
          y: yPosition,
          size: line.isHeading ? styling.fontSize * 1.2 : styling.fontSize,
          font: line.isHeading ? boldFont : font,
          color: line.isHeading
            ? this.hexToRgb(styling.colors.primary)
            : this.hexToRgb(styling.colors.text),
        });

        yPosition -= lineHeight * (line.isHeading ? 1.5 : 1);
      }

      // 为最后一页添加水印
      if (watermarkConfig?.enabled !== false) {
        await this.addWatermarkToPage(currentPage, watermarkConfig, pdfDoc);
      }

      const pdfBytes = await pdfDoc.save();

      const writeFile = promisify(fs.writeFile);
      await writeFile(outputPath, pdfBytes);

      return {
        success: true,
        outputPath,
        fileSize: pdfBytes.length,
        format: 'pdf',
      };
    } catch (error: any) {
      return {
        success: false,
        error: `PDF转换失败: ${error.message}`,
        format: 'pdf',
      };
    }
  }

  /**
   * 为PDF生成创建HTML内容
   */
  private createHTMLForPDF(content: DocumentContent, styling: any): string {
    const { fontSize, fontFamily, colors } = styling;

    let html = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${content.title ?? 'Document'}</title>
    <style>
        /* 使用本地字体，避免外部资源引用的安全风险 */
        body {
            font-family: '微软雅黑', 'Microsoft YaHei', 'SimHei', 'Noto Sans SC', Arial, sans-serif;
            font-size: ${this.escapeHtml(fontSize.toString())}px;
            line-height: ${this.escapeHtml(styling.lineHeight.toString())};
            color: ${this.escapeHtml(colors.text)};
            margin: 0;
            padding: 20px;
            background: white;
            word-wrap: break-word;
            word-break: break-all;
        }
        
        p {
            margin: 12px 0;
            line-height: 1.6;
            text-align: justify;
            word-spacing: 0.1em;
            padding: 4px 0;
        }
        
        h1, h2, h3, h4, h5, h6 {
            margin: 20px 0 12px 0;
            line-height: 1.4;
            padding: 6px 0;
        }
        
        pre {
            white-space: pre-wrap;
            word-wrap: break-word;
            margin: 12px 0;
            padding: 12px;
            background-color: #f8f9fa;
            border-radius: 4px;
        }
        
        br {
            line-height: 1.6;
        }
        
        h1 {
            color: ${this.escapeHtml(colors.primary)};
            font-size: ${fontSize * 1.8}px;
            font-weight: 700;
            margin-bottom: 20px;
            line-height: 1.2;
        }
        
        h2 {
            color: ${this.escapeHtml(colors.primary)};
            font-size: ${fontSize * 1.4}px;
            font-weight: 700;
            margin-top: 30px;
            margin-bottom: 15px;
            line-height: 1.3;
        }
        
        h3 {
            color: ${this.escapeHtml(colors.primary)};
            font-size: ${fontSize * 1.2}px;
            font-weight: 700;
            margin-top: 25px;
            margin-bottom: 12px;
            line-height: 1.3;
        }
        
        p {
            margin-bottom: 15px;
            text-align: justify;
        }
        
        ul, ol {
            margin-bottom: 15px;
            padding-left: 30px;
        }
        
        li {
            margin-bottom: 8px;
        }
        
        strong {
            font-weight: 700;
        }
        
        .author {
            color: ${this.escapeHtml(colors.secondary)};
            font-size: ${fontSize * 0.9}px;
            margin-bottom: 10px;
        }
        
        .description {
            color: ${this.escapeHtml(colors.secondary)};
            font-size: ${fontSize * 0.9}px;
            margin-bottom: 20px;
            font-style: italic;
        }
        
        @media print {
            body {
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
        }
    </style>
</head>
<body>
`;

    // 添加标题
    if (content.title) {
      html += `    <h1>${this.escapeHtml(content.title)}</h1>\n`;
    }

    // 添加作者信息
    if (content.author) {
      html += `    <div class="author">作者: ${this.escapeHtml(content.author)}</div>\n`;
    }

    // 添加描述
    if (content.description) {
      html += `    <div class="description">${this.escapeHtml(content.description)}</div>\n`;
    }

    // 添加内容
    html += this.parseContentToHTML(content.content);

    html += `
</body>
</html>`;

    return html;
  }

  /**
   * HTML转义
   */
  private escapeHtml(text: string): string {
    const div = { innerHTML: '' } as any;
    div.textContent = text;
    return (
      div.innerHTML ||
      text.replace(/[&<>"']/g, (match: string) => {
        const escapeMap: { [key: string]: string } = {
          '&': '&amp;',
          '<': '&lt;',
          '>': '&gt;',
          '"': '&quot;',
          "'": '&#39;',
        };
        return escapeMap[match];
      })
    );
  }

  /**
   * 旧的PDF生成方法（使用pdf-lib，保留作为备用）
   */
  private async convertToPDFLegacy(
    content: DocumentContent,
    outputPath: string,
    styling: any
  ): Promise<DocumentConversionResult> {
    try {
      const { PDFDocument, rgb, StandardFonts } = await import('pdf-lib');

      const pdfDoc = await PDFDocument.create();
      const page = pdfDoc.addPage();
      const { width, height } = page.getSize();

      // 设置字体
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

      let yPosition = height - styling.margins.top;
      const lineHeight = styling.fontSize * styling.lineHeight;

      // 绘制标题
      if (content.title) {
        page.drawText(content.title, {
          x: styling.margins.left,
          y: yPosition,
          size: styling.fontSize * 1.5,
          font: boldFont,
          color: this.hexToRgb(styling.colors.primary),
        });
        yPosition -= lineHeight * 2;
      }

      // 绘制作者信息
      if (content.author) {
        page.drawText(`Author: ${content.author}`, {
          x: styling.margins.left,
          y: yPosition,
          size: styling.fontSize * 0.9,
          font: font,
          color: this.hexToRgb(styling.colors.secondary),
        });
        yPosition -= lineHeight;
      }

      // 绘制描述
      if (content.description) {
        page.drawText(`Description: ${content.description}`, {
          x: styling.margins.left,
          y: yPosition,
          size: styling.fontSize * 0.9,
          font: font,
          color: this.hexToRgb(styling.colors.secondary),
        });
        yPosition -= lineHeight * 2;
      }

      // 绘制内容
      const lines = this.splitTextIntoLines(
        content.content,
        width - styling.margins.left - styling.margins.right,
        font,
        styling.fontSize
      );

      let currentPage = page;

      for (const line of lines) {
        if (yPosition < styling.margins.bottom + lineHeight) {
          // 添加新页面
          currentPage = pdfDoc.addPage();
          yPosition = currentPage.getSize().height - styling.margins.top;
        }

        // 处理中文文本以避免编码问题
        const processedLine = this.processChineseTextForPdfLib(line);

        currentPage.drawText(processedLine, {
          x: styling.margins.left,
          y: yPosition,
          size: styling.fontSize,
          font: font,
          color: this.hexToRgb(styling.colors.text),
        });

        yPosition -= lineHeight;
      }

      const pdfBytes = await pdfDoc.save();

      const writeFile = promisify(fs.writeFile);
      await writeFile(outputPath, pdfBytes);

      return {
        success: true,
        outputPath,
        fileSize: pdfBytes.length,
        format: 'pdf',
      };
    } catch (error: any) {
      return {
        success: false,
        error: `PDF转换失败: ${error.message}`,
        format: 'pdf',
      };
    }
  }

  /**
   * 转换为DOCX格式
   */
  private async convertToDocx(
    content: DocumentContent,
    outputPath: string,
    styling: any
  ): Promise<DocumentConversionResult> {
    try {
      const children: any[] = [];

      // 添加标题
      if (content.title) {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: content.title,
                bold: true,
                size: Math.round(styling.fontSize * 1.5 * 2), // Word使用半点单位
                color: styling.colors.primary.replace('#', ''),
              }),
            ],
            heading: HeadingLevel.TITLE,
            spacing: { after: 400 },
          })
        );
      }

      // 添加作者信息
      if (content.author) {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `作者: ${content.author}`,
                size: Math.round(styling.fontSize * 0.9 * 2),
                color: styling.colors.secondary.replace('#', ''),
              }),
            ],
            spacing: { after: 200 },
          })
        );
      }

      // 添加描述
      if (content.description) {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `描述: ${content.description}`,
                size: Math.round(styling.fontSize * 0.9 * 2),
                color: styling.colors.secondary.replace('#', ''),
              }),
            ],
            spacing: { after: 400 },
          })
        );
      }

      // 添加内容段落
      const paragraphs = content.content.split('\n\n');
      for (const paragraph of paragraphs) {
        if (paragraph.trim()) {
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: paragraph.trim(),
                  size: styling.fontSize * 2,
                  color: styling.colors.text.replace('#', ''),
                }),
              ],
              spacing: { after: 200 },
            })
          );
        }
      }

      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: styling.margins.top * 20, // Word使用twips单位
                  bottom: styling.margins.bottom * 20,
                  left: styling.margins.left * 20,
                  right: styling.margins.right * 20,
                },
              },
            },
            children,
          },
        ],
      });

      const buffer = await Packer.toBuffer(doc);

      const writeFile = promisify(fs.writeFile);
      await writeFile(outputPath, buffer);

      return {
        success: true,
        outputPath,
        fileSize: buffer.length,
        format: 'docx',
      };
    } catch (error: any) {
      return {
        success: false,
        error: `DOCX转换失败: ${error.message}`,
        format: 'docx',
      };
    }
  }

  /**
   * 解析内容为HTML
   */
  private parseContentToHTML(content: string): string {
    // 简单的文本到HTML转换
    return content
      .split('\n\n')
      .map(paragraph => {
        const trimmed = paragraph.trim();
        if (!trimmed) return '';

        // 检查是否是标题
        if (trimmed.startsWith('# ')) {
          return `<h1>${this.escapeHtml(trimmed.substring(2))}</h1>`;
        } else if (trimmed.startsWith('## ')) {
          return `<h2>${this.escapeHtml(trimmed.substring(3))}</h2>`;
        } else if (trimmed.startsWith('### ')) {
          return `<h3>${this.escapeHtml(trimmed.substring(4))}</h3>`;
        } else if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
          // 简单的列表处理
          const items = trimmed
            .split('\n')
            .map(line => {
              if (line.startsWith('- ') || line.startsWith('* ')) {
                return `<li>${this.escapeHtml(line.substring(2))}</li>`;
              }
              return this.escapeHtml(line);
            })
            .join('\n');
          return `<ul>\n${items}\n</ul>`;
        } else {
          return `<p>${this.escapeHtml(trimmed).replace(/\n/g, '<br>')}</p>`;
        }
      })
      .filter(p => p)
      .join('\n');
  }

  /**
   * 将内容分割成适合PDF的行，支持标题识别
   */
  private splitContentIntoLines(
    content: string,
    maxWidth: number,
    font: any,
    fontSize: number
  ): Array<{ text: string; isHeading: boolean }> {
    const lines: Array<{ text: string; isHeading: boolean }> = [];
    const contentLines = content.split('\n');

    for (const line of contentLines) {
      const processedLines = this.processContentLine(line, maxWidth, fontSize);
      lines.push(...processedLines);
    }

    return lines;
  }

  /**
   * 处理单行内容
   */
  private processContentLine(
    line: string,
    maxWidth: number,
    fontSize: number
  ): Array<{ text: string; isHeading: boolean }> {
    const trimmedLine = line.trim();
    
    if (!trimmedLine) {
      return [{ text: '', isHeading: false }];
    }

    const isHeading = this.isMarkdownHeading(trimmedLine);
    const text = this.removeMarkdownHeadingMarkers(trimmedLine, isHeading);
    
    return this.wrapTextToLines(text, maxWidth, fontSize, isHeading);
  }

  /**
   * 检查是否是Markdown标题
   */
  private isMarkdownHeading(line: string): boolean {
    return line.startsWith('#') || line.startsWith('##');
  }

  /**
   * 移除Markdown标题标记
   */
  private removeMarkdownHeadingMarkers(text: string, isHeading: boolean): string {
    return isHeading ? text.replace(/^#+\s*/, '') : text;
  }

  /**
   * 将文本包装成多行
   */
  private wrapTextToLines(
    text: string,
    maxWidth: number,
    fontSize: number,
    isHeading: boolean
  ): Array<{ text: string; isHeading: boolean }> {
    const lines: Array<{ text: string; isHeading: boolean }> = [];
    const words = text.split(' ');
    let currentLine = '';

    for (const word of words) {
      const testLine = currentLine ? `${currentLine} ${word}` : word;
      const estimatedWidth = this.estimateTextWidth(testLine, fontSize);

      if (estimatedWidth <= maxWidth) {
        currentLine = testLine;
      } else {
        if (currentLine) {
          lines.push({ text: currentLine, isHeading });
          currentLine = word;
        } else {
          lines.push({ text: word, isHeading });
        }
      }
    }

    if (currentLine) {
      lines.push({ text: currentLine, isHeading });
    }

    return lines;
  }

  /**
   * 估算文本宽度
   */
  private estimateTextWidth(text: string, fontSize: number): number {
    return text.length * fontSize * 0.6;
  }

  /**
   * 将文本分割成适合PDF页面宽度的行
   */
  private splitTextIntoLines(
    text: string,
    maxWidth: number,
    font: any,
    fontSize: number
  ): string[] {
    const lines: string[] = [];
    const paragraphs = text.split('\n');

    for (const paragraph of paragraphs) {
      const paragraphLines = this.processParagraphForPDF(paragraph, maxWidth, font, fontSize);
      lines.push(...paragraphLines);
    }

    return lines;
  }

  /**
   * 处理单个段落用于PDF
   */
  private processParagraphForPDF(
    paragraph: string,
    maxWidth: number,
    font: any,
    fontSize: number
  ): string[] {
    if (!paragraph.trim()) {
      return [''];
    }

    return this.wrapParagraphToLines(paragraph, maxWidth, font, fontSize);
  }

  /**
   * 将段落包装成多行
   */
  private wrapParagraphToLines(
    paragraph: string,
    maxWidth: number,
    font: any,
    fontSize: number
  ): string[] {
    const lines: string[] = [];
    const words = paragraph.split(' ');
    let currentLine = '';

    for (const word of words) {
      const testLine = currentLine ? `${currentLine} ${word}` : word;
      const textWidth = font.widthOfTextAtSize(testLine, fontSize);

      if (textWidth <= maxWidth) {
        currentLine = testLine;
      } else {
        if (currentLine) {
          lines.push(currentLine);
          currentLine = word;
        } else {
          lines.push(word);
        }
      }
    }

    if (currentLine) {
      lines.push(currentLine);
    }

    return lines;
  }

  /**
   * 处理中文文本，替换为ASCII字符以避免PDF编码问题
   */
  private processChineseText(text: string): string {
    // 对于新的PDF生成方法（使用playwright-mcp），保持中文字符不变
    // 只处理一些可能导致问题的特殊字符
    return (
      text
        // 标准化中文标点符号（可选，保持原样也可以）
        .replace(/\u3000/g, ' ') // 中文空格转换为普通空格
        .replace(/\u2014/g, '—') // 标准化破折号
        .replace(/\u2026/g, '…')
    ); // 标准化省略号
  }

  /**
   * 处理中文文本用于pdf-lib（旧方法）
   */
  private processChineseTextForPdfLib(text: string): string {
    return (
      text
        // 替换中文字符
        .replace(/[\u4e00-\u9fa5]/g, '?')
        // 替换所有中文标点符号和特殊字符
        .replace(/[\u3000-\u303f\uff00-\uffef]/g, function (match) {
          const charCode = match.charCodeAt(0);
          switch (charCode) {
            case 0xff0c:
              return ','; // 中文逗号
            case 0xff1a:
              return ':'; // 中文冒号
            case 0xff1b:
              return ';'; // 中文分号
            case 0xff1f:
              return '?'; // 中文问号
            case 0xff01:
              return '!'; // 中文感叹号
            case 0xff08:
              return '('; // 中文左括号
            case 0xff09:
              return ')'; // 中文右括号
            case 0xff0e:
              return '.'; // 中文句号
            case 0x3001:
              return ','; // 中文顿号
            case 0x3002:
              return '.'; // 中文句号
            case 0x300c:
              return '"'; // 中文左引号
            case 0x300d:
              return '"'; // 中文右引号
            case 0x300e:
              return '<'; // 中文左书名号
            case 0x300f:
              return '>'; // 中文右书名号
            case 0x3000:
              return ' '; // 中文空格
            case 0x2014:
              return '-'; // 破折号
            case 0x2026:
              return '...'; // 省略号
            default:
              return match.charCodeAt(0) > 127 ? '?' : match;
          }
        })
        // 最后再次检查，替换任何剩余的非ASCII字符
        .replace(/[^\x00-\x7F]/g, '?')
    );
  }

  /**
   * 将十六进制颜色转换为RGB
   */
  private hexToRgb(hex: string): any {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    if (result) {
      return rgb(
        parseInt(result[1], 16) / 255,
        parseInt(result[2], 16) / 255,
        parseInt(result[3], 16) / 255
      );
    }
    return rgb(0, 0, 0);
  }

  /**
   * 为PDF页面添加水印
   */
  private async addWatermarkToPage(page: any, watermarkConfig: any, pdfDoc: any): Promise<void> {
    if (!watermarkConfig || watermarkConfig.enabled === false) {
      return;
    }

    const { width, height } = page.getSize();
    const { StandardFonts, rgb } = await import('pdf-lib');
    
    // 水印配置默认值
    const config = {
      text: watermarkConfig.text || 'doc-ops-mcp',
      opacity: watermarkConfig.opacity || 0.1,
      fontSize: watermarkConfig.fontSize || 48,
      rotation: watermarkConfig.rotation || -45,
      spacing: {
        x: watermarkConfig.spacing?.x || 200,
        y: watermarkConfig.spacing?.y || 150
      }
    };

    // 如果提供了图片路径，优先使用图片水印
    if (watermarkConfig.imagePath && fs.existsSync(watermarkConfig.imagePath)) {
      try {
        const readFile = promisify(fs.readFile);
        const imageBytes = await readFile(watermarkConfig.imagePath);
        let image;
        
        // 根据文件扩展名判断图片类型
        const ext = watermarkConfig.imagePath.toLowerCase().split('.').pop();
        if (ext === 'png') {
          image = await pdfDoc.embedPng(imageBytes);
        } else if (ext === 'jpg' || ext === 'jpeg') {
          image = await pdfDoc.embedJpg(imageBytes);
        } else {
           console.warn('不支持的图片格式，使用文字水印');
           await this.addTextWatermark(page, config, width, height, pdfDoc);
           return;
         }

        // 绘制图片水印
        const imageSize = Math.min(width, height) * 0.3; // 图片大小为页面最小边的30%
        
        // 计算水印位置，规则铺满页面
        for (let x = -imageSize; x < width + imageSize; x += config.spacing.x) {
          for (let y = -imageSize; y < height + imageSize; y += config.spacing.y) {
            page.drawImage(image, {
              x: x,
              y: y,
              width: imageSize,
              height: imageSize,
              opacity: config.opacity,
              rotate: {
                type: 'degrees',
                angle: config.rotation,
              },
            });
          }
        }
      } catch (error) {
         console.warn('图片水印添加失败，使用文字水印:', error);
         await this.addTextWatermark(page, config, width, height, pdfDoc);
       }
    } else {
       // 使用文字水印
       await this.addTextWatermark(page, config, width, height, pdfDoc);
     }
  }

  /**
   * 添加文字水印
   */
  private async addTextWatermark(page: any, config: any, width: number, height: number, pdfDoc?: any): Promise<void> {
    const { StandardFonts, rgb } = await import('pdf-lib');
    // 如果没有传入pdfDoc，尝试从page获取，否则创建新的字体
    let font;
    try {
      font = await (pdfDoc || page.doc).embedFont(StandardFonts.Helvetica);
    } catch {
      // 如果无法获取字体，使用默认字体
      font = StandardFonts.Helvetica;
    }
    
    // 计算文字水印的位置，斜着规则铺满页面
    const textWidth = config.text.length * config.fontSize * 0.6; // 估算文字宽度
    const textHeight = config.fontSize;
    
    // 计算旋转后的实际占用空间
    const radians = (config.rotation * Math.PI) / 180;
    const cos = Math.abs(Math.cos(radians));
    const sin = Math.abs(Math.sin(radians));
    const rotatedWidth = textWidth * cos + textHeight * sin;
    const rotatedHeight = textWidth * sin + textHeight * cos;
    
    // 规则铺满页面
    for (let x = -rotatedWidth; x < width + rotatedWidth; x += config.spacing.x) {
      for (let y = -rotatedHeight; y < height + rotatedHeight; y += config.spacing.y) {
        page.drawText(config.text, {
          x: x,
          y: y,
          size: config.fontSize,
          font: font,
          color: rgb(0.5, 0.5, 0.5), // 灰色
          opacity: config.opacity,
          rotate: {
            type: 'degrees',
            angle: config.rotation,
          },
        });
      }
    }
  }

  /**
   * 生成输出路径
   */
  private generateOutputPath(filename: string, format: string): string {
    const sanitizedFilename = filename.replace(/[^a-zA-Z0-9\u4e00-\u9fa5_-]/g, '_');
    return `${sanitizedFilename}.${format}`;
  }

  /**
   * 创建Docker简介内容
   */
  createDockerIntroContent(): DocumentContent {
    return {
      title: 'Docker 简介',
      author: 'AI助手',
      description: '一份关于Docker容器技术的详细介绍文档',
      content: `Docker 是一个开源的应用容器引擎，它允许开发者将应用及其依赖打包到一个轻量级、可移植的容器中，然后发布到任何流行的 Linux、Windows 或 macOS 机器上，也可以实现虚拟化。容器是完全使用沙箱机制，相互之间不会有任何接口。

## 核心概念

Docker 的核心在于"容器化"技术。与传统的虚拟机相比，容器共享主机的操作系统内核，因此它们更加轻量、启动速度更快、占用的系统资源也更少。这使得在一台物理服务器上运行更多的应用成为可能，极大地提高了资源利用率。

容器与虚拟机的主要区别在于：容器直接运行在宿主机的内核上，而虚拟机需要完整的操作系统。这使得容器更加高效和快速。

## 主要优势

### 一致的环境
Docker 解决了"在我的机器上可以运行"的经典问题。通过将应用和其环境一起打包，确保了从开发到测试再到生产环境的一致性。

### 快速部署与扩展
容器的轻量级特性使其可以秒级启动和停止，方便进行快速的应用部署、回滚和弹性伸缩。

### 隔离与安全
每个容器都在一个独立的环境中运行，应用之间互不影响，提供了良好的隔离性。

### 资源效率
相比传统虚拟机，Docker 容器占用更少的系统资源，可以在同一台机器上运行更多的应用实例。

### 可移植性
容器可以在任何支持 Docker 的平台上运行，无论是开发环境、测试环境还是生产环境。

## 核心组件

### Docker Engine
Docker Engine 是 Docker 的核心组件，负责创建和管理容器。它包括一个服务器端守护进程、REST API 以及命令行接口客户端。

### Docker Image
Docker 镜像是一个只读的模板，用于创建 Docker 容器。镜像包含了运行应用所需的所有内容：代码、运行时、库、环境变量和配置文件。

### Docker Container
容器是镜像的运行实例。可以启动、停止、移动和删除容器。每个容器都是相互隔离的、保证安全的平台。

## 应用场景

### 微服务架构
Docker 非常适合微服务架构，每个服务可以独立打包和部署。

### 持续集成/持续部署
在 CI/CD 流水线中，Docker 可以确保构建和部署环境的一致性。

### 云原生应用
Docker 是云原生应用的基础，与 Kubernetes 等编排工具完美配合。

### 开发环境标准化
团队成员可以使用相同的 Docker 环境进行开发，避免环境差异问题。

## 总结

总而言之，Docker 通过提供一种标准化的打包和运行应用的方式，简化了应用的开发、测试和部署流程，是现代 DevOps 和微服务架构中不可或缺的关键工具。它不仅提高了开发效率，还增强了应用的可移植性和可扩展性，为现代软件开发带来了革命性的变化。

随着云计算和微服务架构的普及，Docker 将继续在软件开发和部署领域发挥重要作用，成为现代应用架构的基石。`,
    };
  }
}

export { DocumentContent, ConversionOptions, DocumentConversionResult };
