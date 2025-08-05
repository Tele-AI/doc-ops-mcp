/**
 * HTML 转换器 - 支持HTML到多种格式的转换
 * 实现HTML到PDF、Markdown、TXT、DOCX的转换功能
 */

import { promises as fs } from 'fs';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from 'docx';
const cheerio = require('cheerio');
const path = require('path');
import { EnhancedHtmlToMarkdownConverter } from './enhancedHtmlToMarkdownConverter';
import { EnhancedHtmlToDocxConverter } from './enhancedHtmlToDocxConverter';

// 转换选项接口
interface HtmlConversionOptions {
  preserveStyles?: boolean;
  outputPath?: string;
  debug?: boolean;
  // PDF特定选项
  pdfOptions?: {
    format?: 'A4' | 'A3' | 'Letter';
    orientation?: 'portrait' | 'landscape';
    margins?: {
      top?: string;
      bottom?: string;
      left?: string;
      right?: string;
    };
  };
  // DOCX特定选项
  docxOptions?: {
    fontSize?: number;
    fontFamily?: string;
    lineSpacing?: number;
  };
}

// 转换结果接口
interface HtmlConversionResult {
  success: boolean;
  outputPath?: string;
  content?: string | Buffer;
  metadata?: {
    originalFormat: string;
    targetFormat: string;
    contentLength: number;
    converter: string;
  };
  error?: string;
  requiresExternalTool?: boolean;
  externalToolInstructions?: string;
}

/**
 * HTML 转换器类
 */
class HtmlConverter {
  private options: HtmlConversionOptions = {};

  constructor() {}

  /**
   * HTML 转 PDF
   */
  async convertHtmlToPdf(
    inputPath: string,
    options: HtmlConversionOptions = {}
  ): Promise<HtmlConversionResult> {
    try {
      this.options = {
        preserveStyles: true,
        debug: false,
        ...options,
      };

      if (this.options.debug) {
        console.log('🚀 开始 HTML 到 PDF 转换...');
        console.log('📄 输入文件:', inputPath);
      }

      // 验证输入文件
      await fs.access(inputPath);
      const htmlContent = await fs.readFile(inputPath, 'utf-8');

      // 生成输出路径
      const outputPath = this.options.outputPath || inputPath.replace(/\.html?$/i, '.pdf');

      // 由于需要浏览器引擎来生成 PDF，返回外部工具指令
      const instructions = `
HTML 转 PDF 转换 - 需要 playwright-mcp 完成

✅ 已完成 (当前 MCP):
  1. HTML 文件验证: ${inputPath}
  2. 输出路径确定: ${outputPath}

🎯 需要执行 (playwright-mcp):
  请运行以下命令完成 PDF 转换:
  
  1. browser_navigate("file://${path.resolve(inputPath)}")
  2. browser_wait_for({ time: 3 })
  3. browser_pdf_save({ filename: "${outputPath}" })

📁 最终输出: ${outputPath}

💡 优势:
  - 完美保留 HTML 样式和布局
  - 支持CSS样式和图片
  - 高质量PDF输出
`;

      if (this.options.debug) {
        console.log(instructions);
      }

      return {
        success: true,
        outputPath,
        requiresExternalTool: true,
        externalToolInstructions: instructions,
        metadata: {
          originalFormat: 'html',
          targetFormat: 'pdf',
          contentLength: htmlContent.length,
          converter: 'html-converter',
        },
      };
    } catch (error: any) {
      console.error('❌ HTML 转 PDF 失败:', error.message);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * HTML 转 Markdown
   */
  async convertHtmlToMarkdown(
    inputPath: string,
    options: HtmlConversionOptions = {}
  ): Promise<HtmlConversionResult> {
    try {
      this.options = {
        preserveStyles: false, // Markdown 不支持复杂样式
        debug: false,
        ...options,
      };

      if (this.options.debug) {
        console.log('🚀 开始 HTML 到 Markdown 转换...');
        console.log('📄 输入文件:', inputPath);
      }

      // 读取HTML文件
      const htmlContent = await fs.readFile(inputPath, 'utf-8');

      // 使用增强的HTML到Markdown转换器
      const enhancedConverter = new EnhancedHtmlToMarkdownConverter();
      const result = await enhancedConverter.convertHtmlToMarkdown(inputPath, {
        preserveStyles: true,
        includeCSS: false,
        debug: true,
      });

      if (!result.success) {
        throw new Error(result.error || 'HTML到Markdown转换失败');
      }

      const markdownContent = result.content || '';

      // 生成输出路径
      const outputPath = this.options.outputPath || inputPath.replace(/\.html?$/i, '.md');

      // 保存文件
      await fs.writeFile(outputPath, markdownContent, 'utf-8');

      if (this.options.debug) {
        console.log('✅ Markdown 转换完成:', outputPath);
      }

      return {
        success: true,
        outputPath,
        content: markdownContent,
        metadata: {
          originalFormat: 'html',
          targetFormat: 'markdown',
          contentLength: markdownContent.length,
          converter: 'html-converter',
        },
      };
    } catch (error: any) {
      console.error('❌ HTML 转 Markdown 失败:', error.message);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * HTML 转 TXT
   */
  async convertHtmlToTxt(
    inputPath: string,
    options: HtmlConversionOptions = {}
  ): Promise<HtmlConversionResult> {
    try {
      this.options = {
        preserveStyles: false, // TXT 不支持样式
        debug: false,
        ...options,
      };

      if (this.options.debug) {
        console.log('🚀 开始 HTML 到 TXT 转换...');
        console.log('📄 输入文件:', inputPath);
      }

      // 读取HTML文件
      const htmlContent = await fs.readFile(inputPath, 'utf-8');

      // 使用cheerio解析HTML并提取纯文本
      const $ = cheerio.load(htmlContent);

      // 移除script和style标签
      $('script, style').remove();

      // 提取纯文本内容
      const textContent = this.htmlToText($);

      // 生成输出路径
      const outputPath = this.options.outputPath || inputPath.replace(/\.html?$/i, '.txt');

      // 保存文件
      await fs.writeFile(outputPath, textContent, 'utf-8');

      if (this.options.debug) {
        console.log('✅ TXT 转换完成:', outputPath);
      }

      return {
        success: true,
        outputPath,
        content: textContent,
        metadata: {
          originalFormat: 'html',
          targetFormat: 'txt',
          contentLength: textContent.length,
          converter: 'html-converter',
        },
      };
    } catch (error: any) {
      console.error('❌ HTML 转 TXT 失败:', error.message);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * HTML 转 DOCX
   */
  async convertHtmlToDocx(
    inputPath: string,
    options: HtmlConversionOptions = {}
  ): Promise<HtmlConversionResult> {
    try {
      this.options = {
        preserveStyles: true,
        debug: false,
        docxOptions: {
          fontSize: 12,
          fontFamily: 'Times New Roman',
          lineSpacing: 1.15,
        },
        ...options,
      };

      if (this.options.debug) {
        console.log('🚀 开始 HTML 到 DOCX 转换...');
        console.log('📄 输入文件:', inputPath);
      }

      // 读取HTML文件
      const htmlContent = await fs.readFile(inputPath, 'utf-8');

      // 使用增强的HTML到DOCX转换器
      const enhancedConverter = new EnhancedHtmlToDocxConverter();
      const docxBuffer = await enhancedConverter.convertHtmlToDocx(htmlContent);

      // 生成输出路径
      const outputPath =
        options.outputPath || this.options.outputPath || inputPath.replace(/\.html?$/i, '.docx');

      // 保存文件
      await fs.writeFile(outputPath, docxBuffer);

      if (this.options.debug) {
        console.log('✅ DOCX 转换完成:', outputPath);
      }

      return {
        success: true,
        outputPath,
        content: docxBuffer,
        metadata: {
          originalFormat: 'html',
          targetFormat: 'docx',
          contentLength: docxBuffer.length,
          converter: 'html-converter',
        },
      };
    } catch (error: any) {
      console.error('❌ HTML 转 DOCX 失败:', error.message);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * 将HTML转换为Markdown格式
   */
  private htmlToMarkdown($: any): string {
    let markdown = '';

    // 处理body内容
    const body = $('body').length > 0 ? $('body') : $.root();

    body.children().each((i: number, element: any) => {
      markdown += this.processElementToMarkdown($, $(element));
    });

    // 清理多余的空行
    return markdown
      .replace(/\n{3,}/g, '\n\n')
      .replace(/^\n+/, '')
      .replace(/\n+$/, '\n');
  }

  /**
   * 处理单个HTML元素转换为Markdown
   */
  private processElementToMarkdown($: any, element: any): string {
    const tagName = element.prop('tagName')?.toLowerCase();
    let text = element.text().trim();

    // 解码HTML实体
    text = this.decodeHtmlEntities(text);

    if (!text && !['br', 'hr', 'img'].includes(tagName)) {
      return '';
    }

    switch (tagName) {
      case 'h1':
        return `# ${text}\n\n`;
      case 'h2':
        return `## ${text}\n\n`;
      case 'h3':
        return `### ${text}\n\n`;
      case 'h4':
        return `#### ${text}\n\n`;
      case 'h5':
        return `##### ${text}\n\n`;
      case 'h6':
        return `###### ${text}\n\n`;
      case 'p':
        return `${this.processInlineElements($, element)}\n\n`;
      case 'strong':
      case 'b':
        return `**${text}**`;
      case 'em':
      case 'i':
        return `*${text}*`;
      case 'code':
        return `\`${text}\``;
      case 'pre':
        // 尝试从code子元素中提取语言信息
        const codeElement = element.find('code').first();
        let language = '';
        if (codeElement.length > 0) {
          const className = codeElement.attr('class') || '';
          const languageMatch = className.match(/language-([\w-]+)/);
          if (languageMatch) {
            language = languageMatch[1];
          }
        }
        return `\`\`\`${language}\n${text}\n\`\`\`\n\n`;
      case 'a':
        const href = element.attr('href');
        return href ? `[${text}](${href})` : text;
      case 'img':
        const src = element.attr('src');
        const alt = element.attr('alt') || 'Image';
        return src ? `![${alt}](${src})\n\n` : '';
      case 'ul':
        let ulMarkdown = '';
        element.children('li').each((i: number, li: any) => {
          const liText = this.decodeHtmlEntities($(li).text().trim());
          ulMarkdown += `- ${liText}\n`;
        });
        return ulMarkdown + '\n';
      case 'ol':
        let olMarkdown = '';
        element.children('li').each((i: number, li: any) => {
          const liText = this.decodeHtmlEntities($(li).text().trim());
          olMarkdown += `${i + 1}. ${liText}\n`;
        });
        return olMarkdown + '\n';
      case 'blockquote':
        return `> ${text}\n\n`;
      case 'hr':
        return `---\n\n`;
      case 'br':
        return '\n';
      case 'div':
      case 'span':
        return this.processInlineElements($, element) + '\n';
      default:
        return text ? `${text}\n` : '';
    }
  }

  /**
   * 处理内联元素
   */
  private processInlineElements($: any, element: any): string {
    let result = '';

    element.contents().each((i: number, node: any) => {
      if (node.type === 'text') {
        result += this.decodeHtmlEntities($(node).text());
      } else if (node.type === 'tag') {
        const $node = $(node);
        const tagName = $node.prop('tagName')?.toLowerCase();
        const text = $node.text();

        switch (tagName) {
          case 'strong':
          case 'b':
            result += `**${text}**`;
            break;
          case 'em':
          case 'i':
            result += `*${text}*`;
            break;
          case 'code':
            result += `\`${text}\``;
            break;
          case 'a':
            const href = $node.attr('href');
            result += href ? `[${text}](${href})` : text;
            break;
          default:
            result += text;
        }
      }
    });

    return result;
  }

  /**
   * 将HTML转换为纯文本
   */
  private htmlToText($: any): string {
    // 移除script和style标签
    $('script, style').remove();

    let text = '';
    const body = $('body').length > 0 ? $('body') : $.root();

    body.find('*').each((i: number, element: any) => {
      const $element = $(element);
      const tagName = $element.prop('tagName')?.toLowerCase();

      // 在块级元素前后添加换行
      if (['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'div', 'br', 'hr'].includes(tagName)) {
        if (tagName === 'br') {
          text += '\n';
        } else if (tagName === 'hr') {
          text += '\n---\n';
        } else {
          const elementText = $element.text().trim();
          if (elementText) {
            text += elementText + '\n';
          }
        }
      }
    });

    // 如果没有找到块级元素，直接提取所有文本
    if (!text.trim()) {
      text = body.text();
    }

    // 清理多余的空行
    return text
      .replace(/\n{3,}/g, '\n\n')
      .replace(/^\n+/, '')
      .replace(/\n+$/, '\n');
  }

  /**
   * 将HTML转换为DOCX文档
   */
  private async htmlToDocx($: any): Promise<Buffer> {
    const paragraphs: any[] = [];
    const docxOptions = this.options.docxOptions!;

    // 处理body内容
    const body = $('body').length > 0 ? $('body') : $.root();

    body.children().each((i: number, element: any) => {
      const paragraph = this.processElementToDocx($, $(element), docxOptions);
      if (paragraph) {
        paragraphs.push(paragraph);
      }
    });

    // 如果没有找到内容，添加默认段落
    if (paragraphs.length === 0) {
      const text = body.text().trim();
      if (text) {
        paragraphs.push(
          new Paragraph({
            children: [
              new TextRun({
                text: text,
                size: docxOptions.fontSize! * 2, // DOCX使用半点单位
                font: docxOptions.fontFamily,
              }),
            ],
          })
        );
      }
    }

    // 创建文档
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: paragraphs,
        },
      ],
    });

    // 生成Buffer
    return await Packer.toBuffer(doc);
  }

  /**
   * 处理单个HTML元素转换为DOCX段落 - 增强版，支持CSS样式解析
   */
  private processElementToDocx($: any, element: any, docxOptions: any): any {
    const tagName = element.prop('tagName')?.toLowerCase();
    const text = element.text().trim();

    if (!text && !['br', 'hr'].includes(tagName)) {
      return null;
    }

    const baseFontSize = docxOptions.fontSize * 2; // DOCX使用半点单位
    const baseFontFamily = docxOptions.fontFamily;

    // 解析元素样式
    const elementStyles = this.parseElementStyles(element);
    const paragraphConfig: any = {};

    // 设置对齐方式
    if (elementStyles.textAlign) {
      switch (elementStyles.textAlign) {
        case 'center':
          paragraphConfig.alignment = AlignmentType.CENTER;
          break;
        case 'right':
          paragraphConfig.alignment = AlignmentType.RIGHT;
          break;
        case 'justify':
          paragraphConfig.alignment = AlignmentType.JUSTIFIED;
          break;
        default:
          paragraphConfig.alignment = AlignmentType.LEFT;
      }
    }

    // 设置间距
    if (elementStyles.marginTop || elementStyles.marginBottom) {
      paragraphConfig.spacing = {};
      if (elementStyles.marginTop) {
        paragraphConfig.spacing.before = elementStyles.marginTop;
      }
      if (elementStyles.marginBottom) {
        paragraphConfig.spacing.after = elementStyles.marginBottom;
      }
    }

    // 设置缩进
    if (elementStyles.paddingLeft || elementStyles.marginLeft) {
      paragraphConfig.indent = {
        left: elementStyles.paddingLeft || elementStyles.marginLeft || 0,
      };
    }

    switch (tagName) {
      case 'h1':
        paragraphConfig.heading = HeadingLevel.HEADING_1;
        paragraphConfig.children = [
          new TextRun({
            text: text,
            bold: true,
            size: elementStyles.fontSize || baseFontSize * 1.5,
            font: elementStyles.fontFamily || baseFontFamily,
            color: elementStyles.color || '000000',
          }),
        ];
        return new Paragraph(paragraphConfig);

      case 'h2':
        paragraphConfig.heading = HeadingLevel.HEADING_2;
        paragraphConfig.children = [
          new TextRun({
            text: text,
            bold: true,
            size: elementStyles.fontSize || baseFontSize * 1.3,
            font: elementStyles.fontFamily || baseFontFamily,
            color: elementStyles.color || '000000',
          }),
        ];
        return new Paragraph(paragraphConfig);

      case 'h3':
        paragraphConfig.heading = HeadingLevel.HEADING_3;
        paragraphConfig.children = [
          new TextRun({
            text: text,
            bold: true,
            size: elementStyles.fontSize || baseFontSize * 1.1,
            font: elementStyles.fontFamily || baseFontFamily,
            color: elementStyles.color || '000000',
          }),
        ];
        return new Paragraph(paragraphConfig);

      case 'h4':
      case 'h5':
      case 'h6':
        paragraphConfig.children = [
          new TextRun({
            text: text,
            bold: true,
            size: elementStyles.fontSize || baseFontSize,
            font: elementStyles.fontFamily || baseFontFamily,
            color: elementStyles.color || '000000',
          }),
        ];
        return new Paragraph(paragraphConfig);

      case 'p':
      case 'div':
        paragraphConfig.children = this.processInlineElementsToDocx($, element, docxOptions);
        return new Paragraph(paragraphConfig);

      case 'br':
        return new Paragraph({
          children: [
            new TextRun({
              text: '',
              size: baseFontSize,
              font: baseFontFamily,
            }),
          ],
        });

      case 'hr':
        return new Paragraph({
          children: [
            new TextRun({
              text: '---',
              size: baseFontSize,
              font: baseFontFamily,
            }),
          ],
          alignment: AlignmentType.CENTER,
        });

      case 'blockquote':
        paragraphConfig.children = this.processInlineElementsToDocx($, element, docxOptions);
        paragraphConfig.indent = { left: 720 }; // 0.5英寸缩进
        paragraphConfig.border = {
          left: {
            color: 'CCCCCC',
            size: 6,
            style: 'single',
          },
        };
        return new Paragraph(paragraphConfig);

      default:
        if (text) {
          paragraphConfig.children = [
            new TextRun({
              text: text,
              size: elementStyles.fontSize || baseFontSize,
              font: elementStyles.fontFamily || baseFontFamily,
              color: elementStyles.color || '000000',
              bold: elementStyles.bold || false,
              italics: elementStyles.italic || false,
              underline: elementStyles.underline ? {} : undefined,
              strike: elementStyles.strikethrough || false,
            }),
          ];
          return new Paragraph(paragraphConfig);
        }
        return null;
    }
  }

  /**
   * 处理内联元素转换为DOCX TextRun - 增强版，支持CSS样式解析
   */
  private processInlineElementsToDocx($: any, element: any, docxOptions: any): any[] {
    const textRuns: any[] = [];
    const baseFontSize = docxOptions.fontSize * 2;
    const baseFontFamily = docxOptions.fontFamily;

    element.contents().each((i: number, node: any) => {
      if (node.type === 'text') {
        const text = $(node).text();
        if (text.trim()) {
          textRuns.push(
            new TextRun({
              text: text,
              size: baseFontSize,
              font: baseFontFamily,
            })
          );
        }
      } else if (node.type === 'tag') {
        const $node = $(node);
        const tagName = $node.prop('tagName')?.toLowerCase();
        const text = $node.text();

        if (text.trim()) {
          // 解析CSS样式
          const styles = this.parseElementStyles($node);

          // 构建TextRun配置
          const textRunConfig: any = {
            text: text,
            size: styles.fontSize || baseFontSize,
            font: styles.fontFamily || baseFontFamily,
          };

          // 应用样式
          if (styles.bold || ['strong', 'b'].includes(tagName)) {
            textRunConfig.bold = true;
          }

          if (styles.italic || ['em', 'i'].includes(tagName)) {
            textRunConfig.italics = true;
          }

          if (styles.underline || tagName === 'u') {
            textRunConfig.underline = {};
          }

          if (styles.strikethrough || ['del', 'strike', 's'].includes(tagName)) {
            textRunConfig.strike = true;
          }

          if (styles.color) {
            textRunConfig.color = styles.color;
          }

          if (styles.highlight) {
            textRunConfig.highlight = styles.highlight;
          }

          textRuns.push(new TextRun(textRunConfig));
        }
      }
    });

    // 如果没有找到内联元素，使用整个元素的文本并应用样式
    if (textRuns.length === 0) {
      const text = element.text().trim();
      if (text) {
        const styles = this.parseElementStyles(element);
        const textRunConfig: any = {
          text: text,
          size: styles.fontSize || baseFontSize,
          font: styles.fontFamily || baseFontFamily,
        };

        if (styles.bold) textRunConfig.bold = true;
        if (styles.italic) textRunConfig.italics = true;
        if (styles.underline) textRunConfig.underline = {};
        if (styles.strikethrough) textRunConfig.strike = true;
        if (styles.color) textRunConfig.color = styles.color;
        if (styles.highlight) textRunConfig.highlight = styles.highlight;

        textRuns.push(new TextRun(textRunConfig));
      }
    }

    return textRuns;
  }

  /**
   * 解析元素的CSS样式
   */
  private parseElementStyles(element: any): any {
    const styles: any = {};

    // 获取style属性
    const styleAttr = element.attr('style');
    if (styleAttr) {
      const styleRules = styleAttr.split(';');

      for (const rule of styleRules) {
        const [property, value] = rule.split(':').map((s: string) => s.trim());
        if (!property || !value) continue;

        switch (property.toLowerCase()) {
          case 'font-weight':
            if (value === 'bold' || parseInt(value) >= 600) {
              styles.bold = true;
            }
            break;
          case 'font-style':
            if (value === 'italic') {
              styles.italic = true;
            }
            break;
          case 'text-decoration':
            if (value.includes('underline')) {
              styles.underline = true;
            }
            if (value.includes('line-through')) {
              styles.strikethrough = true;
            }
            break;
          case 'font-size':
            const fontSize = this.parseFontSize(value);
            if (fontSize) {
              styles.fontSize = fontSize * 2; // DOCX使用半点单位
            }
            break;
          case 'font-family':
            styles.fontFamily = value.replace(/["']/g, '').split(',')[0].trim();
            break;
          case 'color':
            const color = this.parseColor(value);
            if (color) {
              styles.color = color;
            }
            break;
          case 'background-color':
            const bgColor = this.parseColor(value);
            if (bgColor) {
              styles.highlight = bgColor;
            }
            break;
          case 'text-align':
            styles.textAlign = value.toLowerCase();
            break;
          case 'margin-top':
            const marginTop = this.parseSpacing(value);
            if (marginTop) {
              styles.marginTop = marginTop;
            }
            break;
          case 'margin-bottom':
            const marginBottom = this.parseSpacing(value);
            if (marginBottom) {
              styles.marginBottom = marginBottom;
            }
            break;
          case 'margin-left':
            const marginLeft = this.parseSpacing(value);
            if (marginLeft) {
              styles.marginLeft = marginLeft;
            }
            break;
          case 'padding-left':
            const paddingLeft = this.parseSpacing(value);
            if (paddingLeft) {
              styles.paddingLeft = paddingLeft;
            }
            break;
          case 'line-height':
            const lineHeight = parseFloat(value);
            if (!isNaN(lineHeight)) {
              styles.lineHeight = lineHeight;
            }
            break;
        }
      }
    }

    // 检查class属性中的常见样式类
    const classAttr = element.attr('class');
    if (classAttr) {
      const classes = classAttr.split(' ');
      for (const cls of classes) {
        switch (cls.toLowerCase()) {
          case 'bold':
          case 'font-weight-bold':
          case 'fw-bold':
            styles.bold = true;
            break;
          case 'italic':
          case 'font-style-italic':
          case 'fst-italic':
            styles.italic = true;
            break;
          case 'underline':
          case 'text-decoration-underline':
            styles.underline = true;
            break;
          case 'strikethrough':
          case 'text-decoration-line-through':
            styles.strikethrough = true;
            break;
        }
      }
    }

    return styles;
  }

  /**
   * 解析字体大小
   */
  private parseFontSize(value: string): number | null {
    const match = value.match(/([\d.]+)(px|pt|em|rem|%)?/);
    if (!match) return null;

    const size = parseFloat(match[1]);
    const unit = match[2] || 'px';

    switch (unit) {
      case 'pt':
        return size;
      case 'px':
        return size * 0.75; // 1px ≈ 0.75pt
      case 'em':
      case 'rem':
        return size * 12; // 假设基础字体为12pt
      case '%':
        return (size / 100) * 12; // 假设基础字体为12pt
      default:
        return size;
    }
  }

  /**
   * 解析颜色值
   */
  private parseColor(value: string): string | null {
    // 移除颜色值前的#号
    if (value.startsWith('#')) {
      return value.substring(1);
    }

    // 处理rgb/rgba颜色
    const rgbMatch = value.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (rgbMatch) {
      const r = parseInt(rgbMatch[1]).toString(16).padStart(2, '0');
      const g = parseInt(rgbMatch[2]).toString(16).padStart(2, '0');
      const b = parseInt(rgbMatch[3]).toString(16).padStart(2, '0');
      return r + g + b;
    }

    // 处理常见颜色名称
    const colorMap: { [key: string]: string } = {
      black: '000000',
      white: 'FFFFFF',
      red: 'FF0000',
      green: '008000',
      blue: '0000FF',
      yellow: 'FFFF00',
      orange: 'FFA500',
      purple: '800080',
      gray: '808080',
      grey: '808080',
    };

    return colorMap[value.toLowerCase()] || null;
  }

  /**
   * 解析间距值（margin、padding等）
   */
  private parseSpacing(value: string): number | null {
    const match = value.match(/([\d.]+)(px|pt|em|rem|%)?/);
    if (!match) return null;

    const size = parseFloat(match[1]);
    const unit = match[2] || 'px';

    switch (unit) {
      case 'pt':
        return size * 20; // 1pt = 20 twips (DOCX单位)
      case 'px':
        return size * 15; // 1px ≈ 15 twips
      case 'em':
      case 'rem':
        return size * 240; // 1em ≈ 12pt ≈ 240 twips
      case '%':
        return (size / 100) * 240; // 相对于基础字体大小
      default:
        return size * 20; // 默认按pt处理
    }
  }

  /**
   * 解码HTML实体
   */
  private decodeHtmlEntities(text: string): string {
    return (
      text
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&')
        .replace(/&#39;/g, "'")
        .replace(/&quot;/g, '"')
        .replace(/&nbsp;/g, ' ')
        .replace(/&copy;/g, '©')
        .replace(/&reg;/g, '®')
        .replace(/&trade;/g, '™')
        // 处理数字HTML实体
        .replace(/&#(\d+);/g, (match, num) => {
          return String.fromCharCode(parseInt(num, 10));
        })
        // 处理十六进制HTML实体
        .replace(/&#x([0-9A-Fa-f]+);/g, (match, hex) => {
          return String.fromCharCode(parseInt(hex, 16));
        })
    );
  }
}

// 导出便捷函数
export async function convertHtmlToPdf(
  inputPath: string,
  options: HtmlConversionOptions = {}
): Promise<HtmlConversionResult> {
  const converter = new HtmlConverter();
  return await converter.convertHtmlToPdf(inputPath, options);
}

export async function convertHtmlToMarkdown(
  inputPath: string,
  options: HtmlConversionOptions = {}
): Promise<HtmlConversionResult> {
  try {
    const enhancedConverter = new EnhancedHtmlToMarkdownConverter();
    const result = await enhancedConverter.convertHtmlToMarkdown(inputPath, {
      preserveStyles: true,
      includeCSS: false,
      outputPath: options.outputPath,
      debug: options.debug || false,
    });

    if (!result.success) {
      return {
        success: false,
        error: result.error || 'HTML到Markdown转换失败',
      };
    }

    return {
      success: true,
      outputPath: result.outputPath,
      content: result.content,
      metadata: result.metadata,
    };
  } catch (error: any) {
    return {
      success: false,
      error: error.message,
    };
  }
}

export async function convertHtmlToTxt(
  inputPath: string,
  options: HtmlConversionOptions = {}
): Promise<HtmlConversionResult> {
  const converter = new HtmlConverter();
  return await converter.convertHtmlToTxt(inputPath, options);
}

export async function convertHtmlToDocx(
  inputPath: string,
  outputPath?: string,
  options: HtmlConversionOptions = {}
): Promise<HtmlConversionResult> {
  const converter = new HtmlConverter();
  const finalOptions = {
    ...options,
    outputPath: outputPath || options.outputPath,
  };
  return await converter.convertHtmlToDocx(inputPath, finalOptions);
}

export { HtmlConverter, HtmlConversionOptions, HtmlConversionResult };
