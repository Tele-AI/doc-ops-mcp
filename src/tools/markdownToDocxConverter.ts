/**
 * Markdown 到 DOCX 转换器 - 支持样式保留和美化
 * 解决 Markdown 转 DOCX 时样式丢失的问题
 */

import { promises as fs } from 'fs';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  UnderlineType,
} from 'docx';

// 转换选项接口
interface MarkdownToDocxOptions {
  preserveStyles?: boolean;
  theme?: 'default' | 'professional' | 'academic' | 'modern';
  includeTableOfContents?: boolean;
  customStyles?: {
    fontSize?: number;
    fontFamily?: string;
    lineSpacing?: number;
    margins?: {
      top?: number;
      bottom?: number;
      left?: number;
      right?: number;
    };
  };
  outputPath?: string;
  debug?: boolean;
}

// 转换结果接口
interface MarkdownToDocxResult {
  success: boolean;
  content?: Buffer;
  docxPath?: string;
  metadata?: {
    originalFormat: string;
    targetFormat: string;
    stylesPreserved: boolean;
    theme: string;
    converter: string;
    contentLength: number;
    headingsCount: number;
    paragraphsCount: number;
    tablesCount: number;
  };
  error?: string;
}

// 解析的内容元素接口
interface ParsedElement {
  type: 'heading' | 'paragraph' | 'list' | 'table' | 'blockquote' | 'code' | 'hr';
  level?: number;
  content: string;
  children?: ParsedElement[];
  attributes?: { [key: string]: any };
}

/**
 * Markdown 到 DOCX 转换器类
 */
class MarkdownToDocxConverter {
  private options: MarkdownToDocxOptions = {};
  private themes: Map<string, any>;

  constructor() {
    this.themes = new Map();
    this.initializeThemes();
  }

  /**
   * 初始化预设主题
   */
  private initializeThemes(): void {
    // 默认主题
    this.themes.set('default', {
      fontSize: 11,
      fontFamily: 'Segoe UI',
      lineSpacing: 1.15,
      headingStyles: {
        h1: { size: 16, color: '2F5496', bold: false },
        h2: { size: 13, color: '2F5496', bold: false },
        h3: { size: 12, color: '1F3763', bold: false },
        h4: { size: 11, color: '2F5496', bold: true },
        h5: { size: 11, color: '2F5496', bold: true },
        h6: { size: 11, color: '2F5496', bold: true },
      },
      margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 }, // 1 inch = 1440 twips
    });

    // 专业主题
    this.themes.set('professional', {
      fontSize: 12,
      fontFamily: 'Segoe UI',
      lineSpacing: 1.5,
      headingStyles: {
        h1: { size: 18, color: '000000', bold: true },
        h2: { size: 16, color: '000000', bold: true },
        h3: { size: 14, color: '000000', bold: true },
        h4: { size: 12, color: '000000', bold: true },
        h5: { size: 12, color: '000000', bold: true },
        h6: { size: 12, color: '000000', bold: true },
      },
      margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
    });

    // 学术主题
    this.themes.set('academic', {
      fontSize: 12,
      fontFamily: 'Segoe UI',
      lineSpacing: 2.0,
      headingStyles: {
        h1: { size: 14, color: '000000', bold: true },
        h2: { size: 13, color: '000000', bold: true },
        h3: { size: 12, color: '000000', bold: true },
        h4: { size: 12, color: '000000', bold: false },
        h5: { size: 12, color: '000000', bold: false },
        h6: { size: 12, color: '000000', bold: false },
      },
      margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
    });

    // 现代主题
    this.themes.set('modern', {
      fontSize: 11,
      fontFamily: 'Segoe UI',
      lineSpacing: 1.2,
      headingStyles: {
        h1: { size: 20, color: '0078D4', bold: true },
        h2: { size: 16, color: '0078D4', bold: true },
        h3: { size: 14, color: '323130', bold: true },
        h4: { size: 12, color: '323130', bold: true },
        h5: { size: 11, color: '323130', bold: true },
        h6: { size: 11, color: '323130', bold: true },
      },
      margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
    });
  }

  /**
   * 主转换函数
   */
  async convertMarkdownToDocx(
    inputPath: string,
    options: MarkdownToDocxOptions = {}
  ): Promise<MarkdownToDocxResult> {
    try {
      this.options = {
        preserveStyles: true,
        theme: 'default',
        includeTableOfContents: false,
        debug: false,
        ...options,
      };

      if (this.options.debug) {
        console.log('🚀 开始 Markdown 到 DOCX 转换...');
        console.log('📄 输入文件:', inputPath);
        console.log('🎨 使用主题:', this.options.theme);
      }

      // 读取 Markdown 文件
      const markdownContent = await fs.readFile(inputPath, 'utf-8');

      // 解析 Markdown 内容
      const parsedElements = await this.parseMarkdown(markdownContent);

      // 生成 DOCX 文档
      const docxDocument = await this.generateDocxDocument(parsedElements);

      // 生成文档缓冲区
      const docxBuffer = await Packer.toBuffer(docxDocument);

      // 保存文件（如果指定了输出路径）
      let docxPath: string | undefined;
      if (this.options.outputPath) {
        const { validateAndSanitizePath } = require('../security/securityConfig');
        const allowedPaths = [process.cwd()];
        const validatedPath = validateAndSanitizePath(this.options.outputPath, allowedPaths);
        if (validatedPath) {
          docxPath = validatedPath;
          await fs.writeFile(validatedPath, docxBuffer);
        }

        if (this.options.debug) {
          console.log('✅ DOCX 文件已保存:', docxPath);
        }
      }

      // 分析内容统计
      const stats = this.analyzeContent(parsedElements);

      if (this.options.debug) {
        console.log('📊 转换统计:', stats);
        console.log('✅ Markdown 转换完成');
      }

      return {
        success: true,
        content: docxBuffer,
        docxPath,
        metadata: {
          originalFormat: 'markdown',
          targetFormat: 'docx',
          stylesPreserved: this.options.preserveStyles || false,
          theme: this.options.theme || 'default',
          converter: 'markdown-to-docx-converter',
          contentLength: docxBuffer.length,
          ...stats,
        },
      };
    } catch (error: any) {
      console.error('❌ Markdown 转换失败:', error.message);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * 解析 Markdown 内容
   */
  private async parseMarkdown(markdownContent: string): Promise<ParsedElement[]> {
    const elements: ParsedElement[] = [];
    const lines = markdownContent.split('\n');
    const state = {
      inCodeBlock: false,
      codeBlockContent: '',
      inTable: false,
      tableRows: [] as string[]
    };

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmedLine = line.trim();

      if (this.processCodeBlock(trimmedLine, line, state, elements)) {
        continue;
      }

      if (this.processTable(trimmedLine, line, state, elements)) {
        continue;
      }

      if (this.processMarkdownElement(trimmedLine, elements)) {
        continue;
      }
    }

    this.finalizeTable(state, elements);
    return elements;
  }

  /**
   * 处理代码块
   */
  private processCodeBlock(
    trimmedLine: string,
    line: string,
    state: any,
    elements: ParsedElement[]
  ): boolean {
    if (trimmedLine.startsWith('```')) {
      if (state.inCodeBlock) {
        elements.push({
          type: 'code',
          content: state.codeBlockContent.trim(),
        });
        state.inCodeBlock = false;
        state.codeBlockContent = '';
      } else {
        state.inCodeBlock = true;
      }
      return true;
    }

    if (state.inCodeBlock) {
      state.codeBlockContent += line + '\n';
      return true;
    }

    return false;
  }

  /**
   * 处理表格
   */
  private processTable(
    trimmedLine: string,
    line: string,
    state: any,
    elements: ParsedElement[]
  ): boolean {
    if (trimmedLine.includes('|') && !state.inTable) {
      state.inTable = true;
      state.tableRows = [line];
      return true;
    }

    if (state.inTable) {
      if (trimmedLine.includes('|')) {
        state.tableRows.push(line);
        return true;
      } else {
        elements.push({
          type: 'table',
          content: state.tableRows.join('\n'),
        });
        state.inTable = false;
        state.tableRows = [];
        return false;
      }
    }

    return false;
  }

  /**
   * 处理其他Markdown元素
   */
  private processMarkdownElement(trimmedLine: string, elements: ParsedElement[]): boolean {
    // 处理标题
    const headingMatch = trimmedLine.match(/^(#{1,6})\s+(.+)$/);
    if (headingMatch) {
      elements.push({
        type: 'heading',
        level: headingMatch[1].length,
        content: headingMatch[2],
      });
      return true;
    }

    // 处理水平线
    if (trimmedLine.match(/^[-*_]{3,}$/)) {
      elements.push({
        type: 'hr',
        content: '',
      });
      return true;
    }

    // 处理引用
    if (trimmedLine.startsWith('>')) {
      elements.push({
        type: 'blockquote',
        content: trimmedLine.substring(1).trim(),
      });
      return true;
    }

    // 处理列表
    if (trimmedLine.match(/^[-*+]\s+/) || trimmedLine.match(/^\d+\.\s+/)) {
      elements.push({
        type: 'list',
        content: trimmedLine,
      });
      return true;
    }

    // 处理普通段落
    if (trimmedLine.length > 0) {
      elements.push({
        type: 'paragraph',
        content: trimmedLine,
      });
      return true;
    }

    return false;
  }

  /**
   * 完成表格处理
   */
  private finalizeTable(state: any, elements: ParsedElement[]): void {
    if (state.inTable && state.tableRows.length > 0) {
      elements.push({
        type: 'table',
        content: state.tableRows.join('\n'),
      });
    }
  }

  /**
   * 生成 DOCX 文档
   */
  private async generateDocxDocument(elements: ParsedElement[]): Promise<Document> {
    const theme = this.themes.get(this.options.theme || 'default')!;
    const children: any[] = [];

    for (const element of elements) {
      switch (element.type) {
        case 'heading':
          children.push(this.createHeading(element, theme));
          break;
        case 'paragraph':
          children.push(this.createParagraph(element, theme));
          break;
        case 'blockquote':
          children.push(this.createBlockquote(element, theme));
          break;
        case 'list':
          children.push(this.createListItem(element, theme));
          break;
        case 'table':
          const table = this.createTable(element, theme);
          if (table) children.push(table);
          break;
        case 'code':
          children.push(this.createCodeBlock(element, theme));
          break;
        case 'hr':
          children.push(this.createHorizontalRule(theme));
          break;
      }
    }

    return new Document({
      sections: [
        {
          properties: {
            page: {
              margin: theme.margins,
            },
          },
          children,
        },
      ],
    });
  }

  /**
   * 创建标题
   */
  private createHeading(element: ParsedElement, theme: any): Paragraph {
    const level = Math.min(element.level || 1, 6);
    const headingKey = `h${level}` as keyof typeof theme.headingStyles;
    const headingStyle = theme.headingStyles[headingKey];
    const hasEmoji = this.detectEmoji(element.content);
    const headingLevel = this.getHeadingLevel(level);

    return new Paragraph({
      heading: headingLevel,
      children: [
        new TextRun({
          text: element.content,
          font: hasEmoji ? 'Segoe UI Emoji' : theme.fontFamily,
          size: headingStyle.size * 2, // DOCX uses half-points
          color: headingStyle.color,
          bold: headingStyle.bold,
        }),
      ],
      spacing: {
        before: 240, // 12pt
        after: 120, // 6pt
      },
    });
  }

  /**
   * 检测emoji字符
   */
  private detectEmoji(content: string): boolean {
    const emojiRegex =
      /[\u{1F600}-\u{1F64F}]|[\u{1F300}-\u{1F5FF}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F900}-\u{1F9FF}]|[\u{1FA70}-\u{1FAFF}]|[\u{2B50}]|[\u{2B55}]|[\u{231A}-\u{231B}]|[\u{23E9}-\u{23EC}]|[\u{23F0}]|[\u{23F3}]|[\u{25FD}-\u{25FE}]|[\u{2614}-\u{2615}]|[\u{2648}-\u{2653}]|[\u{267F}]|[\u{2693}]|[\u{26A1}]|[\u{26AA}-\u{26AB}]|[\u{26BD}-\u{26BE}]|[\u{26C4}-\u{26C5}]|[\u{26CE}]|[\u{26D4}]|[\u{26EA}]|[\u{26F2}-\u{26F3}]|[\u{26F5}]|[\u{26FA}]|[\u{26FD}]|[\u{2705}]|[\u{270A}-\u{270B}]|[\u{2728}]|[\u{274C}]|[\u{274E}]|[\u{2753}-\u{2755}]|[\u{2757}]|[\u{2795}-\u{2797}]|[\u{27B0}]|[\u{27BF}]|[\u{2934}-\u{2935}]|[\u{2B05}-\u{2B07}]|[\u{2B1B}-\u{2B1C}]|[\u{3030}]|[\u{303D}]|[\u{3297}]|[\u{3299}]/gu;
    return emojiRegex.test(content);
  }

  /**
   * 获取标题级别
   */
  private getHeadingLevel(level: number) {
    const headingLevels = [
      HeadingLevel.HEADING_1,
      HeadingLevel.HEADING_2,
      HeadingLevel.HEADING_3,
      HeadingLevel.HEADING_4,
      HeadingLevel.HEADING_5,
      HeadingLevel.HEADING_6,
    ];
    return headingLevels[level - 1] || HeadingLevel.HEADING_6;
  }

  /**
   * 创建段落
   */
  private createParagraph(element: ParsedElement, theme: any): Paragraph {
    const textRuns = this.parseInlineFormatting(element.content, theme);

    return new Paragraph({
      children: textRuns,
      spacing: {
        line: Math.round(theme.lineSpacing * 240), // 240 = 12pt in twips
        after: 120, // 6pt
      },
    });
  }

  /**
   * 创建引用块
   */
  private createBlockquote(element: ParsedElement, theme: any): Paragraph {
    // 检测emoji字符 - 扩展Unicode范围以支持更多emoji
    const emojiRegex =
      /[\u{1F600}-\u{1F64F}]|[\u{1F300}-\u{1F5FF}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F900}-\u{1F9FF}]|[\u{1FA70}-\u{1FAFF}]|[\u{2B50}]|[\u{2B55}]|[\u{231A}-\u{231B}]|[\u{23E9}-\u{23EC}]|[\u{23F0}]|[\u{23F3}]|[\u{25FD}-\u{25FE}]|[\u{2614}-\u{2615}]|[\u{2648}-\u{2653}]|[\u{267F}]|[\u{2693}]|[\u{26A1}]|[\u{26AA}-\u{26AB}]|[\u{26BD}-\u{26BE}]|[\u{26C4}-\u{26C5}]|[\u{26CE}]|[\u{26D4}]|[\u{26EA}]|[\u{26F2}-\u{26F3}]|[\u{26F5}]|[\u{26FA}]|[\u{26FD}]|[\u{2705}]|[\u{270A}-\u{270B}]|[\u{2728}]|[\u{274C}]|[\u{274E}]|[\u{2753}-\u{2755}]|[\u{2757}]|[\u{2795}-\u{2797}]|[\u{27B0}]|[\u{27BF}]|[\u{2934}-\u{2935}]|[\u{2B05}-\u{2B07}]|[\u{2B1B}-\u{2B1C}]|[\u{3030}]|[\u{303D}]|[\u{3297}]|[\u{3299}]/gu;
    const hasEmoji = emojiRegex.test(element.content);

    return new Paragraph({
      children: [
        new TextRun({
          text: element.content,
          font: hasEmoji ? 'Segoe UI Emoji' : theme.fontFamily,
          size: theme.fontSize * 2,
          italics: true,
        }),
      ],
      indent: {
        left: 720, // 0.5 inch
      },
      spacing: {
        line: Math.round(theme.lineSpacing * 240),
        after: 120,
      },
    });
  }

  /**
   * 创建列表项
   */
  private createListItem(element: ParsedElement, theme: any): Paragraph {
    const isNumbered = element.content.match(/^\d+\.\s+/);
    const content = element.content.replace(/^[-*+\d+\.\s]+/, '');

    // 检测emoji字符 - 扩展Unicode范围以支持更多emoji
    const emojiRegex =
      /[\u{1F600}-\u{1F64F}]|[\u{1F300}-\u{1F5FF}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F900}-\u{1F9FF}]|[\u{1FA70}-\u{1FAFF}]|[\u{2B50}]|[\u{2B55}]|[\u{231A}-\u{231B}]|[\u{23E9}-\u{23EC}]|[\u{23F0}]|[\u{23F3}]|[\u{25FD}-\u{25FE}]|[\u{2614}-\u{2615}]|[\u{2648}-\u{2653}]|[\u{267F}]|[\u{2693}]|[\u{26A1}]|[\u{26AA}-\u{26AB}]|[\u{26BD}-\u{26BE}]|[\u{26C4}-\u{26C5}]|[\u{26CE}]|[\u{26D4}]|[\u{26EA}]|[\u{26F2}-\u{26F3}]|[\u{26F5}]|[\u{26FA}]|[\u{26FD}]|[\u{2705}]|[\u{270A}-\u{270B}]|[\u{2728}]|[\u{274C}]|[\u{274E}]|[\u{2753}-\u{2755}]|[\u{2757}]|[\u{2795}-\u{2797}]|[\u{27B0}]|[\u{27BF}]|[\u{2934}-\u{2935}]|[\u{2B05}-\u{2B07}]|[\u{2B1B}-\u{2B1C}]|[\u{3030}]|[\u{303D}]|[\u{3297}]|[\u{3299}]/gu;
    const hasEmoji = emojiRegex.test(content);

    return new Paragraph({
      children: [
        new TextRun({
          text: content,
          font: hasEmoji ? 'Segoe UI Emoji' : theme.fontFamily,
          size: theme.fontSize * 2,
        }),
      ],
      bullet: isNumbered
        ? undefined
        : {
            level: 0,
          },
      numbering: isNumbered
        ? {
            reference: 'default-numbering',
            level: 0,
          }
        : undefined,
      indent: {
        left: 720, // 0.5 inch
      },
      spacing: {
        line: Math.round(theme.lineSpacing * 240),
        after: 60,
      },
    });
  }

  /**
   * 创建表格
   */
  private createTable(element: ParsedElement, theme: any): Table | null {
    const lines = element.content.split('\n').filter(line => line.trim().length > 0);
    if (lines.length < 2) return null;

    const rows: TableRow[] = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (i === 1 && line.includes('---')) continue; // Skip separator line

      const cells = line
        .split('|')
        .map(cell => cell.trim())
        .filter(cell => cell.length > 0);
      if (cells.length === 0) continue;

      const tableCells = cells.map(cellContent => {
        return new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: cellContent,
                  font: /[\u{1F600}-\u{1F64F}]|[\u{1F300}-\u{1F5FF}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F900}-\u{1F9FF}]|[\u{1FA70}-\u{1FAFF}]|[\u{2B50}]|[\u{2B55}]|[\u{231A}-\u{231B}]|[\u{23E9}-\u{23EC}]|[\u{23F0}]|[\u{23F3}]|[\u{25FD}-\u{25FE}]|[\u{2614}-\u{2615}]|[\u{2648}-\u{2653}]|[\u{267F}]|[\u{2693}]|[\u{26A1}]|[\u{26AA}-\u{26AB}]|[\u{26BD}-\u{26BE}]|[\u{26C4}-\u{26C5}]|[\u{26CE}]|[\u{26D4}]|[\u{26EA}]|[\u{26F2}-\u{26F3}]|[\u{26F5}]|[\u{26FA}]|[\u{26FD}]|[\u{2705}]|[\u{270A}-\u{270B}]|[\u{2728}]|[\u{274C}]|[\u{274E}]|[\u{2753}-\u{2755}]|[\u{2757}]|[\u{2795}-\u{2797}]|[\u{27B0}]|[\u{27BF}]|[\u{2934}-\u{2935}]|[\u{2B05}-\u{2B07}]|[\u{2B1B}-\u{2B1C}]|[\u{3030}]|[\u{303D}]|[\u{3297}]|[\u{3299}]/gu.test(
                    cellContent
                  )
                    ? 'Segoe UI Emoji'
                    : theme.fontFamily,
                  size: theme.fontSize * 2,
                  bold: i === 0, // Header row
                }),
              ],
            }),
          ],
          width: {
            size: 100 / cells.length,
            type: WidthType.PERCENTAGE,
          },
        });
      });

      rows.push(new TableRow({ children: tableCells }));
    }

    return new Table({
      rows,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
    });
  }

  /**
   * 创建代码块
   */
  private createCodeBlock(element: ParsedElement, theme: any): Paragraph {
    return new Paragraph({
      children: [
        new TextRun({
          text: element.content,
          font: 'Consolas',
          size: (theme.fontSize - 1) * 2,
          color: '000000',
        }),
      ],
      shading: {
        fill: 'F5F5F5',
      },
      indent: {
        left: 360, // 0.25 inch
      },
      spacing: {
        before: 120,
        after: 120,
      },
    });
  }

  /**
   * 创建水平线
   */
  private createHorizontalRule(theme: any): Paragraph {
    return new Paragraph({
      children: [
        new TextRun({
          text: '',
          font: theme.fontFamily,
        }),
      ],
      border: {
        bottom: {
          color: 'auto',
          space: 1,
          style: 'single',
          size: 6,
        },
      },
      spacing: {
        before: 120,
        after: 120,
      },
    });
  }

  /**
   * 解析内联格式
   */
  private parseInlineFormatting(text: string, theme: any): TextRun[] {
    const runs: TextRun[] = [];
    let currentText = text;

    // 检测emoji字符 - 扩展Unicode范围以支持更多emoji
    const emojiRegex =
      /[\u{1F600}-\u{1F64F}]|[\u{1F300}-\u{1F5FF}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F900}-\u{1F9FF}]|[\u{1FA70}-\u{1FAFF}]|[\u{2B50}]|[\u{2B55}]|[\u{231A}-\u{231B}]|[\u{23E9}-\u{23EC}]|[\u{23F0}]|[\u{23F3}]|[\u{25FD}-\u{25FE}]|[\u{2614}-\u{2615}]|[\u{2648}-\u{2653}]|[\u{267F}]|[\u{2693}]|[\u{26A1}]|[\u{26AA}-\u{26AB}]|[\u{26BD}-\u{26BE}]|[\u{26C4}-\u{26C5}]|[\u{26CE}]|[\u{26D4}]|[\u{26EA}]|[\u{26F2}-\u{26F3}]|[\u{26F5}]|[\u{26FA}]|[\u{26FD}]|[\u{2705}]|[\u{270A}-\u{270B}]|[\u{2728}]|[\u{274C}]|[\u{274E}]|[\u{2753}-\u{2755}]|[\u{2757}]|[\u{2795}-\u{2797}]|[\u{27B0}]|[\u{27BF}]|[\u{2934}-\u{2935}]|[\u{2B05}-\u{2B07}]|[\u{2B1B}-\u{2B1C}]|[\u{3030}]|[\u{303D}]|[\u{3297}]|[\u{3299}]/gu;
    const hasEmoji = emojiRegex.test(currentText);

    // 处理粗体 **text**
    const boldRegex = /\*\*(.*?)\*\*/g;
    const italicRegex = /\*(.*?)\*/g;
    const codeRegex = /`(.*?)`/g;
    const linkRegex = /\[([^\]]+)\]\(([^)]+)\)/g;

    // 简化处理：先处理所有格式，然后创建单个 TextRun
    // 在实际应用中，可能需要更复杂的解析来正确处理嵌套格式

    let hasBold = boldRegex.test(currentText);
    let hasItalic = italicRegex.test(currentText);
    let hasCode = codeRegex.test(currentText);
    let hasLink = linkRegex.test(currentText);

    if (!hasBold && !hasItalic && !hasCode && !hasLink) {
      // 纯文本
      runs.push(
        new TextRun({
          text: currentText,
          font: hasEmoji ? 'Segoe UI Emoji' : theme.fontFamily,
          size: theme.fontSize * 2,
        })
      );
    } else {
      // 简化处理：移除 Markdown 标记并应用基本格式
      const cleanText = currentText
        .replace(/\*\*(.*?)\*\*/g, '$1')
        .replace(/\*(.*?)\*/g, '$1')
        .replace(/`(.*?)`/g, '$1')
        .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '$1');

      runs.push(
        new TextRun({
          text: cleanText,
          font: hasEmoji ? 'Segoe UI Emoji' : theme.fontFamily,
          size: theme.fontSize * 2,
          bold: hasBold,
          italics: hasItalic,
        })
      );
    }

    return runs;
  }

  /**
   * 分析内容统计信息
   */
  private analyzeContent(elements: ParsedElement[]): {
    headingsCount: number;
    paragraphsCount: number;
    tablesCount: number;
  } {
    const headingsCount = elements.filter(e => e.type === 'heading').length;
    const paragraphsCount = elements.filter(e => e.type === 'paragraph').length;
    const tablesCount = elements.filter(e => e.type === 'table').length;

    return {
      headingsCount,
      paragraphsCount,
      tablesCount,
    };
  }

  /**
   * 获取可用主题列表
   */
  getAvailableThemes(): string[] {
    return Array.from(this.themes.keys());
  }

  /**
   * 添加自定义主题
   */
  addCustomTheme(name: string, theme: any): void {
    this.themes.set(name, theme);
  }
}

// 导出便捷函数
export async function convertMarkdownToDocx(
  inputPath: string,
  options: MarkdownToDocxOptions = {}
): Promise<MarkdownToDocxResult> {
  const converter = new MarkdownToDocxConverter();
  return await converter.convertMarkdownToDocx(inputPath, options);
}

export { MarkdownToDocxConverter, MarkdownToDocxOptions, MarkdownToDocxResult };
