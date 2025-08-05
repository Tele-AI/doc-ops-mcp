/**
 * 增强的 HTML 到 DOCX 转换器
 * 解决样式丢失问题，实现精确的样式映射
 */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  UnderlineType,
} from 'docx';
import { promises as fs } from 'fs';
const cheerio = require('cheerio');
const path = require('path');

interface StyleMapping {
  heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel];
  size?: number;
  bold?: boolean;
  italics?: boolean;
  underline?: any;
  color?: string;
  highlight?: "none" | "yellow" | "green" | "cyan" | "magenta" | "blue" | "red" | "darkBlue" | "darkCyan" | "darkGreen" | "darkMagenta" | "darkRed" | "darkYellow" | "darkGray" | "lightGray" | "black" | "white";
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
}

interface ParsedElement {
  tag: string;
  text: string;
  html: string;
  styles: any;
  children: ParsedElement[];
}

class EnhancedHtmlToDocxConverter {
  private styleMap: Map<string, StyleMapping> = new Map();

  constructor() {
    this.initializeStyles();
  }

  private initializeStyles() {
    // 预定义样式映射
    this.styleMap.set('h1', {
      heading: HeadingLevel.HEADING_1,
      size: 32,
      bold: true,
      color: '2F5496',
    });

    this.styleMap.set('h2', {
      heading: HeadingLevel.HEADING_2,
      size: 28,
      bold: true,
      color: '2F5496',
    });

    this.styleMap.set('h3', {
      heading: HeadingLevel.HEADING_3,
      size: 24,
      bold: true,
      color: '1F3763',
    });

    this.styleMap.set('h4', {
      heading: HeadingLevel.HEADING_4,
      size: 22,
      bold: true,
      color: '1F3763',
    });

    this.styleMap.set('h5', {
      heading: HeadingLevel.HEADING_5,
      size: 20,
      bold: true,
      color: '1F3763',
    });

    this.styleMap.set('h6', {
      heading: HeadingLevel.HEADING_6,
      size: 18,
      bold: true,
      color: '1F3763',
    });

    this.styleMap.set('p', {
      size: 22,
      color: '000000',
    });

    this.styleMap.set('strong', {
      bold: true,
    });

    this.styleMap.set('b', {
      bold: true,
    });

    this.styleMap.set('em', {
      italics: true,
    });

    this.styleMap.set('i', {
      italics: true,
    });

    this.styleMap.set('u', {
      underline: {
        type: UnderlineType.SINGLE,
      },
    });

    this.styleMap.set('blockquote', {
      size: 22,
      italics: true,
      color: '666666',
    });

    this.styleMap.set('pre', {
      size: 18,
      color: '000000',
    });

    this.styleMap.set('code', {
      size: 18,
      color: 'd73a49',
    });
  }

  async convertHtmlToDocx(htmlContent: string): Promise<Buffer> {
    const $ = cheerio.load(htmlContent);
    const docElements: any[] = [];

    // 解析 HTML 结构
    const elements = this.parseHtmlElements($);

    // 转换为 DOCX 元素
    for (const element of elements) {
      const docxElement = this.createDocxElement(element, $);
      if (docxElement) {
        if (Array.isArray(docxElement)) {
          docElements.push(...docxElement);
        } else {
          docElements.push(docxElement);
        }
      }
    }

    // 创建文档
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: docElements,
        },
      ],
    });

    return await Packer.toBuffer(doc);
  }

  private parseHtmlElements($: any): ParsedElement[] {
    const elements: ParsedElement[] = [];

    $('body')
      .children()
      .each((i: number, elem: any) => {
        const $elem = $(elem);
        const tagName = elem.tagName.toLowerCase();

        // 提取样式信息
        const styles = this.extractStyles($elem);

        elements.push({
          tag: tagName,
          text: $elem.text(),
          html: $elem.html(),
          styles: styles,
          children: this.parseChildren($elem),
        });
      });

    return elements;
  }

  private parseChildren($elem: any): ParsedElement[] {
    const children: ParsedElement[] = [];

    $elem.children().each((i: number, child: any) => {
      const $ = cheerio.load('');
      const $child = $(child);
      const tagName = child.tagName ? child.tagName.toLowerCase() : 'text';

      children.push({
        tag: tagName,
        text: $child.text(),
        html: $child.html(),
        styles: this.extractStyles($child),
        children: [],
      });
    });

    return children;
  }

  private extractStyles($elem: any): any {
    const styles: any = {};

    // 提取内联样式
    const inlineStyle = $elem.attr('style');
    if (inlineStyle) {
      const cssRules = inlineStyle.split(';');
      cssRules.forEach((rule: string) => {
        const [property, value] = rule.split(':').map((s: string) => s.trim());
        if (property && value) {
          styles[property] = value;
        }
      });
    }

    // 提取类样式
    const className = $elem.attr('class');
    if (className) {
      styles.className = className;
    }

    return styles;
  }

  private createDocxElement(element: ParsedElement, $: any): any {
    const baseStyle = this.styleMap.get(element.tag) || {};
    const customStyle = this.convertCssToDocx(element.styles);
    const finalStyle = { ...baseStyle, ...customStyle };

    switch (element.tag) {
      case 'h1':
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6':
        return new Paragraph({
          heading: finalStyle.heading,
          alignment: finalStyle.alignment || AlignmentType.LEFT,
          spacing: {
            before: 480, // 标题前间距
            after: 240, // 标题后间距
            line: 360, // 1.5倍行距
            lineRule: 'auto',
          },
          children: [
            new TextRun({
              text: element.text,
              bold: finalStyle.bold !== false, // 默认粗体
              size: finalStyle.size,
              color: finalStyle.color || '2c3e50',
              italics: finalStyle.italics,
              font: {
                name: 'Segoe UI Emoji, Apple Color Emoji, Noto Color Emoji, Microsoft YaHei, SimHei, Arial, sans-serif',
              },
            }),
          ],
        });

      case 'p':
        return new Paragraph({
          alignment: finalStyle.alignment || AlignmentType.LEFT,
          spacing: {
            line: 360, // 1.5倍行距
            lineRule: 'auto',
            after: 240, // 段后间距
          },
          children: this.createTextRuns(element, finalStyle, $),
        });

      case 'pre':
        // 代码块处理
        return this.createCodeBlock(element, $);

      case 'blockquote':
        return new Paragraph({
          alignment: finalStyle.alignment || AlignmentType.LEFT,
          spacing: {
            line: 360,
            lineRule: 'auto',
            before: 240,
            after: 240,
          },
          indent: {
            left: 720, // 0.5 inch
            right: 360,
          },
          border: {
            left: {
              style: 'single',
              size: 12,
              color: '3498db',
            },
          },
          children: [
            new TextRun({
              text: element.text,
              italics: true,
              size: finalStyle.size || 22,
              color: finalStyle.color || '5a6c7d',
              font: {
                name: 'Segoe UI Emoji, Apple Color Emoji, Noto Color Emoji, Microsoft YaHei, SimHei, Arial, sans-serif',
              },
            }),
          ],
        });

      case 'ul':
      case 'ol':
        return this.createListElements(element, finalStyle, $);

      case 'table':
        // 表格处理（简化版）
        return this.createTableElements(element, finalStyle, $);

      default:
        if (element.text.trim()) {
          const runOptions: any = {
            text: element.text,
            bold: finalStyle.bold,
            italics: finalStyle.italics,
            size: finalStyle.size,
            color: finalStyle.color,
            font: {
              name: 'Segoe UI Emoji, Apple Color Emoji, Noto Color Emoji, Microsoft YaHei, SimHei, Arial, sans-serif',
            },
          };
          if (finalStyle.underline) {
            runOptions.underline = finalStyle.underline;
          }
          return new Paragraph({
            children: [new TextRun(runOptions)],
          });
        }
        return null;
    }
  }

  private createTextRuns(element: ParsedElement, baseStyle: StyleMapping, $: any): any[] {
    const runs: any[] = [];
    const html = element.html;

    if (!html) {
      return [this.createSimpleTextRun(element.text, baseStyle)];
    }

    const $content = $(html);

    if ($content.length === 0) {
      return [this.createSimpleTextRun(element.text, baseStyle)];
    }

    $content.contents().each((i: number, node: any) => {
      this.processHtmlNode(node, baseStyle, runs, $);
    });

    return runs.length > 0 ? runs : [this.createSimpleTextRun(element.text, baseStyle)];
  }

  private createSimpleTextRun(text: string, baseStyle: StyleMapping): any {
    const { highlight, ...safeStyle } = baseStyle;
    return new TextRun({
      text,
      ...safeStyle,
      ...(highlight && { highlight }),
    });
  }

  private processHtmlNode(node: any, baseStyle: StyleMapping, runs: any[], $: any): void {
    if (node.type === 'text') {
      this.processTextNode(node, baseStyle, runs);
    } else if (node.type === 'tag') {
      this.processTagNode(node, baseStyle, runs, $);
    }
  }

  private processTextNode(node: any, baseStyle: StyleMapping, runs: any[]): void {
    if (node.data.trim()) {
      const textOptions: any = {
        text: node.data,
        bold: baseStyle.bold,
        italics: baseStyle.italics,
        size: baseStyle.size,
        color: baseStyle.color,
      };
      if (baseStyle.underline) {
        textOptions.underline = baseStyle.underline;
      }
      runs.push(new TextRun(textOptions));
    }
  }

  private processTagNode(node: any, baseStyle: StyleMapping, runs: any[], $: any): void {
    const tagStyle = this.applyTagStyles(node, baseStyle, $);
    const text = $(node).text();
    
    if (text.trim()) {
      const decodedText = this.decodeHtmlEntities(text);
      const nodeOptions = this.createNodeOptions(decodedText, tagStyle, node.name);
      runs.push(new TextRun(nodeOptions));
    }
  }

  private applyTagStyles(node: any, baseStyle: StyleMapping, $: any): StyleMapping {
    const tagStyle = { ...baseStyle };
    const $node = $(node);

    // 根据标签添加样式
    this.applyBasicTagStyles(node.name, tagStyle);

    // 提取该节点的样式
    const nodeStyles = this.extractStyles($node);
    const nodeDocxStyle = this.convertCssToDocx(nodeStyles);
    Object.assign(tagStyle, nodeDocxStyle);

    return tagStyle;
  }

  private applyBasicTagStyles(tagName: string, tagStyle: StyleMapping): void {
    switch (tagName) {
      case 'strong':
      case 'b':
        tagStyle.bold = true;
        break;
      case 'em':
      case 'i':
        tagStyle.italics = true;
        break;
      case 'u':
        tagStyle.underline = {
          type: UnderlineType.SINGLE,
        };
        break;
      case 'del':
      case 'strike':
        (tagStyle as any).strike = true;
        break;
      case 'code':
        tagStyle.size = 18;
        tagStyle.color = 'd73a49';
        break;
    }
  }

  private createNodeOptions(text: string, tagStyle: StyleMapping, tagName: string): any {
    const nodeOptions: any = {
      text,
      bold: tagStyle.bold,
      italics: tagStyle.italics,
      size: tagStyle.size,
      color: tagStyle.color,
    };

    // 设置字体
    if (tagName === 'code') {
      nodeOptions.font = { name: 'Consolas' };
    } else {
      nodeOptions.font = {
        name: 'Segoe UI Emoji, Apple Color Emoji, Noto Color Emoji, Microsoft YaHei, SimHei, Arial, sans-serif',
      };
    }

    // 添加其他样式
    if (tagStyle.underline) {
      nodeOptions.underline = tagStyle.underline;
    }
    if ((tagStyle as any).strike) {
      nodeOptions.strike = (tagStyle as any).strike;
    }

    return nodeOptions;
  }

  private createListElements(element: ParsedElement, baseStyle: StyleMapping, $: any): any[] {
    const paragraphs: any[] = [];
    const $list = $(element.html);

    $list.find('li').each((i: number, li: any) => {
      const $li = $(li);
      const text = $li.text();
      const bullet = element.tag === 'ul' ? '• ' : `${i + 1}. `;

      const listRunOptions: any = {
        text: bullet + text,
        size: baseStyle.size,
        color: baseStyle.color,
      };
      paragraphs.push(
        new Paragraph({
          children: [new TextRun(listRunOptions)],
          indent: {
            left: 720, // 0.5 inch
          },
        })
      );
    });

    return paragraphs;
  }

  private createTableElements(element: ParsedElement, baseStyle: StyleMapping, $: any): any[] {
    const paragraphs: any[] = [];
    const $table = $(element.html);

    // 简化的表格处理 - 转换为段落
    $table.find('tr').each((i: number, tr: any) => {
      const $tr = $(tr);
      const cells: string[] = [];

      $tr.find('td, th').each((j: number, cell: any) => {
        const $cell = $(cell);
        cells.push($cell.text());
      });

      if (cells.length > 0) {
        const tableRunOptions: any = {
          text: cells.join(' | '),
          size: baseStyle.size,
          color: baseStyle.color,
          bold: $tr.find('th').length > 0, // 表头加粗
        };
        paragraphs.push(
          new Paragraph({
            children: [new TextRun(tableRunOptions)],
          })
        );
      }
    });

    return paragraphs;
  }

  private createCodeBlock(element: ParsedElement, $: any): any[] {
    const paragraphs: any[] = [];
    const codeContent = element.text;

    // 处理代码块内容，保持格式
    const lines = codeContent.split('\n');

    // 添加代码块前的空行
    paragraphs.push(
      new Paragraph({
        children: [new TextRun({ text: ' ', size: 4 })],
        spacing: { after: 120 },
      })
    );

    lines.forEach((line: string, index: number) => {
      // 解码HTML实体
      const decodedLine = this.decodeHtmlEntities(line);

      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: decodedLine || ' ', // 空行用空格代替
              font: {
                name: 'Consolas',
              },
              size: 20, // 稍微增大字体
              color: '24292f',
            }),
          ],
          spacing: {
            line: 276, // 1.15倍行距
            lineRule: 'auto',
            before: 0,
            after: 0,
          },
          indent: {
            left: 432, // 0.3 inch
            right: 432,
          },
          border: {
            left: {
              style: 'single',
              size: 4,
              color: 'e1e4e8',
            },
          },
          shading: {
            type: 'solid',
            color: 'f6f8fa',
          },
        })
      );
    });

    // 添加代码块后的空行
    paragraphs.push(
      new Paragraph({
        children: [new TextRun({ text: ' ', size: 4 })],
        spacing: { before: 120 },
      })
    );

    return paragraphs;
  }

  private convertCssToDocx(styles: any): StyleMapping {
    const docxStyle: StyleMapping = {};

    this.convertFontSize(styles, docxStyle);
    this.convertColor(styles, docxStyle);
    this.convertBackgroundColor(styles, docxStyle);
    this.convertFontWeight(styles, docxStyle);
    this.convertFontStyle(styles, docxStyle);
    this.convertTextDecoration(styles, docxStyle);
    this.convertTextAlign(styles, docxStyle);

    return docxStyle;
  }

  private convertFontSize(styles: any, docxStyle: StyleMapping): void {
    if (styles['font-size']) {
      const fontSize = this.parseFontSize(styles['font-size']);
      if (fontSize) {
        docxStyle.size = fontSize;
      }
    }
  }

  private convertColor(styles: any, docxStyle: StyleMapping): void {
    if (styles['color']) {
      const color = this.parseColor(styles['color']);
      if (color) {
        docxStyle.color = color;
      }
    }
  }

  private convertBackgroundColor(styles: any, docxStyle: StyleMapping): void {
    if (styles['background-color']) {
      const bgColor = this.parseColor(styles['background-color']);
      if (bgColor) {
        docxStyle.highlight = this.mapColorToHighlight(bgColor);
      }
    }
  }

  private mapColorToHighlight(bgColor: string): StyleMapping['highlight'] {
    const highlightMap: { [key: string]: StyleMapping['highlight'] } = {
      '#ffff00': 'yellow',
      '#00ff00': 'green',
      '#00ffff': 'cyan',
      '#ff00ff': 'magenta',
      '#0000ff': 'blue',
      '#ff0000': 'red',
      '#000080': 'darkBlue',
      '#008080': 'darkCyan',
      '#008000': 'darkGreen',
      '#800080': 'darkMagenta',
      '#800000': 'darkRed',
      '#808000': 'darkYellow',
      '#808080': 'darkGray',
      '#c0c0c0': 'lightGray',
      '#000000': 'black',
      '#ffffff': 'white'
    };
    return highlightMap[bgColor.toLowerCase()] || 'yellow';
  }

  private convertFontWeight(styles: any, docxStyle: StyleMapping): void {
    if (styles['font-weight']) {
      const weight = styles['font-weight'].toLowerCase();
      if (weight === 'bold' || weight === 'bolder' || parseInt(weight) >= 600) {
        docxStyle.bold = true;
      }
    }
  }

  private convertFontStyle(styles: any, docxStyle: StyleMapping): void {
    if (styles['font-style']) {
      if (styles['font-style'].toLowerCase() === 'italic') {
        docxStyle.italics = true;
      }
    }
  }

  private convertTextDecoration(styles: any, docxStyle: StyleMapping): void {
    if (styles['text-decoration']) {
      const decoration = styles['text-decoration'].toLowerCase();
      if (decoration.includes('underline')) {
        docxStyle.underline = {
          type: UnderlineType.SINGLE,
        };
      }
    }
  }

  private convertTextAlign(styles: any, docxStyle: StyleMapping): void {
    if (styles['text-align']) {
      const align = styles['text-align'].toLowerCase();
      switch (align) {
        case 'center':
          docxStyle.alignment = AlignmentType.CENTER;
          break;
        case 'right':
          docxStyle.alignment = AlignmentType.RIGHT;
          break;
        case 'justify':
          docxStyle.alignment = AlignmentType.JUSTIFIED;
          break;
        default:
          docxStyle.alignment = AlignmentType.LEFT;
      }
    }
  }

  private parseFontSize(value: string): number | null {
    if (!value) return null;

    // 移除单位并转换为数字
    const numValue = parseFloat(value.replace(/[^0-9.]/g, ''));
    if (isNaN(numValue)) return null;

    // 根据单位转换
    if (value.includes('pt')) {
      return numValue * 2; // pt to half-points
    } else if (value.includes('px')) {
      return Math.round(numValue * 1.5); // px to half-points (approximate)
    } else if (value.includes('em')) {
      return Math.round(numValue * 24); // em to half-points (assuming 12pt base)
    }

    return Math.round(numValue * 2); // 默认当作pt处理
  }

  private parseColor(value: string): string | null {
    if (!value) return null;

    // 移除空格
    value = value.trim();

    // 处理十六进制颜色
    if (value.startsWith('#')) {
      let hex = value.substring(1);
      // 处理3位十六进制颜色（如#abc -> #aabbcc）
      if (hex.length === 3) {
        hex = hex
          .split('')
          .map(c => c + c)
          .join('');
      }
      return hex.toUpperCase();
    }

    // 处理rgb()格式
    const rgbMatch = value.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
    if (rgbMatch) {
      const r = parseInt(rgbMatch[1]).toString(16).padStart(2, '0');
      const g = parseInt(rgbMatch[2]).toString(16).padStart(2, '0');
      const b = parseInt(rgbMatch[3]).toString(16).padStart(2, '0');
      return (r + g + b).toUpperCase();
    }

    // 处理rgba()格式（忽略透明度）
    const rgbaMatch = value.match(/rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*[\d.]+\s*\)/);
    if (rgbaMatch) {
      const r = parseInt(rgbaMatch[1]).toString(16).padStart(2, '0');
      const g = parseInt(rgbaMatch[2]).toString(16).padStart(2, '0');
      const b = parseInt(rgbaMatch[3]).toString(16).padStart(2, '0');
      return (r + g + b).toUpperCase();
    }

    // 处理hsl()格式
    const hslMatch = value.match(/hsl\(\s*(\d+)\s*,\s*(\d+)%\s*,\s*(\d+)%\s*\)/);
    if (hslMatch) {
      const h = parseInt(hslMatch[1]) / 360;
      const s = parseInt(hslMatch[2]) / 100;
      const l = parseInt(hslMatch[3]) / 100;
      const rgb = this.hslToRgb(h, s, l);
      return rgb
        .map(c => Math.round(c).toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase();
    }

    // 处理命名颜色
    const colorMap: { [key: string]: string } = {
      red: 'FF0000',
      green: '008000',
      blue: '0000FF',
      black: '000000',
      white: 'FFFFFF',
      yellow: 'FFFF00',
      orange: 'FFA500',
      purple: '800080',
      gray: '808080',
      grey: '808080',
      pink: 'FFC0CB',
      brown: 'A52A2A',
      cyan: '00FFFF',
      magenta: 'FF00FF',
      lime: '00FF00',
      navy: '000080',
      maroon: '800000',
      olive: '808000',
      teal: '008080',
      silver: 'C0C0C0',
      darkred: '8B0000',
      darkgreen: '006400',
      darkblue: '00008B',
      lightgray: 'D3D3D3',
      lightgrey: 'D3D3D3',
      darkgray: 'A9A9A9',
      darkgrey: 'A9A9A9',
    };

    return colorMap[value.toLowerCase()] || null;
  }

  private hslToRgb(h: number, s: number, l: number): number[] {
    let r, g, b;

    if (s === 0) {
      r = g = b = l; // achromatic
    } else {
      const hue2rgb = (p: number, q: number, t: number) => {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
      };

      const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
      const p = 2 * l - q;
      r = hue2rgb(p, q, h + 1 / 3);
      g = hue2rgb(p, q, h);
      b = hue2rgb(p, q, h - 1 / 3);
    }

    return [r * 255, g * 255, b * 255];
  }

  private decodeHtmlEntities(text: string): string {
    // 首先处理HTML实体
    let decoded = text
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&#39;/g, "'")
      .replace(/&quot;/g, '"')
      .replace(/&nbsp;/g, ' ')
      .replace(/&copy;/g, '©')
      .replace(/&reg;/g, '®')
      .replace(/&trade;/g, '™');

    // 处理数字HTML实体（如&#8203;）
    decoded = decoded.replace(/&#(\d+);/g, (match, num) => {
      return String.fromCharCode(parseInt(num, 10));
    });

    // 处理十六进制HTML实体（如&#x200B;）
    decoded = decoded.replace(/&#x([0-9A-Fa-f]+);/g, (match, hex) => {
      return String.fromCharCode(parseInt(hex, 16));
    });

    return decoded;
  }
}

export { EnhancedHtmlToDocxConverter };
