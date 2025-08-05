/**
 * StyleExtractor - XML 样式提取器
 * 解析 word/styles.xml 和 word/document.xml 获取完整样式信息
 * 构建样式映射表，支持段落样式（pPr）和文字样式（rPr）
 */

const JSZip = require('jszip');
const xml2js = require('xml2js');
const fs = require('fs/promises');

interface StyleDefinition {
  styleId: string;
  name: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering';
  css: Record<string, string>;
  mammothMapping: {
    selector: string;
    element: string;
    className?: string;
  };
  basedOn?: string;
  next?: string;
  isDefault?: boolean;
}

interface DocumentStyle {
  elementType: 'paragraph' | 'run' | 'table';
  styleId?: string;
  directFormatting?: Record<string, any>;
  css: Record<string, string>;
}

export class StyleExtractor {
  private styles: Map<string, StyleDefinition> = new Map();
  private documentStyles: DocumentStyle[] = [];
  private relationships: Map<string, string> = new Map();
  private fonts: Map<string, string> = new Map();
  private themes: Map<string, any> = new Map();

  /**
   * 主要提取函数
   */
  async extractStyles(docxPath: string): Promise<{
    styles: Map<string, StyleDefinition>;
    documentStyles: DocumentStyle[];
    css: string;
  }> {
    try {
      console.log('🎨 开始提取 DOCX 样式信息...');

      const docxBuffer = await fs.readFile(docxPath);
      const zip = await JSZip.loadAsync(docxBuffer);

      // 提取各种样式信息
      await this.extractStylesXml(zip);
      await this.extractDocumentXml(zip);
      await this.extractFonts(zip);
      await this.extractThemes(zip);
      await this.extractRelationships(zip);

      // 生成完整的 CSS
      const css = this.generateCompleteCSS();

      console.log(`✅ 提取完成: ${this.styles.size} 个样式定义`);

      return {
        styles: this.styles,
        documentStyles: this.documentStyles,
        css,
      };
    } catch (error: any) {
      console.error('❌ 样式提取失败:', error);
      throw error;
    }
  }

  /**
   * 解析 word/styles.xml
   */
  private async extractStylesXml(zip: any): Promise<void> {
    try {
      const stylesFile = zip.file('word/styles.xml');
      if (!stylesFile) {
        console.warn('⚠️ 未找到 styles.xml 文件');
        return;
      }

      const stylesXml = await stylesFile.async('text');
      const stylesData = await xml2js.parseStringPromise(stylesXml, {
        explicitArray: false,
        mergeAttrs: true,
      });

      if (stylesData['w:styles'] && stylesData['w:styles']['w:style']) {
        const styles = Array.isArray(stylesData['w:styles']['w:style'])
          ? stylesData['w:styles']['w:style']
          : [stylesData['w:styles']['w:style']];

        for (const style of styles) {
          const styleDefinition = this.parseStyleDefinition(style);
          this.styles.set(styleDefinition.styleId, styleDefinition);
        }
      }

      console.log(`📋 解析了 ${this.styles.size} 个样式定义`);
    } catch (error: any) {
      console.error('❌ 解析 styles.xml 失败:', error);
    }
  }

  /**
   * 解析单个样式定义
   */
  private parseStyleDefinition(style: any): StyleDefinition {
    const styleId = style.styleId || style['w:styleId'] || 'unknown';
    const name = style['w:name']?.val || style.name?.val || styleId;
    const type = this.getStyleType(style.type?.val || style['w:type']?.val);

    const css: Record<string, string> = {};

    // 解析段落属性 (pPr)
    if (style['w:pPr']) {
      this.parseParagraphProperties(style['w:pPr'], css);
    }

    // 解析文字属性 (rPr)
    if (style['w:rPr']) {
      this.parseRunProperties(style['w:rPr'], css);
    }

    // 解析表格属性 (tblPr)
    if (style['w:tblPr']) {
      this.parseTableProperties(style['w:tblPr'], css);
    }

    // 构建 mammoth 映射
    const mammothMapping = this.buildMammothMapping(styleId, name, type, css);

    return {
      styleId,
      name,
      type,
      css,
      mammothMapping,
      basedOn: style['w:basedOn']?.val,
      next: style['w:next']?.val,
      isDefault: style.default === '1' || style['w:default'] === '1',
    };
  }

  /**
   * 解析段落属性
   */
  private parseParagraphProperties(pPr: any, css: Record<string, string>): void {
    // 对齐方式
    if (pPr['w:jc']) {
      const alignmentValue = pPr['w:jc'].val || pPr['w:jc'];
      if (alignmentValue && typeof alignmentValue === 'string') {
        css['text-align'] = this.convertAlignment(alignmentValue);
      }
    }

    // 间距
    if (pPr['w:spacing']) {
      const spacing = pPr['w:spacing'];
      if (spacing.before && !isNaN(parseInt(spacing.before))) {
        const beforePt = this.convertTwipsToPoints(spacing.before);
        if (beforePt >= 0 && beforePt <= 144) {
          // 合理的间距范围
          css['margin-top'] = beforePt + 'pt';
        }
      }
      if (spacing.after && !isNaN(parseInt(spacing.after))) {
        const afterPt = this.convertTwipsToPoints(spacing.after);
        if (afterPt >= 0 && afterPt <= 144) {
          css['margin-bottom'] = afterPt + 'pt';
        }
      }
      if (spacing.line && !isNaN(parseInt(spacing.line))) {
        const lineHeight = this.convertLineSpacing(spacing.line, spacing.lineRule);
        if (lineHeight && lineHeight !== '0') {
          css['line-height'] = lineHeight;
        }
      }
    }

    // 缩进
    if (pPr['w:ind']) {
      const indent = pPr['w:ind'];
      if (indent.left && !isNaN(parseInt(indent.left))) {
        const leftPt = this.convertTwipsToPoints(indent.left);
        if (leftPt >= 0 && leftPt <= 720) {
          // 合理的缩进范围（0-10英寸）
          css['margin-left'] = leftPt + 'pt';
        }
      }
      if (indent.right && !isNaN(parseInt(indent.right))) {
        const rightPt = this.convertTwipsToPoints(indent.right);
        if (rightPt >= 0 && rightPt <= 720) {
          css['margin-right'] = rightPt + 'pt';
        }
      }
      if (indent.firstLine && !isNaN(parseInt(indent.firstLine))) {
        const firstLinePt = this.convertTwipsToPoints(indent.firstLine);
        if (firstLinePt >= -360 && firstLinePt <= 360) {
          // 合理的首行缩进范围
          css['text-indent'] = firstLinePt + 'pt';
        }
      }
      if (indent.hanging && !isNaN(parseInt(indent.hanging))) {
        const hangingPt = this.convertTwipsToPoints(indent.hanging);
        if (hangingPt >= 0 && hangingPt <= 360) {
          css['text-indent'] = '-' + hangingPt + 'pt';
        }
      }
    }

    // 边框
    if (pPr['w:pBdr']) {
      this.parseBorders(pPr['w:pBdr'], css, 'paragraph');
    }

    // 背景色
    if (pPr['w:shd']) {
      const shading = pPr['w:shd'];
      if (shading.fill && shading.fill !== 'auto') {
        css['background-color'] = '#' + shading.fill;
      }
    }
  }

  /**
   * 解析文字属性
   */
  private parseRunProperties(rPr: any, css: Record<string, string>): void {
    // 字体
    if (rPr['w:rFonts']) {
      const fonts = rPr['w:rFonts'];
      const fontFamily = fonts.ascii || fonts.eastAsia || fonts.hAnsi || fonts.cs;
      if (fontFamily && typeof fontFamily === 'string') {
        css['font-family'] = `"${fontFamily}", sans-serif`;
      }
    }

    // 字号
    if (rPr['w:sz']) {
      const sizeValue = rPr['w:sz'].val || rPr['w:sz'];
      if (sizeValue && !isNaN(parseInt(sizeValue))) {
        const size = parseInt(sizeValue) / 2; // Word 中是半点
        if (size > 0 && size <= 72) {
          // 合理的字体大小范围
          css['font-size'] = size + 'pt';
        }
      }
    }

    // 字号（东亚字体）
    if (rPr['w:szCs']) {
      const sizeValue = rPr['w:szCs'].val || rPr['w:szCs'];
      if (sizeValue && !isNaN(parseInt(sizeValue))) {
        const size = parseInt(sizeValue) / 2;
        if (size > 0 && size <= 72) {
          css['font-size'] = size + 'pt';
        }
      }
    }

    // 颜色
    if (rPr['w:color']) {
      const colorElement = rPr['w:color'];
      const colorValue = colorElement.val || colorElement['w:val'] || colorElement;
      const themeColor = colorElement.themeColor || colorElement['w:themeColor'];

      if (themeColor && typeof themeColor === 'string') {
        // 处理主题颜色
        css['color'] = this.convertThemeColor(themeColor);
      } else if (colorValue && typeof colorValue === 'string' && colorValue !== 'auto') {
        // 确保颜色值是有效的十六进制
        const cleanColor = colorValue.replace(/[^0-9A-Fa-f]/g, '').toUpperCase();
        if (cleanColor.length === 6) {
          css['color'] = '#' + cleanColor;
        }
      }
    }

    // 背景色/高亮
    if (rPr['w:highlight']) {
      const highlightValue = rPr['w:highlight'].val || rPr['w:highlight'];
      if (highlightValue && typeof highlightValue === 'string') {
        css['background-color'] = this.convertHighlightColor(highlightValue);
      }
    }

    // 阴影/背景色（更全面的处理）
    if (rPr['w:shd']) {
      const shading = rPr['w:shd'];

      // 处理填充颜色
      const fillColor = shading.fill || shading['w:fill'];
      if (
        fillColor &&
        typeof fillColor === 'string' &&
        fillColor !== 'auto' &&
        fillColor !== '000000'
      ) {
        const cleanFill = fillColor.replace(/[^0-9A-Fa-f]/g, '').toUpperCase();
        if (cleanFill.length === 6) {
          css['background-color'] = '#' + cleanFill;
        }
      }

      // 处理前景色（如果没有其他颜色定义）
      if (!css['color']) {
        const fgColor = shading.color || shading['w:color'];
        if (fgColor && typeof fgColor === 'string' && fgColor !== 'auto') {
          const cleanColor = fgColor.replace(/[^0-9A-Fa-f]/g, '').toUpperCase();
          if (cleanColor.length === 6) {
            css['color'] = '#' + cleanColor;
          }
        }
      }
    }

    // 文字修饰
    if (rPr['w:b']) css['font-weight'] = 'bold';
    if (rPr['w:i']) css['font-style'] = 'italic';
    if (rPr['w:u']) {
      const underline = rPr['w:u'].val || rPr['w:u'];
      css['text-decoration'] = underline === 'none' ? 'none' : 'underline';
    }
    if (rPr['w:strike']) css['text-decoration'] = 'line-through';
    if (rPr['w:dstrike']) css['text-decoration'] = 'line-through';

    // 上下标
    if (rPr['w:vertAlign']) {
      const vertAlign = rPr['w:vertAlign'].val || rPr['w:vertAlign'];
      if (vertAlign === 'superscript') {
        css['vertical-align'] = 'super';
        css['font-size'] = '0.8em';
      } else if (vertAlign === 'subscript') {
        css['vertical-align'] = 'sub';
        css['font-size'] = '0.8em';
      }
    }

    // 字符间距
    if (rPr['w:spacing']) {
      const spacingValue = rPr['w:spacing'].val || rPr['w:spacing'];
      if (spacingValue && !isNaN(parseInt(spacingValue))) {
        const spacing = parseInt(spacingValue);
        css['letter-spacing'] = spacing / 20 + 'pt'; // 转换为点
      }
    }
  }

  /**
   * 解析表格属性
   */
  private parseTableProperties(tblPr: any, css: Record<string, string>): void {
    // 表格宽度
    if (tblPr['w:tblW']) {
      const width = tblPr['w:tblW'];
      if (width.type === 'pct' && width.w && !isNaN(parseInt(width.w))) {
        css['width'] = parseInt(width.w) / 50 + '%'; // Word 百分比是 50 的倍数
      } else if (width.type === 'dxa' && width.w && !isNaN(parseInt(width.w))) {
        css['width'] = this.convertTwipsToPoints(width.w) + 'pt';
      }
    }

    // 表格边框
    if (tblPr['w:tblBorders']) {
      this.parseBorders(tblPr['w:tblBorders'], css, 'table');
    }

    // 表格对齐
    if (tblPr['w:jc']) {
      const alignmentValue = tblPr['w:jc'].val || tblPr['w:jc'];
      if (alignmentValue && typeof alignmentValue === 'string') {
        css['margin-left'] = alignmentValue === 'center' ? 'auto' : '0';
        css['margin-right'] = alignmentValue === 'center' ? 'auto' : '0';
      }
    }

    // 表格间距
    if (tblPr['w:tblCellSpacing']) {
      const spacingValue = tblPr['w:tblCellSpacing'].w;
      if (spacingValue && !isNaN(parseInt(spacingValue))) {
        css['border-spacing'] = this.convertTwipsToPoints(spacingValue) + 'pt';
      }
    }
  }

  /**
   * 解析边框
   */
  private parseBorders(borders: any, css: Record<string, string>, context: string): void {
    const borderSides = ['top', 'right', 'bottom', 'left'];
    const wordSides = ['w:top', 'w:right', 'w:bottom', 'w:left'];

    for (let i = 0; i < borderSides.length; i++) {
      const side = borderSides[i];
      const wordSide = wordSides[i];

      if (borders[wordSide]) {
        const border = borders[wordSide];
        const width =
          border.sz && !isNaN(parseInt(border.sz)) ? parseInt(border.sz) / 8 + 'pt' : '1pt';
        const style = this.convertBorderStyle(border.val || 'single');
        let color = '#000000';

        if (border.color && typeof border.color === 'string' && border.color !== 'auto') {
          const cleanColor = border.color.replace(/[^0-9A-Fa-f]/g, '');
          if (cleanColor.length === 6) {
            color = '#' + cleanColor;
          }
        }

        css[`border-${side}`] = `${width} ${style} ${color}`;
      }
    }
  }

  /**
   * 解析 word/document.xml 获取内联样式
   */
  private async extractDocumentXml(zip: any): Promise<void> {
    try {
      const documentFile = zip.file('word/document.xml');
      if (!documentFile) {
        console.warn('⚠️ 未找到 document.xml 文件');
        return;
      }

      const documentXml = await documentFile.async('text');
      const documentData = await xml2js.parseStringPromise(documentXml, {
        explicitArray: false,
        mergeAttrs: true,
      });

      // 遍历文档结构，提取内联样式
      this.extractInlineStyles(documentData['w:document']['w:body']);

      console.log(`📄 提取了 ${this.documentStyles.length} 个文档样式`);
    } catch (error: any) {
      console.error('❌ 解析 document.xml 失败:', error);
    }
  }

  /**
   * 提取内联样式
   */
  private extractInlineStyles(body: any): void {
    if (!body) return;

    // 处理段落
    if (body['w:p']) {
      const paragraphs = Array.isArray(body['w:p']) ? body['w:p'] : [body['w:p']];
      for (const paragraph of paragraphs) {
        this.extractParagraphStyles(paragraph);
      }
    }

    // 处理表格
    if (body['w:tbl']) {
      const tables = Array.isArray(body['w:tbl']) ? body['w:tbl'] : [body['w:tbl']];
      for (const table of tables) {
        this.extractTableStyles(table);
      }
    }

    // 递归处理其他元素
    for (const [key, value] of Object.entries(body)) {
      if (key !== 'w:p' && key !== 'w:tbl' && typeof value === 'object' && value !== null) {
        if (Array.isArray(value)) {
          for (const item of value) {
            if (typeof item === 'object' && item !== null) {
              this.extractInlineStyles(item);
            }
          }
        } else {
          this.extractInlineStyles(value);
        }
      }
    }
  }

  /**
   * 提取段落样式
   */
  private extractParagraphStyles(paragraph: any): void {
    const css: Record<string, string> = {};
    let styleId: string | undefined;

    // 段落属性
    if (paragraph['w:pPr']) {
      const pPr = paragraph['w:pPr'];

      // 样式引用
      if (pPr['w:pStyle']) {
        styleId = pPr['w:pStyle'].val || pPr['w:pStyle'];
      }

      // 直接格式化
      this.parseParagraphProperties(pPr, css);

      // 处理段落级别的文字属性（rPr）
      if (pPr['w:rPr']) {
        this.parseRunProperties(pPr['w:rPr'], css);
      }
    }

    // 只有当CSS不为空时才添加
    if (Object.keys(css).length > 0 || styleId) {
      this.documentStyles.push({
        elementType: 'paragraph',
        styleId,
        css,
      });
    }

    // 处理段落中的文字
    if (paragraph['w:r']) {
      const runs = Array.isArray(paragraph['w:r']) ? paragraph['w:r'] : [paragraph['w:r']];
      for (const run of runs) {
        this.extractRunStyles(run);
      }
    }
  }

  /**
   * 提取文字样式
   */
  private extractRunStyles(run: any): void {
    const css: Record<string, string> = {};
    let styleId: string | undefined;

    if (run['w:rPr']) {
      const rPr = run['w:rPr'];

      // 样式引用
      if (rPr['w:rStyle']) {
        styleId = rPr['w:rStyle'].val || rPr['w:rStyle'];
      }

      // 直接格式化
      this.parseRunProperties(rPr, css);
    }

    // 只有当CSS不为空时才添加
    if (Object.keys(css).length > 0 || styleId) {
      this.documentStyles.push({
        elementType: 'run',
        styleId,
        css,
      });
    }
  }

  /**
   * 提取表格样式
   */
  private extractTableStyles(table: any): void {
    const css: Record<string, string> = {};
    let styleId: string | undefined;

    if (table['w:tblPr']) {
      const tblPr = table['w:tblPr'];

      // 样式引用
      if (tblPr['w:tblStyle']) {
        styleId = tblPr['w:tblStyle'].val || tblPr['w:tblStyle'];
      }

      // 直接格式化
      this.parseTableProperties(tblPr, css);
    }

    this.documentStyles.push({
      elementType: 'table',
      styleId,
      css,
    });
  }

  /**
   * 提取字体信息
   */
  private async extractFonts(zip: any): Promise<void> {
    try {
      const fontTableFile = zip.file('word/fontTable.xml');
      if (!fontTableFile) return;

      const fontTableXml = await fontTableFile.async('text');
      const fontData = await xml2js.parseStringPromise(fontTableXml);

      // 解析字体表
      if (fontData['w:fonts'] && fontData['w:fonts']['w:font']) {
        const fonts = Array.isArray(fontData['w:fonts']['w:font'])
          ? fontData['w:fonts']['w:font']
          : [fontData['w:fonts']['w:font']];

        for (const font of fonts) {
          const name = font.$.name || font.$['w:name'];
          if (name) {
            this.fonts.set(name, font);
          }
        }
      }
    } catch (error: any) {
      console.warn('⚠️ 字体表解析失败:', error.message);
    }
  }

  /**
   * 提取主题信息
   */
  private async extractThemes(zip: any): Promise<void> {
    try {
      const themeFile = zip.file('word/theme/theme1.xml');
      if (!themeFile) return;

      const themeXml = await themeFile.async('text');
      const themeData = await xml2js.parseStringPromise(themeXml);

      // 解析主题颜色和字体
      // 这里可以根据需要扩展
    } catch (error: any) {
      console.warn('⚠️ 主题解析失败:', error.message);
    }
  }

  /**
   * 提取关系信息
   */
  private async extractRelationships(zip: any): Promise<void> {
    try {
      const relsFile = zip.file('word/_rels/document.xml.rels');
      if (!relsFile) return;

      const relsXml = await relsFile.async('text');
      const relsData = await xml2js.parseStringPromise(relsXml);

      if (relsData.Relationships && relsData.Relationships.Relationship) {
        const relationships = Array.isArray(relsData.Relationships.Relationship)
          ? relsData.Relationships.Relationship
          : [relsData.Relationships.Relationship];

        for (const rel of relationships) {
          this.relationships.set(rel.$.Id, rel.$.Target);
        }
      }
    } catch (error: any) {
      console.warn('⚠️ 关系文件解析失败:', error.message);
    }
  }

  /**
   * 生成完整的 CSS
   */
  private generateCompleteCSS(): string {
    const cssRules: string[] = [];

    // 基础样式
    cssRules.push(`
/* 基础文档样式 */
body {
  font-family: Calibri, sans-serif;
  font-size: 11pt;
  line-height: 1.15;
  margin: 0;
  padding: 20pt;
  color: #000000;
}

/* 段落基础样式 */
p {
  margin: 0 0 8pt 0;
}

/* 表格基础样式 */
table {
  border-collapse: collapse;
  width: 100%;
}

td, th {
  border: 0.5pt solid #000000;
  padding: 4pt;
  vertical-align: top;
}
`);

    // 样式定义的 CSS
    for (const [styleId, style] of this.styles) {
      if (Object.keys(style.css).length > 0) {
        // 确保类名不以数字开头，CSS规范要求类名必须以字母、下划线或连字符开头
        const sanitizedId = /^\d/.test(styleId) ? `style-${styleId}` : styleId;
        const selector = `.${sanitizedId}`;
        const cssText = this.cssObjectToString(style.css);
        cssRules.push(`${selector} {\n${cssText}\n}`);
      }
    }

    return cssRules.join('\n\n');
  }

  /**
   * 构建 mammoth 样式映射
   */
  private buildMammothMapping(
    styleId: string,
    name: string,
    type: string,
    css: Record<string, string>
  ) {
    let element = 'span';
    let selector = `r[style-name='${name}']`;

    // 确保类名不以数字开头，CSS规范要求类名必须以字母、下划线或连字符开头
    const sanitizedId = /^\d/.test(styleId) ? `style-${styleId}` : styleId;

    if (type === 'paragraph') {
      if (name.toLowerCase().includes('heading')) {
        const level = this.extractHeadingLevel(name);
        element = `h${level}`;
      } else {
        element = 'p';
      }
      selector = `p[style-name='${name}']`;
    } else if (type === 'table') {
      element = 'table';
      selector = `table[style-name='${name}']`;
    }

    return {
      selector,
      element: `${element}.${sanitizedId}:fresh`,
      className: sanitizedId,
    };
  }

  // 辅助方法
  private getStyleType(type: string): 'paragraph' | 'character' | 'table' | 'numbering' {
    switch (type) {
      case 'paragraph':
        return 'paragraph';
      case 'character':
        return 'character';
      case 'table':
        return 'table';
      case 'numbering':
        return 'numbering';
      default:
        return 'paragraph';
    }
  }

  private convertAlignment(alignment: string): string {
    switch (alignment) {
      case 'center':
        return 'center';
      case 'right':
        return 'right';
      case 'justify':
        return 'justify';
      case 'both':
        return 'justify';
      default:
        return 'left';
    }
  }

  private convertTwipsToPoints(twips: string | number): number {
    return Math.round((parseInt(twips.toString()) / 20) * 100) / 100;
  }

  private convertLineSpacing(line: string, lineRule?: string): string {
    const lineValue = parseInt(line);
    if (lineRule === 'exact') {
      return this.convertTwipsToPoints(lineValue) + 'pt';
    } else {
      return (lineValue / 240).toFixed(2); // 240 = 单倍行距
    }
  }

  private convertHighlightColor(color: string): string {
    const colorMap: Record<string, string> = {
      yellow: '#FFFF00',
      green: '#00FF00',
      cyan: '#00FFFF',
      magenta: '#FF00FF',
      blue: '#0000FF',
      red: '#FF0000',
      darkBlue: '#000080',
      darkCyan: '#008080',
      darkGreen: '#008000',
      darkMagenta: '#800080',
      darkRed: '#800000',
      darkYellow: '#808000',
      darkGray: '#808080',
      lightGray: '#C0C0C0',
      black: '#000000',
    };
    return colorMap[color] || '#FFFF00';
  }

  /**
   * 转换主题颜色
   */
  private convertThemeColor(themeColor: string): string {
    const themeColorMap: Record<string, string> = {
      hyperlink: '#0563C1',
      followedHyperlink: '#954F72',
      accent1: '#4F81BD',
      accent2: '#F79646',
      accent3: '#9BBB59',
      accent4: '#8064A2',
      accent5: '#4BACC6',
      accent6: '#F596AA',
      dark1: '#000000',
      dark2: '#1F497D',
      light1: '#FFFFFF',
      light2: '#EEECE1',
      text1: '#000000',
      text2: '#1F497D',
      background1: '#FFFFFF',
      background2: '#EEECE1',
    };

    return themeColorMap[themeColor] || '#000000';
  }

  private convertBorderStyle(style: string): string {
    switch (style) {
      case 'single':
        return 'solid';
      case 'double':
        return 'double';
      case 'dotted':
        return 'dotted';
      case 'dashed':
        return 'dashed';
      case 'thick':
        return 'solid';
      default:
        return 'solid';
    }
  }

  private extractHeadingLevel(styleName: string): number {
    const match = styleName.match(/(\d+)/);
    return match ? parseInt(match[1]) : 1;
  }

  private cssObjectToString(css: Record<string, string>): string {
    return Object.entries(css)
      .map(([property, value]) => `  ${property}: ${value};`)
      .join('\n');
  }
}

export { StyleDefinition, DocumentStyle };
