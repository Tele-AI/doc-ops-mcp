/**
 * CSSGenerator - CSS 规则生成器
 * 将解析的样式信息转换为 CSS 规则，生成基础样式和自定义样式
 * 处理样式优先级和继承，优化 CSS 输出
 */

import { StyleDefinition, DocumentStyle } from './styleExtractor';

interface CSSGeneratorOptions {
  minify?: boolean;
  includeComments?: boolean;
  generateResponsive?: boolean;
  generatePrintStyles?: boolean;
  customPrefix?: string;
  fontFallbacks?: { [key: string]: string[] };
}

interface GeneratedCSS {
  baseStyles: string;
  customStyles: string;
  responsiveStyles: string;
  printStyles: string;
  complete: string;
}

export class CSSGenerator {
  private styles: Map<string, StyleDefinition> = new Map();
  private documentStyles: DocumentStyle[] = [];
  private options: CSSGeneratorOptions;
  private cssRules: string[] = [];
  private usedSelectors: Set<string> = new Set();

  constructor(options: CSSGeneratorOptions = {}) {
    this.options = {
      minify: false,
      includeComments: true,
      generateResponsive: true,
      generatePrintStyles: true,
      customPrefix: '',
      fontFallbacks: {
        Calibri: ['Calibri', 'Arial', 'sans-serif'],
        'Times New Roman': ['Times New Roman', 'Times', 'serif'],
        Arial: ['Arial', 'Helvetica', 'sans-serif'],
        'Courier New': ['Courier New', 'Courier', 'monospace'],
        'Microsoft YaHei': ['Microsoft YaHei', 'SimHei', 'Arial', 'sans-serif'],
        SimSun: ['SimSun', 'NSimSun', 'serif'],
        SimHei: ['SimHei', 'Microsoft YaHei', 'Arial', 'sans-serif'],
      },
      ...options,
    };
  }

  /**
   * 设置样式数据
   */
  setStyles(styles: Map<string, StyleDefinition>, documentStyles: DocumentStyle[] = []): void {
    this.styles = styles;
    this.documentStyles = documentStyles;
    this.cssRules = [];
    this.usedSelectors.clear();
  }

  /**
   * 生成完整的 CSS
   */
  generateCSS(): GeneratedCSS {
    console.log('🎨 开始生成 CSS 规则...');

    const baseStyles = this.generateBaseStyles();
    const customStyles = this.generateCustomStyles();
    const responsiveStyles = this.options.generateResponsive ? this.generateResponsiveStyles() : '';
    const printStyles = this.options.generatePrintStyles ? this.generatePrintStyles() : '';

    const complete = this.combineStyles(baseStyles, customStyles, responsiveStyles, printStyles);

    console.log(`✅ CSS 生成完成，共 ${this.cssRules.length} 条规则`);

    return {
      baseStyles,
      customStyles,
      responsiveStyles,
      printStyles,
      complete,
    };
  }

  /**
   * 生成基础样式
   */
  private generateBaseStyles(): string {
    const rules: string[] = [];

    if (this.options.includeComments) {
      rules.push('/* 基础文档样式 */');
    }

    // 文档基础样式
    rules.push(
      this.createRule('body', {
        'font-family': this.getFontFamily('Calibri'),
        'font-size': '11pt',
        'line-height': '1.15',
        margin: '0',
        padding: '20pt',
        color: '#000000',
        'background-color': '#ffffff',
      })
    );

    // 段落基础样式
    rules.push(
      this.createRule('p', {
        margin: '0 0 8pt 0',
        padding: '0',
      })
    );

    // 标题基础样式
    for (let i = 1; i <= 6; i++) {
      const fontSize = this.calculateHeadingSize(i);
      rules.push(
        this.createRule(`h${i}`, {
          'font-size': fontSize,
          'font-weight': 'bold',
          margin: '12pt 0 6pt 0',
          padding: '0',
        })
      );
    }

    // 列表基础样式
    rules.push(
      this.createRule('ul, ol', {
        margin: '0 0 8pt 0',
        'padding-left': '24pt',
      })
    );

    rules.push(
      this.createRule('li', {
        margin: '0 0 4pt 0',
      })
    );

    // 表格基础样式
    rules.push(
      this.createRule('table', {
        'border-collapse': 'collapse',
        width: '100%',
        margin: '12pt 0',
      })
    );

    rules.push(
      this.createRule('td, th', {
        border: '0.5pt solid #000000',
        padding: '4pt 8pt',
        'vertical-align': 'top',
      })
    );

    rules.push(
      this.createRule('th', {
        'background-color': '#f0f0f0',
        'font-weight': 'bold',
      })
    );

    // 文本格式基础样式
    rules.push(
      this.createRule('strong, b', {
        'font-weight': 'bold',
      })
    );

    rules.push(
      this.createRule('em, i', {
        'font-style': 'italic',
      })
    );

    rules.push(
      this.createRule('u', {
        'text-decoration': 'underline',
      })
    );

    rules.push(
      this.createRule('del, strike', {
        'text-decoration': 'line-through',
      })
    );

    rules.push(
      this.createRule('sup', {
        'vertical-align': 'super',
        'font-size': '0.8em',
      })
    );

    rules.push(
      this.createRule('sub', {
        'vertical-align': 'sub',
        'font-size': '0.8em',
      })
    );

    // 链接样式
    rules.push(
      this.createRule('a', {
        color: '#0066cc',
        'text-decoration': 'underline',
      })
    );

    rules.push(
      this.createRule('a:hover', {
        color: '#004499',
      })
    );

    return this.joinRules(rules);
  }

  /**
   * 生成自定义样式
   */
  private generateCustomStyles(): string {
    const rules: string[] = [];

    if (this.options.includeComments) {
      rules.push('/* 自定义样式 */');
    }

    // 处理样式定义
    for (const [styleId, style] of this.styles) {
      const selector = this.generateSelector(styleId, style);
      if (selector && Object.keys(style.css).length > 0) {
        const processedCSS = this.processStyleCSS(style.css);
        rules.push(this.createRule(selector, processedCSS));
        this.usedSelectors.add(selector);
      }
    }

    // 处理文档内联样式
    const inlineStyles = this.generateInlineStyles();
    if (inlineStyles) {
      rules.push(inlineStyles);
    }

    return this.joinRules(rules);
  }

  /**
   * 生成响应式样式
   */
  private generateResponsiveStyles(): string {
    const rules: string[] = [];

    if (this.options.includeComments) {
      rules.push('/* 响应式样式 */');
    }

    // 平板样式
    rules.push('@media screen and (max-width: 1024px) {');
    rules.push(
      this.createRule('  body', {
        padding: '16pt',
      })
    );
    rules.push(
      this.createRule('  table', {
        'font-size': '10pt',
      })
    );
    rules.push('}');

    // 手机样式
    rules.push('@media screen and (max-width: 768px) {');
    rules.push(
      this.createRule('  body', {
        padding: '12pt',
        'font-size': '10pt',
      })
    );

    rules.push(
      this.createRule('  table', {
        'font-size': '9pt',
        'overflow-x': 'auto',
        display: 'block',
        'white-space': 'nowrap',
      })
    );

    rules.push(
      this.createRule('  td, th', {
        padding: '2pt 4pt',
      })
    );

    for (let i = 1; i <= 6; i++) {
      const mobileSize = this.calculateMobileHeadingSize(i);
      rules.push(
        this.createRule(`  h${i}`, {
          'font-size': mobileSize,
        })
      );
    }

    rules.push('}');

    return this.joinRules(rules);
  }

  /**
   * 生成打印样式
   */
  private generatePrintStyles(): string {
    const rules: string[] = [];

    if (this.options.includeComments) {
      rules.push('/* 打印样式 */');
    }

    rules.push('@media print {');

    rules.push(
      this.createRule('  body', {
        'font-size': '10pt',
        padding: '0',
        margin: '0',
      })
    );

    rules.push(
      this.createRule('  table', {
        'page-break-inside': 'avoid',
      })
    );

    rules.push(
      this.createRule('  h1, h2, h3, h4, h5, h6', {
        'page-break-after': 'avoid',
      })
    );

    rules.push(
      this.createRule('  tr', {
        'page-break-inside': 'avoid',
      })
    );

    rules.push(
      this.createRule('  a', {
        color: '#000000',
        'text-decoration': 'none',
      })
    );

    rules.push(
      this.createRule('  a:after', {
        content: '" (" attr(href) ")"',
        'font-size': '0.8em',
      })
    );

    rules.push('}');

    return this.joinRules(rules);
  }

  /**
   * 生成选择器
   */
  private generateSelector(styleId: string, style: StyleDefinition): string {
    const prefix = this.options.customPrefix || '';
    // 确保类名不以数字开头，CSS规范要求类名必须以字母、下划线或连字符开头
    const sanitizedId = /^\d/.test(styleId) ? `style-${styleId}` : styleId;
    const className = `${prefix}${sanitizedId}`;

    switch (style.type) {
      case 'paragraph':
        if (style.name.toLowerCase().includes('heading')) {
          const level = this.extractHeadingLevel(style.name);
          return `h${level}.${className}, .${className}`;
        } else if (style.name.toLowerCase().includes('list')) {
          return `li.${className}, .${className}`;
        } else {
          return `p.${className}, .${className}`;
        }

      case 'character':
        return `span.${className}, .${className}`;

      case 'table':
        return `table.${className}, .${className}`;

      default:
        return `.${className}`;
    }
  }

  /**
   * 处理样式 CSS
   */
  private processStyleCSS(css: Record<string, string>): Record<string, string> {
    const processed: Record<string, string> = {};

    for (const [property, value] of Object.entries(css)) {
      // 跳过无效值
      if (
        !value ||
        value === 'undefined' ||
        value === 'null' ||
        value.includes('[object Object]') ||
        value.includes('NaN')
      ) {
        continue;
      }

      let processedValue = value;

      // 处理字体
      if (property === 'font-family' && typeof value === 'string') {
        processedValue = this.getFontFamily(value);
      }

      // 处理颜色
      if (property.includes('color') && typeof value === 'string') {
        if (!value.startsWith('#') && !value.startsWith('rgb')) {
          processedValue = this.normalizeColor(value);
        } else {
          processedValue = value;
        }
      }

      // 处理尺寸
      if (this.isSizeProperty(property) && typeof value === 'string') {
        processedValue = this.normalizeSize(value);
      }

      // 最终验证：确保值是有效的
      if (
        processedValue &&
        typeof processedValue === 'string' &&
        !processedValue.includes('NaN') &&
        !processedValue.includes('[object Object]')
      ) {
        processed[property] = processedValue;
      }
    }

    return processed;
  }

  /**
   * 生成内联样式
   */
  private generateInlineStyles(): string {
    const rules: string[] = [];
    const inlineStyleMap: Map<string, Record<string, string>> = new Map();

    // 收集内联样式
    for (const docStyle of this.documentStyles) {
      if (Object.keys(docStyle.css).length > 0) {
        const key = `${docStyle.elementType}-${JSON.stringify(docStyle.css)}`;
        if (!inlineStyleMap.has(key)) {
          inlineStyleMap.set(key, docStyle.css);
        }
      }
    }

    // 生成内联样式规则
    let counter = 1;
    for (const [key, css] of inlineStyleMap) {
      const className = `inline-style-${counter++}`;
      const processedCSS = this.processStyleCSS(css);
      rules.push(this.createRule(`.${className}`, processedCSS));
    }

    return this.joinRules(rules);
  }

  /**
   * 创建 CSS 规则
   */
  private createRule(selector: string, properties: Record<string, string>): string {
    const props = Object.entries(properties)
      .map(([prop, value]) => `  ${prop}: ${value};`)
      .join(this.options.minify ? '' : '\n');

    if (this.options.minify) {
      return `${selector}{${props}}`;
    } else {
      return `${selector} {\n${props}\n}`;
    }
  }

  /**
   * 合并样式
   */
  private combineStyles(...styles: string[]): string {
    const separator = this.options.minify ? '' : '\n\n';
    return styles.filter(style => style.trim()).join(separator);
  }

  /**
   * 连接规则
   */
  private joinRules(rules: string[]): string {
    const separator = this.options.minify ? '' : '\n\n';
    return rules.filter(rule => rule.trim()).join(separator);
  }

  /**
   * 获取字体族
   */
  private getFontFamily(fontName: string): string {
    const cleanName = fontName.replace(/["']/g, '');
    const fallbacks = this.options.fontFallbacks?.[cleanName] || [cleanName, 'sans-serif'];
    return fallbacks.map(font => (font.includes(' ') ? `"${font}"` : font)).join(', ');
  }

  /**
   * 标准化颜色
   */
  private normalizeColor(color: string): string {
    // 验证输入
    if (!color || typeof color !== 'string') {
      return '#000000';
    }

    // 如果已经是有效的十六进制颜色，直接返回
    if (/^#[0-9A-Fa-f]{6}$/.test(color)) {
      return color;
    }

    // 如果是颜色名称，转换为十六进制
    const colorMap: Record<string, string> = {
      black: '#000000',
      white: '#ffffff',
      red: '#ff0000',
      green: '#008000',
      blue: '#0000ff',
      yellow: '#ffff00',
      cyan: '#00ffff',
      magenta: '#ff00ff',
      gray: '#808080',
      grey: '#808080',
      auto: '#000000',
    };

    const lowerColor = color.toLowerCase();
    if (colorMap[lowerColor]) {
      return colorMap[lowerColor];
    }

    // 尝试清理并验证十六进制颜色
    const cleanColor = color.replace(/[^0-9A-Fa-f]/g, '');
    if (cleanColor.length === 6) {
      return '#' + cleanColor;
    }

    // 如果都不匹配，返回默认黑色
    return '#000000';
  }

  /**
   * 标准化尺寸
   */
  private normalizeSize(size: string): string {
    // 验证输入
    if (!size || typeof size !== 'string') {
      return '0pt';
    }

    // 如果已经有有效的单位，直接返回
    if (/^\d+(\.\d+)?(pt|px|em|rem|%|in|cm|mm)$/.test(size)) {
      return size;
    }

    // 如果是纯数字，添加pt单位
    if (/^\d+(\.\d+)?$/.test(size)) {
      return size + 'pt';
    }

    // 尝试提取数字部分
    const match = size.match(/^(\d+(\.\d+)?)/);
    if (match) {
      return match[1] + 'pt';
    }

    // 如果都不匹配，返回默认值
    return '0pt';
  }

  /**
   * 判断是否为尺寸属性
   */
  private isSizeProperty(property: string): boolean {
    const sizeProperties = [
      'font-size',
      'width',
      'height',
      'margin',
      'margin-top',
      'margin-right',
      'margin-bottom',
      'margin-left',
      'padding',
      'padding-top',
      'padding-right',
      'padding-bottom',
      'padding-left',
      'border-width',
      'text-indent',
      'line-height',
      'letter-spacing',
      'word-spacing',
    ];
    return sizeProperties.includes(property);
  }

  /**
   * 计算标题尺寸
   */
  private calculateHeadingSize(level: number): string {
    const baseSizes = ['18pt', '16pt', '14pt', '12pt', '11pt', '10pt'];
    return baseSizes[level - 1] || '10pt';
  }

  /**
   * 计算移动端标题尺寸
   */
  private calculateMobileHeadingSize(level: number): string {
    const mobileSizes = ['16pt', '14pt', '12pt', '11pt', '10pt', '9pt'];
    return mobileSizes[level - 1] || '9pt';
  }

  /**
   * 提取标题级别
   */
  private extractHeadingLevel(styleName: string): number {
    const match = styleName.match(/(\d+)/);
    return match ? Math.min(parseInt(match[1]), 6) : 1;
  }

  /**
   * 优化 CSS
   */
  optimizeCSS(css: string): string {
    let optimized = css;

    // 移除重复的规则
    optimized = this.removeDuplicateRules(optimized);

    // 合并相似的选择器
    optimized = this.mergeSimilarSelectors(optimized);

    // 移除未使用的规则
    optimized = this.removeUnusedRules(optimized);

    return optimized;
  }

  /**
   * 移除重复规则
   */
  private removeDuplicateRules(css: string): string {
    const rules = css.split('}').filter(rule => rule.trim());
    const uniqueRules = new Set();
    const result: string[] = [];

    for (const rule of rules) {
      const normalized = rule.trim().replace(/\s+/g, ' ');
      if (!uniqueRules.has(normalized)) {
        uniqueRules.add(normalized);
        result.push(rule + '}');
      }
    }

    return result.join('\n');
  }

  /**
   * 合并相似选择器
   */
  private mergeSimilarSelectors(css: string): string {
    // 这里可以实现更复杂的选择器合并逻辑
    return css;
  }

  /**
   * 移除未使用的规则
   */
  private removeUnusedRules(css: string): string {
    // 这里可以实现未使用规则的检测和移除
    return css;
  }

  /**
   * 获取生成统计信息
   */
  getStats(): {
    totalRules: number;
    usedSelectors: number;
    baseStylesSize: number;
    customStylesSize: number;
  } {
    return {
      totalRules: this.cssRules.length,
      usedSelectors: this.usedSelectors.size,
      baseStylesSize: this.generateBaseStyles().length,
      customStylesSize: this.generateCustomStyles().length,
    };
  }
}

export { CSSGeneratorOptions, GeneratedCSS };
