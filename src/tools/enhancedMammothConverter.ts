/**
 * 结合 mammoth.js + 高级样式解析的转换器
 * 基于样式还原解决方案.txt 中的增强型实现
 */

const mammoth = require('mammoth');
const fs = require('fs').promises;
const JSZip = require('jszip');
const xml2js = require('xml2js');

/**
 * 结合 mammoth.js 和高级样式解析的转换器
 */
class EnhancedMammothConverter {
  private styles: Map<string, any>;
  private relationships: Map<string, string>;
  private media: Map<string, any>;
  private cssRules: string[];
  private documentStructure: any;
  private runStyleMap: Map<string, any>; // 文字样式映射
  private paragraphStyleMap: Map<string, any>; // 段落样式映射

  constructor() {
    this.styles = new Map();
    this.relationships = new Map();
    this.media = new Map();
    this.cssRules = [];
    this.documentStructure = null;
    this.runStyleMap = new Map(); // 文字样式映射
    this.paragraphStyleMap = new Map(); // 段落样式映射
  }

  /**
   * 主转换函数 - 基于测试脚本验证的完整样式保留方案
   */
  async convertDocxToHtml(inputPath, options = {}) {
    try {
      console.log('🚀 开始增强型 mammoth 转换（完整样式保留版）...');

      // 步骤1: 深度解析 DOCX 文件结构（基于测试脚本验证的方法）
      await this.deepParseDocxStructure(inputPath);

      // 步骤2: 解析样式信息
      await this.parseDocxStyles(inputPath);

      // 步骤3: 使用完整的样式映射配置转换
      const result = await this.convertWithCompleteStyleMapping(inputPath, options);

      console.log('✅ 增强型转换完成');
      return result;
    } catch (error: any) {
      console.error('❌ 转换失败:', error);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * 解析 DOCX 样式信息
   */
  async parseDocxStyles(inputPath) {
    const docxBuffer = await fs.readFile(inputPath);
    const zip = await JSZip.loadAsync(docxBuffer);

    // 解析样式
    await this.extractStyles(zip);

    // 解析关系和媒体
    await this.extractRelationships(zip);
    await this.extractMedia(zip);

    // 解析文档结构以获取样式引用
    await this.extractDocumentStructure(zip);

    console.log(`📊 解析了 ${this.styles.size} 个样式`);
  }

  /**
   * 提取样式信息
   */
  async extractStyles(zip) {
    try {
      console.log('🔍 开始解析样式文件...');
      const stylesXml = await zip.file('word/styles.xml').async('text');
      console.log('📄 样式XML长度:', stylesXml.length);

      const stylesData = await xml2js.parseStringPromise(stylesXml);
      console.log('📊 解析的样式数据结构:', Object.keys(stylesData));

      if (stylesData['w:styles'] && stylesData['w:styles']['w:style']) {
        const styles = stylesData['w:styles']['w:style'];
        console.log(`🎨 发现 ${styles.length} 个样式定义`);

        for (const style of styles) {
          const styleId = style.$.styleId;
          const styleName = style['w:name'] ? style['w:name'][0].$.val : styleId;
          const styleType = style.$.type;

          console.log(`🔧 处理样式: ${styleName} (${styleId}) - 类型: ${styleType}`);

          const parsedStyle = this.parseStyleDefinition(style);
          this.styles.set(styleId, parsedStyle);

          // 生成 CSS 规则
          if (Object.keys(parsedStyle.css).length > 0) {
            this.generateCssRule(`.${styleId}`, parsedStyle.css);
            console.log(`✅ 生成CSS规则: .${styleId}`, parsedStyle.css);
          } else {
            console.log(`⚠️ 样式 ${styleId} 没有CSS属性`);
          }

          // 为 mammoth 构建样式映射
          this.buildMammothStyleMap(styleId, parsedStyle);
        }

        console.log(
          `✅ 样式解析完成: ${this.styles.size} 个样式, ${this.cssRules.length} 个CSS规则`
        );
        console.log(
          `📋 段落样式: ${this.paragraphStyleMap.size}, 文字样式: ${this.runStyleMap.size}`
        );
      } else {
        console.log('⚠️ 未找到样式定义');
      }
    } catch (error: any) {
      console.error('❌ 样式文件解析失败:', error.message);
      console.error('🔍 错误详情:', error);
      // 不要静默失败，确保调用者知道样式解析失败
      throw new Error(`样式解析失败: ${error.message}`);
    }
  }

  /**
   * 解析样式定义（与之前相同，但更简洁）
   */
  parseStyleDefinition(style) {
    // 安全地提取样式属性，提供默认值
    const styleName =
      style['w:name'] && style['w:name'][0] && style['w:name'][0].$
        ? style['w:name'][0].$.val
        : 'unknown-style';
    const styleType = style.$ && style.$.type ? style.$.type : 'paragraph';
    const styleId = style.$ && style.$.styleId ? style.$.styleId : `style-${Date.now()}`;

    const styleObj = {
      name: styleName,
      type: styleType,
      css: {},
      mammothMap: {},
    };

    console.log(`🎯 解析样式定义: ${styleName} (${styleId}) - 类型: ${styleType}`);

    // 解析段落属性
    if (style['w:pPr']) {
      console.log(`📝 解析段落属性: ${styleName}`);
      this.parseParagraphProperties(style['w:pPr'][0], styleObj);
    }

    // 解析文字属性
    if (style['w:rPr']) {
      console.log(`✏️ 解析文字属性: ${styleName}`);
      this.parseRunProperties(style['w:rPr'][0], styleObj);
    }

    console.log(`🔍 样式解析结果: ${styleName}`, {
      css: styleObj.css,
      mammothMap: styleObj.mammothMap,
    });

    return styleObj;
  }

  /**
   * 解析段落属性
   */
  parseParagraphProperties(pPr, styleObj) {
    const css = styleObj.css;
    const mammoth = styleObj.mammothMap;

    // 间距
    if (pPr['w:spacing']) {
      const spacing = pPr['w:spacing'][0].$;
      if (spacing.before) {
        css.marginTop = `${spacing.before / 20}pt`;
        mammoth.marginTop = `${spacing.before / 20}pt`;
      }
      if (spacing.after) {
        css.marginBottom = `${spacing.after / 20}pt`;
        mammoth.marginBottom = `${spacing.after / 20}pt`;
      }
      if (spacing.line) {
        if (spacing.lineRule === 'exact') {
          css.lineHeight = `${spacing.line / 240}pt`;
          mammoth.lineHeight = `${spacing.line / 240}pt`;
        } else {
          css.lineHeight = (spacing.line / 240).toString();
          mammoth.lineHeight = (spacing.line / 240).toString();
        }
      }
    }

    // 对齐
    if (pPr['w:jc']) {
      const align = pPr['w:jc'][0].$.val;
      const alignMap = { left: 'left', center: 'center', right: 'right', both: 'justify' };
      css.textAlign = alignMap[align] || 'left';
      mammoth.textAlign = alignMap[align] || 'left';
    }

    // 缩进
    if (pPr['w:ind']) {
      const ind = pPr['w:ind'][0].$;
      if (ind.left) {
        css.marginLeft = `${ind.left / 20}pt`;
        mammoth.marginLeft = `${ind.left / 20}pt`;
      }
      if (ind.right) {
        css.marginRight = `${ind.right / 20}pt`;
        mammoth.marginRight = `${ind.right / 20}pt`;
      }
      if (ind.firstLine) {
        css.textIndent = `${ind.firstLine / 20}pt`;
        mammoth.textIndent = `${ind.firstLine / 20}pt`;
      }
      if (ind.hanging) {
        css.textIndent = `-${ind.hanging / 20}pt`;
        mammoth.textIndent = `-${ind.hanging / 20}pt`;
      }
    }
  }

  /**
   * 解析文字属性
   */
  parseRunProperties(rPr, styleObj) {
    const css = styleObj.css;
    const mammoth = styleObj.mammothMap;

    // 字体
    if (rPr['w:rFonts']) {
      const fonts = rPr['w:rFonts'][0].$;
      const fontFamily = `"${fonts.ascii || fonts.hAnsi || 'Calibri'}", sans-serif`;
      css.fontFamily = fontFamily;
      mammoth.fontFamily = fontFamily;
    }

    // 字号
    if (rPr['w:sz']) {
      const fontSize = `${rPr['w:sz'][0].$.val / 2}pt`;
      css.fontSize = fontSize;
      mammoth.fontSize = fontSize;
    }

    // 颜色
    if (rPr['w:color']) {
      const color = rPr['w:color'][0].$.val;
      if (color !== 'auto') {
        css.color = `#${color}`;
        mammoth.color = `#${color}`;
      }
    }

    // 背景色/高亮
    if (rPr['w:highlight']) {
      const highlight = rPr['w:highlight'][0].$.val;
      const colorMap = {
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
      if (colorMap[highlight]) {
        css.backgroundColor = colorMap[highlight];
        mammoth.backgroundColor = colorMap[highlight];
      }
    }

    // 阴影背景
    if (rPr['w:shd']) {
      const shd = rPr['w:shd'][0].$;
      if (shd.fill && shd.fill !== 'auto') {
        css.backgroundColor = `#${shd.fill}`;
        mammoth.backgroundColor = `#${shd.fill}`;
      }
    }

    // 粗体
    if (rPr['w:b']) {
      css.fontWeight = 'bold';
      mammoth.fontWeight = 'bold';
    }

    // 斜体
    if (rPr['w:i']) {
      css.fontStyle = 'italic';
      mammoth.fontStyle = 'italic';
    }

    // 下划线
    if (rPr['w:u']) {
      const uType = rPr['w:u'][0].$.val;
      if (uType === 'single' || uType === 'thick' || uType === 'double') {
        css.textDecoration = 'underline';
        mammoth.textDecoration = 'underline';
      }
    }

    // 删除线
    if (rPr['w:strike']) {
      css.textDecoration = 'line-through';
      mammoth.textDecoration = 'line-through';
    }
  }

  /**
   * 为 mammoth 构建样式映射
   */
  buildMammothStyleMap(styleId, parsedStyle) {
    const styleName = parsedStyle.name;

    if (parsedStyle.type === 'paragraph') {
      this.paragraphStyleMap.set(styleName, {
        styleId,
        css: parsedStyle.css,
        mammoth: parsedStyle.mammothMap,
      });
    } else if (parsedStyle.type === 'character') {
      this.runStyleMap.set(styleName, {
        styleId,
        css: parsedStyle.css,
        mammoth: parsedStyle.mammothMap,
      });
    }
  }

  /**
   * 提取关系信息
   */
  async extractRelationships(zip) {
    try {
      const relsXml = await zip.file('word/_rels/document.xml.rels').async('text');
      const relsData = await xml2js.parseStringPromise(relsXml);

      if (relsData.Relationships && relsData.Relationships.Relationship) {
        for (const rel of relsData.Relationships.Relationship) {
          this.relationships.set(rel.$.Id, rel.$.Target);
        }
      }
    } catch (error: any) {
      console.log('⚠️ 关系文件解析失败');
    }
  }

  /**
   * 提取媒体文件
   */
  async extractMedia(zip: any) {
    for (const [fileName, file] of Object.entries(zip.files)) {
      if (fileName.startsWith('word/media/')) {
        const mediaBuffer = await (file as any).async('base64');
        const mediaType = this.getMediaType(fileName);
        this.media.set(fileName, {
          data: `data:${mediaType};base64,${mediaBuffer}`,
          type: mediaType,
        });
      }
    }
  }

  /**
   * 提取文档结构
   */
  async extractDocumentStructure(zip) {
    try {
      const docXml = await zip.file('word/document.xml').async('text');
      this.documentStructure = await xml2js.parseStringPromise(docXml);
    } catch (error: any) {
      console.log('⚠️ 文档结构解析失败');
    }
  }

  /**
   * 深度解析 DOCX 结构 - 基于测试脚本验证的方法
   */
  async deepParseDocxStructure(inputPath) {
    console.log('🔍 深度解析 DOCX 结构...');

    const docxBuffer = await fs.readFile(inputPath);
    const zip = await JSZip.loadAsync(docxBuffer);

    // 解析文档主体，提取所有样式引用
    await this.extractDocumentWithStyleReferences(zip);

    // 解析编号样式
    await this.extractNumberingStyles(zip);

    // 解析主题样式
    await this.extractThemeStyles(zip);

    console.log('✅ DOCX 结构深度解析完成');
  }

  /**
   * 提取文档中的样式引用
   */
  async extractDocumentWithStyleReferences(zip) {
    try {
      const docXml = await zip.file('word/document.xml').async('text');
      const docData = await xml2js.parseStringPromise(docXml);

      // 遍历文档，收集所有样式引用
      this.collectStyleReferences(docData);
    } catch (error: any) {
      console.log('⚠️ 文档样式引用提取失败:', error.message);
    }
  }

  /**
   * 收集样式引用
   */
  collectStyleReferences(node, path = '') {
    if (!node || typeof node !== 'object') return;

    if (Array.isArray(node)) {
      node.forEach((item, index) => {
        this.collectStyleReferences(item, `${path}[${index}]`);
      });
      return;
    }

    // 收集段落样式引用
    if (node['w:pStyle']) {
      const styleId = node['w:pStyle'][0].$.val;
      console.log(`📝 发现段落样式引用: ${styleId}`);
    }

    // 收集字符样式引用
    if (node['w:rStyle']) {
      const styleId = node['w:rStyle'][0].$.val;
      console.log(`✏️ 发现字符样式引用: ${styleId}`);
    }

    // 收集表格样式引用
    if (node['w:tblStyle']) {
      const styleId = node['w:tblStyle'][0].$.val;
      console.log(`📊 发现表格样式引用: ${styleId}`);
    }

    // 递归处理子节点
    Object.keys(node).forEach(key => {
      if (key !== '$' && node[key]) {
        this.collectStyleReferences(node[key], `${path}.${key}`);
      }
    });
  }

  /**
   * 提取编号样式
   */
  async extractNumberingStyles(zip) {
    try {
      const numberingXml = await zip.file('word/numbering.xml')?.async('text');
      if (numberingXml) {
        const numberingData = await xml2js.parseStringPromise(numberingXml);
        console.log('📋 编号样式解析完成');
        // 这里可以进一步处理编号样式
      }
    } catch (error: any) {
      console.log('⚠️ 编号样式解析失败');
    }
  }

  /**
   * 提取主题样式
   */
  async extractThemeStyles(zip) {
    try {
      const themeXml = await zip.file('word/theme/theme1.xml')?.async('text');
      if (themeXml) {
        const themeData = await xml2js.parseStringPromise(themeXml);
        console.log('🎨 主题样式解析完成');
        // 这里可以进一步处理主题样式
      }
    } catch (error: any) {
      console.log('⚠️ 主题样式解析失败');
    }
  }

  /**
   * 使用完整样式映射配置进行转换
   */
  async convertWithCompleteStyleMapping(inputPath, options = {}) {
    console.log('🔄 使用修复后的mammoth配置转换...');

    // 构建强化的样式映射
    const styleMapping = this.buildFixedStyleMapping();

    const mammothOptions = {
      styleMap: styleMapping,

      // 图片处理
      convertImage:
        (options as any).preserveImages !== false
          ? mammoth.images.imgElement(function (image) {
              return image.read('base64').then(function (imageBuffer) {
                return {
                  src: `data:${image.contentType};base64,${imageBuffer}`,
                  alt: image.altText || 'Document Image',
                };
              });
            })
          : mammoth.images.ignore,

      // 包含默认样式映射
      includeDefaultStyleMap: true,

      // 包含嵌入样式
      includeEmbeddedStyleMap: true,

      // 不忽略空段落
      ignoreEmptyParagraphs: false,

      // 转换文档时保留更多信息
      transformDocument: mammoth.transforms.paragraph(element => {
        // 为每个段落添加样式类和ID
        if (element.styleName) {
          return {
            ...element,
            styleId: element.styleName,
            className: this.sanitizeClassName(element.styleName),
          };
        }
        return element;
      }),
    };

    console.log('📋 修复后的mammoth配置:', {
      styleMappingCount: styleMapping.length,
      preserveImages: (options as any).preserveImages !== false,
      includeDefaults: true,
      stylesExtracted: this.styles.size,
    });

    // 使用 mammoth 转换
    const result = await mammoth.convertToHtml({ path: inputPath }, mammothOptions);

    console.log('📄 mammoth 转换结果:', {
      contentLength: result.value.length,
      messagesCount: result.messages.length,
      hasContent: !!result.value,
      hasStyleTags: result.value.includes('<style>'),
      hasInlineStyles: result.value.includes('style='),
    });

    if (result.messages.length > 0) {
      console.log('⚠️ mammoth 消息:', result.messages.slice(0, 5));
    }

    // 强制注入Word样式
    const styledHtml = this.forceInjectWordStyles(result.value);

    // 生成完整 HTML
    const fullHtml = this.generateCompleteHtml(styledHtml, options);

    return {
      success: true,
      content: fullHtml,
      rawContent: styledHtml,
      styles: this.cssRules.join('\n'),
      messages: result.messages,
      metadata: {
        converter: 'enhanced-mammoth-fixed',
        stylesCount: this.styles.size,
        mediaCount: this.media.size,
        cssRulesGenerated: this.cssRules.length,
        warnings: result.messages.filter(m => m.type === 'warning').length,
      },
    };
  }

  /**
   * 构建修复后的样式映射
   */
  buildFixedStyleMapping() {
    console.log('🎨 构建修复后的样式映射...');

    const styleMapping = [
      // 基础段落样式 - 保留样式信息
      "p[style-name='Normal'] => p.normal:fresh",
      "p[style-name='正文'] => p.normal:fresh",
      "p[style-name='Body Text'] => p.body-text:fresh",

      // 标题样式 - 保留层级和样式
      "p[style-name='Heading 1'] => h1.heading-1:fresh",
      "p[style-name='标题 1'] => h1.heading-1:fresh",
      "p[style-name='Heading 2'] => h2.heading-2:fresh",
      "p[style-name='标题 2'] => h2.heading-2:fresh",
      "p[style-name='Heading 3'] => h3.heading-3:fresh",
      "p[style-name='标题 3'] => h3.heading-3:fresh",
      "p[style-name='Heading 4'] => h4.heading-4:fresh",
      "p[style-name='标题 4'] => h4.heading-4:fresh",
      "p[style-name='Heading 5'] => h5.heading-5:fresh",
      "p[style-name='标题 5'] => h5.heading-5:fresh",
      "p[style-name='Heading 6'] => h6.heading-6:fresh",
      "p[style-name='标题 6'] => h6.heading-6:fresh",

      // 特殊样式
      "p[style-name='Title'] => h1.title:fresh",
      "p[style-name='Subtitle'] => h2.subtitle:fresh",
      "p[style-name='标题'] => h1.title:fresh",
      "p[style-name='副标题'] => h2.subtitle:fresh",

      // 字符样式 - 保留格式
      "r[style-name='Strong'] => strong.word-strong",
      "r[style-name='Emphasis'] => em.word-emphasis",
      "r[style-name='加粗'] => strong.word-strong",
      "r[style-name='斜体'] => em.word-emphasis",

      // 列表样式
      "p[style-name='List Paragraph'] => li.list-item:fresh",
      "p[style-name='列表段落'] => li.list-item:fresh",
      "p[style-name='Bullet List'] => li.bullet-item:fresh",
      "p[style-name='Numbered List'] => li.numbered-item:fresh",

      // 表格样式
      'table => table.docx-table',
      'tr => tr.docx-row',
      'td => td.docx-cell',
      'th => th.docx-header',

      // 通用映射 - 保留所有样式信息
      'p => p.docx-paragraph:fresh',
      'r => span.docx-run',
      'b => strong.docx-bold',
      'i => em.docx-italic',
      'u => span.docx-underline',
      'strike => span.docx-strikethrough',
      'sup => sup.docx-superscript',
      'sub => sub.docx-subscript',
    ];

    // 为每个提取的样式创建映射
    for (const [styleId, styleData] of this.styles) {
      // 跳过无效的样式ID
      if (!styleId || typeof styleId !== 'string') {
        console.log(`⚠️ 跳过无效样式ID: ${styleId}`);
        continue;
      }

      const styleName = styleData.name || styleId;
      const className = this.sanitizeClassName(styleId);

      if (styleData.type === 'paragraph') {
        styleMapping.push(`p[style-name='${styleName}'] => p.${className}:fresh`);

        // 检查是否是标题样式
        const headingLevel = this.extractHeadingLevel(styleName);
        if (headingLevel && headingLevel > 0) {
          styleMapping.push(`p[style-name='${styleName}'] => h${headingLevel}.${className}:fresh`);
        }
      } else if (styleData.type === 'character') {
        styleMapping.push(`r[style-name='${styleName}'] => span.${className}`);
      } else if (styleData.type === 'table') {
        styleMapping.push(`table[style-name='${styleName}'] => table.${className}`);
      }

      // 生成对应的 CSS 规则
      this.generateCssRule(`.${className}`, styleData.css);
    }

    console.log(`📋 生成了 ${styleMapping.length} 个样式映射`);
    return styleMapping;
  }

  /**
   * 强制注入Word样式
   */
  forceInjectWordStyles(html) {
    console.log('💉 强制注入Word样式...');

    // 生成完整的Word样式CSS
    const wordStyles = this.generateWordStylesCSS();

    // 将样式添加到CSS规则中
    this.cssRules.unshift(wordStyles);

    return html;
  }

  /**
   * 生成Word样式CSS
   */
  generateWordStylesCSS() {
    return `
/* ===== Word样式强制覆盖 ===== */

/* 全局重置 */
* {
  box-sizing: border-box !important;
  -webkit-print-color-adjust: exact !important;
  color-adjust: exact !important;
  print-color-adjust: exact !important;
}

/* 字体强制设置 */
html, body, p, div, span, h1, h2, h3, h4, h5, h6, 
table, td, th, tr, ul, ol, li, strong, em, b, i {
  font-family: "Calibri", "Microsoft YaHei", "SimSun", "宋体", sans-serif !important;
  color: #000000 !important;
}

/* 根元素样式 */
html, body {
  margin: 0 !important;
  padding: 0 !important;
  background-color: #ffffff !important;
  font-size: 11pt !important;
  line-height: 1.08 !important;
}

/* 段落样式 */
p, p.normal, p.docx-paragraph, .normal {
  font-size: 11pt !important;
  line-height: 1.08 !important;
  margin: 0pt 0pt 8pt 0pt !important;
  text-align: left !important;
  color: #000000 !important;
}

/* 标题样式 */
h1, h1.heading-1, h1.title, .heading-1, .title {
  font-family: "Calibri Light", "Calibri", "Microsoft YaHei", sans-serif !important;
  font-size: 16pt !important;
  font-weight: normal !important;
  color: #2F5496 !important;
  margin: 24pt 0pt 0pt 0pt !important;
  line-height: 1.15 !important;
}

h2, h2.heading-2, h2.subtitle, .heading-2, .subtitle {
  font-family: "Calibri Light", "Calibri", "Microsoft YaHei", sans-serif !important;
  font-size: 13pt !important;
  font-weight: normal !important;
  color: #2F5496 !important;
  margin: 10pt 0pt 0pt 0pt !important;
  line-height: 1.15 !important;
}

h3, h3.heading-3, .heading-3 {
  font-family: "Calibri Light", "Calibri", "Microsoft YaHei", sans-serif !important;
  font-size: 12pt !important;
  font-weight: normal !important;
  color: #1F3763 !important;
  margin: 10pt 0pt 0pt 0pt !important;
  line-height: 1.15 !important;
}

h4, h4.heading-4, .heading-4,
h5, h5.heading-5, .heading-5,
h6, h6.heading-6, .heading-6 {
  font-size: 11pt !important;
  font-weight: bold !important;
  color: #2F5496 !important;
  margin: 10pt 0pt 0pt 0pt !important;
  line-height: 1.15 !important;
}

/* 文本格式 */
strong, strong.word-strong, strong.docx-bold, b {
  font-weight: bold !important;
}

em, em.word-emphasis, em.docx-italic, i {
  font-style: italic !important;
}

span.docx-underline, u {
  text-decoration: underline !important;
}

span.docx-strikethrough, del, strike {
  text-decoration: line-through !important;
}

/* 表格样式 */
table, table.docx-table {
  width: 100% !important;
  border-collapse: collapse !important;
  margin: 0pt 0pt 8pt 0pt !important;
  font-size: 11pt !important;
  line-height: 1.08 !important;
}

td, td.docx-cell, th, th.docx-header {
  border: 0.5pt solid #000000 !important;
  padding: 0pt 5.4pt !important;
  vertical-align: top !important;
  font-size: 11pt !important;
  line-height: 1.08 !important;
}

/* 列表样式 */
ul, ol {
  margin: 0pt 0pt 8pt 0pt !important;
  padding-left: 36pt !important;
}

li, li.list-item, li.bullet-item, li.numbered-item {
  margin: 0pt !important;
  font-size: 11pt !important;
  line-height: 1.08 !important;
}

/* 图片样式 */
img {
  max-width: 100% !important;
  height: auto !important;
  display: inline-block !important;
}

/* 链接样式 */
a {
  color: #0563C1 !important;
  text-decoration: underline !important;
}

a:visited {
  color: #954F72 !important;
}

/* 打印样式 */
@media print {
  * {
    -webkit-print-color-adjust: exact !important;
    color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
  
  body {
    background-color: white !important;
  }
}

@page {
  size: A4 portrait !important;
  margin: 2.54cm 1.91cm 2.54cm 1.91cm !important;
}
`;
  }

  /**
   * 清理类名
   */
  sanitizeClassName(name) {
    // 添加空值检查
    if (!name || typeof name !== 'string') {
      console.log(`⚠️ sanitizeClassName收到无效参数: ${name}`);
      return 'default-style';
    }

    let sanitized = name
      .toLowerCase()
      .replace(/[^a-z0-9\u4e00-\u9fff]/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-|-$/g, '');

    // 确保类名不以数字开头，CSS规范要求类名必须以字母、下划线或连字符开头
    if (/^\d/.test(sanitized)) {
      sanitized = `style-${sanitized}`;
    }

    return sanitized;
  }

  /**
   * 提取标题级别
   */
  extractHeadingLevel(styleName) {
    const match = styleName.match(/(?:Heading|标题)\s*(\d+)/i);
    return match ? Math.min(parseInt(match[1]), 6) : null;
  }

  /**
   * 转换文档并注入样式
   */
  transformDocumentWithStyles(document) {
    // 这里可以添加文档转换的自定义逻辑
    // mammoth 会在转换过程中调用这个函数
    return document;
  }

  /**
   * 增强 HTML 样式
   */
  enhanceHtmlWithStyles(html) {
    // 后处理 HTML，添加额外的样式信息
    let enhancedHtml = html;

    // 处理表格样式
    enhancedHtml = enhancedHtml.replace(
      /<table class="docx-table">/g,
      '<table class="docx-table" style="border-collapse: collapse; width: 100%; margin: 8pt 0;">'
    );

    // 处理表格单元格
    enhancedHtml = enhancedHtml.replace(
      /<td>/g,
      '<td style="border: 0.5pt solid #000; padding: 4pt; vertical-align: top;">'
    );

    return enhancedHtml;
  }

  /**
   * 工具函数
   */
  getMediaType(fileName) {
    const ext = fileName.split('.').pop().toLowerCase();
    const typeMap = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      svg: 'image/svg+xml',
    };
    return typeMap[ext] || 'image/png';
  }

  generateCssRule(selector, styles) {
    if (Object.keys(styles).length === 0) return;
    const cssText = this.cssObjectToString(styles);
    this.cssRules.push(`${selector} { ${cssText} }`);
  }

  cssObjectToString(styles) {
    return Object.entries(styles)
      .map(([key, value]) => `${this.camelToKebab(key)}: ${value}`)
      .join('; ');
  }

  camelToKebab(str) {
    return str.replace(/[A-Z]/g, letter => `-${letter.toLowerCase()}`);
  }

  /**
   * 生成完整 HTML
   */
  generateCompleteHtml(content, options = {}) {
    const baseStyles = `
      body {
        font-family: "Calibri", "Microsoft YaHei", "SimSun", sans-serif;
        font-size: 11pt;
        line-height: 1.08;
        margin: 0;
        padding: 20px;
        background: white;
        color: black;
      }
      
      .docx-table {
        border-collapse: collapse;
        width: 100%;
        margin: 8pt 0;
      }
      
      .docx-table td, .docx-table th {
        border: 0.5pt solid #000;
        padding: 4pt;
        vertical-align: top;
      }
      
      p {
        margin: 0 0 8pt 0;
      }
      
      h1, h2, h3, h4, h5, h6 {
        margin: 12pt 0 6pt 0;
        font-weight: bold;
      }
      
      /* 打印优化 */
      @media print {
        * {
          -webkit-print-color-adjust: exact !important;
          color-adjust: exact !important;
        }
      }
    `;

    const allStyles = baseStyles + '\n' + this.cssRules.join('\n');

    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <style>
${allStyles}
    </style>
</head>
<body>
${content}
</body>
</html>`;
  }
}

/**
 * 导出函数
 */
export async function convertDocxToHtmlEnhanced(inputPath: string, options: any = {}) {
  const converter = new EnhancedMammothConverter();
  return await converter.convertDocxToHtml(inputPath, options);
}

export { EnhancedMammothConverter };
