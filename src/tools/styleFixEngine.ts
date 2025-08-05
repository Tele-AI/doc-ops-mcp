/**
 * 样式修复引擎 - 解决DOCX到HTML转换中样式丢失的问题
 *
 * 主要问题分析：
 * 1. mammoth的样式映射配置不完整
 * 2. 样式提取器没有正确生成CSS
 * 3. 样式注入时机不对
 * 4. CSS选择器优先级不够
 */

const mammoth = require('mammoth');
const fs = require('fs').promises;
const JSZip = require('jszip');
const xml2js = require('xml2js');
const cheerio = require('cheerio');

interface StyleFixOptions {
  preserveImages?: boolean;
  enableDebug?: boolean;
  forceWordStyles?: boolean;
  customFontFamily?: string;
}

interface FixedConversionResult {
  success: boolean;
  content?: string;
  error?: string;
  metadata?: {
    stylesExtracted: number;
    cssRulesGenerated: number;
    elementsStyled: number;
    converter: string;
  };
}

export class StyleFixEngine {
  private docxStyles: Map<string, any> = new Map();
  private cssRules: string[] = [];
  private options: StyleFixOptions;

  constructor(options: StyleFixOptions = {}) {
    this.options = {
      preserveImages: true,
      enableDebug: false,
      forceWordStyles: true,
      customFontFamily: 'Microsoft YaHei',
      ...options,
    };
  }

  /**
   * 主要修复函数 - 解决样式丢失问题
   */
  async fixDocxToHtml(docxPath: string): Promise<FixedConversionResult> {
    try {
      this.log('🔧 开始样式修复转换...');

      // 步骤1: 提取DOCX中的原始样式信息
      await this.extractDocxStyles(docxPath);

      // 步骤2: 使用修复后的mammoth配置进行转换
      const mammothResult = await this.convertWithFixedMammoth(docxPath);

      // 步骤3: 强制注入Word样式
      const styledHtml = this.injectWordStyles(mammothResult.value);

      // 步骤4: 验证和优化最终HTML
      const finalHtml = this.optimizeFinalHtml(styledHtml);

      this.log('✅ 样式修复完成');

      return {
        success: true,
        content: finalHtml,
        metadata: {
          stylesExtracted: this.docxStyles.size,
          cssRulesGenerated: this.cssRules.length,
          elementsStyled: this.countStyledElements(finalHtml),
          converter: 'style-fix-engine',
        },
      };
    } catch (error: any) {
      this.log('❌ 样式修复失败:', error.message);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * 提取DOCX中的样式信息
   */
  private async extractDocxStyles(docxPath: string): Promise<void> {
    this.log('📊 提取DOCX样式信息...');

    const docxBuffer = await fs.readFile(docxPath);
    const zip = await JSZip.loadAsync(docxBuffer);

    // 提取styles.xml
    const stylesFile = zip.file('word/styles.xml');
    if (stylesFile) {
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
          const styleId = style.styleId || style['w:styleId'];
          if (styleId) {
            this.docxStyles.set(styleId, style);
            this.log(`  发现样式: ${styleId}`);
          }
        }
      }
    }

    this.log(`📊 提取了 ${this.docxStyles.size} 个样式定义`);
  }

  /**
   * 使用修复后的mammoth配置进行转换
   */
  private async convertWithFixedMammoth(docxPath: string): Promise<any> {
    this.log('🔄 使用修复后的mammoth配置转换...');

    // 创建强化的样式映射
    const styleMap = this.createEnhancedStyleMap();

    const config = {
      styleMap: styleMap,

      // 图片处理
      convertImage: this.options.preserveImages
        ? mammoth.images.imgElement((image: any) => {
            return image.read('base64').then((imageBuffer: string) => {
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
      transformDocument: mammoth.transforms.paragraph((element: any) => {
        // 为每个段落添加样式类
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

    const result = await mammoth.convertToHtml({ path: docxPath }, config);

    this.log('📄 mammoth转换结果:');
    this.log(`  - 内容长度: ${result.value.length}`);
    this.log(`  - 消息数量: ${result.messages.length}`);

    return result;
  }

  /**
   * 创建增强的样式映射
   */
  private createEnhancedStyleMap(): string[] {
    const styleMap = [
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
    for (const [styleId, styleData] of this.docxStyles) {
      const className = this.sanitizeClassName(styleId);
      const styleName = styleData['w:name']?.[0]?.val || styleId;

      if (styleData.type === 'paragraph') {
        styleMap.push(`p[style-name='${styleName}'] => p.${className}:fresh`);
      } else if (styleData.type === 'character') {
        styleMap.push(`r[style-name='${styleName}'] => span.${className}`);
      }
    }

    this.log(`📋 创建了 ${styleMap.length} 个样式映射`);
    return styleMap;
  }

  /**
   * 强制注入Word样式
   */
  private injectWordStyles(html: string): string {
    this.log('💉 注入Word样式...');

    // 生成完整的Word样式CSS
    const wordStyles = this.generateWordStylesCSS();

    // 如果HTML已经有完整结构，在head中注入样式
    if (html.includes('<html>') || html.includes('<!DOCTYPE')) {
      if (html.includes('</head>')) {
        return html.replace('</head>', `<style type="text/css">\n${wordStyles}\n</style>\n</head>`);
      } else if (html.includes('<head>')) {
        return html.replace('<head>', `<head>\n<style type="text/css">\n${wordStyles}\n</style>`);
      } else {
        // 在html标签后添加head
        return html.replace(
          '<html>',
          `<html>\n<head>\n<style type="text/css">\n${wordStyles}\n</style>\n</head>`
        );
      }
    } else {
      // 创建完整的HTML结构
      return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Document</title>
  <style type="text/css">
${wordStyles}
  </style>
</head>
<body>
${html}
</body>
</html>`;
    }
  }

  /**
   * 生成Word样式CSS
   */
  private generateWordStylesCSS(): string {
    const fontFamily = this.options.customFontFamily || 'Microsoft YaHei';

    return `
/* ===== Word样式修复 - 强制应用 ===== */

/* 全局重置和基础样式 */
* {
  box-sizing: border-box !important;
  -webkit-print-color-adjust: exact !important;
  color-adjust: exact !important;
  print-color-adjust: exact !important;
}

/* 字体强制设置 */
html, body, p, div, span, h1, h2, h3, h4, h5, h6, 
table, td, th, tr, ul, ol, li, strong, em, b, i {
  font-family: "Calibri", "${fontFamily}", "SimSun", "宋体", sans-serif !important;
  color: #000000 !important;
  background-color: transparent !important;
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
  font-family: "Calibri Light", "Calibri", "${fontFamily}", sans-serif !important;
  font-size: 16pt !important;
  font-weight: normal !important;
  color: #2F5496 !important;
  margin: 24pt 0pt 0pt 0pt !important;
  line-height: 1.15 !important;
}

h2, h2.heading-2, h2.subtitle, .heading-2, .subtitle {
  font-family: "Calibri Light", "Calibri", "${fontFamily}", sans-serif !important;
  font-size: 13pt !important;
  font-weight: normal !important;
  color: #2F5496 !important;
  margin: 10pt 0pt 0pt 0pt !important;
  line-height: 1.15 !important;
}

h3, h3.heading-3, .heading-3 {
  font-family: "Calibri Light", "Calibri", "${fontFamily}", sans-serif !important;
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
   * 优化最终HTML
   */
  private optimizeFinalHtml(html: string): string {
    this.log('🔧 优化最终HTML...');

    // 使用cheerio处理HTML
    const $ = cheerio.load(html, { decodeEntities: false });

    // 确保有DOCTYPE
    if (!html.includes('<!DOCTYPE')) {
      // cheerio会自动添加html结构，但我们需要确保DOCTYPE
      const finalHtml = $.html();
      return `<!DOCTYPE html>\n${finalHtml}`;
    }

    // 为没有class的元素添加适当的class
    $('p').each((i, el) => {
      const $el = $(el);
      if (!$el.attr('class')) {
        $el.addClass('docx-paragraph');
      }
    });

    $('span').each((i, el) => {
      const $el = $(el);
      if (!$el.attr('class')) {
        $el.addClass('docx-run');
      }
    });

    $('table').each((i, el) => {
      const $el = $(el);
      if (!$el.attr('class')) {
        $el.addClass('docx-table');
      }
    });

    return $.html();
  }

  /**
   * 统计带样式的元素数量
   */
  private countStyledElements(html: string): number {
    const $ = cheerio.load(html, { decodeEntities: false });
    return $('[class], [style]').length;
  }

  /**
   * 清理类名
   */
  private sanitizeClassName(name: string): string {
    return name
      .toLowerCase()
      .replace(/[^a-z0-9\u4e00-\u9fff]/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-|-$/g, '');
  }

  /**
   * 调试日志
   */
  private log(...args: any[]): void {
    if (this.options.enableDebug) {
      console.error(...args);
    }
  }
}

/**
 * 导出修复函数
 */
export async function fixDocxToHtmlStyles(
  docxPath: string,
  options: StyleFixOptions = {}
): Promise<FixedConversionResult> {
  const engine = new StyleFixEngine(options);
  return await engine.fixDocxToHtml(docxPath);
}
