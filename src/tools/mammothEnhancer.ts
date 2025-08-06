/**
 * MammothEnhancer - mammoth.js 增强器
 * 构建 mammoth.js 样式映射配置，集成样式信息到 mammoth 转换流程
 * 自定义图片转换逻辑，文档转换预处理和后处理
 */

const mammoth = require('mammoth');
const path = require('path');
const fs = require('fs/promises');
import { StyleExtractor, StyleDefinition, DocumentStyle } from './styleExtractor';

// 路径安全验证函数
function validatePath(inputPath: string): string {
  const resolvedPath = path.resolve(inputPath);
  const normalizedPath = path.normalize(resolvedPath);
  
  // 检查路径遍历攻击
  if (normalizedPath.includes('..') || normalizedPath !== resolvedPath) {
    throw new Error('Invalid path: Path traversal detected');
  }
  
  return normalizedPath;
}

interface MammothEnhancerOptions {
  preserveImages?: boolean;
  imageOutputDir?: string;
  convertImagesToBase64?: boolean;
  includeDefaultStyles?: boolean;
  customStyleMappings?: string[];
  transformDocument?: boolean;
}

interface ConversionResult {
  html: string;
  css: string;
  messages: any[];
  success: boolean;
  error?: string;
  images?: { [key: string]: string };
}

export class MammothEnhancer {
  private styleExtractor: StyleExtractor;
  private extractedStyles: Map<string, StyleDefinition> = new Map();
  private documentStyles: DocumentStyle[] = [];
  private generatedCSS: string = '';
  private imageCounter: number = 0;
  private images: { [key: string]: string } = {};

  constructor() {
    this.styleExtractor = new StyleExtractor();
  }

  /**
   * 主转换函数 - 增强的 DOCX 到 HTML 转换
   */
  async convertDocxToHtml(
    inputPath: string,
    options: MammothEnhancerOptions = {}
  ): Promise<ConversionResult> {
    try {
      console.log('🚀 开始增强型 mammoth 转换...');

      // 步骤1: 提取样式信息
      const styleResult = await this.styleExtractor.extractStyles(inputPath);
      this.extractedStyles = styleResult.styles;
      this.documentStyles = styleResult.documentStyles;
      this.generatedCSS = styleResult.css;

      console.log(`📊 提取了 ${this.extractedStyles.size} 个样式定义`);

      // 步骤2: 构建 mammoth 配置
      const mammothConfig = this.buildMammothConfig(options);

      // 步骤3: 执行 mammoth 转换
      const result = await mammoth.convertToHtml({ path: inputPath }, mammothConfig);

      // 步骤4: 后处理 HTML
      const enhancedHtml = this.postProcessHtml(result.value, options);

      // 步骤5: 生成完整的 CSS
      const completeCSS = this.generateCompleteCSS(options);

      console.log('✅ 增强型转换完成');

      return {
        html: enhancedHtml,
        css: completeCSS,
        messages: result.messages,
        success: true,
        images: this.images,
      };
    } catch (error: any) {
      console.error('❌ 转换失败:', error);
      return {
        html: '',
        css: '',
        messages: [],
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * 构建 mammoth 配置
   */
  private buildMammothConfig(options: MammothEnhancerOptions): any {
    const config: any = {
      // 样式映射
      styleMap: this.buildStyleMappings(options),

      // 图片转换
      convertImage: this.buildImageConverter(options),

      // 包含默认样式映射
      includeDefaultStyleMap: options.includeDefaultStyles !== false,

      // 包含嵌入样式
      includeEmbeddedStyleMap: true,

      // 忽略空段落
      ignoreEmptyParagraphs: false,
    };

    // 文档转换器
    if (options.transformDocument !== false) {
      config.transformDocument = this.buildDocumentTransformer();
    }

    return config;
  }

  /**
   * 构建样式映射
   */
  private buildStyleMappings(options: MammothEnhancerOptions): string[] {
    const mappings: string[] = [];

    // 基础样式映射
    const baseMappings = [
      // 段落样式映射
      "p[style-name='Normal'] => p.normal:fresh",
      "p[style-name='Heading 1'] => h1.heading1:fresh",
      "p[style-name='Heading 2'] => h2.heading2:fresh",
      "p[style-name='Heading 3'] => h3.heading3:fresh",
      "p[style-name='Heading 4'] => h4.heading4:fresh",
      "p[style-name='Heading 5'] => h5.heading5:fresh",
      "p[style-name='Heading 6'] => h6.heading6:fresh",
      "p[style-name='Title'] => h1.title:fresh",
      "p[style-name='Subtitle'] => h2.subtitle:fresh",
      "p[style-name='List Paragraph'] => li.list-paragraph:fresh",

      // 字符样式映射
      "r[style-name='Strong'] => strong.strong:fresh",
      "r[style-name='Emphasis'] => em.emphasis:fresh",
      "r[style-name='Hyperlink'] => a.hyperlink:fresh",
      "r[style-name='Subtle Emphasis'] => em.subtle-emphasis:fresh",
      "r[style-name='Intense Emphasis'] => strong.intense-emphasis:fresh",

      // 表格样式映射
      'table => table.docx-table:fresh',
      'tr => tr.docx-row:fresh',
      'td => td.docx-cell:fresh',
      'th => th.docx-header:fresh',

      // 通用映射
      'p => p:fresh',
      'r => span:fresh',
      'b => strong',
      'i => em',
      'u => u',
      'strike => del',
      'sup => sup',
      'sub => sub',
    ];

    mappings.push(...baseMappings);

    // 从提取的样式生成映射
    for (const [styleId, style] of this.extractedStyles) {
      const mapping = this.generateStyleMapping(styleId, style);
      if (mapping) {
        mappings.push(mapping);
      }
    }

    // 自定义样式映射
    if (options.customStyleMappings) {
      mappings.push(...options.customStyleMappings);
    }

    console.log(`🎨 生成了 ${mappings.length} 个样式映射`);
    return mappings;
  }

  /**
   * 生成单个样式映射
   */
  private generateStyleMapping(styleId: string, style: StyleDefinition): string | null {
    const { name, type, mammothMapping } = style;

    if (!name || !mammothMapping) return null;

    // 确保类名不以数字开头，CSS规范要求类名必须以字母、下划线或连字符开头
    const sanitizedId = /^\d/.test(styleId) ? `style-${styleId}` : styleId;

    let mapping = '';

    switch (type) {
      case 'paragraph':
        if (name.toLowerCase().includes('heading')) {
          const level = this.extractHeadingLevel(name);
          mapping = `p[style-name='${name}'] => h${level}.${sanitizedId}:fresh`;
        } else if (name.toLowerCase().includes('list')) {
          mapping = `p[style-name='${name}'] => li.${sanitizedId}:fresh`;
        } else {
          mapping = `p[style-name='${name}'] => p.${sanitizedId}:fresh`;
        }
        break;

      case 'character':
        if (name.toLowerCase().includes('strong') || name.toLowerCase().includes('bold')) {
          mapping = `r[style-name='${name}'] => strong.${sanitizedId}:fresh`;
        } else if (
          name.toLowerCase().includes('emphasis') ||
          name.toLowerCase().includes('italic')
        ) {
          mapping = `r[style-name='${name}'] => em.${sanitizedId}:fresh`;
        } else {
          mapping = `r[style-name='${name}'] => span.${sanitizedId}:fresh`;
        }
        break;

      case 'table':
        mapping = `table[style-name='${name}'] => table.${sanitizedId}:fresh`;
        break;

      default:
        mapping = `p[style-name='${name}'] => p.${sanitizedId}:fresh`;
    }

    return mapping;
  }

  /**
   * 构建图片转换器
   */
  private buildImageConverter(options: MammothEnhancerOptions): any {
    return mammoth.images.imgElement((image: any) => {
      return image.read('base64').then((imageBuffer: string) => {
        this.imageCounter++;
        const extension = this.getImageExtension(image.contentType);
        const imageName = `image_${this.imageCounter}.${extension}`;

        if (options.convertImagesToBase64 === true) {
          // 转换为 base64
          const dataUrl = `data:${image.contentType};base64,${imageBuffer}`;
          this.images[imageName] = dataUrl;

          return {
            src: dataUrl,
            alt: image.altText || 'Document Image',
            title: image.title || '',
          };
        } else if (options.imageOutputDir) {
          // 导入安全配置函数
          const { safePathJoin, validateAndSanitizePath } = require('../security/securityConfig');
          const allowedPaths = [options.imageOutputDir, process.cwd()];
          
          // 保存到文件并使用相对路径
          const rawImagePath = safePathJoin(options.imageOutputDir, imageName);
          const imagePath = validateAndSanitizePath(rawImagePath, allowedPaths);
          this.saveImageToFile(imagePath, imageBuffer);

          // 使用相对路径或缓存路径标识符
          const relativePath = `./images/${imageName}`;
          this.images[imageName] = imagePath; // 保存实际路径用于文件操作

          return {
            src: relativePath, // HTML中使用相对路径
            alt: image.altText || 'Document Image',
            title: image.title || '',
            'data-original-path': imagePath, // 保存原始路径作为数据属性
          };
        } else {
          // 默认使用占位符，避免base64
          const placeholderPath = `./images/${imageName}`;
          this.images[imageName] = placeholderPath;

          return {
            src: placeholderPath,
            alt: image.altText || 'Document Image',
            title: image.title || '',
            'data-image-buffer': imageBuffer, // 保存图片数据用于后续处理
            'data-content-type': image.contentType,
          };
        }
      });
    });
  }

  /**
   * 构建文档转换器
   */
  private buildDocumentTransformer(): any {
    return mammoth.transforms.paragraph((element: any) => {
      // 保留段落的对齐方式和样式信息
      if (element.alignment || element.styleName) {
        return {
          ...element,
          styleName: element.styleName || 'Normal',
        };
      }
      return element;
    });
  }

  /**
   * 后处理 HTML
   */
  private postProcessHtml(html: string, options: MammothEnhancerOptions): string {
    let processedHtml = html;

    // 1. 注入样式类名
    processedHtml = this.injectStyleClasses(processedHtml);

    // 2. 修复表格格式
    processedHtml = this.fixTableFormatting(processedHtml);

    // 3. 处理列表格式
    processedHtml = this.fixListFormatting(processedHtml);

    // 4. 清理冗余标签
    processedHtml = this.cleanupRedundantTags(processedHtml);

    // 5. 添加语义化标签
    processedHtml = this.addSemanticTags(processedHtml);

    return processedHtml;
  }

  /**
   * 注入样式类名
   */
  private injectStyleClasses(html: string): string {
    // 为没有类名的元素添加默认类名
    let processedHtml = html;

    // 为段落添加类名
    processedHtml = processedHtml.replace(/<p(?![^>]*class=)([^>]*)>/g, '<p class="normal"$1>');

    // 为表格添加类名
    processedHtml = processedHtml.replace(
      /<table(?![^>]*class=)([^>]*)>/g,
      '<table class="docx-table"$1>'
    );

    processedHtml = processedHtml.replace(
      /<td(?![^>]*class=)([^>]*)>/g,
      '<td class="docx-cell"$1>'
    );

    processedHtml = processedHtml.replace(
      /<th(?![^>]*class=)([^>]*)>/g,
      '<th class="docx-header"$1>'
    );

    return processedHtml;
  }

  /**
   * HTML转义函数
   */
  private escapeHtml(unsafe: string): string {
    return unsafe
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  /**
   * 修复表格格式
   */
  private fixTableFormatting(html: string): string {
    // 确保表格有正确的结构
    return html.replace(/<table([^>]*)>([\s\S]*?)<\/table>/g, (match, attrs, content) => {
      // 如果没有 tbody，添加它
      if (!content.includes('<tbody>')) {
        const rows = content.match(/<tr[^>]*>[\s\S]*?<\/tr>/g) || [];
        if (rows.length > 0) {
          content = `<tbody>${rows.join('')}</tbody>`;
        }
      }
      return `<table${this.escapeHtml(attrs)}>${content}</table>`;
    });
  }

  /**
   * 修复列表格式
   */
  private fixListFormatting(html: string): string {
    // 将连续的列表项包装在 ul 或 ol 中
    let processedHtml = html;

    // 查找连续的 li 元素并包装 - 修复ReDoS风险
    processedHtml = processedHtml.replace(/(<li[^>]*>[\s\S]{0,5000}?<\/li>\s*)+/g, match => {
      // 检查是否已经在列表中
      if (match.includes('<ul>') || match.includes('<ol>')) {
        return match;
      }
      return `<ul>${match}</ul>`;
    });

    return processedHtml;
  }

  /**
   * 清理冗余标签
   */
  private cleanupRedundantTags(html: string): string {
    let processedHtml = html;

    // 移除空的段落
    processedHtml = processedHtml.replace(/<p[^>]*>\s*<\/p>/g, '');

    // 移除空的 span
    processedHtml = processedHtml.replace(/<span[^>]*>\s*<\/span>/g, '');

    // 合并相邻的相同格式标签
    processedHtml = processedHtml.replace(/<\/strong>\s*<strong[^>]*>/g, '');
    processedHtml = processedHtml.replace(/<\/em>\s*<em[^>]*>/g, '');

    return processedHtml;
  }

  /**
   * 添加语义化标签
   */
  private addSemanticTags(html: string): string {
    // 这里可以根据内容添加语义化标签
    // 例如：将标题序列包装在 header 中，将主要内容包装在 main 中等
    return html;
  }

  /**
   * 生成完整的 CSS
   */
  private generateCompleteCSS(options: MammothEnhancerOptions): string {
    const cssRules: string[] = [];

    // 添加提取的基础 CSS
    cssRules.push(this.generatedCSS);

    // 添加增强样式
    cssRules.push(this.generateEnhancedStyles());

    // 添加响应式样式
    cssRules.push(this.generateResponsiveStyles());

    return cssRules.join('\n\n');
  }

  /**
   * 生成增强样式
   */
  private generateEnhancedStyles(): string {
    return `
/* 增强样式 */
.docx-table {
  border-collapse: collapse;
  width: 100%;
  margin: 12pt 0;
}

.docx-cell, .docx-header {
  border: 0.5pt solid #000000;
  padding: 4pt 8pt;
  vertical-align: top;
}

.docx-header {
  background-color: #f0f0f0;
  font-weight: bold;
}

.normal {
  margin: 0 0 8pt 0;
}

.heading1, .heading2, .heading3, .heading4, .heading5, .heading6 {
  margin: 12pt 0 6pt 0;
  font-weight: bold;
}

.title {
  font-size: 18pt;
  font-weight: bold;
  text-align: center;
  margin: 24pt 0 12pt 0;
}

.subtitle {
  font-size: 14pt;
  font-weight: bold;
  text-align: center;
  margin: 12pt 0 6pt 0;
  color: #666666;
}

.list-paragraph {
  margin: 0 0 4pt 0;
}

.strong {
  font-weight: bold;
}

.emphasis {
  font-style: italic;
}

.hyperlink {
  color: #0066cc;
  text-decoration: underline;
}

.subtle-emphasis {
  font-style: italic;
  color: #666666;
}

.intense-emphasis {
  font-weight: bold;
  color: #000000;
}
`;
  }

  /**
   * 生成响应式样式
   */
  private generateResponsiveStyles(): string {
    return `
/* 响应式样式 */
@media screen and (max-width: 768px) {
  .docx-table {
    font-size: 10pt;
  }
  
  .docx-cell, .docx-header {
    padding: 2pt 4pt;
  }
  
  .title {
    font-size: 16pt;
  }
  
  .subtitle {
    font-size: 12pt;
  }
}

@media print {
  .docx-table {
    page-break-inside: avoid;
  }
  
  .heading1, .heading2, .heading3 {
    page-break-after: avoid;
  }
}
`;
  }

  /**
   * 保存图片到文件
   */
  private async saveImageToFile(imagePath: string, imageBuffer: string): Promise<void> {
    try {
      const buffer = Buffer.from(imageBuffer, 'base64');
      await fs.writeFile(imagePath, buffer);
      console.log(`💾 图片已保存: ${imagePath}`);
    } catch (error: any) {
      console.error(`❌ 保存图片失败: ${error.message}`);
    }
  }

  /**
   * 获取图片扩展名
   */
  private getImageExtension(contentType: string): string {
    const typeMap: { [key: string]: string } = {
      'image/jpeg': 'jpg',
      'image/jpg': 'jpg',
      'image/png': 'png',
      'image/gif': 'gif',
      'image/bmp': 'bmp',
      'image/webp': 'webp',
      'image/svg+xml': 'svg',
    };

    return typeMap[contentType] || 'png';
  }

  /**
   * 提取标题级别
   */
  private extractHeadingLevel(styleName: string): number {
    const match = styleName.match(/(\d+)/);
    return match ? Math.min(parseInt(match[1]), 6) : 1;
  }

  /**
   * 获取样式统计信息
   */
  getStyleStats(): {
    totalStyles: number;
    paragraphStyles: number;
    characterStyles: number;
    tableStyles: number;
    documentStyles: number;
  } {
    const paragraphStyles = Array.from(this.extractedStyles.values()).filter(
      style => style.type === 'paragraph'
    ).length;
    const characterStyles = Array.from(this.extractedStyles.values()).filter(
      style => style.type === 'character'
    ).length;
    const tableStyles = Array.from(this.extractedStyles.values()).filter(
      style => style.type === 'table'
    ).length;

    return {
      totalStyles: this.extractedStyles.size,
      paragraphStyles,
      characterStyles,
      tableStyles,
      documentStyles: this.documentStyles.length,
    };
  }
}

export { MammothEnhancerOptions, ConversionResult };
