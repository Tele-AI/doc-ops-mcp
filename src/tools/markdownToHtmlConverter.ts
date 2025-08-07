/**
 * Markdown 到 HTML 转换器 - 支持样式保留和美化
 * 解决 Markdown 转 HTML 时样式丢失的问题
 */

const fs = require('fs').promises;
const path = require('path');
const marked = require('marked');
const cheerio = require('cheerio');

// 转换选项接口
interface MarkdownToHtmlOptions {
  preserveStyles?: boolean;
  theme?: 'default' | 'github' | 'academic' | 'modern';
  includeTableOfContents?: boolean;
  enableSyntaxHighlighting?: boolean;
  customCSS?: string;
  outputPath?: string;
  standalone?: boolean;
  debug?: boolean;
}

// 转换结果接口
interface MarkdownConversionResult {
  success: boolean;
  content?: string;
  htmlPath?: string;
  cssPath?: string;
  metadata?: {
    originalFormat: string;
    targetFormat: string;
    stylesPreserved: boolean;
    theme: string;
    converter: string;
    contentLength: number;
    headingsCount: number;
    linksCount: number;
    imagesCount: number;
  };
  error?: string;
}

/**
 * Markdown 到 HTML 转换器类
 */
class MarkdownToHtmlConverter {
  private options: MarkdownToHtmlOptions = {};
  private themes: Map<string, string>;

  constructor() {
    this.themes = new Map();
    this.initializeThemes();
  }

  /**
   * 初始化预设主题
   */
  private initializeThemes(): void {
    // GitHub 风格主题
    this.themes.set(
      'github',
      `
      body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans', Helvetica, Arial, sans-serif;
        font-size: 16px;
        line-height: 1.5;
        color: #24292f;
        background-color: #ffffff;
        max-width: 980px;
        margin: 0 auto;
        padding: 45px;
      }
      
      h1, h2, h3, h4, h5, h6 {
        margin-top: 24px;
        margin-bottom: 16px;
        font-weight: 600;
        line-height: 1.25;
      }
      
      h1 {
        font-size: 2em;
        border-bottom: 1px solid #d0d7de;
        padding-bottom: 0.3em;
      }
      
      h2 {
        font-size: 1.5em;
        border-bottom: 1px solid #d0d7de;
        padding-bottom: 0.3em;
      }
      
      h3 { font-size: 1.25em; }
      h4 { font-size: 1em; }
      h5 { font-size: 0.875em; }
      h6 { font-size: 0.85em; color: #656d76; }
      
      p {
        margin-top: 0;
        margin-bottom: 16px;
      }
      
      blockquote {
        padding: 0 1em;
        color: #656d76;
        border-left: 0.25em solid #d0d7de;
        margin: 0 0 16px 0;
      }
      
      ul, ol {
        margin-top: 0;
        margin-bottom: 16px;
        padding-left: 2em;
      }
      
      li {
        margin-bottom: 0.25em;
      }
      
      code {
        padding: 0.2em 0.4em;
        margin: 0;
        font-size: 85%;
        background-color: rgba(175,184,193,0.2);
        border-radius: 6px;
        font-family: ui-monospace, SFMono-Regular, 'SF Mono', Consolas, 'Liberation Mono', Menlo, monospace;
      }
      
      pre {
        padding: 16px;
        overflow: auto;
        font-size: 85%;
        line-height: 1.45;
        background-color: #f6f8fa;
        border-radius: 6px;
        margin-bottom: 16px;
      }
      
      pre code {
        background-color: transparent;
        border: 0;
        padding: 0;
        margin: 0;
        font-size: 100%;
      }
      
      table {
        border-spacing: 0;
        border-collapse: collapse;
        margin-bottom: 16px;
        width: 100%;
      }
      
      table th, table td {
        padding: 6px 13px;
        border: 1px solid #d0d7de;
      }
      
      table th {
        font-weight: 600;
        background-color: #f6f8fa;
      }
      
      table tr:nth-child(2n) {
        background-color: #f6f8fa;
      }
      
      a {
        color: #0969da;
        text-decoration: none;
      }
      
      a:hover {
        text-decoration: underline;
      }
      
      img {
        max-width: 100%;
        height: auto;
        margin: 16px 0;
      }
      
      hr {
        height: 0.25em;
        padding: 0;
        margin: 24px 0;
        background-color: #d0d7de;
        border: 0;
      }
    `
    );

    // 学术风格主题
    this.themes.set(
      'academic',
      `
      body {
        font-family: 'Times New Roman', Times, serif;
        font-size: 12pt;
        line-height: 1.6;
        color: #000000;
        background-color: #ffffff;
        max-width: 210mm;
        margin: 0 auto;
        padding: 25mm;
      }
      
      h1, h2, h3, h4, h5, h6 {
        font-family: 'Times New Roman', Times, serif;
        font-weight: bold;
        margin-top: 18pt;
        margin-bottom: 12pt;
        text-align: left;
      }
      
      h1 {
        font-size: 18pt;
        text-align: center;
        margin-bottom: 24pt;
      }
      
      h2 {
        font-size: 14pt;
        margin-top: 24pt;
      }
      
      h3 {
        font-size: 12pt;
        font-style: italic;
      }
      
      p {
        margin: 0 0 12pt 0;
        text-align: justify;
        text-indent: 2em;
      }
      
      blockquote {
        margin: 12pt 2em;
        padding: 0;
        font-style: italic;
        border-left: none;
      }
      
      ul, ol {
        margin: 12pt 0;
        padding-left: 2em;
      }
      
      li {
        margin-bottom: 6pt;
      }
      
      table {
        border-collapse: collapse;
        margin: 12pt auto;
        width: 100%;
      }
      
      table th, table td {
        border: 1pt solid #000;
        padding: 6pt;
        text-align: left;
      }
      
      table th {
        font-weight: bold;
        background-color: #f0f0f0;
      }
      
      code {
        font-family: 'Courier New', Courier, monospace;
        font-size: 10pt;
      }
      
      pre {
        font-family: 'Courier New', Courier, monospace;
        font-size: 10pt;
        margin: 12pt 0;
        padding: 12pt;
        border: 1pt solid #ccc;
        background-color: #f9f9f9;
      }
      
      a {
        color: #000;
        text-decoration: underline;
      }
      
      img {
        max-width: 100%;
        height: auto;
        display: block;
        margin: 12pt auto;
      }
    `
    );

    // 现代风格主题
    this.themes.set(
      'modern',
      `
      body {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        font-size: 16px;
        line-height: 1.7;
        color: #2d3748;
        background-color: #ffffff;
        max-width: 768px;
        margin: 0 auto;
        padding: 2rem;
      }
      
      h1, h2, h3, h4, h5, h6 {
        font-weight: 700;
        margin-top: 2rem;
        margin-bottom: 1rem;
        color: #1a202c;
      }
      
      h1 {
        font-size: 2.5rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 2rem;
      }
      
      h2 {
        font-size: 2rem;
        border-left: 4px solid #667eea;
        padding-left: 1rem;
      }
      
      h3 {
        font-size: 1.5rem;
        color: #4a5568;
      }
      
      p {
        margin-bottom: 1.5rem;
        color: #4a5568;
      }
      
      blockquote {
        border-left: 4px solid #e2e8f0;
        padding: 1rem 1.5rem;
        margin: 1.5rem 0;
        background-color: #f7fafc;
        border-radius: 0 8px 8px 0;
        font-style: italic;
      }
      
      ul, ol {
        margin: 1.5rem 0;
        padding-left: 2rem;
      }
      
      li {
        margin-bottom: 0.5rem;
        color: #4a5568;
      }
      
      code {
        background-color: #edf2f7;
        color: #e53e3e;
        padding: 0.25rem 0.5rem;
        border-radius: 4px;
        font-family: 'Fira Code', 'Consolas', monospace;
        font-size: 0.875rem;
      }
      
      pre {
        background-color: #2d3748;
        color: #e2e8f0;
        padding: 1.5rem;
        border-radius: 8px;
        overflow-x: auto;
        margin: 1.5rem 0;
      }
      
      pre code {
        background-color: transparent;
        color: inherit;
        padding: 0;
      }
      
      table {
        width: 100%;
        border-collapse: collapse;
        margin: 1.5rem 0;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        border-radius: 8px;
        overflow: hidden;
      }
      
      table th, table td {
        padding: 1rem;
        text-align: left;
        border-bottom: 1px solid #e2e8f0;
      }
      
      table th {
        background-color: #f7fafc;
        font-weight: 600;
        color: #2d3748;
      }
      
      table tr:hover {
        background-color: #f7fafc;
      }
      
      a {
        color: #667eea;
        text-decoration: none;
        border-bottom: 1px solid transparent;
        transition: border-color 0.2s;
      }
      
      a:hover {
        border-bottom-color: #667eea;
      }
      
      img {
        max-width: 100%;
        height: auto;
        border-radius: 8px;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
      }
      
      hr {
        border: none;
        height: 2px;
        background: linear-gradient(90deg, #667eea, #764ba2);
        margin: 2rem 0;
        border-radius: 1px;
      }
    `
    );

    // 默认主题
    this.themes.set(
      'default',
      `
      body {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        font-size: 16px;
        line-height: 1.6;
        color: #333;
        background-color: #fff;
        max-width: 800px;
        margin: 0 auto;
        padding: 2rem;
      }
      
      h1, h2, h3, h4, h5, h6 {
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        font-weight: bold;
      }
      
      h1 { font-size: 2.5rem; color: #2c3e50; }
      h2 { font-size: 2rem; color: #34495e; }
      h3 { font-size: 1.5rem; color: #34495e; }
      h4 { font-size: 1.25rem; }
      h5 { font-size: 1rem; }
      h6 { font-size: 0.875rem; }
      
      p {
        margin-bottom: 1rem;
      }
      
      blockquote {
        border-left: 4px solid #3498db;
        padding-left: 1rem;
        margin: 1rem 0;
        color: #7f8c8d;
        font-style: italic;
      }
      
      ul, ol {
        margin: 1rem 0;
        padding-left: 2rem;
      }
      
      code {
        background-color: #f8f9fa;
        padding: 0.25rem 0.5rem;
        border-radius: 3px;
        font-family: 'Consolas', 'Monaco', monospace;
      }
      
      pre {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 5px;
        overflow-x: auto;
        margin: 1rem 0;
      }
      
      table {
        border-collapse: collapse;
        width: 100%;
        margin: 1rem 0;
      }
      
      table th, table td {
        border: 1px solid #ddd;
        padding: 0.75rem;
        text-align: left;
      }
      
      table th {
        background-color: #f8f9fa;
        font-weight: bold;
      }
      
      a {
        color: #3498db;
        text-decoration: none;
      }
      
      a:hover {
        text-decoration: underline;
      }
      
      img {
        max-width: 100%;
        height: auto;
        margin: 1rem 0;
      }
    `
    );
  }

  /**
   * 主转换函数
   */
  async convertMarkdownToHtml(
    inputPath: string,
    options: MarkdownToHtmlOptions = {}
  ): Promise<MarkdownConversionResult> {
    try {
      this.options = {
        preserveStyles: true,
        theme: 'default',
        includeTableOfContents: false,
        enableSyntaxHighlighting: false,
        standalone: true,
        debug: false,
        ...options,
      };

      if (this.options.debug) {
        console.log('🚀 开始 Markdown 到 HTML 转换...');
        console.log('📄 输入文件:', inputPath);
        console.log('🎨 使用主题:', this.options.theme);
      }

      // 读取 Markdown 文件
      const markdownContent = await fs.readFile(inputPath, 'utf-8');

      // 配置 marked
      this.configureMarked();

      // 转换为 HTML
      const htmlContent = marked.parse(markdownContent);

      // 分析内容统计
      const stats = this.analyzeContent(htmlContent);

      // 生成完整的 HTML 文档
      const completeHtml = this.generateCompleteHtml(htmlContent);

      // 保存文件（如果指定了输出路径）
      let htmlPath: string | undefined;
      if (this.options.outputPath) {
        const { validateAndSanitizePath } = require('../security/securityConfig');
        const allowedPaths = [process.cwd()];
        htmlPath = validateAndSanitizePath(this.options.outputPath, allowedPaths);
        await fs.writeFile(htmlPath, completeHtml, 'utf-8');

        if (this.options.debug) {
          console.log('✅ HTML 文件已保存:', htmlPath);
        }
      }

      if (this.options.debug) {
        console.log('📊 转换统计:', stats);
        console.log('✅ Markdown 转换完成');
      }

      return {
        success: true,
        content: completeHtml,
        htmlPath,
        metadata: {
          originalFormat: 'markdown',
          targetFormat: 'html',
          stylesPreserved: this.options.preserveStyles ?? false,
      theme: this.options.theme ?? 'default',
          converter: 'markdown-to-html-converter',
          contentLength: completeHtml.length,
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
   * 配置 marked 解析器
   */
  private configureMarked(): void {
    marked.setOptions({
      gfm: true, // GitHub Flavored Markdown
      breaks: true, // 支持换行
      pedantic: false,
      sanitize: false,
      smartLists: true,
      smartypants: true,
    });

    // 使用更简单的配置，避免自定义渲染器的问题
    // 暂时移除自定义渲染器，使用默认的marked渲染
  }

  /**
   * 分析内容统计信息
   */
  private analyzeContent(htmlContent: string): {
    headingsCount: number;
    linksCount: number;
    imagesCount: number;
  } {
    const $ = cheerio.load(htmlContent);

    return {
      headingsCount: $('h1, h2, h3, h4, h5, h6').length,
      linksCount: $('a').length,
      imagesCount: $('img').length,
    };
  }

  /**
   * 生成完整的 HTML 文档
   */
  private generateCompleteHtml(content: string): string {
    if (!this.options.standalone) {
      return content;
    }

    const theme = this.options.theme ?? 'default';
    const themeCSS = this.themes.get(theme) ?? this.themes.get('default')!;
    const customCSS = this.options.customCSS ?? '';

    // 生成目录（如果启用）
    let tocHtml = '';
    if (this.options.includeTableOfContents) {
      tocHtml = this.generateTableOfContents(content);
    }

    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Markdown Document</title>
    <style>
${themeCSS}
${customCSS}
    </style>
</head>
<body>
${tocHtml}
${content}
</body>
</html>`;
  }

  /**
   * 生成目录
   */
  private generateTableOfContents(content: string): string {
    const $ = cheerio.load(content);
    const headings = $('h1, h2, h3, h4, h5, h6');

    if (headings.length === 0) {
      return '';
    }

    let toc = '<div class="table-of-contents">\n<h2>目录</h2>\n<ul>\n';

    headings.each((index, element) => {
      const $heading = $(element);
      const level = parseInt(element.tagName.substring(1));
      const text = $heading.text();
      const id = $heading.attr('id') ?? text.toLowerCase().replace(/[^\w]+/g, '-');

      const indent = '  '.repeat(level - 1);
      toc += `${indent}<li><a href="#${id}">${text}</a></li>\n`;
    });

    toc += '</ul>\n</div>\n\n';
    return toc;
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
  addCustomTheme(name: string, css: string): void {
    this.themes.set(name, css);
  }
}

// 导出便捷函数
export async function convertMarkdownToHtml(
  inputPath: string,
  options: MarkdownToHtmlOptions = {}
): Promise<MarkdownConversionResult> {
  const converter = new MarkdownToHtmlConverter();
  return await converter.convertMarkdownToHtml(inputPath, options);
}

export { MarkdownToHtmlConverter, MarkdownToHtmlOptions, MarkdownConversionResult };
