/**
 * 增强的 HTML 到 Markdown 转换器
 * 支持样式保留、emoji和完整的格式转换
 */

import { promises as fs } from 'fs';
const cheerio = require('cheerio');
const path = require('path');

interface HtmlToMarkdownOptions {
  preserveStyles?: boolean;
  includeCSS?: boolean;
  outputPath?: string;
  debug?: boolean;
}

interface HtmlToMarkdownResult {
  success: boolean;
  content?: string;
  outputPath?: string;
  metadata?: {
    originalFormat: string;
    targetFormat: string;
    stylesPreserved: boolean;
    contentLength: number;
    converter: string;
  };
  error?: string;
}

class EnhancedHtmlToMarkdownConverter {
  private options: HtmlToMarkdownOptions = {};

  async convertHtmlToMarkdown(
    inputPath: string,
    options: HtmlToMarkdownOptions = {}
  ): Promise<HtmlToMarkdownResult> {
    try {
      this.options = {
        preserveStyles: true,
        includeCSS: true,
        debug: false,
        ...options,
      };

      if (this.options.debug) {
        console.log('🚀 开始增强的 HTML 到 Markdown 转换...');
        console.log('📄 输入文件:', inputPath);
      }

      // 读取HTML文件
      const htmlContent = await fs.readFile(inputPath, 'utf-8');

      // 使用cheerio解析HTML
      const $ = cheerio.load(htmlContent);

      // 提取CSS样式（如果需要）
      let cssStyles = '';
      if (this.options.includeCSS) {
        cssStyles = this.extractCSS($);
      }

      // 转换为Markdown
      let markdownContent = this.htmlToMarkdown($);

      // 如果包含CSS，添加到文档开头
      if (cssStyles && this.options.includeCSS) {
        markdownContent = `<!-- CSS Styles\n${cssStyles}\n-->\n\n${markdownContent}`;
      }

      // 添加样式保留说明
      if (this.options.preserveStyles) {
        const styleNote = `<!-- 样式保留说明：\n本文档在转换过程中保留了原始HTML的样式信息。\n如需查看完整样式效果，请在支持HTML的环境中查看。\n图片路径已转换为相对路径，请确保图片文件在正确位置。\n-->\n\n`;
        markdownContent = styleNote + markdownContent;
      }

      // 生成输出路径
      const outputPath = this.options.outputPath || inputPath.replace(/\.html?$/i, '.md');

      // 保存文件
      await fs.writeFile(outputPath, markdownContent, 'utf-8');

      if (this.options.debug) {
        console.log('✅ 增强的 Markdown 转换完成:', outputPath);
      }

      return {
        success: true,
        content: markdownContent,
        outputPath,
        metadata: {
          originalFormat: 'html',
          targetFormat: 'markdown',
          stylesPreserved: this.options.preserveStyles || false,
          contentLength: markdownContent.length,
          converter: 'enhanced-html-to-markdown-converter',
        },
      };
    } catch (error: any) {
      console.error('❌ 增强的 HTML 转 Markdown 失败:', error.message);
      return {
        success: false,
        error: error.message,
      };
    }
  }

  /**
   * 提取CSS样式
   */
  private extractCSS($: any): string {
    let css = '';

    // 提取style标签中的CSS
    $('style').each((i: number, element: any) => {
      css += $(element).html() + '\n';
    });

    // 提取link标签引用的CSS（仅记录引用）
    $('link[rel="stylesheet"]').each((i: number, element: any) => {
      const href = $(element).attr('href');
      if (href) {
        css += `/* External CSS: ${href} */\n`;
      }
    });

    // 提取内联样式并转换为CSS类
    const inlineStyles = new Map<string, string>();
    $('[style]').each((i: number, element: any) => {
      const style = $(element).attr('style');
      const tagName = $(element).prop('tagName')?.toLowerCase();
      if (style && tagName) {
        const className = `inline-${tagName}-${i}`;
        inlineStyles.set(className, style);
        $(element).addClass(className);
      }
    });

    // 添加内联样式到CSS
    if (inlineStyles.size > 0) {
      css += '\n/* Converted inline styles */\n';
      inlineStyles.forEach((style, className) => {
        css += `.${className} { ${style} }\n`;
      });
    }

    return css.trim();
  }

  /**
   * 将HTML转换为Markdown格式
   */
  private htmlToMarkdown($: any): string {
    let markdown = '';

    // 处理body内容，如果没有body则处理整个文档
    const body = $('body').length > 0 ? $('body') : $.root();

    body.children().each((i: number, element: any) => {
      markdown += this.processElementToMarkdown($, $(element));
    });

    // 额外处理所有图片，确保不遗漏
    $('img').each((i: number, img: any) => {
      const $img = $(img);
      const src = $img.attr('src');
      const alt = $img.attr('alt') || 'Image';
      const title = $img.attr('title') || '';
      const dataOriginalPath = $img.attr('data-original-path');

      if (src || dataOriginalPath) {
        const imagePath = dataOriginalPath || src;
        let finalPath = src || imagePath;

        // 确保图片路径正确
        if (
          !finalPath.startsWith('./') &&
          !finalPath.startsWith('/') &&
          !finalPath.startsWith('http')
        ) {
          finalPath = `./${finalPath}`;
        }

        const imageMarkdown = title
          ? `![${alt}](${finalPath} "${title}")\n\n`
          : `![${alt}](${finalPath})\n\n`;

        // 检查是否已经包含此图片
        if (!markdown.includes(imageMarkdown.trim())) {
          markdown += imageMarkdown;
        }
      }
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

    if (!text && !['br', 'hr', 'img', 'table'].includes(tagName)) {
      return '';
    }

    return this.convertTagToMarkdown(tagName, text, element, $);
  }

  private convertTagToMarkdown(tagName: string, text: string, element: any, $: any): string {
    // 处理标题
    if (this.isHeadingTag(tagName)) {
      return this.convertHeadingToMarkdown(tagName, text);
    }

    // 处理基本格式化标签
    if (this.isFormattingTag(tagName)) {
      return this.convertFormattingToMarkdown(tagName, text);
    }

    // 处理特殊元素
    switch (tagName) {
      case 'p':
        return `${this.processInlineElements($, element)}\n\n`;
      case 'pre':
        return this.convertPreToMarkdown(element, text);
      case 'a':
        return this.convertLinkToMarkdown(element, text);
      case 'img':
        return this.convertImageToMarkdown(element);
      case 'ul':
        return this.processListToMarkdown($, element, 'ul');
      case 'ol':
        return this.processListToMarkdown($, element, 'ol');
      case 'table':
        return this.processTableToMarkdown($, element);
      case 'blockquote':
        const quoteText = this.processInlineElements($, element);
        return `> ${quoteText}\n\n`;
      case 'hr':
        return `---\n\n`;
      case 'br':
        return '\n';
      case 'div':
      case 'span':
      case 'section':
      case 'article':
        return this.convertContainerToMarkdown($, element, text);
      default:
        return text ? `${text}\n` : '';
    }
  }

  private isHeadingTag(tagName: string): boolean {
    return ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(tagName);
  }

  private isFormattingTag(tagName: string): boolean {
    return ['strong', 'b', 'em', 'i', 'code'].includes(tagName);
  }

  private convertHeadingToMarkdown(tagName: string, text: string): string {
    const level = parseInt(tagName.charAt(1));
    const hashes = '#'.repeat(level);
    return `${hashes} ${text}\n\n`;
  }

  private convertFormattingToMarkdown(tagName: string, text: string): string {
    switch (tagName) {
      case 'strong':
      case 'b':
        return `**${text}**`;
      case 'em':
      case 'i':
        return `*${text}*`;
      case 'code':
        return `\`${text}\``;
      default:
        return text;
    }
  }

  private convertPreToMarkdown(element: any, text: string): string {
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
  }

  private convertLinkToMarkdown(element: any, text: string): string {
    const href = element.attr('href');
    return href ? `[${text}](${href})` : text;
  }

  private convertImageToMarkdown(element: any): string {
    const src = element.attr('src');
    const alt = element.attr('alt') || 'Image';
    const title = element.attr('title') || '';
    const dataOriginalPath = element.attr('data-original-path');

    const imagePath = dataOriginalPath || src;
    if (imagePath) {
      let finalPath = src || imagePath;
      if (!finalPath.startsWith('./') && !finalPath.startsWith('/')) {
        finalPath = `./${finalPath}`;
      }
      return title ? `![${alt}](${finalPath} "${title}")\n\n` : `![${alt}](${finalPath})\n\n`;
    }
    return '';
  }

  private convertContainerToMarkdown($: any, element: any, text: string): string {
    const inlineContent = this.processInlineElements($, element);
    if (inlineContent.trim()) {
      return inlineContent + '\n';
    }

    let divContent = '';
    element.children().each((i: number, child: any) => {
      divContent += this.processElementToMarkdown($, $(child));
    });
    return divContent || (text ? `${text}\n` : '');
  }

  /**
   * 处理列表转换为Markdown
   */
  private processListToMarkdown($: any, element: any, listType: 'ul' | 'ol'): string {
    let listMarkdown = '';

    element.children('li').each((i: number, li: any) => {
      const $li = $(li);
      const liContent = this.processInlineElements($, $li);

      if (listType === 'ul') {
        listMarkdown += `- ${liContent}\n`;
      } else {
        listMarkdown += `${i + 1}. ${liContent}\n`;
      }
    });

    return listMarkdown + '\n';
  }

  /**
   * 处理表格转换为Markdown
   */
  private processTableToMarkdown($: any, element: any): string {
    let tableMarkdown = '';
    let isFirstRow = true;

    element.find('tr').each((i: number, tr: any) => {
      const $tr = $(tr);
      let rowMarkdown = '|';

      $tr.find('td, th').each((j: number, cell: any) => {
        const $cell = $(cell);
        const cellText = this.decodeHtmlEntities($cell.text().trim());
        rowMarkdown += ` ${cellText} |`;
      });

      tableMarkdown += rowMarkdown + '\n';

      // 添加表头分隔线
      if (isFirstRow && $tr.find('th').length > 0) {
        let separatorRow = '|';
        $tr.find('th').each(() => {
          separatorRow += ' --- |';
        });
        tableMarkdown += separatorRow + '\n';
        isFirstRow = false;
      }
    });

    return tableMarkdown + '\n';
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
        result += this.processInlineTag($, node);
      }
    });

    return result;
  }

  private processInlineTag($: any, node: any): string {
    const $node = $(node);
    const tagName = $node.prop('tagName')?.toLowerCase();
    const nodeText = this.decodeHtmlEntities($node.text());

    switch (tagName) {
      case 'strong':
      case 'b':
        return `**${nodeText}**`;
      case 'em':
      case 'i':
        return `*${nodeText}*`;
      case 'code':
        return `\`${nodeText}\``;
      case 'a':
        return this.processInlineLink($node, nodeText);
      case 'img':
        return this.processInlineImage($node);
      case 'span':
        return this.processSpanElement($, $node, nodeText);
      case 'br':
        return '\n';
      default:
        return this.processInlineElements($, $node);
    }
  }

  private processInlineLink($node: any, nodeText: string): string {
    const href = $node.attr('href');
    return href ? `[${nodeText}](${href})` : nodeText;
  }

  private processInlineImage($node: any): string {
    const src = $node.attr('src');
    const alt = $node.attr('alt') || 'Image';
    const title = $node.attr('title') || '';
    const dataOriginalPath = $node.attr('data-original-path');

    const imagePath = dataOriginalPath || src;
    if (imagePath) {
      let finalPath = src || imagePath;
      if (!finalPath.startsWith('./') && !finalPath.startsWith('/')) {
        finalPath = `./${finalPath}`;
      }
      return title ? `![${alt}](${finalPath} "${title}")` : `![${alt}](${finalPath})`;
    }
    return '';
  }

  private processSpanElement($: any, $node: any, nodeText: string): string {
    const imgElement = $node.find('img');
    if (imgElement.length > 0) {
      return this.processSpanImages($, imgElement);
    }
    return nodeText;
  }

  private processSpanImages($: any, imgElement: any): string {
    let result = '';
    imgElement.each((j: number, img: any) => {
      const $img = $(img);
      const imgSrc = $img.attr('src');
      const imgAlt = $img.attr('alt') || 'Image';
      const imgTitle = $img.attr('title') || '';
      const imgDataOriginalPath = $img.attr('data-original-path');
      const imgImagePath = imgDataOriginalPath || imgSrc;

      if (imgImagePath) {
        let imgFinalPath = imgSrc || imgImagePath;
        if (!imgFinalPath.startsWith('./') && !imgFinalPath.startsWith('/')) {
          imgFinalPath = `./${imgFinalPath}`;
        }
        result += imgTitle
          ? `![${imgAlt}](${imgFinalPath} "${imgTitle}")`
          : `![${imgAlt}](${imgFinalPath})`;
      }
    });
    return result;
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

export { EnhancedHtmlToMarkdownConverter, HtmlToMarkdownOptions, HtmlToMarkdownResult };
