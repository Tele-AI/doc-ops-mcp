/**
 * 优化的 DOCX 转换器
 * 重构转换逻辑，专注于样式完整性和格式还原
 * 转换路径：DOCX -> 样式完整的 HTML/MD -> PDF
 */

import { SafeErrorHandler } from '../security/errorHandler';
import { DualParsingEngine, DualParsingOptions, DualParsingResult } from './dualParsingEngine';
import { dualParsingDocxToHtml } from './dualParsingDocxTools';
const fs = require('fs/promises');
const path = require('path');
const os = require('os');

interface OptimizedConversionOptions {
  // 输出格式选项
  outputFormat?: 'html' | 'markdown' | 'pdf';

  // 样式保留选项
  preserveStyles?: boolean;
  preserveImages?: boolean;
  preserveLayout?: boolean;

  // 输出路径选项
  outputPath?: string;
  htmlOutputPath?: string;
  markdownOutputPath?: string;

  // 调试选项
  debug?: boolean;
  saveIntermediateFiles?: boolean;

  // PDF 转换选项（当输出格式为 PDF 时）
  pdfOptions?: {
    format?: 'A4' | 'A3' | 'A5' | 'Letter' | 'Legal';
    margin?: string;
    printBackground?: boolean;
    landscape?: boolean;
  };
}

interface ConversionResult {
  success: boolean;
  outputPath?: string;
  htmlPath?: string;
  markdownPath?: string;

  // 转换详情
  details: {
    originalFormat: string;
    targetFormat: string;
    stylesPreserved: boolean;
    imagesPreserved: boolean;
    conversionTime: number;
  };

  // 如果需要外部工具完成转换
  requiresExternalTool?: boolean;
  externalToolInstructions?: string;

  error?: string;
}

export class OptimizedDocxConverter {
  private dualParsingEngine: DualParsingEngine;

  constructor() {
    // 初始化双重解析引擎，配置最佳样式保留选项
    const dualParsingOptions: DualParsingOptions = {
      extractStyles: true,
      mammothOptions: {
        preserveImages: true,
        convertImagesToBase64: false, // 保存为文件以避免 token 限制
        includeDefaultStyles: true,
        transformDocument: true,
      },
      cssOptions: {
        minify: false,
        includeComments: true,
        generateResponsive: true,
        generatePrintStyles: true,
      },
      postProcessOptions: {
        injectStyles: true,
        optimizeStructure: true,
        fixFormatting: true,
        addSemanticTags: true,
        removeEmptyElements: true,
        mergeAdjacentElements: true,
        addAccessibility: true,
      },
      outputOptions: {
        includeCSS: true,
        inlineCSS: false, // 先保持分离，便于调试
        generateCompleteHTML: true,
        preserveOriginalStructure: false,
        addDocumentMetadata: true,
      },
      debugOptions: {
        logProgress: true,
        saveIntermediateResults: true,
      },
    };

    this.dualParsingEngine = new DualParsingEngine(dualParsingOptions);
  }

  /**
   * 主要转换方法
   */
  async convertDocx(
    inputPath: string,
    options: OptimizedConversionOptions = {}
  ): Promise<ConversionResult> {
    const startTime = Date.now();

    try {
      console.log('🚀 开始优化的 DOCX 转换...');
      console.log(`📄 输入文件: ${inputPath}`);
      console.log(`🎯 目标格式: ${options.outputFormat ?? 'html'}`);

      // 验证输入文件
      await this.validateInput(inputPath);

      // 第一步：使用双重解析引擎转换为高质量 HTML
      console.log('\n📖 第一步：转换为样式完整的 HTML...');
      const htmlResult = await this.convertToStyledHtml(inputPath, options);

      if (!htmlResult.success) {
        throw new Error(`HTML 转换失败: ${htmlResult.error}`);
      }

      console.log(`✅ HTML 转换成功: ${path.basename(htmlResult.htmlPath ?? 'output.html')}`);

      // 根据目标格式进行后续处理
      const targetFormat = options.outputFormat ?? 'html';

      switch (targetFormat) {
        case 'html':
          return this.finalizeHtmlOutput(htmlResult, options, startTime);

        case 'markdown':
          return await this.convertHtmlToMarkdown(htmlResult, options, startTime);

        case 'pdf':
          return await this.convertHtmlToPdf(htmlResult, options, startTime);

        default:
          throw new Error(`不支持的输出格式: ${targetFormat}`);
      }
    } catch (error: any) {
      SafeErrorHandler.logError('转换失败', error);
      return {
        success: false,
        error: SafeErrorHandler.sanitizeErrorMessage(error),
        details: {
          originalFormat: 'docx',
          targetFormat: options.outputFormat ?? 'html',
          stylesPreserved: false,
          imagesPreserved: false,
          conversionTime: Date.now() - startTime,
        },
      };
    }
  }

  /**
   * 转换为样式完整的 HTML
   */
  private async convertToStyledHtml(
    inputPath: string,
    options: OptimizedConversionOptions
  ): Promise<{
    success: boolean;
    htmlPath?: string;
    cssPath?: string;
    error?: string;
    dualParsingResult?: DualParsingResult;
  }> {
    try {
      // 使用双重解析引擎
      const result = await this.dualParsingEngine.convertDocxToHtml(inputPath);

      if (!result.success) {
        throw new Error(result.error ?? '双重解析引擎转换失败');
      }

      // 导入安全配置函数
      const { safePathJoin, validateAndSanitizePath } = require('../security/securityConfig');
      
      // 生成输出路径
      const outputDir = options.htmlOutputPath
        ? path.dirname(options.htmlOutputPath)
        : safePathJoin(process.cwd(), 'output');
      
      const allowedPaths = [outputDir, process.cwd()];

      await fs.mkdir(outputDir, { recursive: true });

      const rawHtmlPath =
        options.htmlOutputPath ??
        safePathJoin(outputDir, `${path.basename(inputPath, '.docx')}_styled.html`);
      const htmlPath = validateAndSanitizePath(rawHtmlPath, allowedPaths);

      const rawCssPath = safePathJoin(outputDir, `${path.basename(inputPath, '.docx')}_styles.css`);
      const cssPath = validateAndSanitizePath(rawCssPath, allowedPaths);

      // 保存 HTML 文件
      await fs.writeFile(htmlPath, result.completeHTML, 'utf8');

      // 保存 CSS 文件（如果需要）
      if (result.css && options.saveIntermediateFiles) {
        await fs.writeFile(cssPath, result.css, 'utf8');
      }

      console.log(`📊 转换统计:`);
      console.log(`  - 样式数量: ${result.details.styleExtraction.totalStyles}`);
      console.log(`  - 媒体文件: ${result.details.mediaExtraction.totalFiles}`);
      console.log(`  - CSS 规则: ${result.details.cssGeneration.totalRules}`);
      console.log(`  - 转换时间: ${result.performance.totalTime}ms`);

      return {
        success: true,
        htmlPath,
        cssPath: options.saveIntermediateFiles ? cssPath : undefined,
        dualParsingResult: result,
      };
    } catch (error: any) {
      SafeErrorHandler.logError('HTML 转换失败', error);
      return {
        success: false,
        error: SafeErrorHandler.sanitizeErrorMessage(error),
      };
    }
  }

  /**
   * 完成 HTML 输出
   */
  private finalizeHtmlOutput(
    htmlResult: any,
    options: OptimizedConversionOptions,
    startTime: number
  ): ConversionResult {
    const finalPath = options.outputPath ?? htmlResult.htmlPath;

    return {
      success: true,
      outputPath: finalPath,
      htmlPath: htmlResult.htmlPath,
      details: {
        originalFormat: 'docx',
        targetFormat: 'html',
        stylesPreserved: true,
        imagesPreserved: true,
        conversionTime: Date.now() - startTime,
      },
    };
  }

  /**
   * 转换 HTML 为 Markdown
   */
  private async convertHtmlToMarkdown(
    htmlResult: any,
    options: OptimizedConversionOptions,
    startTime: number
  ): Promise<ConversionResult> {
    try {
      console.log('\n📝 第二步：转换 HTML 为 Markdown...');

      // 读取 HTML 内容
      const htmlContent = await fs.readFile(htmlResult.htmlPath, 'utf8');

      // 使用简单的 HTML 到 Markdown 转换
      // 注意：这里会丢失一些样式信息，但保留结构
      const markdownContent = await this.htmlToMarkdown(htmlContent);

      // 导入安全配置函数
      const { validateAndSanitizePath } = require('../security/securityConfig');
      const allowedPaths = [path.dirname(htmlResult.htmlPath), process.cwd()];
      
      // 生成输出路径
      const rawOutputPath =
        options.outputPath ??
        options.markdownOutputPath ??
        htmlResult.htmlPath.replace('.html', '.md');
      const outputPath = validateAndSanitizePath(rawOutputPath, allowedPaths);

      await fs.writeFile(outputPath, markdownContent, 'utf8');

      console.log(`✅ Markdown 转换完成: ${path.basename(outputPath)}`);

      return {
        success: true,
        outputPath,
        htmlPath: htmlResult.htmlPath,
        markdownPath: outputPath,
        details: {
          originalFormat: 'docx',
          targetFormat: 'markdown',
          stylesPreserved: false, // Markdown 不完全支持样式
          imagesPreserved: true,
          conversionTime: Date.now() - startTime,
        },
      };
    } catch (error: any) {
      SafeErrorHandler.logError('PDF 转换准备失败', error);
      return {
        success: false,
        error: SafeErrorHandler.sanitizeErrorMessage(error),
        details: {
          originalFormat: 'docx',
          targetFormat: 'pdf',
          stylesPreserved: false,
          imagesPreserved: false,
          conversionTime: Date.now() - startTime,
        },
      };
    }
  }

  /**
   * 转换 HTML 为 PDF
   */
  private async convertHtmlToPdf(
    htmlResult: any,
    options: OptimizedConversionOptions,
    startTime: number
  ): Promise<ConversionResult> {
    try {
      console.log('\n📄 第二步：准备 PDF 转换...');

      const pdfPath = options.outputPath ?? htmlResult.htmlPath.replace('.html', '.pdf');

      // 由于需要浏览器引擎来生成 PDF，我们返回指令给外部工具
      const instructions = `
优化的 DOCX 转 PDF 转换 - 需要 playwright-mcp 完成

✅ 已完成 (当前 MCP):
  1. DOCX 文件解析和样式提取
  2. 双重解析引擎处理
  3. 样式完整的 HTML 文件生成: ${path.basename(htmlResult.htmlPath)}

🎯 需要执行 (playwright-mcp):
  请运行以下命令完成 PDF 转换:
  
  1. browser_navigate("file://${htmlResult.htmlPath}")
  2. browser_wait_for({ time: 3 })
  3. browser_pdf_save({ filename: "${path.basename(pdfPath)}" })

📁 最终输出: ${path.basename(pdfPath)}

💡 优势:
  - 使用双重解析引擎确保样式完整性
  - 完美保留 Word 文档的格式和布局
  - 支持图片和复杂样式
  - 避免样式丢失问题
`;

      console.log(instructions);

      return {
        success: true,
        outputPath: pdfPath,
        htmlPath: htmlResult.htmlPath,
        requiresExternalTool: true,
        externalToolInstructions: instructions,
        details: {
          originalFormat: 'docx',
          targetFormat: 'pdf',
          stylesPreserved: true,
          imagesPreserved: true,
          conversionTime: Date.now() - startTime,
        },
      };
    } catch (error: any) {
      console.error('❌ PDF 转换准备失败:', error.message);
      return {
        success: false,
        error: error.message,
        details: {
          originalFormat: 'docx',
          targetFormat: 'pdf',
          stylesPreserved: false,
          imagesPreserved: false,
          conversionTime: Date.now() - startTime,
        },
      };
    }
  }

  /**
   * 简单的 HTML 到 Markdown 转换
   */
  private async htmlToMarkdown(htmlContent: string): Promise<string> {
    // 移除样式标签和脚本
    let markdown = htmlContent
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/<head[^>]*>[\s\S]*?<\/head>/gi, '');

    // 基本的 HTML 标签转换
    markdown = markdown
      .replace(/<h1[^>]*>(.*?)<\/h1>/gi, '# $1\n\n')
      .replace(/<h2[^>]*>(.*?)<\/h2>/gi, '## $1\n\n')
      .replace(/<h3[^>]*>(.*?)<\/h3>/gi, '### $1\n\n')
      .replace(/<h4[^>]*>(.*?)<\/h4>/gi, '#### $1\n\n')
      .replace(/<h5[^>]*>(.*?)<\/h5>/gi, '##### $1\n\n')
      .replace(/<h6[^>]*>(.*?)<\/h6>/gi, '###### $1\n\n')
      .replace(/<p[^>]*>(.*?)<\/p>/gi, '$1\n\n')
      .replace(/<strong[^>]*>(.*?)<\/strong>/gi, '**$1**')
      .replace(/<b[^>]*>(.*?)<\/b>/gi, '**$1**')
      .replace(/<em[^>]*>(.*?)<\/em>/gi, '*$1*')
      .replace(/<i[^>]*>(.*?)<\/i>/gi, '*$1*')
      .replace(/<u[^>]*>(.*?)<\/u>/gi, '_$1_')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<hr\s*\/?>/gi, '\n---\n\n')
      .replace(/<img[^>]*src=["']([^"']*)["'][^>]*>/gi, '![Image]($1)\n\n')
      .replace(/<a[^>]*href=["']([^"']*)["'][^>]*>(.*?)<\/a>/gi, '[$2]($1)')
      .replace(/<li[^>]*>(.*?)<\/li>/gi, '- $1\n')
      .replace(/<ul[^>]*>/gi, '\n')
      .replace(/<\/ul>/gi, '\n')
      .replace(/<ol[^>]*>/gi, '\n')
      .replace(/<\/ol>/gi, '\n');

    // 清理多余的空行
    markdown = markdown
      .replace(/\n{3,}/g, '\n\n')
      .replace(/^\n+/, '')
      .replace(/\n+$/, '\n');

    return markdown;
  }

  /**
   * 验证输入文件
   */
  private async validateInput(inputPath: string): Promise<void> {
    try {
      await fs.access(inputPath);

      if (!inputPath.toLowerCase().endsWith('.docx')) {
        throw new Error('输入文件必须是 .docx 格式');
      }
    } catch (error: any) {
      if (error.code === 'ENOENT') {
        throw new Error(`文件不存在: ${inputPath}`);
      }
      throw error;
    }
  }

  /**
   * 快速测试转换（仅转换为 HTML）
   */
  async quickTest(
    inputPath: string
  ): Promise<{ success: boolean; htmlPath?: string; error?: string }> {
    try {
      console.log('🧪 快速测试转换...');

      const result = await this.convertDocx(inputPath, {
        outputFormat: 'html',
        debug: true,
        saveIntermediateFiles: true,
      });

      return {
        success: result.success,
        htmlPath: result.htmlPath,
        error: result.error,
      };
    } catch (error: any) {
      return {
        success: false,
        error: error.message,
      };
    }
  }
}

// 导出便捷函数
export async function optimizedDocxToHtml(
  inputPath: string,
  options: OptimizedConversionOptions = {}
): Promise<ConversionResult> {
  const converter = new OptimizedDocxConverter();
  return await converter.convertDocx(inputPath, { ...options, outputFormat: 'html' });
}

export async function optimizedDocxToMarkdown(
  inputPath: string,
  options: OptimizedConversionOptions = {}
): Promise<ConversionResult> {
  const converter = new OptimizedDocxConverter();
  return await converter.convertDocx(inputPath, { ...options, outputFormat: 'markdown' });
}

export async function optimizedDocxToPdf(
  inputPath: string,
  options: OptimizedConversionOptions = {}
): Promise<ConversionResult> {
  const converter = new OptimizedDocxConverter();
  return await converter.convertDocx(inputPath, { ...options, outputFormat: 'pdf' });
}

export async function testDocxConversion(
  inputPath: string
): Promise<{ success: boolean; htmlPath?: string; error?: string }> {
  const converter = new OptimizedDocxConverter();
  return await converter.quickTest(inputPath);
}

export { OptimizedConversionOptions, ConversionResult };
