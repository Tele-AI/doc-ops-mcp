// doc-ops-mcp - Document Operations MCP Server
// Enhanced DOCX to PDF conversion with perfect Word style replication

import { SafeErrorHandler } from './security/errorHandler';
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const fsSync = require('fs');
const os = require('os');
const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} = require('@modelcontextprotocol/sdk/types.js');
const fs = require('fs/promises');
const WordExtractor = require('word-extractor');
import * as path from 'path';
const cheerio = require('cheerio');
const mammoth = require('mammoth');
const { createSecureTempPath, escapeHtml, sanitizeCssProperty, defaultSecurityConfig, validateAndSanitizePath, safePathJoin } = require('./security/securityConfig.js');

// 路径安全验证函数 - 使用更安全的实现
function validatePath(inputPath: string, allowedBasePaths: string[] = []): string {
  return validateAndSanitizePath(inputPath, allowedBasePaths);
}

import { WEB_SCRAPING_TOOL, STRUCTURED_DATA_TOOL } from './tools/webScrapingTools';
import { convertDocxToHtmlWithStyles } from './tools/enhancedMammothConfig';
import { convertDocxToHtmlEnhanced } from './tools/enhancedMammothConverter';

// 安全的HTML内容处理函数，防止XSS攻击
function sanitizeHtmlForOutput(html: string): string {
  // 对于已经包含HTML标签的内容，我们假设它是来自可信源的处理过的内容
  // 但仍然需要确保没有恶意脚本
  if (/<[a-z][\s\S]*>/i.test(html)) {
    // 移除潜在的危险标签和属性
    return html
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/<iframe[^>]*>[\s\S]*?<\/iframe>/gi, '')
      .replace(/on\w+\s*=\s*(?:"[^"]*"|'[^']*')/gi, '') // 修复ReDoS风险
      .replace(/javascript:/gi, '');
  }
  // 对于纯文本内容，进行HTML转义
  return html
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
import {
  DUAL_PARSING_TOOLS,
  dualParsingDocxToHtml,
  dualParsingDocxToPdf,
  analyzeDocxStyles as analyzeDualParsingStyles,
} from './tools/dualParsingDocxTools';
import { optimizedDocxToPdf, optimizedDocxToMarkdown } from './tools/optimizedDocxConverter';
import { convertMarkdownToHtml } from './tools/markdownToHtmlConverter';
import { convertMarkdownToDocx } from './tools/markdownToDocxConverter';
import {
  convertHtmlToPdf,
  convertHtmlToMarkdown,
  convertHtmlToTxt,
  convertHtmlToDocx,
} from './tools/htmlConverter';
import { EnhancedHtmlToDocxConverter } from './tools/enhancedHtmlToDocxConverter';
import { EnhancedHtmlToMarkdownConverter } from './tools/enhancedHtmlToMarkdownConverter';
import { ConversionPlanner } from './tools/conversionPlanner';

// 导出关键函数，使其可以被外部模块直接引用
// 注意：PDF 转换需要先转换为 HTML，然后使用 playwright-mcp 完成最终的 PDF 转换
// 任何格式 → PDF 的转换流程：原格式 → HTML → PDF (通过 playwright-mcp)
export {
  convertDocxToPdf,
  convertDocxToHtmlEnhanced,
  convertDocxToMarkdown,
  convertDocxToTxt,
  convertMarkdownToDocx,
  convertHtmlToPdf,
  convertHtmlToMarkdown,
  convertHtmlToTxt,
  convertHtmlToDocx,
  convertMarkdownToHtml,
  convertMarkdownToPdf,
  convertMarkdownToTxt,
  convertHtmlToDocxEnhanced,
  EnhancedHtmlToMarkdownConverter,
  ConversionPlanner,
};

// 辅助函数：准备转换路径
async function prepareHtmlToDocxPaths(inputPath: string, outputPath?: string) {
  const allowedInputPaths = [process.cwd(), defaultResourcePaths.tempDir];
  const allowedOutputPaths = [defaultResourcePaths.outputDir, process.cwd()];
  
  const validatedInputPath = validatePath(inputPath, allowedInputPaths);
  let finalOutputPath = outputPath;
  
  if (!finalOutputPath) {
    const baseName = path.basename(validatedInputPath, path.extname(validatedInputPath));
    finalOutputPath = safePathJoin(defaultResourcePaths.outputDir, `${baseName}.docx`);
  } else if (!path.isAbsolute(finalOutputPath)) {
    finalOutputPath = safePathJoin(defaultResourcePaths.outputDir, finalOutputPath);
  }
  
  const validatedOutputPath = validatePath(finalOutputPath ?? '', allowedOutputPaths);
  await fs.mkdir(path.dirname(validatedOutputPath), { recursive: true });
  
  return { validatedInputPath, validatedOutputPath };
}

// 辅助函数：读取HTML内容
async function readHtmlContent(filePath: string): Promise<string> {
  return await fs.readFile(filePath, 'utf-8');
}

// 辅助函数：转换HTML到DOCX缓冲区
async function convertHtmlToDocxBuffer(htmlContent: string): Promise<Buffer> {
  const converter = new EnhancedHtmlToDocxConverter();
  return await converter.convertHtmlToDocx(htmlContent);
}

// 辅助函数：写入输出文件
async function writeDocxOutputFile(outputPath: string, content: Buffer): Promise<void> {
  await fs.writeFile(outputPath, content);
}

// 辅助函数：创建成功响应
function createHtmlToDocxSuccessResponse(outputPath: string, docxBuffer: Buffer) {
  return {
    success: true,
    outputPath,
    message: '增强的 HTML 到 DOCX 转换完成',
    metadata: {
      converter: 'EnhancedHtmlToDocxConverter',
      stylesPreserved: true,
      fileSize: docxBuffer.length,
    },
  };
}

// 辅助函数：创建错误响应
function createHtmlToDocxErrorResponse(error: Error) {
  return {
    success: false,
    error: error.message,
  };
}

// 增强的 HTML 到 DOCX 转换函数 - 使用新的增强转换器
async function convertHtmlToDocxEnhanced(
  inputPath: string,
  outputPath?: string,
  options: any = {}
) {
  try {
    const { validatedInputPath, validatedOutputPath } = await prepareHtmlToDocxPaths(inputPath, outputPath);
    const htmlContent = await readHtmlContent(validatedInputPath);
    const docxBuffer = await convertHtmlToDocxBuffer(htmlContent);
    
    await writeDocxOutputFile(validatedOutputPath, docxBuffer);
    return createHtmlToDocxSuccessResponse(validatedOutputPath, docxBuffer);
  } catch (error: any) {
    return createHtmlToDocxErrorResponse(error);
  }
}

function getDefaultResourcePaths() {
  const homeDir = os.homedir();

  // 从环境变量获取输出目录，如果没有设置则使用默认值
  const outputDir = process.env.OUTPUT_DIR ?? path.join(homeDir, 'Documents');
  const cacheDir = process.env.CACHE_DIR ?? path.join(homeDir, '.cache', 'doc-ops-mcp');

  // 从环境变量获取水印和二维码图片路径
  const watermarkImagePath = process.env.WATERMARK_IMAGE;
  const qrCodeImagePath = process.env.QR_CODE_IMAGE;

  try {
    return {
      outputDir: path.resolve(outputDir),
      cacheDir: path.resolve(cacheDir),
      defaultQrCodePath: qrCodeImagePath ?? null,
      defaultWatermarkPath: watermarkImagePath ?? null,
      tempDir: os.tmpdir(),
    };
  } catch (error) {
    // 生产环境移除调试输出
    return {
      outputDir: '/tmp',
      cacheDir: '/tmp/cache',
      defaultQrCodePath: null,
      defaultWatermarkPath: null,
      tempDir: '/tmp',
    };
  }
}

const defaultResourcePaths = getDefaultResourcePaths();
// Using hybrid mode - browser operations handled by external playwright-mcp

// Interfaces
interface ReadDocumentOptions {
  extractMetadata?: boolean;
  preserveFormatting?: boolean;
  imageOutputDir?: string;
  saveImages?: boolean;
}

interface WriteDocumentOptions {
  encoding?: string;
}

interface ConvertDocumentOptions {
  preserveFormatting?: boolean;
  pdfOptions?: {
    format?: string;
    margin?: string;
  };
  textReplacements?: Array<{
    oldText: string;
    newText: string;
    useRegex?: boolean;
    preserveCase?: boolean;
  }>;
  useInternalPlaywright?: boolean;
}

interface WatermarkOptions {
  watermarkImage?: string;
  watermarkText?: string;
  watermarkImageScale?: number;
  watermarkImageOpacity?: number;
  watermarkImagePosition?: string;
  watermarkFontSize?: number;
  watermarkTextOpacity?: number;
  watermarkTextPosition?: string;
}

interface QRCodeOptions {
  qrScale?: number;
  qrOpacity?: number;
  qrPosition?: string;
  addText?: boolean;
  customText?: string;
  textSize?: number;
  textColor?: string;
}

// Helper functions for readDocument
async function processDocxWithFormatting(filePath: string, options: ReadDocumentOptions) {
  // Using enhanced mammoth converter for style preservation
  
  // Try enhanced mammoth converter first
  const enhancedResult = await tryEnhancedMammothConverter(filePath, options);
  if (enhancedResult) {
    return enhancedResult;
  }
  
  // Fallback to basic mammoth configuration
  return await fallbackToBasicMammoth(filePath, options);
}

async function tryEnhancedMammothConverter(filePath: string, options: ReadDocumentOptions) {
  try {
    // Trying enhanced mammoth converter with fixes
    const enhancedResult = await convertDocxToHtmlEnhanced(filePath, {
      preserveImages: options.saveImages !== false,
      enableExperimentalFeatures: true,
      debug: true,
    });

    logEnhancedMammothResult(enhancedResult);

    if (enhancedResult.success && (enhancedResult as any).content) {
      const content = (enhancedResult as any).content;
      const metadata = (enhancedResult as any).metadata;
      
      if (validateStylesInContent(content)) {
        return createSuccessfulDocxResult(content, metadata, options);
      }
    }
  } catch (enhancedError: any) {
    SafeErrorHandler.logError('Enhanced mammoth converter failed', enhancedError);
  }
  
  return null;
}

function logEnhancedMammothResult(enhancedResult: any) {
  // Debug output removed for production
}

function validateStylesInContent(content: string): boolean {
  const hasValidStyles = content.includes('<style>') && 
                        content.includes('Calibri') && 
                        content.includes('font-family');

  // Debug output removed for production

  if (hasValidStyles) {
    // Debug output removed for production
    return true;
  } else {
    // Debug output removed for production
    return false;
  }
}

function createSuccessfulDocxResult(content: string, metadata: any, options: ReadDocumentOptions) {
  return {
    success: true,
    content: content,
    metadata: {
      ...metadata,
      format: 'html',
      originalFormat: 'docx',
      converter: 'enhanced-mammoth-fixed',
      stylesPreserved: true,
      stylesValidated: true,
      imagesPreserved: options.saveImages !== false,
      standalone: true,
    },
  };
}

async function fallbackToBasicMammoth(filePath: string, options: ReadDocumentOptions) {
  console.error('🔄 回退到基础增强配置...');
  console.error('🔄 使用基础 mammoth 配置作为回退方案...');
  
  const result = await convertDocxToHtmlWithStyles(filePath, {
    saveImagesToFiles: true,
    imageOutputDir: options.imageOutputDir ?? path.join(process.cwd(), 'output', 'images'),
  });

  console.error('🔍 回退转换结果分析:', {
    success: result.success,
    contentLength: result.content?.length ?? 0,
    hasStyle: result.content?.includes('<style>') ?? false,
    hasDoctype: result.content?.includes('<!DOCTYPE') ?? false,
    contentPreview: (result.content?.substring(0, 300) ?? 'No content') + '...',
  });

  return {
    success: result.success,
    content: result.content,
    metadata: result.metadata
      ? {
          ...result.metadata,
          converter: 'enhanced-mammoth',
          fallbackUsed: true,
          standalone: true,
        }
      : {
          format: 'html',
          originalFormat: 'docx',
          stylesPreserved: true,
          converter: 'enhanced-mammoth',
          fallbackUsed: true,
          standalone: true,
        },
  };
}

async function processDocxAsText(filePath: string) {
  let result: { value: string };

  try {
    result = await mammoth.extractRawText({ path: filePath });
  } catch (mammothError) {
    SafeErrorHandler.logError('Mammoth text extraction failed', mammothError);
    throw mammothError;
  }

  return {
    success: true,
    content: result.value,
    metadata: { format: 'text', originalFormat: 'docx', converter: 'mammoth' },
  };
}

async function processDocFile(filePath: string) {
  const extractor = new WordExtractor();
  const extracted = await extractor.extract(filePath);
  return {
    success: true,
    content: extracted.getBody(),
    metadata: { format: 'text', originalFormat: 'doc' },
  };
}

async function processMarkdownWithFormatting(filePath: string) {
  console.error('🚀 使用增强型 Markdown 转换器进行样式保留转换...');

  try {
    const result = await convertMarkdownToHtml(filePath, {
      preserveStyles: true,
      theme: 'github', // 使用 GitHub 风格主题
      standalone: true,
      debug: true,
    });

    console.error('📊 Markdown 转换结果:', {
      success: result.success,
      contentLength: result.content?.length ?? 0,
      hasStyle: result.content?.includes('<style>') ?? false,
      hasDoctype: result.content?.includes('<!DOCTYPE') ?? false,
    });

    return {
      success: result.success,
      content: result.content,
      metadata: {
        ...result.metadata,
        format: 'html',
        originalFormat: 'markdown',
        converter: 'enhanced-markdown',
        stylesPreserved: true,
        standalone: true,
      },
    };
  } catch (markdownError: any) {
    console.error('❌ 增强型 Markdown 转换失败:', markdownError.message);
    throw markdownError;
  }
}

async function processMarkdownAsText(filePath: string) {
  const content = await fs.readFile(filePath, 'utf-8');
  return {
    success: true,
    content,
    metadata: { format: 'markdown', originalFormat: 'markdown' },
  };
}

// 辅助函数：处理DOCX文件
async function handleDocxFile(filePath: string, options: ReadDocumentOptions) {
  return options.preserveFormatting 
    ? await processDocxWithFormatting(filePath, options)
    : await processDocxAsText(filePath);
}

// 辅助函数：处理Markdown文件
async function handleMarkdownFile(filePath: string, options: ReadDocumentOptions) {
  return options.preserveFormatting 
    ? await processMarkdownWithFormatting(filePath)
    : await processMarkdownAsText(filePath);
}

// 辅助函数：处理其他文件类型
async function handleOtherFile(filePath: string, ext: string) {
  const content = await fs.readFile(filePath, 'utf-8');
  return {
    success: true,
    content,
    metadata: { format: 'text', originalFormat: ext.slice(1) },
  };
}

// 辅助函数：根据文件扩展名选择处理器
async function processFileByExtension(filePath: string, ext: string, options: ReadDocumentOptions) {
  switch (ext) {
    case '.docx':
      return await handleDocxFile(filePath, options);
    case '.doc':
      return await processDocFile(filePath);
    case '.md':
    case '.markdown':
      return await handleMarkdownFile(filePath, options);
    default:
      return await handleOtherFile(filePath, ext);
  }
}

// Read document function
async function readDocument(filePath: string, options: ReadDocumentOptions = {}) {
  try {
    const ext = path.extname(filePath).toLowerCase();
    return await processFileByExtension(filePath, ext, options);
  } catch (error: any) {
    SafeErrorHandler.logError('读取文档失败', error);
    return {
      success: false,
      content: '',
      metadata: { 
        error: SafeErrorHandler.sanitizeErrorMessage(error) 
      },
    };
  }
}



// Helper functions for writeDocument
function resolveFinalOutputPath(outputPath?: string): string {
  if (!outputPath) {
    const outputDir = defaultResourcePaths.outputDir;
    return path.join(outputDir, `output_${Date.now()}.txt`);
  } else if (!path.isAbsolute(outputPath)) {
    // 如果是相对路径，基于环境变量的输出目录
    return path.join(defaultResourcePaths.outputDir, outputPath);
  }
  return outputPath;
}

async function writeFileWithEncoding(finalPath: string, content: string, encoding: string) {
  // 验证路径安全性
  const allowedPaths = [defaultResourcePaths.outputDir, defaultResourcePaths.tempDir, process.cwd()];
  const validatedPath = validatePath(finalPath, allowedPaths);
  
  await fs.mkdir(path.dirname(validatedPath), { recursive: true });
  await fs.writeFile(validatedPath, content, encoding);
}

// Write document function
async function writeDocument(
  content: string,
  outputPath?: string,
  options: WriteDocumentOptions = {}
) {
  try {
    const finalPath = resolveFinalOutputPath(outputPath);
    const encoding = options.encoding ?? 'utf-8';

    const validatedFinalPath = validatePath(finalPath);
    await writeFileWithEncoding(validatedFinalPath, content, encoding);

    return {
      success: true,
      outputPath: validatedFinalPath,
      message: `Document written successfully to ${finalPath}. Output directory controlled by OUTPUT_DIR environment variable: ${defaultResourcePaths.outputDir}`,
    };
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

// Helper functions for convertDocument
function resolveConvertOutputPath(inputPath: string, outputPath?: string): { finalOutputPath: string, inputExt: string, outputExt: string } {
  const validatedInputPath = validatePath(inputPath);
  const inputExt = path.extname(validatedInputPath).toLowerCase();
  const allowedOutputPaths = [defaultResourcePaths.outputDir, process.cwd()];
  
  const outputExt = outputPath ? path.extname(validatePath(outputPath, allowedOutputPaths)).toLowerCase() : '.html';

  let finalOutputPath: string;
  if (!outputPath) {
    const baseName = path.basename(validatedInputPath, inputExt);
    finalOutputPath = safePathJoin(defaultResourcePaths.outputDir, `${baseName}${outputExt}`);
  } else if (!path.isAbsolute(outputPath)) {
    finalOutputPath = safePathJoin(defaultResourcePaths.outputDir, outputPath);
  } else {
    finalOutputPath = validatePath(outputPath, allowedOutputPaths);
  }

  return { finalOutputPath, inputExt, outputExt };
}

async function convertHtmlToMarkdownSpecial(inputPath: string, finalOutputPath: string, options: ConvertDocumentOptions) {
  console.error('🔄 检测到 HTML 转 Markdown 转换...');
  console.error(`📄 输入: ${inputPath}`);
  console.error(`📁 输出: ${finalOutputPath}`);

  try {
    const result = await convertHtmlToMarkdown(inputPath, {
      outputPath: finalOutputPath,
      preserveStyles: options.preserveFormatting !== false,
      debug: true,
    });

    console.error('📊 HTML 转 Markdown 结果:', {
      success: result.success,
      outputPath: result.outputPath,
      contentLength: result.content?.length ?? 0,
      error: result.error,
    });

    if (result.success) {
      return {
        success: true,
        outputPath: result.outputPath,
        content: result.content,
        metadata: {
          ...result.metadata,
          converter: 'html-to-markdown',
        },
      };
    } else {
      throw new Error(result.error ?? 'HTML 转 Markdown 失败');
    }
  } catch (conversionError: any) {
    console.error('❌ HTML 转 Markdown 转换失败:', conversionError.message);
    return {
      success: false,
      error: `HTML 转 Markdown 失败: ${conversionError.message}`,
    };
  }
}

async function convertHtmlToDocxSpecial(inputPath: string, finalOutputPath: string, options: ConvertDocumentOptions) {
  console.error('🔄 检测到 HTML 转 DOCX 转换...');
  console.error(`📄 输入: ${inputPath}`);
  console.error(`📁 输出: ${finalOutputPath}`);

  try {
    const result = await convertHtmlToDocx(inputPath, finalOutputPath, {
      preserveStyles: options.preserveFormatting !== false,
      debug: true,
    });

    console.error('📊 HTML 转 DOCX 结果:', {
      success: result.success,
      outputPath: result.outputPath,
      error: result.error,
    });

    if (result.success) {
      return {
        success: true,
        outputPath: result.outputPath,
        metadata: {
          ...result.metadata,
          converter: 'html-to-docx',
        },
      };
    } else {
      throw new Error(result.error ?? 'HTML 转 DOCX 失败');
    }
  } catch (conversionError: any) {
    console.error('❌ HTML 转 DOCX 转换失败:', conversionError.message);
    return {
      success: false,
      error: `HTML 转 DOCX 失败: ${conversionError.message}`,
    };
  }
}

async function convertDocxToMarkdownSpecial(inputPath: string, finalOutputPath: string, options: ConvertDocumentOptions) {
  console.error('🔄 检测到 DOCX 转 Markdown 转换...');
  console.error(`📄 输入: ${inputPath}`);
  console.error(`📁 输出: ${finalOutputPath}`);

  try {
    const result = await convertDocxToMarkdown(inputPath, finalOutputPath, {
      preserveFormatting: options.preserveFormatting !== false,
    });

    console.error('📊 DOCX 转 Markdown 结果:', {
      success: result.success,
      outputPath: result.outputPath,
      error: result.error,
    });

    if (result.success) {
      return {
        success: true,
        outputPath: result.outputPath,
        metadata: {
          originalFormat: 'docx',
          targetFormat: 'markdown',
          converter: 'docx-to-markdown',
        },
      };
    } else {
      throw new Error(result.error ?? 'DOCX 转 Markdown 失败');
    }
  } catch (conversionError: any) {
    console.error('❌ DOCX 转 Markdown 转换失败:', conversionError.message);
    return {
      success: false,
      error: `DOCX 转 Markdown 失败: ${conversionError.message}`,
    };
  }
}

async function convertDocxToHtmlSpecial(inputPath: string, finalOutputPath: string) {
  console.error('🔄 检测到 DOCX 转 HTML 转换...');
  console.error(`📄 输入: ${inputPath}`);
  console.error(`📁 输出: ${finalOutputPath}`);

  try {
    const validatedInputPath = validatePath(inputPath);
    const result = await convertDocxToHtmlEnhanced(validatedInputPath, {
      preserveImages: true,
      enableExperimentalFeatures: true,
      debug: true,
    });

    console.error('📊 DOCX 转 HTML 结果:', {
      success: result.success,
      hasContent: !!(result as any).content,
      contentLength: (result as any).content?.length ?? 0,
      error: (result as any).error,
    });

    if (result.success && (result as any).content) {
      // 写入HTML文件
      await fs.writeFile(finalOutputPath, (result as any).content, 'utf-8');

      return {
        success: true,
        outputPath: finalOutputPath,
        metadata: {
          originalFormat: 'docx',
          targetFormat: 'html',
          converter: 'docx-to-html-enhanced',
          stylesPreserved: true,
        },
      };
    } else {
      throw new Error((result as any).error ?? 'DOCX 转 HTML 失败');
    }
  } catch (conversionError: any) {
    console.error('❌ DOCX 转 HTML 转换失败:', conversionError.message);
    return {
      success: false,
      error: `DOCX 转 HTML 失败: ${conversionError.message}`,
    };
  }
}

function applyRegexReplacement(content: string, replacement: any): string {
  if (replacement.oldText.length > 100) {
    console.warn('正则表达式过长，跳过处理');
    return content;
  }
  
  try {
    const flags = replacement.preserveCase ? 'g' : 'gi';
    const regex = new RegExp(replacement.oldText, flags);
    return content.replace(regex, replacement.newText);
  } catch (error: any) {
    console.warn('正则表达式无效，跳过处理:', error.message);
    return content;
  }
}

function applyPlainTextReplacement(content: string, replacement: any): string {
  if (replacement.oldText.length > 100) {
    console.warn('搜索文本过长，跳过处理');
    return content;
  }

  try {
    const searchValue = replacement.preserveCase
      ? replacement.oldText
      : new RegExp(replacement.oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    return content.replace(searchValue, replacement.newText);
  } catch (error: any) {
    console.warn('文本替换失败，跳过处理:', error.message);
    return content;
  }
}

function applyTextReplacements(content: string, textReplacements: any[]): string {
  return textReplacements.reduce((processedContent, replacement) => {
    return replacement.useRegex
      ? applyRegexReplacement(processedContent, replacement)
      : applyPlainTextReplacement(processedContent, replacement);
  }, content);
}

async function performGenericConversion(inputPath: string, finalOutputPath: string, options: ConvertDocumentOptions) {
  // Read the input document
  const readResult = await readDocument(inputPath, {
    preserveFormatting: options.preserveFormatting,
  });
  if (!readResult.success) {
    return readResult;
  }

  let content = readResult.content;

  // Apply text replacements if specified
  if (options.textReplacements) {
    content = applyTextReplacements(content, options.textReplacements);
  }

  // Write the converted content
  return await writeDocument(content, finalOutputPath);
}

// Convert document function
async function convertDocument(
  inputPath: string,
  outputPath?: string,
  options: ConvertDocumentOptions = {}
) {
  try {
    const { finalOutputPath, inputExt, outputExt } = resolveConvertOutputPath(inputPath, outputPath);

    // Handle special conversion cases
    if (inputExt === '.html' && outputExt === '.md') {
      return await convertHtmlToMarkdownSpecial(inputPath, finalOutputPath, options);
    }

    if (inputExt === '.html' && outputExt === '.docx') {
      return await convertHtmlToDocxSpecial(inputPath, finalOutputPath, options);
    }

    if (inputExt === '.docx' && outputExt === '.md') {
      return await convertDocxToMarkdownSpecial(inputPath, finalOutputPath, options);
    }

    if (inputExt === '.docx' && outputExt === '.html') {
      return await convertDocxToHtmlSpecial(inputPath, finalOutputPath);
    }

    // Handle generic conversions
    return await performGenericConversion(inputPath, finalOutputPath, options);
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

// Helper functions for convertDocxToPdf
function resolvePdfOutputPath(inputPath: string, outputPath?: string): string {
  const allowedPaths = [defaultResourcePaths.outputDir, process.cwd()];
  
  if (!outputPath) {
    const baseName = path.basename(inputPath, path.extname(inputPath));
    return safePathJoin(defaultResourcePaths.outputDir, `${baseName}.pdf`);
  } else if (!path.isAbsolute(outputPath)) {
    return safePathJoin(defaultResourcePaths.outputDir, outputPath);
  }
  return validatePath(outputPath, allowedPaths);
}

function checkWatermarkAndQRConfig(options: any): { hasWatermark: boolean, hasQRCode: boolean } {
  const hasWatermark =
    defaultResourcePaths.defaultWatermarkPath &&
    fsSync.existsSync(defaultResourcePaths.defaultWatermarkPath);
  const hasQRCode =
    options.addQRCode &&
    defaultResourcePaths.defaultQrCodePath &&
    fsSync.existsSync(defaultResourcePaths.defaultQrCodePath);

  if (hasWatermark) {
    console.error(`🎨 检测到水印图片: ${defaultResourcePaths.defaultWatermarkPath}`);
  }
  if (hasQRCode) {
    console.error(`📱 检测到二维码图片: ${defaultResourcePaths.defaultQrCodePath}`);
  }

  return { hasWatermark, hasQRCode };
}

function logConversionStart(inputPath: string, finalOutputPath: string) {
  console.error(`🚀 开始优化的 DOCX 到 PDF 转换...`);
  console.error(`📄 输入文件: ${inputPath}`);
  console.error(`📁 输出路径: ${finalOutputPath}`);
  console.error(`🌍 输出目录由环境变量控制: OUTPUT_DIR=${defaultResourcePaths.outputDir}`);
}

function createMcpCommands(result: any, hasWatermark: boolean, hasQRCode: boolean): string[] {
  const mcpCommands = [
    `browser_navigate("file://${result.htmlPath}")`,
    `browser_wait_for({ time: 3 })`,
    `browser_pdf_save({ filename: "${result.outputPath}" })`,
  ];

  // 如果需要添加水印或二维码，添加后处理步骤
  if (hasWatermark || hasQRCode) {
    mcpCommands.push('# PDF后处理步骤:');
    if (hasWatermark) {
      mcpCommands.push(`# 自动添加水印: ${defaultResourcePaths.defaultWatermarkPath}`);
    }
    if (hasQRCode) {
      mcpCommands.push(`# 添加二维码: ${defaultResourcePaths.defaultQrCodePath}`);
    }
  }

  return mcpCommands;
}

function createPostProcessingConfig(hasWatermark: boolean, hasQRCode: boolean) {
  return {
    watermark: hasWatermark
      ? {
          enabled: true,
          imagePath: defaultResourcePaths.defaultWatermarkPath,
          options: {
            scale: 0.25,
            opacity: 0.6,
            position: 'top-right',
          },
        }
      : { enabled: false },
    qrCode: hasQRCode
      ? {
          enabled: true,
          imagePath: defaultResourcePaths.defaultQrCodePath,
          options: {
            scale: 0.15,
            opacity: 1.0,
            position: 'bottom-center',
            addText: true,
          },
        }
      : { enabled: false },
  };
}

async function addWatermarkToPdf(outputPath: string): Promise<{ success: boolean, error?: string }> {
  console.error(`🎨 自动添加水印...`);
  try {
    // @ts-ignore
    const watermarkResult = await addWatermark(outputPath, {
      // @ts-ignore
      watermarkImage: defaultResourcePaths.defaultWatermarkPath,
      watermarkImageScale: 1.0, // fullscreen模式下会自动计算尺寸
      watermarkImageOpacity: 0.15, // 降低透明度，避免影响阅读
      watermarkImagePosition: 'fullscreen',
    });

    if (watermarkResult.success) {
      console.error(`✅ 水印添加成功`);
      return { success: true };
    } else {
      console.error(`⚠️ 水印添加失败: ${watermarkResult.error}`);
      return { success: false, error: watermarkResult.error };
    }
  } catch (watermarkError: any) {
    console.error(`❌ 水印添加异常: ${watermarkError.message}`);
    return { success: false, error: watermarkError.message };
  }
}

async function addQRCodeToPdf(outputPath: string): Promise<{ success: boolean, error?: string }> {
  console.error(`📱 添加二维码...`);
  try {
    // @ts-ignore
    const qrResult = await addQRCode(outputPath, defaultQrCodePath, {
      qrScale: 0.15,
      qrOpacity: 1.0,
      qrPosition: 'bottom-center',
      addText: true,
    });

    if (qrResult.success) {
      console.error(`✅ 二维码添加成功`);
      return { success: true };
    } else {
      console.error(`⚠️ 二维码添加失败: ${qrResult.error}`);
      return { success: false, error: qrResult.error };
    }
  } catch (qrError: any) {
    console.error(`❌ 二维码添加异常: ${qrError.message}`);
    return { success: false, error: qrError.message };
  }
}

async function handleDirectConversionSuccess(result: any, hasWatermark: boolean, hasQRCode: boolean) {
  let finalResult = {
    success: true,
    outputPath: result.outputPath,
    stats: {
      conversionTime: result.details.conversionTime,
      stylesPreserved: result.details.stylesPreserved,
      imagesPreserved: result.details.imagesPreserved,
    },
  };

  // 自动添加水印
  if (hasWatermark && defaultResourcePaths.defaultWatermarkPath) {
    const watermarkResult = await addWatermarkToPdf(result.outputPath);
    if (watermarkResult.success) {
      (finalResult as any).watermarkAdded = true;
    } else {
      (finalResult as any).watermarkError = watermarkResult.error;
    }
  }

  // 添加二维码（仅在用户明确要求时）
  //@ts-ignore
  if (hasQRCode && defaultQrCodePath) {
    const qrResult = await addQRCodeToPdf(result.outputPath);
    if (qrResult.success) {
      (finalResult as any).qrCodeAdded = true;
    } else {
      (finalResult as any).qrCodeError = qrResult.error;
    }
  }

  return finalResult;
}

// 重构的 DOCX to PDF 转换 - 使用优化的转换器
async function convertDocxToPdf(inputPath: string, outputPath?: string, options: any = {}) {
  const startTime = Date.now();
  const finalOutputPath = resolvePdfOutputPath(inputPath, outputPath);
  const { hasWatermark, hasQRCode } = checkWatermarkAndQRConfig(options);
  
  logConversionStart(inputPath, finalOutputPath);

  try {
    // 使用新的优化转换器
    const result = await optimizedDocxToPdf(inputPath, {
      outputPath: finalOutputPath,
      preserveStyles: true,
      preserveImages: true,
      preserveLayout: true,
      debug: true,
      saveIntermediateFiles: true,
      pdfOptions: {
        format: options.pdfOptions?.format ?? 'A4',
        margin: options.pdfOptions?.margin ?? '1in',
        printBackground: true,
        landscape: false,
      },
    });

    if (result.success) {
      console.error(`✅ 转换准备完成!`);
      console.error(`📊 转换统计:`);
      console.error(`  - 原始格式: ${result.details.originalFormat}`);
      console.error(`  - 目标格式: ${result.details.targetFormat}`);
      console.error(`  - 样式保留: ${result.details.stylesPreserved ? '✅' : '❌'}`);
      console.error(`  - 图片保留: ${result.details.imagesPreserved ? '✅' : '❌'}`);
      console.error(`  - 转换时间: ${result.details.conversionTime}ms`);

      if (result.requiresExternalTool) {
        const mcpCommands = createMcpCommands(result, hasWatermark, hasQRCode);
        const postProcessingConfig = createPostProcessingConfig(hasWatermark, hasQRCode);

        // 返回需要外部工具完成的结果
        return {
          success: false,
          useMcpMode: true,
          mode: 'optimized_pdf_conversion',
          currentMcp: {
            completed: ['DOCX 文件解析和样式提取', '双重解析引擎处理', '样式完整的 HTML 文件生成'],
            htmlFile: result.htmlPath,
            tempFiles: [result.htmlPath],
          },
          playwrightMcp: {
            required: true,
            commands: [
              `browser_navigate("file://${result.htmlPath}")`,
              `browser_wait_for({ time: 3 })`,
              `browser_pdf_save({ filename: "${result.outputPath}" })`,
            ],
            reason: 'PDF 转换需要 playwright-mcp 浏览器实例',
          },
          finalOutput: result.outputPath,
          instructions: result.externalToolInstructions,
          mcpCommands: mcpCommands,
          postProcessing: postProcessingConfig,
          stats: {
            conversionTime: result.details.conversionTime,
            stylesPreserved: result.details.stylesPreserved,
            imagesPreserved: result.details.imagesPreserved,
          },
        };
      } else {
        // 直接转换成功，需要添加水印和二维码
        return await handleDirectConversionSuccess(result, hasWatermark, hasQRCode);
      }
    } else {
      throw new Error(result.error ?? '优化转换器转换失败');
    }
  } catch (error: any) {
    console.error('❌ 优化转换失败，尝试回退方案:', error.message);

    // 回退到原有的转换逻辑（简化版）
    return await fallbackConvertDocxToPdf(inputPath, finalOutputPath, options);
  }
}

// Helper functions for fallbackConvertDocxToPdf
function resolveFallbackOutputPath(inputPath: string, outputPath?: string): string {
  const allowedPaths = [defaultResourcePaths.outputDir, process.cwd()];
  
  if (!outputPath) {
    const baseName = path.basename(inputPath, path.extname(inputPath));
    return safePathJoin(defaultResourcePaths.outputDir, `${baseName}.pdf`);
  } else if (!path.isAbsolute(outputPath)) {
    return safePathJoin(defaultResourcePaths.outputDir, outputPath);
  }
  return validatePath(outputPath, allowedPaths);
}

function logFallbackConversionStart(docxPath: string, finalOutputPath: string) {
  console.error(`🔄 使用回退方案进行 DOCX 到 PDF 转换...`);
  console.error(`📄 输入: ${docxPath}`);
  console.error(`📁 输出: ${finalOutputPath}`);
  console.error(`🌍 输出目录由环境变量控制: OUTPUT_DIR=${defaultResourcePaths.outputDir}`);
}

async function tryDualParsingEngine(docxPath: string): Promise<{ success: boolean, content?: string }> {
  try {
    console.error('🚀 尝试双重解析引擎...');
    const dualResult = await dualParsingDocxToHtml({
      docxPath: docxPath,
      options: {
        extractStyles: true,
        preserveImages: true,
        generateCompleteHTML: true,
        inlineCSS: true,
      },
    });

    if (dualResult.success && dualResult.content) {
      console.error('✅ 双重解析引擎转换成功！');
      return { success: true, content: dualResult.content };
    }
    return { success: false };
  } catch (dualError: any) {
    console.error('❌ 双重解析引擎失败:', dualError.message);
    return { success: false };
  }
}

async function tryEnhancedMammoth(docxPath: string): Promise<{ success: boolean, content?: string }> {
  try {
    console.error('🔄 使用增强型 mammoth 转换器...');
    const enhancedResult = await convertDocxToHtmlEnhanced(docxPath, {
      preserveImages: true,
      enableExperimentalFeatures: true,
      debug: true,
    });

    if (enhancedResult.success && (enhancedResult as any).content) {
      console.error('✅ 增强型 mammoth 转换成功！');
      return { success: true, content: (enhancedResult as any).content };
    }
    return { success: false };
  } catch (enhancedError: any) {
    console.error('❌ 增强型 mammoth 失败:', enhancedError.message);
    return { success: false };
  }
}

async function useBasicMammothFallback(docxPath: string, options: any): Promise<string> {
  console.error('🔄 使用最终回退方案...');
  const basicResult = await mammoth.convertToHtml({ path: docxPath });
  return createPerfectWordHtml(basicResult.value, options);
}

// 回退转换方法（简化的原有逻辑）
async function fallbackConvertDocxToPdf(inputPath: string, outputPath?: string, options: any = {}) {
  const docxPath = inputPath;
  const finalOutputPath = resolveFallbackOutputPath(inputPath, outputPath);
  
  logFallbackConversionStart(docxPath, finalOutputPath);
  await fs.mkdir(path.dirname(finalOutputPath), { recursive: true });

  // 生成HTML内容
  const perfectWordHtml = await generateHtmlContent(docxPath, options);
  console.error(`🎨 HTML 生成完成 (长度: ${perfectWordHtml.length})`);

  // 处理样式和HTML结构
  const finalHtml = await processHtmlStyles(perfectWordHtml, options);
  
  // 创建临时文件并验证
  const tempHtmlPath = await createAndValidateHtmlFile(finalHtml, options);
  
  // 返回Playwright指令
  return createPlaywrightInstructions(finalOutputPath, tempHtmlPath, options);
}

// 生成HTML内容的辅助函数
async function generateHtmlContent(docxPath: string, options: any): Promise<string> {
  const dualResult = await tryDualParsingEngine(docxPath);
  if (dualResult.success && dualResult.content) {
    return dualResult.content;
  }
  
  const enhancedResult = await tryEnhancedMammoth(docxPath);
  if (enhancedResult.success && enhancedResult.content) {
    return enhancedResult.content;
  }
  
  return await useBasicMammothFallback(docxPath, options);
}

// 处理HTML样式的辅助函数
async function processHtmlStyles(html: string, options: any): Promise<string> {
  let finalHtml = html;
  
  // 确保包含完整的Word样式
  if (!finalHtml.includes('<style>')) {
    finalHtml = addMissingStyles(finalHtml, options);
  }
  
  // 确保DOCTYPE和完整的HTML结构
  if (!finalHtml.includes('<!DOCTYPE')) {
    finalHtml = ensureCompleteHtmlStructure(finalHtml, options);
  }
  
  // 处理样式合并和优化
  finalHtml = await mergeAndOptimizeStyles(finalHtml);
  
  return finalHtml;
}

// 添加缺失样式的辅助函数
function addMissingStyles(html: string, options: any): string {
  console.error('⚠️ 检测到样式缺失，强制注入Word样式...');
  
  const contentWithoutWrapper = html
    .replace(/<!DOCTYPE[^>]*>/gi, '')
    .replace(/<html[^>]*>/gi, '')
    .replace(/<\/html>/gi, '')
    .replace(/<head[^>]*>.*?<\/head>/gi, '')
    .replace(/<body[^>]*>/gi, '')
    .replace(/<\/body>/gi, '');

  const safeContent = sanitizeHtmlForOutput(contentWithoutWrapper);
  return createPerfectWordHtml(safeContent, options);
}

// 确保完整HTML结构的辅助函数
function ensureCompleteHtmlStructure(html: string, options: any): string {
  const cleanedContent = html
    .replace(/<!DOCTYPE[^>]*>/gi, '')
    .replace(/<html[^>]*>/gi, '')
    .replace(/<\/html>/gi, '')
    .replace(/<head[^>]*>.*?<\/head>/gi, '')
    .replace(/<body[^>]*>/gi, '')
    .replace(/<\/body>/gi, '');
  
  const safeContent = sanitizeHtmlForOutput(cleanedContent);
  const styleContent = html.includes('<style>') ? '' : 
    createPerfectWordHtml('', options).match(/<style[^>]*>[\s\S]*?<\/style>/gi)?.[0] || '';
  
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    ${styleContent}
</head>
<body>
${safeContent}
</body>
</html>`;
}

// 合并和优化样式的辅助函数
async function mergeAndOptimizeStyles(html: string): Promise<string> {
  const styleTagsContent = extractStyleContent(html);
  
  if (styleTagsContent.length === 0) {
    styleTagsContent.push(getBasicWordStyles());
  }
  
  let combinedStyles = styleTagsContent.join('\n');
  combinedStyles = addImportantToStyles(combinedStyles);
  
  let finalHtml = html.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');
  finalHtml = insertCombinedStyles(finalHtml, combinedStyles);
  finalHtml = enhanceInlineStyles(finalHtml);
  
  return await processWithCheerio(finalHtml);
}

// 提取样式内容的辅助函数
function extractStyleContent(html: string): string[] {
  const styleTagsContent: string[] = [];
  let styleTagMatch;
  const styleTagRegex = /<style[^>]*>([\s\S]*?)<\/style>/gi;
  
  while ((styleTagMatch = styleTagRegex.exec(html)) !== null) {
    const styleContent = styleTagMatch[1].trim();
    if (styleContent) {
      styleTagsContent.push(styleContent);
    }
  }
  
  return styleTagsContent;
}

// 获取基本Word样式的辅助函数
function getBasicWordStyles(): string {
  console.error('⚠️ 未找到有效样式内容，添加基本Word样式');
  return `
    body { font-family: "Calibri", "Microsoft YaHei", "SimSun", sans-serif !important; }
    p { margin-bottom: 8pt !important; line-height: 1.08 !important; }
    table { border-collapse: collapse !important; }
    td, th { border: 1px solid #000 !important; padding: 5pt !important; }
    * { -webkit-print-color-adjust: exact !important; color-adjust: exact !important; print-color-adjust: exact !important; }
  `;
}

// 为样式添加!important的辅助函数
function addImportantToStyles(styles: string): string {
  return styles.replace(/([^;{}:]+:[^;{}!]+)(?=;|})/g, match => {
    return match.includes('!important') ? match : match + ' !important';
  });
}

// 插入合并样式的辅助函数
function insertCombinedStyles(html: string, combinedStyles: string): string {
  // Sanitize CSS content before injection to prevent XSS
  const sanitizedStyles = sanitizeCssProperty(combinedStyles);
  
  if (html.includes('</head>')) {
    return html.replace(/<\/head>/i, `<style type="text/css">\n${sanitizedStyles}\n</style>\n</head>`);
  } else if (html.includes('<head')) {
    return html.replace(/<head[^>]*>/i, `$&\n<style type="text/css">\n${sanitizedStyles}\n</style>`);
  } else if (html.includes('<html')) {
    return html.replace(/<html[^>]*>/i, `$&\n<head>\n<style type="text/css">\n${sanitizedStyles}\n</style>\n</head>`);
  } else {
    return `<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="UTF-8">\n<style type="text/css">\n${sanitizedStyles}\n</style>\n</head>\n<body>\n${html}\n</body>\n</html>`;
  }
}

// HTML转义和CSS清理函数现在从安全配置模块导入

// 增强内联样式的辅助函数
function enhanceInlineStyles(html: string): string {
  // 对输入HTML进行基本验证
  if (!html || typeof html !== 'string') {
    return '';
  }
  
  return html.replace(/<([a-z][a-z0-9]*)([^>]*?)>/gi, (match, tag, attrs) => {
    if (attrs.includes('style=')) {
      return match.replace(/style=(["'])(.*?)\1/gi, (styleMatch, quote, styleContent) => {
        // 清理样式内容
        const cleanStyleContent = sanitizeCssProperty(styleContent);
        const enhancedStyle = cleanStyleContent.replace(/([^;]+)(?=;|$)/g, prop => {
          const cleanProp = sanitizeCssProperty(prop);
          return cleanProp.includes('!important') ? cleanProp : `${cleanProp} !important`;
        });
        return `style=${quote}${enhancedStyle}${quote}`;
      });
    }
    return match;
  });
}

// 使用Cheerio处理HTML的辅助函数
async function processWithCheerio(html: string): Promise<string> {
  const $ = cheerio.load(html, { decodeEntities: false });
  
  let allStyles = '';
  $('style').each((i, el) => {
    // 使用text()而不是html()来避免XSS攻击
    allStyles += $(el).text() + '\n';
  });
  
  allStyles = addImportantToStyles(allStyles);
  $('style').remove();
  
  if (!$('head').length) {
    $('html').prepend('<head></head>');
  }
  // Sanitize CSS content before injection to prevent XSS
  const sanitizedStyles = sanitizeCssProperty(allStyles);
  $('head').append(`<style type="text/css">${sanitizedStyles}</style>`);
  
  $('[style]').each((i, el) => {
    const style = $(el).attr('style');
    if (style) {
      const importantStyle = style
        .split(';')
        .map(rule => {
          rule = rule.trim();
          return rule && !rule.includes('!important') ? rule + ' !important' : rule;
        })
        .join(';');
      $(el).attr('style', importantStyle);
    }
  });
  
  return $.html();
}

// 创建和验证HTML文件的辅助函数
async function createAndValidateHtmlFile(finalHtml: string, options: any): Promise<string> {
  // Create secure temporary file with proper permissions
  const tempHtmlPath = createSecureTempPath('docx-conversion', '.html');
  // Write file with secure permissions
  await fs.writeFile(tempHtmlPath, finalHtml, { encoding: 'utf8', mode: defaultSecurityConfig.tempFilePermissions });
  // HTML file with style fixes created
  
  const writtenContent = await fs.readFile(tempHtmlPath, 'utf8');
  const validationResult = validateHtmlContent(writtenContent);
  
  console.error('🔍 样式修复验证:', {
    filePath: tempHtmlPath,
    fileSize: writtenContent.length,
    ...validationResult,
    contentPreview: writtenContent.substring(0, 500) + '...',
  });
  
  if (!validationResult.isValid) {
    await forceInjectWordStyles(tempHtmlPath, writtenContent, options);
  }
  
  return tempHtmlPath;
}

// 验证HTML内容的辅助函数
function validateHtmlContent(content: string) {
  const styleTagsMatch = content.match(/<style[^>]*>[\s\S]*?<\/style>/gi) || [];
  const hasStyle = styleTagsMatch.length > 0;
  const hasDoctype = content.includes('<!DOCTYPE');
  const hasWordStyles = content.includes('Microsoft YaHei') || content.includes('Calibri');
  
  const hasStyleContent = styleTagsMatch.some(tag => {
    const styleContent = tag.replace(/<style[^>]*>/i, '').replace(/<\/style>/i, '');
    return styleContent.trim().length > 0;
  });
  
  const hasCssRules = styleTagsMatch.some(tag => {
    const styleContent = tag.replace(/<style[^>]*>/i, '').replace(/<\/style>/i, '');
    return styleContent.includes('{') && styleContent.includes('}');
  });
  
  return {
    hasStyle,
    hasStyleContent,
    hasCssRules,
    hasDoctype,
    hasWordStyles,
    styleTagCount: styleTagsMatch.length,
    isValid: hasStyle && hasStyleContent && hasCssRules && hasWordStyles
  };
}

// 强制注入Word样式的辅助函数
// Clean up temporary files securely
async function cleanupTempFile(filePath: string): Promise<void> {
  try {
    await fs.unlink(filePath);
  } catch (error) {
    // File may already be deleted, ignore error
  }
}

async function forceInjectWordStyles(tempHtmlPath: string, content: string, options: any): Promise<void> {
  // Style fix failed, manually injecting Word styles
  
  const perfectHtml = createPerfectWordHtml('', options);
  const wordStylesRaw = perfectHtml.match(/<style[^>]*>[\s\S]*?<\/style>/gi)?.[0] ?? '';
  
  // Extract and sanitize CSS content from style tags
  const cssContent = wordStylesRaw.replace(/<\/?style[^>]*>/gi, '');
  const sanitizedCss = sanitizeCssProperty(cssContent);
  const wordStyles = `<style type="text/css">${sanitizedCss}</style>`;
  
  let bodyContent = extractBodyContent(content);
  bodyContent = bodyContent.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');
  // Escape HTML content to prevent XSS
  const sanitizedBodyContent = escapeHtml(bodyContent);
  
  const enforcedHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    ${wordStyles}
</head>
<body>
${sanitizedBodyContent}
</body>
</html>`;
  
  // Write file with secure permissions
  await fs.writeFile(tempHtmlPath, enforcedHtml, { encoding: 'utf8', mode: defaultSecurityConfig.tempFilePermissions });
  // Style injection completed
}

// 提取body内容的辅助函数
function extractBodyContent(content: string): string {
  if (content.includes('<body') && content.includes('</body>')) {
    const bodyMatch = content.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    return bodyMatch?.[1] ?? content;
  } else if (content.includes('<html') || content.includes('<!DOCTYPE')) {
    return content
      .replace(/<!DOCTYPE[^>]*>/gi, '')
      .replace(/<html[^>]*>/gi, '')
      .replace(/<\/html>/gi, '')
      .replace(/<head[^>]*>[\s\S]*?<\/head>/gi, '')
      .replace(/<body[^>]*>/gi, '')
      .replace(/<\/body>/gi, '');
  }
  return content;
}

// 创建Playwright指令的辅助函数
function createPlaywrightInstructions(finalOutputPath: string, tempHtmlPath: string, options: any) {
  return {
    success: true,
    requiresPlaywright: true,
    message: '请调用 playwright-mcp 工具来完成 PDF 生成',
    nextAction: {
      callPlaywrightMcp: true,
      mcpTool: 'playwright-mcp',
      instructions: 'Call playwright-mcp tools with the following sequence:',
      sequence: [
        '1. browser_launch - Launch headless browser',
        '2. browser_new_page - Create new page',
        '3. page_goto - Navigate to HTML file',
        '4. page_wait_for_timeout - Wait for page load',
        '5. page_pdf - Generate PDF',
        '6. browser_close - Close browser',
      ],
    },
    playwrightInstructions: createPlaywrightSteps(finalOutputPath, tempHtmlPath),
    outputPath: finalOutputPath,
    tempHtmlPath: tempHtmlPath,
    cleanupRequired: !options.useExistingHtml,
  };
}

// 创建Playwright步骤的辅助函数
function createPlaywrightSteps(finalOutputPath: string, tempHtmlPath: string) {
  return [
    {
      action: 'browser_launch',
      params: {
        headless: true,
        args: [
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage',
          '--disable-gpu',
          '--font-render-hinting=none',
          '--disable-font-subpixel-positioning',
        ],
      },
    },
    {
      action: 'browser_new_page',
      params: {},
    },
    {
      action: 'page_goto',
      params: {
        url: `file://${tempHtmlPath}`,
        options: {
          waitUntil: 'networkidle',
          timeout: 30000,
        },
      },
    },
    {
      action: 'page_add_script_tag',
      params: {
        content: `
          function ensureStylesApplied() {
            const styles = document.querySelectorAll('style');
            if (styles.length === 0) {
              console.error('警告: 未找到样式标签');
              const style = document.createElement('style');
              style.textContent = 'body { font-family: "Calibri", "Microsoft YaHei", sans-serif !important; } * { -webkit-print-color-adjust: exact !important; }';
              document.head.appendChild(style);
            }
            styles.forEach(style => {
              if (style.sheet) {
                // Style sheet loaded with rules
              }
            });
          }
          ensureStylesApplied();
        `,
      },
    },
    {
      action: 'page_wait_for_timeout',
      params: {
        timeout: 3000,
      },
    },
    {
      action: 'page_evaluate',
      params: {
        function: () => {
          document.querySelectorAll('style').forEach(style => {
            if (style.sheet) {
              // Style sheet loaded with rules
            }
          });
          document.body.style.visibility = 'hidden';
          setTimeout(() => {
            document.body.style.visibility = 'visible';
          }, 50);
          return 'CSS样式已确认加载';
        },
      },
    },
    {
      action: 'page_pdf',
      params: {
        path: finalOutputPath,
        options: {
          format: 'A4',
          printBackground: true,
          preferCSSPageSize: true,
          margin: {
            top: '0.75in',
            right: '0.75in',
            bottom: '0.75in',
            left: '0.75in',
          },
          styleTagsTimeout: 5000,
        },
      },
    },
    {
      action: 'browser_close',
      params: {},
    },
  ];
}

// 终极版Word样式HTML生成器 - 确保100%样式还原
function createPerfectWordHtml(content: string, options: any = {}): string {
  const chineseFont = options.chineseFont ?? 'Microsoft YaHei';

  // 终极版Word样式 - 包含所有可能的样式强制覆盖
  const ultimateWordStyles = `
    /* ===== 终极Word样式强制覆盖 - 100%还原 ===== */
    
    /* 打印设置 - 增强版 */
    @media print {
      html, body {
        font-size: 11pt !important;
        line-height: 1.08 !important;
        color: #000000 !important;
        background: #ffffff !important;
        -webkit-print-color-adjust: exact !important;
        color-adjust: exact !important;
        print-color-adjust: exact !important;
        margin: 0 !important;
        padding: 0 !important;
      }
      * {
        -webkit-print-color-adjust: exact !important;
        color-adjust: exact !important;
        print-color-adjust: exact !important;
      }
      img, table, figure {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
      }
      h1, h2, h3, h4, h5, h6 {
        page-break-after: avoid !important;
        break-after: avoid !important;
      }
    }
    
    @page { 
      size: A4 portrait !important; 
      margin: 2.54cm 1.91cm 2.54cm 1.91cm !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* 强制全局重置 */
    * {
      box-sizing: border-box !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* 强制字体和基础样式 */
    html, body, div, span, applet, object, iframe,
    h1, h2, h3, h4, h5, h6, p, blockquote, pre,
    a, abbr, acronym, address, big, cite, code,
    del, dfn, em, img, ins, kbd, q, s, samp,
    small, strike, strong, sub, sup, tt, var,
    b, u, i, center, input, textarea, select,
    dl, dt, dd, ol, ul, li,
    fieldset, form, label, legend,
    table, caption, tbody, tfoot, thead, tr, th, td,
    article, aside, canvas, details, embed, 
    figure, figcaption, footer, header, hgroup, 
    menu, nav, output, ruby, section, summary,
    time, mark, audio, video {
      font-family: "Calibri", "${chineseFont}", "SimSun", "宋体", "Microsoft YaHei UI", "Microsoft YaHei", sans-serif !important;
      font-size: 11pt !important;
      line-height: 1.08 !important;
      color: #000000 !important;
      margin: 0 !important;
      padding: 0 !important;
      background: #ffffff !important;
      text-rendering: optimizeLegibility !important;
      -webkit-font-smoothing: antialiased !important;
      -moz-osx-font-smoothing: grayscale !important;
      word-wrap: break-word !important;
      overflow-wrap: break-word !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* 根元素样式强制 */
    html, body {
      width: 100% !important;
      height: 100% !important;
      margin: 0 !important;
      padding: 0 !important;
      background-color: #ffffff !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* 标题样式 - Word 2019标准 */
    h1, h1.heading-1, h1.title, .Title, .title {
      font-family: "Calibri Light", "Calibri", "${chineseFont}", "Microsoft YaHei UI", sans-serif !important;
      font-size: 16pt !important;
      font-weight: normal !important;
      margin: 24pt 0 0pt 0 !important;
      color: #2F5496 !important;
      line-height: 1.15 !important;
      page-break-after: avoid !important;
      break-after: avoid !important;
      text-align: left !important;
      letter-spacing: 0 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    h2, h2.heading-2, h2.subtitle, .Heading2, .subtitle {
      font-family: "Calibri Light", "Calibri", "${chineseFont}", "Microsoft YaHei UI", sans-serif !important;
      font-size: 13pt !important;
      font-weight: normal !important;
      margin: 10pt 0 0pt 0 !important;
      color: #2F5496 !important;
      line-height: 1.15 !important;
      page-break-after: avoid !important;
      break-after: avoid !important;
      text-align: left !important;
      letter-spacing: 0 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    h3, h3.heading-3, .Heading3 {
      font-family: "Calibri Light", "Calibri", "${chineseFont}", "Microsoft YaHei UI", sans-serif !important;
      font-size: 12pt !important;
      font-weight: normal !important;
      margin: 10pt 0 0pt 0 !important;
      color: #1F3763 !important;
      line-height: 1.15 !important;
      page-break-after: avoid !important;
      break-after: avoid !important;
      text-align: left !important;
      letter-spacing: 0 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    h4, h4.heading-4, .Heading4 {
      font-family: "Calibri", "${chineseFont}", "Microsoft YaHei UI", sans-serif !important;
      font-size: 11pt !important;
      font-weight: bold !important;
      margin: 10pt 0 0pt 0 !important;
      color: #2F5496 !important;
      line-height: 1.15 !important;
      text-align: left !important;
      letter-spacing: 0 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    h5, h5.heading-5, .Heading5 {
      font-family: "Calibri", "${chineseFont}", "Microsoft YaHei UI", sans-serif !important;
      font-size: 11pt !important;
      font-weight: bold !important;
      margin: 10pt 0 0pt 0 !important;
      color: #2F5496 !important;
      line-height: 1.15 !important;
      text-align: left !important;
      letter-spacing: 0 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    h6, h6.heading-6, .Heading6 {
      font-family: "Calibri", "${chineseFont}", "Microsoft YaHei UI", sans-serif !important;
      font-size: 11pt !important;
      font-weight: bold !important;
      margin: 10pt 0 0pt 0 !important;
      color: #2F5496 !important;
      line-height: 1.15 !important;
      text-align: left !important;
      letter-spacing: 0 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* 段落样式 - Word标准 */
    p, p.normal, p.body-text, div, span, .Normal, .MsoNormal {
      font-family: "Calibri", "${chineseFont}", "SimSun", "宋体", "Microsoft YaHei UI", sans-serif !important;
      font-size: 11pt !important;
      line-height: 1.08 !important;
      margin: 0pt 0pt 8pt 0pt !important;
      text-align: left !important;
      text-indent: 0pt !important;
      orphans: 2 !important;
      widows: 2 !important;
      color: #000000 !important;
      letter-spacing: 0 !important;
      page-break-inside: auto !important;
      break-inside: auto !important;
      page-break-after: auto !important;
      break-after: auto !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* 文本格式化 */
    strong, strong.word-bold, b, .Bold {
      font-weight: bold !important;
      color: inherit !important;
    }
    
    em, em.word-italic, i, .Italic {
      font-style: italic !important;
      color: inherit !important;
    }
    
    span.word-underline, u, .Underline {
      text-decoration: underline !important;
      color: inherit !important;
    }
    
    span.word-strikethrough, del, strike, .Strikethrough {
      text-decoration: line-through !important;
      color: inherit !important;
    }
    
    /* 表格样式 */
    table, table.word-table, .MsoTableGrid {
      width: 100% !important;
      border-collapse: collapse !important;
      margin: 0pt 0pt 8pt 0pt !important;
      font-family: "Calibri", "${chineseFont}", "SimSun", "宋体", "Microsoft YaHei UI", sans-serif !important;
      font-size: 11pt !important;
      page-break-inside: avoid !important;
      break-inside: avoid !important;
      border: none !important;
      line-height: 1.08 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    th, th.word-header, td, td.word-cell {
      border: 0.5pt solid #000000 !important;
      padding: 0pt 5.4pt !important;
      background-color: transparent !important;
      font-weight: normal !important;
      text-align: left !important;
      vertical-align: top !important;
      font-family: "Calibri", "${chineseFont}", "SimSun", "宋体", "Microsoft YaHei UI", sans-serif !important;
      font-size: 11pt !important;
      line-height: 1.08 !important;
      color: #000000 !important;
      letter-spacing: 0 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    tr, tr.word-row {
      page-break-inside: avoid !important;
      break-inside: avoid !important;
    }
    
    /* 列表样式 */
    ul, ul.word-list, ol, ol.word-list {
      margin: 0pt 0pt 8pt 0pt !important;
      padding-left: 36pt !important;
      font-family: "Calibri", "${chineseFont}", "SimSun", "宋体", "Microsoft YaHei UI", sans-serif !important;
      font-size: 11pt !important;
      line-height: 1.08 !important;
      color: #000000 !important;
      letter-spacing: 0 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    li, li.word-list-item {
      margin: 0pt 0pt 0pt 0pt !important;
      padding: 0pt !important;
      font-family: "Calibri", "${chineseFont}", "SimSun", "宋体", "Microsoft YaHei UI", sans-serif !important;
      font-size: 11pt !important;
      line-height: 1.08 !important;
      color: #000000 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    ul li, ul.word-list li.word-list-item {
      list-style-type: disc !important;
    }
    
    ol li, ol.word-list li.word-list-item {
      list-style-type: decimal !important;
    }
    
    /* 图片样式 */
    img {
      max-width: 100% !important;
      height: auto !important;
      display: inline-block !important;
      margin: 0pt !important;
      vertical-align: baseline !important;
      page-break-inside: avoid !important;
      break-inside: avoid !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* 链接样式 */
    a, a:link, a:visited {
      color: #0563C1 !important;
      text-decoration: underline !important;
      font-family: inherit !important;
      font-size: inherit !important;
    }
    
    a:visited {
      color: #954F72 !important;
    }
    
    /* 特殊样式类 */
    .MsoNormal {
      margin: 0pt 0pt 8pt 0pt !important;
      line-height: 1.08 !important;
    }
    
    .MsoBodyText {
      margin: 0pt 0pt 8pt 0pt !important;
      line-height: 1.08 !important;
    }
    
    /* 防止样式被覆盖 */
    [style*="font-family"], [style*="font-size"], [style*="color"] {
      font-family: inherit !important;
      font-size: inherit !important;
      color: inherit !important;
    }
    
    /* 全局强制样式 */
    * {
      font-family: "Calibri", "${chineseFont}", "SimSun", "宋体", "Microsoft YaHei UI", sans-serif !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* 强制Word样式 */
    body * {
      font-family: "Calibri", "${chineseFont}", "SimSun", "宋体", "Microsoft YaHei UI", sans-serif !important;
      line-height: 1.08 !important;
      -webkit-print-color-adjust: exact !important;
      color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* 强制打印背景色和图片 */
    @media print {
      body {
        -webkit-print-color-adjust: exact !important;
        color-adjust: exact !important;
        print-color-adjust: exact !important;
        background-color: white !important;
      }
    }
  `;

  // 使用全局的安全HTML处理函数
  const safeContent = sanitizeHtmlForOutput(content);

  // 生成完整的HTML文档，确保包含所有必要的元素和样式
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <title>Word Document</title>
  <style>${ultimateWordStyles}</style>
</head>
<body>
  ${safeContent}
</body>
</html>`;
}

// Watermark function
async function addWatermark(pdfPath: string, options: WatermarkOptions = {}) {
  try {
    const pdfBytes = await fs.readFile(pdfPath);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const pages = pdfDoc.getPages();

    for (const page of pages) {
      const { width, height } = page.getSize();

      if (options.watermarkImage) {
        const imageBytes = await fs.readFile(options.watermarkImage);
        const image = await pdfDoc.embedPng(imageBytes);
        const scale = options.watermarkImageScale ?? 0.25;
        const opacity = options.watermarkImageOpacity ?? 0.6;

        let x = 0,
          y = 0;
        const position = options.watermarkImagePosition ?? 'top-right';

        switch (position) {
          case 'top-left':
            x = 20;
            y = height - image.height * scale - 20;
            break;
          case 'top-right':
            x = width - image.width * scale - 20;
            y = height - image.height * scale - 20;
            break;
          case 'bottom-left':
            x = 20;
            y = 20;
            break;
          case 'bottom-right':
            x = width - image.width * scale - 20;
            y = 20;
            break;
          case 'center':
            x = (width - image.width * scale) / 2;
            y = (height - image.height * scale) / 2;
            break;
          case 'fullscreen':
            // 自适应铺满屏幕，居中显示
            const imageAspectRatio = image.width / image.height;
            const pageAspectRatio = width / height;
            
            let newWidth, newHeight;
            if (imageAspectRatio > pageAspectRatio) {
              // 图片更宽，以页面宽度为准
              newWidth = width * 0.8; // 留出一些边距
              newHeight = newWidth / imageAspectRatio;
            } else {
              // 图片更高，以页面高度为准
              newHeight = height * 0.8; // 留出一些边距
              newWidth = newHeight * imageAspectRatio;
            }
            
            x = (width - newWidth) / 2;
            y = (height - newHeight) / 2;
            
            // 重新设置绘制参数，忽略原始的scale
            page.drawImage(image, {
              x,
              y,
              width: newWidth,
              height: newHeight,
              opacity: opacity * 0.3, // 降低透明度，避免影响阅读
            });
            continue; // 跳过后面的默认绘制
        }

        page.drawImage(image, {
          x,
          y,
          width: image.width * scale,
          height: image.height * scale,
          opacity,
        });
      }

      if (options.watermarkText) {
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const fontSize = options.watermarkFontSize ?? 48;
        const opacity = options.watermarkTextOpacity ?? 0.3;

        page.drawText(options.watermarkText, {
          x: width / 2 - (options.watermarkText.length * fontSize) / 4,
          y: height / 2,
          size: fontSize,
          font,
          color: rgb(0.5, 0.5, 0.5),
          opacity,
          rotate: { angle: Math.PI / 4, origin: { x: width / 2, y: height / 2 } },
        });
      }
    }

    const modifiedPdfBytes = await pdfDoc.save();
    await fs.writeFile(pdfPath, modifiedPdfBytes);

    return {
      success: true,
      message: `Watermark added to ${pdfPath}`,
    };
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

// Helper function for QR code position calculation
function calculateQRPosition(position: string, width: number, height: number, qrWidth: number, qrHeight: number): { x: number, y: number } {
  let x = 0, y = 0;

  switch (position) {
    case 'top-left':
      x = 20;
      y = height - qrHeight - 20;
      break;
    case 'top-right':
      x = width - qrWidth - 20;
      y = height - qrHeight - 20;
      break;
    case 'top-center':
      x = (width - qrWidth) / 2;
      y = height - qrHeight - 20;
      break;
    case 'bottom-left':
      x = 20;
      y = 20;
      break;
    case 'bottom-right':
      x = width - qrWidth - 20;
      y = 20;
      break;
    case 'bottom-center':
      x = (width - qrWidth) / 2;
      y = 20;
      break;
    case 'center':
      x = (width - qrWidth) / 2;
      y = (height - qrHeight) / 2;
      break;
  }

  return { x, y };
}

// QR Code function
async function addQRCode(pdfPath: string, qrCodePath?: string, options: QRCodeOptions = {}) {
  try {
    const finalQrPath = qrCodePath ?? defaultResourcePaths.defaultQrCodePath;

    if (!fsSync.existsSync(finalQrPath)) {
      return {
        success: false,
        error: `QR code image not found: ${finalQrPath}`,
      };
    }

    const pdfBytes = await fs.readFile(pdfPath);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const pages = pdfDoc.getPages();

    const qrImageBytes = await fs.readFile(finalQrPath);
    const qrImage = await pdfDoc.embedPng(qrImageBytes);

    for (const page of pages) {
      const { width, height } = page.getSize();
      const scale = options.qrScale ?? 0.15;
      const opacity = options.qrOpacity ?? 1.0;
      const position = options.qrPosition ?? 'bottom-center';

      const qrWidth = qrImage.width * scale;
      const qrHeight = qrImage.height * scale;
      const { x, y } = calculateQRPosition(position, width, height, qrWidth, qrHeight);

      page.drawImage(qrImage, {
        x,
        y,
        width: qrWidth,
        height: qrHeight,
        opacity,
      });

      if (options.addText !== false) {
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const text = options.customText ?? 'Scan QR code for more information';
        const textSize = options.textSize ?? 8;
        const textColor = options.textColor ?? '#000000';

        page.drawText(text, {
          x: x + (qrWidth - text.length * textSize * 0.6) / 2,
          y: y - 15,
          size: textSize,
          font,
          color: rgb(0, 0, 0),
        });
      }
    }

    const modifiedPdfBytes = await pdfDoc.save();
    await fs.writeFile(pdfPath, modifiedPdfBytes);

    return {
      success: true,
      message: `QR code added to ${pdfPath}`,
    };
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

// Helper functions for processPdfPostConversion
function resolvePostProcessingPath(playwrightPdfPath: string, targetPath?: string): string {
  const outputDir = process.env.OUTPUT_DIR ?? path.dirname(playwrightPdfPath);
  const allowedPaths = [outputDir, defaultResourcePaths.outputDir, process.cwd()];
  
  if (!targetPath) {
    // Extract original filename from playwright path
    const playwrightFileName = path.basename(playwrightPdfPath);
    // Remove timestamp prefix and decode the filename
    const decodedName = playwrightFileName
      .replace(/^\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}\.\d{3}Z-/, '')
      .replace(/-/g, '/');
    
    // Sanitize the decoded name to prevent path traversal attacks
    const sanitizedName = path.basename(decodedName).replace(/[^a-zA-Z0-9._-]/g, '_');
    return safePathJoin(outputDir, sanitizedName);
  } else if (!path.isAbsolute(targetPath)) {
    // Use safe path joining for relative paths
    return safePathJoin(outputDir, path.basename(targetPath));
  }

  // For absolute paths, use our validation function
  return validatePath(targetPath, allowedPaths);
}

async function processWatermarkAddition(finalPath: string, options: any): Promise<any> {
  if (!(options.addWatermark || process.env.WATERMARK_IMAGE)) {
    return null;
  }

  const watermarkOptions = {
    watermarkImage: options.watermarkImage ?? process.env.WATERMARK_IMAGE,
    watermarkText: options.watermarkText,
    watermarkImageScale: options.watermarkImageScale ?? 0.25,
    watermarkImageOpacity: options.watermarkImageOpacity ?? 0.6,
    watermarkImagePosition: options.watermarkImagePosition ?? 'top-right',
    watermarkFontSize: options.watermarkFontSize ?? 48,
    watermarkTextOpacity: options.watermarkTextOpacity ?? 0.3,
  };

  try {
    return await addWatermark(finalPath, watermarkOptions);
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

async function processQRCodeAddition(finalPath: string, options: any): Promise<any> {
  if (!(options.addQrCode || (options.addQrCode !== false && process.env.QR_CODE_IMAGE))) {
    return null;
  }

  const qrOptions = {
    qrScale: options.qrScale ?? 0.15,
    qrOpacity: options.qrOpacity ?? 1.0,
    qrPosition: options.qrPosition ?? 'bottom-center',
    addText: options.addText !== false,
    customText: options.customText ?? 'Scan QR code for more information',
    textSize: options.textSize ?? 8,
    textColor: options.textColor ?? '#000000',
  };

  try {
    const qrCodePath = options.qrCodePath || process.env.QR_CODE_IMAGE;
    if (qrCodePath) {
      return await addQRCode(finalPath, qrCodePath, qrOptions);
    } else {
      return { success: false, error: 'QR code path not provided' };
    }
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

// Unified PDF post-processing function
async function processPdfPostConversion(
  playwrightPdfPath: string,
  targetPath?: string,
  options: any = {}
) {
  try {
    const finalPath = resolvePostProcessingPath(playwrightPdfPath, targetPath);

    // Ensure output directory exists
    await fs.mkdir(path.dirname(finalPath), { recursive: true });

    // Copy the PDF from playwright temp location to target location
    await fs.copyFile(playwrightPdfPath, finalPath);

    const results: any[] = [];

    // Add watermark if specified
    const watermarkResult = await processWatermarkAddition(finalPath, options);
    if (watermarkResult) {
      results.push(watermarkResult);
    }

    // Add QR code if specified
    const qrResult = await processQRCodeAddition(finalPath, options);
    if (qrResult) {
      results.push(qrResult);
    }

    // Clean up temporary playwright file if it's different from final path
    if (playwrightPdfPath !== finalPath && fsSync.existsSync(playwrightPdfPath)) {
      await cleanupTempFile(playwrightPdfPath);
    }

    return {
      success: true,
      finalPath,
      originalPath: playwrightPdfPath,
      processedFeatures: results
        .filter(r => r && r.success)
        .map(r => r.message || 'Feature processed'),
      errors: results.filter(r => r && !r.success).map(r => r.error || 'Unknown error'),
    };
  } catch (error: any) {
    return {
      success: false,
      error: error.message,
      originalPath: playwrightPdfPath,
    };
  }
}

// Helper functions for convertMarkdownToPdf
function resolveMarkdownPdfOutputPath(inputPath: string, outputPath?: string): string {
  if (!outputPath) {
    const baseName = path.basename(inputPath, path.extname(inputPath));
    return path.join(defaultResourcePaths.outputDir, `${baseName}.pdf`);
  } else if (!path.isAbsolute(outputPath)) {
    return path.join(defaultResourcePaths.outputDir, outputPath);
  }
  return outputPath;
}

function logMarkdownConversionStart(inputPath: string, outputPath: string): void {
  console.error(`🔄 Markdown 到 PDF 转换...`);
  console.error(`📄 输入: ${inputPath}`);
  console.error(`📁 输出: ${outputPath}`);
}

function createMarkdownPlaywrightCommands(htmlOutputPath: string, finalOutputPath: string): string[] {
  return [
    `browser_navigate("file://${htmlOutputPath}")`,
    `browser_wait_for({ time: 3 })`,
    `browser_pdf_save({ filename: "${finalOutputPath}" })`,
  ];
}

function createMarkdownPostProcessingSteps(options: any): string[] {
  const postProcessingSteps: string[] = [];
  const defaultWatermarkPath = process.env.WATERMARK_IMAGE ?? null;
  const defaultQrCodePath = process.env.QR_CODE_IMAGE ?? null;
  const addQrCode = options.addQrCode ?? false;

  if (defaultWatermarkPath) {
    postProcessingSteps.push(`添加水印: ${defaultWatermarkPath}`);
  }
  if (addQrCode && defaultQrCodePath) {
    postProcessingSteps.push(`添加二维码: ${defaultQrCodePath}`);
  }
  return postProcessingSteps;
}

// Markdown 转 PDF 函数
// 注意：此函数实际上是先将 Markdown 转换为 HTML，然后需要使用 playwright-mcp 完成最终的 PDF 转换
// 转换流程：Markdown → HTML → PDF (通过 playwright-mcp)
async function convertMarkdownToPdf(inputPath: string, outputPath?: string, options: any = {}) {
  try {
    const finalOutputPath = resolveMarkdownPdfOutputPath(inputPath, outputPath);
    logMarkdownConversionStart(inputPath, finalOutputPath);

    // 确保输出目录存在
    await fs.mkdir(path.dirname(finalOutputPath), { recursive: true });

    // 第一步：转换 Markdown 到 HTML
    const htmlOutputPath = finalOutputPath!.replace('.pdf', '.html');
    const htmlResult = await convertMarkdownToHtml(inputPath, {
      outputPath: htmlOutputPath,
      theme: 'github',
      standalone: true,
    });

    if (!htmlResult.success) {
      throw new Error(`Markdown 到 HTML 转换失败: ${htmlResult.error}`);
    }

    console.error(`✅ Markdown 到 HTML 转换成功: ${htmlOutputPath}`);
    console.error(`📋 下一步：使用 playwright-mcp 将 HTML 转换为 PDF`);

    // 获取水印和二维码配置
    const defaultWatermarkPath = process.env.WATERMARK_IMAGE ?? null;
    const defaultQrCodePath = process.env.QR_CODE_IMAGE ?? null;
    const addQrCode = options.addQrCode ?? false;

    // 构建 playwright 命令，包含水印和二维码处理
    const playwrightCommands = createMarkdownPlaywrightCommands(htmlOutputPath, finalOutputPath);

    // 如果有水印或二维码需要添加，在 playwright 命令后添加处理步骤
    const postProcessingSteps = createMarkdownPostProcessingSteps(options);

    return {
      success: true,
      outputPath: htmlOutputPath,
      finalPdfPath: finalOutputPath,
      requiresPlaywrightMcp: true,
      message: 'Markdown 已转换为 HTML，需要使用 playwright-mcp 完成 PDF 转换',
      watermarkPath: defaultWatermarkPath,
      qrCodePath: addQrCode ? defaultQrCodePath : null,
      instructions: {
        step1: '已完成：Markdown → HTML',
        step2: `待完成：使用 playwright-mcp 打开 ${htmlOutputPath} 并保存为 PDF`,
        step3: postProcessingSteps.length > 0 ? `后处理：${postProcessingSteps.join(', ')}` : null,
        playwrightCommands: playwrightCommands,
      },
    };
  } catch (error: any) {
    console.error('❌ Markdown 到 PDF 转换失败:', error.message);
    return {
      success: false,
      error: error.message,
    };
  }
}

// Helper functions for convertDocxToMarkdown
function resolveDocxToMarkdownOutputPath(inputPath: string, outputPath?: string): string {
  if (!outputPath) {
    const baseName = path.basename(inputPath, path.extname(inputPath));
    return path.join(defaultResourcePaths.outputDir, `${baseName}.md`);
  } else if (!path.isAbsolute(outputPath)) {
    return path.join(defaultResourcePaths.outputDir, outputPath);
  }
  return outputPath;
}

function logDocxToMarkdownConversionStart(inputPath: string, outputPath: string): void {
  console.error(`🔄 DOCX 到 Markdown 转换...`);
  console.error(`📄 输入: ${inputPath}`);
  console.error(`📁 输出: ${outputPath}`);
}

// DOCX 转 Markdown 函数
async function convertDocxToMarkdown(inputPath: string, outputPath?: string, options: any = {}) {
  try {
    const finalOutputPath = resolveDocxToMarkdownOutputPath(inputPath, outputPath);
    logDocxToMarkdownConversionStart(inputPath, finalOutputPath);

    // 确保输出目录存在
    await fs.mkdir(path.dirname(finalOutputPath), { recursive: true });

    // 使用优化的转换器
    const result = await optimizedDocxToMarkdown(inputPath, {
      outputPath: finalOutputPath,
      ...options,
    });

    if (result.success) {
      console.error(`✅ DOCX 到 Markdown 转换成功: ${finalOutputPath}`);
      return {
        success: true,
        outputPath: finalOutputPath,
        message: 'DOCX 到 Markdown 转换完成',
      };
    } else {
      throw new Error(result.error || 'DOCX 到 Markdown 转换失败');
    }
  } catch (error: any) {
    console.error('❌ DOCX 到 Markdown 转换失败:', error.message);
    return {
      success: false,
      error: error.message,
    };
  }
}

// Helper functions for convertDocxToTxt
function resolveDocxToTxtOutputPath(inputPath: string, outputPath?: string): string {
  if (!outputPath) {
    const baseName = path.basename(inputPath, path.extname(inputPath));
    return path.join(defaultResourcePaths.outputDir, `${baseName}.txt`);
  } else if (!path.isAbsolute(outputPath)) {
    return path.join(defaultResourcePaths.outputDir, outputPath);
  }
  return outputPath;
}

function logDocxToTxtConversionStart(inputPath: string, outputPath: string): void {
  console.error(`🔄 DOCX 到 TXT 转换...`);
  console.error(`📄 输入: ${inputPath}`);
  console.error(`📁 输出: ${outputPath}`);
}

// DOCX 转 TXT 函数
async function convertDocxToTxt(inputPath: string, outputPath?: string, options: any = {}) {
  try {
    const finalOutputPath = resolveDocxToTxtOutputPath(inputPath, outputPath);
    logDocxToTxtConversionStart(inputPath, finalOutputPath);

    // 确保输出目录存在
    await fs.mkdir(path.dirname(finalOutputPath), { recursive: true });

    // 先转换为 HTML，然后提取纯文本
    const htmlResult = await convertDocxToHtmlEnhanced(inputPath);

    if (!htmlResult.success) {
      throw new Error(`DOCX 到 HTML 转换失败: ${(htmlResult as any).error || '未知错误'}`);
    }

    // 从 HTML 中提取纯文本
    const $ = cheerio.load((htmlResult as any).content || '');

    // 移除样式和脚本标签
    $('style, script').remove();

    // 获取纯文本内容
    let textContent = $.text();

    // 清理多余的空白字符
    textContent = textContent
      .replace(/\n{3,}/g, '\n\n') // 将多个连续换行符替换为最多两个
      .replace(/[ \t]+/g, ' ') // 将多个空格或制表符替换为单个空格
      .replace(/^\s+|\s+$/gm, '') // 移除每行开头和结尾的空白字符
      .trim(); // 移除整个文本开头和结尾的空白字符

    // 写入文件
    await fs.writeFile(finalOutputPath, textContent, 'utf-8');

    console.error(`✅ DOCX 到 TXT 转换成功: ${finalOutputPath}`);
    return {
      success: true,
      outputPath: finalOutputPath,
      message: 'DOCX 到 TXT 转换完成',
    };
  } catch (error: any) {
    console.error('❌ DOCX 到 TXT 转换失败:', error.message);
    return {
      success: false,
      error: error.message,
    };
  }
}

// Markdown 转 TXT 函数
async function convertMarkdownToTxt(inputPath: string, outputPath?: string, options: any = {}) {
  try {
    // 使用环境变量控制的输出路径
    let finalOutputPath = outputPath;
    if (!finalOutputPath) {
      const baseName = path.basename(inputPath, path.extname(inputPath));
      finalOutputPath = path.join(defaultResourcePaths.outputDir, `${baseName}.txt`);
    } else if (!path.isAbsolute(finalOutputPath)) {
      finalOutputPath = path.join(defaultResourcePaths.outputDir, finalOutputPath);
    }

    console.error(`🔄 Markdown 到 TXT 转换...`);
    console.error(`📄 输入: ${inputPath}`);
    console.error(`📁 输出: ${finalOutputPath}`);

    // 确保输出目录存在
    await fs.mkdir(path.dirname(finalOutputPath), { recursive: true });

    // 读取 Markdown 文件
    const markdownContent = await fs.readFile(inputPath, 'utf-8');

    // 简单的 Markdown 到纯文本转换
    let textContent = markdownContent
      // 移除 Markdown 语法
      .replace(/^#{1,6}\s+/gm, '') // 移除标题标记
      .replace(/\*\*(.*?)\*\*/g, '$1') // 移除粗体标记
      .replace(/\*(.*?)\*/g, '$1') // 移除斜体标记
      .replace(/`(.*?)`/g, '$1') // 移除行内代码标记
      .replace(/```[\s\S]*?```/g, '') // 移除代码块
      .replace(/\[([^\]]+)\]\([^\)]+\)/g, '$1') // 移除链接，保留文本
      .replace(/!\[([^\]]*)\]\([^\)]*\)/g, '$1') // 移除图片，保留alt文本
      .replace(/^[-*+]\s+/gm, '') // 移除列表标记
      .replace(/^\d+\.\s+/gm, '') // 移除有序列表标记
      .replace(/^>\s+/gm, '') // 移除引用标记
      .replace(/\|.*\|/g, '') // 移除表格
      .replace(/^-{3,}$/gm, '') // 移除分隔线
      .replace(/\n{3,}/g, '\n\n') // 合并多个换行
      .trim();

    // 写入文件
    await fs.writeFile(finalOutputPath, textContent, 'utf-8');

    console.error(`✅ Markdown 到 TXT 转换成功: ${finalOutputPath}`);
    return {
      success: true,
      outputPath: finalOutputPath,
      message: 'Markdown 到 TXT 转换完成',
    };
  } catch (error: any) {
    console.error('❌ Markdown 到 TXT 转换失败:', error.message);
    return {
      success: false,
      error: error.message,
    };
  }
}

// MCP Server setup
const server = new Server(
  {
    name: 'doc-ops-mcp',
    version: '0.0.8',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Tool definitions
const TOOL_DEFINITIONS = {
  read_document: {
    name: 'read_document',
    description: 'Read various document formats (DOCX, DOC, TXT, MD, HTML, etc.)',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Document path to read' },
        extractMetadata: {
          type: 'boolean',
          description: 'Extract document metadata',
          default: false,
        },
        preserveFormatting: {
          type: 'boolean',
          description: 'Preserve formatting (HTML output)',
          default: false,
        },
      },
      required: ['filePath'],
    },
  },

  write_document: {
    name: 'write_document',
    description:
      'Write content to document files in specified formats. Output directory is controlled by OUTPUT_DIR environment variable. If outputPath is not provided, files will be saved to OUTPUT_DIR with auto-generated names. If outputPath is relative, it will be resolved relative to OUTPUT_DIR.',
    inputSchema: {
      type: 'object',
      properties: {
        content: { type: 'string', description: 'Content to write' },
        outputPath: {
          type: 'string',
          description:
            'Output file path (optional, auto-generated). If not absolute, will be resolved relative to OUTPUT_DIR environment variable.',
        },
        encoding: { type: 'string', description: 'File encoding', default: 'utf-8' },
      },
      required: ['content'],
    },
  },

  convert_document: {
    name: 'convert_document',
    description:
      "Convert documents between formats with enhanced style preservation. Output directory is controlled by OUTPUT_DIR environment variable. All output files will be saved to the directory specified by OUTPUT_DIR. ⚠️ IMPORTANT: For Markdown to HTML conversion with style preservation, use 'convert_markdown_to_html' tool instead for better results with themes and styling.",
    inputSchema: {
      type: 'object',
      properties: {
        inputPath: { type: 'string', description: 'Input file path' },
        outputPath: {
          type: 'string',
          description:
            'Output file path (optional, auto-generated). If not absolute, will be resolved relative to OUTPUT_DIR environment variable.',
        },
        preserveFormatting: { type: 'boolean', description: 'Preserve formatting', default: true },
        useInternalPlaywright: {
          type: 'boolean',
          description: 'Use built-in Playwright for PDF conversion',
          default: false,
        },
      },
      required: ['inputPath'],
    },
  },

  add_watermark: {
    name: 'add_watermark',
    description: 'Add watermarks (image or text) to PDF documents',
    inputSchema: {
      type: 'object',
      properties: {
        pdfPath: { type: 'string', description: 'PDF file path' },
        watermarkImage: { type: 'string', description: 'Watermark image path (PNG/JPG)' },
        watermarkText: { type: 'string', description: 'Watermark text (ASCII only)' },
        watermarkImageScale: { type: 'number', description: 'Image scale ratio', default: 0.25 },
        watermarkImageOpacity: { type: 'number', description: 'Image opacity', default: 0.6 },
        watermarkImagePosition: {
          type: 'string',
          enum: ['top-left', 'top-right', 'bottom-left', 'bottom-right', 'center', 'fullscreen'],
          default: 'fullscreen',
        },
      },
      required: ['pdfPath'],
    },
  },

  add_qrcode: {
    name: 'add_qrcode',
    description: 'Add QR code to PDF documents with friendly text below',
    inputSchema: {
      type: 'object',
      properties: {
        pdfPath: { type: 'string', description: 'PDF file path' },
        qrCodePath: { type: 'string', description: 'QR code image path (optional, uses default)' },
        qrScale: { type: 'number', description: 'QR code scale ratio', default: 0.15 },
        qrOpacity: { type: 'number', description: 'QR code opacity', default: 1.0 },
        qrPosition: {
          type: 'string',
          enum: [
            'top-left',
            'top-right',
            'top-center',
            'bottom-left',
            'bottom-right',
            'bottom-center',
            'center',
          ],
          default: 'bottom-center',
        },
        addText: { type: 'boolean', description: 'Add friendly text below QR code', default: true },
        customText: {
          type: 'string',
          description: "Custom text (default: 'Scan QR code for more information')",
        },
      },
      required: ['pdfPath'],
    },
  },

  convert_docx_to_pdf: {
    name: 'convert_docx_to_pdf',
    description:
      'Enhanced DOCX to PDF conversion with perfect Word style replication and Playwright integration. Automatically adds watermark if WATERMARK_IMAGE environment variable is set. QR code can be added by setting addQrCode=true (requires QR_CODE_IMAGE environment variable). Output directory is controlled by OUTPUT_DIR environment variable.',
    inputSchema: {
      type: 'object',
      properties: {
        docxPath: { type: 'string', description: 'DOCX file path to convert' },
        outputPath: {
          type: 'string',
          description:
            'PDF output path (optional, auto-generated). If not absolute, will be resolved relative to OUTPUT_DIR environment variable.',
        },
        preserveFormatting: {
          type: 'boolean',
          description: 'Preserve original formatting',
          default: true,
        },
        chineseFont: {
          type: 'string',
          description: 'Chinese font family to use',
          default: 'Microsoft YaHei',
        },
        addQrCode: {
          type: 'boolean',
          description: 'Add QR code to PDF (requires QR_CODE_IMAGE environment variable)',
          default: false,
        },
      },
      required: ['docxPath'],
    },
  },

  convert_markdown_to_html: {
    name: 'convert_markdown_to_html',
    description:
      'Enhanced Markdown to HTML conversion with beautiful styling and theme support. Supports GitHub, Academic, Modern, and Default themes with complete style preservation.',
    inputSchema: {
      type: 'object',
      properties: {
        markdownPath: { type: 'string', description: 'Markdown file path to convert' },
        outputPath: {
          type: 'string',
          description:
            'HTML output path (optional, auto-generated). If not absolute, will be resolved relative to OUTPUT_DIR environment variable.',
        },
        theme: {
          type: 'string',
          enum: ['default', 'github', 'academic', 'modern'],
          description: 'Theme to apply',
          default: 'github',
        },
        includeTableOfContents: {
          type: 'boolean',
          description: 'Generate table of contents',
          default: false,
        },
        customCSS: { type: 'string', description: 'Additional custom CSS styles' },
      },
      required: ['markdownPath'],
    },
  },

  convert_markdown_to_docx: {
    name: 'convert_markdown_to_docx',
    description:
      'Enhanced Markdown to DOCX conversion with professional styling and theme support. Preserves formatting, supports tables, lists, headings, and inline formatting with beautiful Word document output.',
    inputSchema: {
      type: 'object',
      properties: {
        markdownPath: { type: 'string', description: 'Markdown file path to convert' },
        outputPath: {
          type: 'string',
          description:
            'DOCX output path (optional, auto-generated). If not absolute, will be resolved relative to OUTPUT_DIR environment variable.',
        },
        theme: {
          type: 'string',
          enum: ['default', 'professional', 'academic', 'modern'],
          description: 'Theme to apply',
          default: 'professional',
        },
        includeTableOfContents: {
          type: 'boolean',
          description: 'Generate table of contents',
          default: false,
        },
        preserveStyles: {
          type: 'boolean',
          description: 'Preserve Markdown formatting and styles',
          default: true,
        },
      },
      required: ['markdownPath'],
    },
  },

  convert_markdown_to_pdf: {
    name: 'convert_markdown_to_pdf',
    description:
      'Enhanced Markdown to PDF conversion with beautiful styling and theme support. Automatically adds watermark if WATERMARK_IMAGE environment variable is set. QR code can be added by setting addQrCode=true (requires QR_CODE_IMAGE environment variable). Requires playwright-mcp for final PDF generation.',
    inputSchema: {
      type: 'object',
      properties: {
        markdownPath: { type: 'string', description: 'Markdown file path to convert' },
        outputPath: {
          type: 'string',
          description:
            'PDF output path (optional, auto-generated). If not absolute, will be resolved relative to OUTPUT_DIR environment variable.',
        },
        theme: {
          type: 'string',
          enum: ['default', 'github', 'academic', 'modern'],
          description: 'Theme to apply',
          default: 'github',
        },
        includeTableOfContents: {
          type: 'boolean',
          description: 'Generate table of contents',
          default: false,
        },
        addQrCode: {
          type: 'boolean',
          description: 'Add QR code to PDF (requires QR_CODE_IMAGE environment variable)',
          default: false,
        },
      },
      required: ['markdownPath'],
    },
  },

  convert_html_to_markdown: {
    name: 'convert_html_to_markdown',
    description:
      'Enhanced HTML to Markdown conversion with style preservation. Converts HTML files to clean Markdown format while preserving structure, links, images, tables, and formatting. Output directory is controlled by OUTPUT_DIR environment variable.',
    inputSchema: {
      type: 'object',
      properties: {
        htmlPath: { type: 'string', description: 'HTML file path to convert' },
        outputPath: {
          type: 'string',
          description:
            'Markdown output path (optional, auto-generated). If not absolute, will be resolved relative to OUTPUT_DIR environment variable.',
        },
        preserveStyles: {
          type: 'boolean',
          description: 'Preserve HTML formatting and styles',
          default: true,
        },
        includeCSS: {
          type: 'boolean',
          description: 'Include CSS styles as comments in Markdown',
          default: false,
        },
        debug: { type: 'boolean', description: 'Enable debug output', default: false },
      },
      required: ['htmlPath'],
    },
  },

  plan_conversion: {
    name: 'plan_conversion',
    description:
      "🎯 **MANDATORY FIRST STEP** - Smart Document Conversion Planner - ⚠️ CRITICAL: This tool MUST be executed BEFORE any document conversion operation! When users request any format conversion (e.g., 'convert MD to PDF', 'DOCX to HTML'), you MUST call this tool first to get the optimal conversion plan, then follow the plan's steps exactly. This ensures proper conversion workflow and handles complex multi-step conversions. For PDF conversions, note that playwright-mcp integration is required for final PDF generation, followed by process_pdf_post_conversion for watermarks/QR codes. Supports all format conversion path planning.",
    inputSchema: {
      type: 'object',
      properties: {
        sourceFormat: {
          type: 'string',
          description: 'Source file format (e.g.: pdf, docx, html, markdown, md, txt)',
          enum: ['pdf', 'docx', 'html', 'markdown', 'md', 'txt', 'doc'],
        },
        targetFormat: {
          type: 'string',
          description: 'Target file format (e.g.: pdf, docx, html, markdown, md, txt)',
          enum: ['pdf', 'docx', 'html', 'markdown', 'md', 'txt', 'doc'],
        },
        sourceFile: { type: 'string', description: 'Source file path (optional, used to generate specific conversion parameters)' },
        preserveStyles: { type: 'boolean', description: 'Whether to preserve style formatting', default: true },
        includeImages: { type: 'boolean', description: 'Whether to include images', default: true },
        theme: {
          type: 'string',
          description: 'Conversion theme (applicable to HTML/DOCX output)',
          enum: ['github', 'academic', 'modern', 'professional', 'default'],
          default: 'github',
        },
        quality: {
          type: 'string',
          description: 'Conversion quality requirements',
          enum: ['fast', 'balanced', 'high'],
          default: 'balanced',
        },
      },
      required: ['sourceFormat', 'targetFormat'],
    },
  },

  process_pdf_post_conversion: {
    name: 'process_pdf_post_conversion',
    description:
      '🔧 PDF Post-Processing Unified Tool - ⚠️ IMPORTANT: This is a necessary follow-up step for playwright-mcp\'s browser_pdf_save command! After using playwright-mcp to generate PDF, this tool must be called immediately to complete final processing. Features include: 1) Automatically move PDF from playwright temporary path to target location 2) Uniformly add watermarks and QR codes 3) Clean up temporary files. Workflow: playwright-mcp:browser_pdf_save → doc-ops-mcp:process_pdf_post_conversion',
    inputSchema: {
      type: 'object',
      properties: {
        playwrightPdfPath: {
          type: 'string',
          description: 'PDF file path generated by playwright-mcp (usually in temporary directory)',
        },
        targetPath: {
          type: 'string',
          description:
            'Target PDF file path (optional, will be auto-generated if not provided). If not absolute path, will be resolved relative to OUTPUT_DIR environment variable',
        },
        addWatermark: {
          type: 'boolean',
          description: 'Whether to add watermark (will be automatically added if WATERMARK_IMAGE environment variable is set)',
          default: false,
        },
        addQrCode: {
          type: 'boolean',
          description: 'Whether to add QR code (will be automatically added if QR_CODE_IMAGE environment variable is set)',
          default: false,
        },
        watermarkImage: { type: 'string', description: 'Watermark image path (overrides environment variable)' },
        watermarkText: { type: 'string', description: 'Watermark text content' },
        watermarkImageScale: { type: 'number', description: 'Watermark image scale ratio', default: 0.25 },
        watermarkImageOpacity: { type: 'number', description: 'Watermark image opacity', default: 0.6 },
        watermarkImagePosition: {
          type: 'string',
          enum: ['top-left', 'top-right', 'bottom-left', 'bottom-right', 'center', 'fullscreen'],
          description: 'Watermark image position',
          default: 'fullscreen',
        },
        qrCodePath: { type: 'string', description: 'QR code image path (overrides environment variable)' },
        qrScale: { type: 'number', description: 'QR code scale ratio', default: 0.15 },
        qrOpacity: { type: 'number', description: 'QR code opacity', default: 1.0 },
        qrPosition: {
          type: 'string',
          enum: [
            'top-left',
            'top-right',
            'top-center',
            'bottom-left',
            'bottom-right',
            'bottom-center',
            'center',
          ],
          description: 'QR code position',
          default: 'bottom-center',
        },
        customText: {
          type: 'string',
          description: 'Custom text below QR code',
          default: 'Scan QR code for more information',
        },
      },
      required: ['playwrightPdfPath'],
    },
  },
};

// List tools handler
server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    TOOL_DEFINITIONS.read_document,
    TOOL_DEFINITIONS.write_document,
    TOOL_DEFINITIONS.convert_document,
    TOOL_DEFINITIONS.add_watermark,
    TOOL_DEFINITIONS.add_qrcode,
    TOOL_DEFINITIONS.convert_docx_to_pdf,
    TOOL_DEFINITIONS.convert_markdown_to_html,
    TOOL_DEFINITIONS.convert_markdown_to_docx,
    TOOL_DEFINITIONS.convert_markdown_to_pdf,
    TOOL_DEFINITIONS.convert_html_to_markdown,
    TOOL_DEFINITIONS.plan_conversion,
    TOOL_DEFINITIONS.process_pdf_post_conversion,
    ...DUAL_PARSING_TOOLS,
    WEB_SCRAPING_TOOL,
    STRUCTURED_DATA_TOOL,
  ],
}));

// Call tool handler
server.setRequestHandler(CallToolRequestSchema, async request => {
  const { name, arguments: args } = request.params;

  try {
    const result = await handleToolCall(name, args);
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  } catch (error: any) {
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ success: false, error: error.message }, null, 2),
        },
      ],
      isError: true,
    };
  }
});

// 主要工具调用处理函数
async function handleToolCall(name: string, args: any) {
  // 基础文档操作
  if (isBasicDocumentOperation(name)) {
    return await handleBasicDocumentOperations(name, args);
  }
  
  // PDF相关操作
  if (isPdfOperation(name)) {
    return await handlePdfOperations(name, args);
  }
  
  // Markdown转换操作
  if (isMarkdownOperation(name)) {
    return await handleMarkdownOperations(name, args);
  }
  
  // HTML转换操作
  if (isHtmlOperation(name)) {
    return await handleHtmlOperations(name, args);
  }
  
  // 双重解析操作
  if (isDualParsingOperation(name)) {
    return await handleDualParsingOperations(name, args);
  }
  
  // 转换规划操作
  if (name === 'plan_conversion') {
    return await handleConversionPlanning(args);
  }
  
  throw new Error(`Unknown tool: ${name}`);
}

// 检查是否为基础文档操作
function isBasicDocumentOperation(name: string): boolean {
  return ['read_document', 'write_document', 'convert_document'].includes(name);
}

// 检查是否为PDF操作
function isPdfOperation(name: string): boolean {
  return ['process_pdf_post_conversion', 'convert_docx_to_pdf'].includes(name);
}

// 检查是否为Markdown操作
function isMarkdownOperation(name: string): boolean {
  return ['convert_markdown_to_html', 'convert_markdown_to_docx', 'convert_markdown_to_pdf'].includes(name);
}

// 检查是否为HTML操作
function isHtmlOperation(name: string): boolean {
  return ['convert_html_to_markdown'].includes(name);
}

// 检查是否为双重解析操作
function isDualParsingOperation(name: string): boolean {
  return ['dual_parsing_docx_to_html', 'dual_parsing_docx_to_pdf', 'docx_style_analysis'].includes(name);
}

// 处理基础文档操作
async function handleBasicDocumentOperations(name: string, args: any) {
  switch (name) {
    case 'read_document':
      return await readDocument(args.filePath, args);
    case 'write_document':
      return await writeDocument(args.content, args.outputPath, args);
    case 'convert_document':
      return await convertDocument(args.inputPath, args.outputPath, args);
    default:
      throw new Error(`Unknown basic document operation: ${name}`);
  }
}

// 处理PDF操作
async function handlePdfOperations(name: string, args: any) {
  switch (name) {
    case 'process_pdf_post_conversion':
      return await processPdfPostConversion(args.playwrightPdfPath, args.targetPath, args);
    case 'convert_docx_to_pdf':
      return await convertDocxToPdf(args.docxPath, args.outputPath, args);
    default:
      throw new Error(`Unknown PDF operation: ${name}`);
  }
}

// 处理Markdown操作
async function handleMarkdownOperations(name: string, args: any) {
  switch (name) {
    case 'convert_markdown_to_html':
      return await handleMarkdownToHtml(args);
    case 'convert_markdown_to_docx':
      return await handleMarkdownToDocx(args);
    case 'convert_markdown_to_pdf':
      return await handleMarkdownToPdf(args);
    default:
      throw new Error(`Unknown Markdown operation: ${name}`);
  }
}

// 处理Markdown到HTML转换
async function handleMarkdownToHtml(args: any) {
  const outputPath = resolveOutputPath(args.outputPath, args.markdownPath, '.html');
  
  return await convertMarkdownToHtml(args.markdownPath, {
    theme: args.theme ?? 'github',
    includeTableOfContents: args.includeTableOfContents ?? false,
    customCSS: args.customCSS,
    outputPath,
    standalone: true,
    debug: true,
  });
}

// 处理Markdown到DOCX转换
async function handleMarkdownToDocx(args: any) {
  const outputPath = resolveOutputPath(args.outputPath, args.markdownPath, '.docx');
  
  return await convertMarkdownToDocx(args.markdownPath, {
    theme: args.theme ?? 'professional',
    includeTableOfContents: args.includeTableOfContents ?? false,
    preserveStyles: args.preserveStyles !== false,
    outputPath,
    debug: true,
  });
}

// 处理Markdown到PDF转换
async function handleMarkdownToPdf(args: any) {
  const outputPath = resolveOutputPath(args.outputPath, args.markdownPath, '.pdf');
  
  return await convertMarkdownToPdf(args.markdownPath, outputPath, {
    theme: args.theme ?? 'github',
    includeTableOfContents: args.includeTableOfContents ?? false,
    addQrCode: args.addQrCode ?? false,
  });
}

// 处理HTML操作
async function handleHtmlOperations(name: string, args: any) {
  switch (name) {
    case 'convert_html_to_markdown':
      return await handleHtmlToMarkdown(args);
    default:
      throw new Error(`Unknown HTML operation: ${name}`);
  }
}

// 处理HTML到Markdown转换
async function handleHtmlToMarkdown(args: any) {
  // Use static import to prevent lazy loading security risks
  const { EnhancedHtmlToMarkdownConverter } = await import('./tools/enhancedHtmlToMarkdownConverter.js');
  const enhancedConverter = new EnhancedHtmlToMarkdownConverter();
  
  const finalOutputPath = resolveOutputPath(args.outputPath, args.htmlPath, '.md');
  
  return await enhancedConverter.convertHtmlToMarkdown(args.htmlPath, {
    preserveStyles: args.preserveStyles !== false,
    includeCSS: args.includeCSS ?? false,
    debug: args.debug ?? false,
    outputPath: finalOutputPath,
  });
}

// 处理双重解析操作
async function handleDualParsingOperations(name: string, args: any) {
  switch (name) {
    case 'dual_parsing_docx_to_html':
      return await dualParsingDocxToHtml(args);
    case 'dual_parsing_docx_to_pdf':
      return await dualParsingDocxToPdf(args);
    case 'docx_style_analysis':
      return await analyzeDualParsingStyles(args);
    default:
      throw new Error(`Unknown dual parsing operation: ${name}`);
  }
}

// 处理转换规划
async function handleConversionPlanning(args: any) {
  const planner = new ConversionPlanner();
  const conversionRequest = {
    sourceFormat: args.sourceFormat,
    targetFormat: args.targetFormat,
    sourceFile: args.sourceFile,
    requirements: {
      preserveStyles: args.preserveStyles !== false,
      includeImages: args.includeImages !== false,
      theme: args.theme ?? 'github',
      quality: args.quality ?? 'balanced',
    },
  };
  
  return await planner.planConversion(conversionRequest);
}

// 解析输出路径的辅助函数
function resolveOutputPath(outputPath: string | undefined, inputPath: string, extension: string): string {
  if (outputPath) {
    return path.isAbsolute(outputPath)
      ? outputPath
      : path.join(defaultResourcePaths.outputDir, outputPath);
  }
  
  return path.join(
    defaultResourcePaths.outputDir,
    `${path.basename(inputPath, path.extname(inputPath))}${extension}`
  );
}

// Start server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('doc-ops-mcp server running on stdio');
}

main().catch(console.error);
