// doc-ops-mcp - Document Operations MCP Server
// Enhanced DOCX to PDF conversion with perfect Word style replication

const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const fsSync = require('fs');
const os = require('os');
const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} = require('@modelcontextprotocol/sdk/types.js');
const fs = require('fs/promises');
const WordExtractor = require('word-extractor');
const path = require('path');
const cheerio = require('cheerio');

import { WEB_SCRAPING_TOOL, STRUCTURED_DATA_TOOL } from './tools/webScrapingTools';
import { convertDocxToHtmlWithStyles } from './tools/enhancedMammothConfig';
import { convertDocxToHtmlEnhanced } from './tools/enhancedMammothConverter';
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

// 增强的 HTML 到 DOCX 转换函数 - 使用新的增强转换器
async function convertHtmlToDocxEnhanced(
  inputPath: string,
  outputPath?: string,
  options: any = {}
) {
  try {
    // 使用环境变量控制的输出路径
    let finalOutputPath = outputPath;
    if (!finalOutputPath) {
      const baseName = path.basename(inputPath, path.extname(inputPath));
      finalOutputPath = path.join(defaultResourcePaths.outputDir, `${baseName}.docx`);
    } else if (!path.isAbsolute(finalOutputPath)) {
      finalOutputPath = path.join(defaultResourcePaths.outputDir, finalOutputPath);
    }

    console.error(`🚀 开始增强的 HTML 到 DOCX 转换...`);
    console.error(`📄 输入: ${inputPath}`);
    console.error(`📁 输出: ${finalOutputPath}`);
    console.error(`🌍 输出目录由环境变量控制: OUTPUT_DIR=${defaultResourcePaths.outputDir}`);

    // 确保输出目录存在
    await fs.mkdir(path.dirname(finalOutputPath), { recursive: true });

    // 读取 HTML 文件
    const htmlContent = await fs.readFile(inputPath, 'utf-8');

    console.error('📝 HTML 文件读取完成，开始使用增强转换器...');

    // 使用新的增强转换器
    const converter = new EnhancedHtmlToDocxConverter();
    const docxBuffer = await converter.convertHtmlToDocx(htmlContent);

    // 写入文件
    await fs.writeFile(finalOutputPath, docxBuffer);

    console.error(`✅ 增强的 HTML 到 DOCX 转换成功: ${finalOutputPath}`);
    return {
      success: true,
      outputPath: finalOutputPath,
      message: '增强的 HTML 到 DOCX 转换完成',
      metadata: {
        converter: 'EnhancedHtmlToDocxConverter',
        stylesPreserved: true,
        fileSize: docxBuffer.length,
      },
    };
  } catch (error: any) {
    console.error('❌ 增强的 HTML 到 DOCX 转换失败:', error.message);
    return {
      success: false,
      error: error.message,
    };
  }
}

function getDefaultResourcePaths() {
  const homeDir = os.homedir();

  // 从环境变量获取输出目录，如果没有设置则使用默认值
  const outputDir = process.env.OUTPUT_DIR || path.join(homeDir, 'Documents');
  const cacheDir = process.env.CACHE_DIR || path.join(homeDir, '.cache', 'doc-ops-mcp');

  // 从环境变量获取水印和二维码图片路径
  const watermarkImagePath = process.env.WATERMARK_IMAGE;
  const qrCodeImagePath = process.env.QR_CODE_IMAGE;

  try {
    return {
      outputDir: path.resolve(outputDir),
      cacheDir: path.resolve(cacheDir),
      defaultQrCodePath: qrCodeImagePath || null,
      defaultWatermarkPath: watermarkImagePath || null,
      tempDir: os.tmpdir(),
    };
  } catch (error) {
    console.error('Error getting default paths:', error);
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
console.error('Using hybrid mode - browser operations handled by external playwright-mcp');

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

// Read document function
async function readDocument(filePath: string, options: ReadDocumentOptions = {}) {
  try {
    const ext = path.extname(filePath).toLowerCase();

    if (ext === '.docx') {
      if (options.preserveFormatting) {
        console.error('🚀 优先使用增强型 mammoth 转换器进行样式保留转换...');

        // 优先使用修复后的增强型 mammoth 转换器
        try {
          console.error('🔍 尝试使用修复后的增强型 mammoth 转换器...');
          const enhancedResult = await convertDocxToHtmlEnhanced(filePath, {
            preserveImages: options.saveImages !== false,
            enableExperimentalFeatures: true,
            debug: true,
          });

          console.error('📊 修复后的增强型 mammoth 结果:', {
            success: enhancedResult.success,
            hasContent: !!(enhancedResult as any).content,
            contentLength: (enhancedResult as any).content?.length || 0,
            hasStyleTags: (enhancedResult as any).content?.includes('<style>') || false,
            hasInlineStyles: (enhancedResult as any).content?.includes('style=') || false,
            converter: (enhancedResult as any).metadata?.converter || 'unknown',
            stylesCount: (enhancedResult as any).metadata?.stylesCount || 0,
            cssRulesGenerated: (enhancedResult as any).metadata?.cssRulesGenerated || 0,
            error: (enhancedResult as any).error,
          });

          if (enhancedResult.success && (enhancedResult as any).content) {
            console.error('✅ 修复后的增强型 mammoth 转换成功！');

            const content = (enhancedResult as any).content;
            const metadata = (enhancedResult as any).metadata;

            // 验证样式是否真的被注入
            const hasValidStyles =
              content.includes('<style>') &&
              content.includes('Calibri') &&
              content.includes('font-family');

            console.error('🔍 样式验证结果:', {
              hasStyleTags: content.includes('<style>'),
              hasCalibriFont: content.includes('Calibri'),
              hasFontFamily: content.includes('font-family'),
              hasValidStyles: hasValidStyles,
              contentLength: content.length,
              contentPreview: content.substring(0, 500) + '...',
            });

            if (hasValidStyles) {
              console.error('✅ 样式验证通过，返回结果');
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
            } else {
              console.error('⚠️ 样式验证失败，内容缺少必要的样式信息');
            }
          } else {
            console.error(
              '⚠️ 修复后的增强型 mammoth 转换失败，错误:',
              (enhancedResult as any).error
            );
          }
        } catch (enhancedError: any) {
          console.error('❌ 修复后的增强型 mammoth 异常:', enhancedError.message);
          console.error('📋 错误堆栈:', enhancedError.stack);
        }

        console.error('🔄 回退到基础增强配置...');

        // 回退方案：使用基础的 mammoth 配置
        console.error('🔄 使用基础 mammoth 配置作为回退方案...');
        const result = await convertDocxToHtmlWithStyles(filePath, {
          saveImagesToFiles: true,
          imageOutputDir: options.imageOutputDir || path.join(process.cwd(), 'output', 'images'),
        });

        console.error('🔍 回退转换结果分析:', {
          success: result.success,
          contentLength: result.content?.length || 0,
          hasStyle: result.content?.includes('<style>') || false,
          hasDoctype: result.content?.includes('<!DOCTYPE') || false,
          contentPreview: result.content?.substring(0, 300) + '...' || 'No content',
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
      } else {
        // 使用 mammoth 进行纯文本提取
        let result: { value: string };

        try {
          const mammoth = require('mammoth');
          result = await mammoth.extractRawText({ path: filePath });
        } catch (mammothError) {
          console.error('⚠️ mammoth 文本提取失败:', mammothError);
          throw mammothError;
        }

        return {
          success: true,
          content: result.value,
          metadata: { format: 'text', originalFormat: 'docx', converter: 'mammoth' },
        };
      }
    } else if (ext === '.doc') {
      const extractor = new WordExtractor();
      const extracted = await extractor.extract(filePath);
      return {
        success: true,
        content: extracted.getBody(),
        metadata: { format: 'text', originalFormat: 'doc' },
      };
    } else if (ext === '.md' || ext === '.markdown') {
      if (options.preserveFormatting) {
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
            hasContent: !!result.content,
            contentLength: result.content?.length || 0,
            hasStyleTags: result.content?.includes('<style>') || false,
            theme: result.metadata?.theme || 'unknown',
            headingsCount: result.metadata?.headingsCount || 0,
            error: result.error,
          });

          if (result.success && result.content) {
            console.error('✅ Markdown 转换成功！');
            return {
              success: true,
              content: result.content,
              metadata: {
                ...result.metadata,
                format: 'html',
                originalFormat: 'markdown',
                stylesPreserved: true,
                standalone: true,
              },
            };
          } else {
            console.error('⚠️ Markdown 转换失败，回退到纯文本模式');
          }
        } catch (markdownError: any) {
          console.error('❌ Markdown 转换异常:', markdownError.message);
        }
      }

      // 回退到纯文本读取
      const content = await fs.readFile(filePath, 'utf-8');
      return {
        success: true,
        content: content,
        metadata: { format: 'text', originalFormat: 'markdown' },
      };
    } else {
      const content = await fs.readFile(filePath, 'utf-8');
      return {
        success: true,
        content: content,
        metadata: { format: 'text', originalFormat: ext.slice(1) },
      };
    }
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

// Write document function
async function writeDocument(
  content: string,
  outputPath?: string,
  options: WriteDocumentOptions = {}
) {
  try {
    // 如果没有指定输出路径，使用环境变量控制的输出目录
    let finalPath = outputPath;
    if (!finalPath) {
      const outputDir = defaultResourcePaths.outputDir;
      finalPath = path.join(outputDir, `output_${Date.now()}.txt`);
    } else if (!path.isAbsolute(finalPath)) {
      // 如果是相对路径，基于环境变量的输出目录
      finalPath = path.join(defaultResourcePaths.outputDir, finalPath);
    }

    const encoding = options.encoding || 'utf-8';

    await fs.mkdir(path.dirname(finalPath), { recursive: true });
    await fs.writeFile(finalPath, content, encoding);

    return {
      success: true,
      outputPath: finalPath,
      message: `Document written successfully to ${finalPath}. Output directory controlled by OUTPUT_DIR environment variable: ${defaultResourcePaths.outputDir}`,
    };
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

// Convert document function
async function convertDocument(
  inputPath: string,
  outputPath?: string,
  options: ConvertDocumentOptions = {}
) {
  try {
    const inputExt = path.extname(inputPath).toLowerCase();
    const outputExt = outputPath ? path.extname(outputPath).toLowerCase() : '.html';

    // 使用环境变量控制的输出路径
    let finalOutputPath = outputPath;
    if (!finalOutputPath) {
      const baseName = path.basename(inputPath, inputExt);
      finalOutputPath = path.join(defaultResourcePaths.outputDir, `${baseName}${outputExt}`);
    } else if (!path.isAbsolute(finalOutputPath)) {
      finalOutputPath = path.join(defaultResourcePaths.outputDir, finalOutputPath);
    }

    // 特殊处理：HTML 转 Markdown
    if (inputExt === '.html' && outputExt === '.md') {
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
          contentLength: result.content?.length || 0,
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
          throw new Error(result.error || 'HTML 转 Markdown 失败');
        }
      } catch (conversionError: any) {
        console.error('❌ HTML 转 Markdown 转换失败:', conversionError.message);
        return {
          success: false,
          error: `HTML 转 Markdown 失败: ${conversionError.message}`,
        };
      }
    }

    // 特殊处理：HTML 转 DOCX
    if (inputExt === '.html' && outputExt === '.docx') {
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
          throw new Error(result.error || 'HTML 转 DOCX 失败');
        }
      } catch (conversionError: any) {
        console.error('❌ HTML 转 DOCX 转换失败:', conversionError.message);
        return {
          success: false,
          error: `HTML 转 DOCX 失败: ${conversionError.message}`,
        };
      }
    }

    // 特殊处理：DOCX 转 Markdown
    if (inputExt === '.docx' && outputExt === '.md') {
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
          throw new Error(result.error || 'DOCX 转 Markdown 失败');
        }
      } catch (conversionError: any) {
        console.error('❌ DOCX 转 Markdown 转换失败:', conversionError.message);
        return {
          success: false,
          error: `DOCX 转 Markdown 失败: ${conversionError.message}`,
        };
      }
    }

    // 特殊处理：DOCX 转 HTML
    if (inputExt === '.docx' && outputExt === '.html') {
      console.error('🔄 检测到 DOCX 转 HTML 转换...');
      console.error(`📄 输入: ${inputPath}`);
      console.error(`📁 输出: ${finalOutputPath}`);

      try {
        const result = await convertDocxToHtmlEnhanced(inputPath, {
          preserveImages: true,
          enableExperimentalFeatures: true,
          debug: true,
        });

        console.error('📊 DOCX 转 HTML 结果:', {
          success: result.success,
          hasContent: !!(result as any).content,
          contentLength: (result as any).content?.length || 0,
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
          throw new Error((result as any).error || 'DOCX 转 HTML 失败');
        }
      } catch (conversionError: any) {
        console.error('❌ DOCX 转 HTML 转换失败:', conversionError.message);
        return {
          success: false,
          error: `DOCX 转 HTML 失败: ${conversionError.message}`,
        };
      }
    }

    // 其他格式转换：使用原有逻辑
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
      for (const replacement of options.textReplacements) {
        if (replacement.useRegex) {
          const flags = replacement.preserveCase ? 'g' : 'gi';
          const regex = new RegExp(replacement.oldText, flags);
          content = content.replace(regex, replacement.newText);
        } else {
          const searchValue = replacement.preserveCase
            ? replacement.oldText
            : new RegExp(replacement.oldText, 'gi');
          content = content.replace(searchValue, replacement.newText);
        }
      }
    }

    // Write the converted content
    const writeResult = await writeDocument(content, finalOutputPath);
    return writeResult;
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}

// 重构的 DOCX to PDF 转换 - 使用优化的转换器
async function convertDocxToPdf(inputPath: string, outputPath?: string, options: any = {}) {
  const startTime = Date.now();

  // 使用环境变量控制的输出路径
  let finalOutputPath = outputPath;
  if (!finalOutputPath) {
    const baseName = path.basename(inputPath, path.extname(inputPath));
    finalOutputPath = path.join(defaultResourcePaths.outputDir, `${baseName}.pdf`);
  } else if (!path.isAbsolute(finalOutputPath)) {
    finalOutputPath = path.join(defaultResourcePaths.outputDir, finalOutputPath);
  }

  console.error(`🚀 开始优化的 DOCX 到 PDF 转换...`);
  console.error(`📄 输入文件: ${inputPath}`);
  console.error(`📁 输出路径: ${finalOutputPath}`);
  console.error(`🌍 输出目录由环境变量控制: OUTPUT_DIR=${defaultResourcePaths.outputDir}`);

  // 检查水印和二维码配置
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
        format: options.pdfOptions?.format || 'A4',
        margin: options.pdfOptions?.margin || '1in',
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
        // 构建MCP命令，包含水印和二维码处理
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
          postProcessing: {
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
          },
          stats: {
            conversionTime: result.details.conversionTime,
            stylesPreserved: result.details.stylesPreserved,
            imagesPreserved: result.details.imagesPreserved,
          },
        };
      } else {
        // 直接转换成功，需要添加水印和二维码
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
          console.error(`🎨 自动添加水印...`);
          try {
            // @ts-ignore
            const watermarkResult = await addWatermark(result.outputPath, {
              watermarkImage: defaultResourcePaths.defaultWatermarkPath,
              watermarkImageScale: 1.0, // fullscreen模式下会自动计算尺寸
              watermarkImageOpacity: 0.15, // 降低透明度，避免影响阅读
              watermarkImagePosition: 'fullscreen',
            });

            if (watermarkResult.success) {
              console.error(`✅ 水印添加成功`);
              (finalResult as any).watermarkAdded = true;
            } else {
              console.error(`⚠️ 水印添加失败: ${watermarkResult.error}`);
              (finalResult as any).watermarkError = watermarkResult.error;
            }
          } catch (watermarkError: any) {
            console.error(`❌ 水印添加异常: ${watermarkError.message}`);
            (finalResult as any).watermarkError = watermarkError.message;
          }
        }

        // 添加二维码（仅在用户明确要求时）
        //@ts-ignore
        if (hasQRCode && defaultQrCodePath) {
          console.error(`📱 添加二维码...`);
          try {
            // @ts-ignore
            const qrResult = await addQRCode(result.outputPath, defaultQrCodePath, {
              qrScale: 0.15,
              qrOpacity: 1.0,
              qrPosition: 'bottom-center',
              addText: true,
            });

            if (qrResult.success) {
              console.error(`✅ 二维码添加成功`);
              (finalResult as any).qrCodeAdded = true;
            } else {
              console.error(`⚠️ 二维码添加失败: ${qrResult.error}`);
              (finalResult as any).qrCodeError = qrResult.error;
            }
          } catch (qrError: any) {
            console.error(`❌ 二维码添加异常: ${qrError.message}`);
            (finalResult as any).qrCodeError = qrError.message;
          }
        }

        return finalResult;
      }
    } else {
      throw new Error(result.error || '优化转换器转换失败');
    }
  } catch (error: any) {
    console.error('❌ 优化转换失败，尝试回退方案:', error.message);

    // 回退到原有的转换逻辑（简化版）
    return await fallbackConvertDocxToPdf(inputPath, finalOutputPath, options);
  }
}

// 回退转换方法（简化的原有逻辑）
async function fallbackConvertDocxToPdf(inputPath: string, outputPath?: string, options: any = {}) {
  const docxPath = inputPath;

  // 使用环境变量控制的输出路径
  let finalOutputPath = outputPath;
  if (!finalOutputPath) {
    const baseName = path.basename(inputPath, path.extname(inputPath));
    finalOutputPath = path.join(defaultResourcePaths.outputDir, `${baseName}.pdf`);
  } else if (!path.isAbsolute(finalOutputPath)) {
    finalOutputPath = path.join(defaultResourcePaths.outputDir, finalOutputPath);
  }

  console.error(`🔄 使用回退方案进行 DOCX 到 PDF 转换...`);
  console.error(`📄 输入: ${docxPath}`);
  console.error(`📁 输出: ${finalOutputPath}`);
  console.error(`🌍 输出目录由环境变量控制: OUTPUT_DIR=${defaultResourcePaths.outputDir}`);

  // Ensure output directory exists
  await fs.mkdir(path.dirname(finalOutputPath), { recursive: true });

  let perfectWordHtml = '';
  let conversionSuccess = false;
  let tempHtmlPath = '';

  // 尝试使用双重解析引擎
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
      perfectWordHtml = dualResult.content;
      conversionSuccess = true;
      console.error('✅ 双重解析引擎转换成功！');
    }
  } catch (dualError: any) {
    console.error('❌ 双重解析引擎失败:', dualError.message);
  }

  // 如果双重解析失败，使用增强型 mammoth
  if (!conversionSuccess) {
    try {
      console.error('🔄 使用增强型 mammoth 转换器...');
      const enhancedResult = await convertDocxToHtmlEnhanced(docxPath, {
        preserveImages: true,
        enableExperimentalFeatures: true,
        debug: true,
      });

      if (enhancedResult.success && (enhancedResult as any).content) {
        perfectWordHtml = (enhancedResult as any).content;
        conversionSuccess = true;
        console.error('✅ 增强型 mammoth 转换成功！');
      }
    } catch (enhancedError: any) {
      console.error('❌ 增强型 mammoth 失败:', enhancedError.message);
    }
  }

  // 最终回退
  if (!conversionSuccess) {
    console.error('🔄 使用最终回退方案...');
    const mammoth = require('mammoth');
    const basicResult = await mammoth.convertToHtml({ path: docxPath });
    perfectWordHtml = createPerfectWordHtml(basicResult.value, options);
    conversionSuccess = true;
  }

  console.error(`🎨 HTML 生成完成 (长度: ${perfectWordHtml.length})`);

  // 确保样式完整性
  let finalHtml = perfectWordHtml;

  // 确保包含完整的Word样式
  if (!finalHtml.includes('<style>')) {
    console.error('⚠️ 检测到样式缺失，强制注入Word样式...');

    // 提取现有内容并包装完整的HTML结构
    const contentWithoutWrapper = finalHtml
      .replace(/<!DOCTYPE[^>]*>/gi, '')
      .replace(/<html[^>]*>/gi, '')
      .replace(/<\/html>/gi, '')
      .replace(/<head[^>]*>.*?<\/head>/gi, '')
      .replace(/<body[^>]*>/gi, '')
      .replace(/<\/body>/gi, '');

    finalHtml = createPerfectWordHtml(contentWithoutWrapper, options);
  }

  // 确保DOCTYPE和完整的HTML结构
  if (!finalHtml.includes('<!DOCTYPE')) {
    finalHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    ${finalHtml.includes('<style>') ? '' : createPerfectWordHtml('', options).match(/<style[^>]*>[\s\S]*?<\/style>/gi)?.[0] || ''}
</head>
<body>
${finalHtml
  .replace(/<!DOCTYPE[^>]*>/gi, '')
  .replace(/<html[^>]*>/gi, '')
  .replace(/<\/html>/gi, '')
  .replace(/<head[^>]*>.*?<\/head>/gi, '')
  .replace(/<body[^>]*>/gi, '')
  .replace(/<\/body>/gi, '')}
</body>
</html>`;
  }

  // 强制确保所有样式都被内联到HTML中，防止样式丢失
  // 提取所有样式标签内容
  const styleTagsContent: string[] = [];
  let styleTagMatch;
  const styleTagRegex = /<style[^>]*>([\s\S]*?)<\/style>/gi;
  while ((styleTagMatch = styleTagRegex.exec(finalHtml)) !== null) {
    const styleContent = styleTagMatch[1].trim();
    if (styleContent) {
      styleTagsContent.push(styleContent);
    }
  }

  // 如果没有找到有效的样式内容，添加基本的Word样式
  if (styleTagsContent.length === 0) {
    console.error('⚠️ 未找到有效样式内容，添加基本Word样式');
    // 从createPerfectWordHtml提取样式
    const perfectWordStyles =
      createPerfectWordHtml('', options).match(/<style[^>]*>([\s\S]*?)<\/style>/i)?.[1] || '';
    if (perfectWordStyles.trim()) {
      styleTagsContent.push(perfectWordStyles.trim());
    } else {
      // 添加最小的基本样式
      const basicWordStyles = `
          body { font-family: "Calibri", "Microsoft YaHei", "SimSun", sans-serif !important; }
          p { margin-bottom: 8pt !important; line-height: 1.08 !important; }
          table { border-collapse: collapse !important; }
          td, th { border: 1px solid #000 !important; padding: 5pt !important; }
          * { -webkit-print-color-adjust: exact !important; color-adjust: exact !important; print-color-adjust: exact !important; }
        `;
      styleTagsContent.push(basicWordStyles);
    }
  }

  // 合并所有样式
  let combinedStyles = styleTagsContent.join('\n');

  // 确保所有CSS规则都有!important
  combinedStyles = combinedStyles.replace(/([^;{}:]+:[^;{}!]+)(?=;|})/g, match => {
    if (!match.includes('!important')) {
      return match + ' !important';
    }
    return match;
  });

  // 移除所有现有样式标签
  finalHtml = finalHtml.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

  // 在head中添加合并后的样式
  if (finalHtml.includes('</head>')) {
    finalHtml = finalHtml.replace(
      /<\/head>/i,
      `<style type="text/css">\n${combinedStyles}\n</style>\n</head>`
    );
  } else if (finalHtml.includes('<head')) {
    // 在已有的head中添加样式
    finalHtml = finalHtml.replace(
      /<head[^>]*>/i,
      `$&\n<style type="text/css">\n${combinedStyles}\n</style>`
    );
  } else if (finalHtml.includes('<html')) {
    // 如果没有head标签，添加head和样式
    finalHtml = finalHtml.replace(
      /<html[^>]*>/i,
      `$&\n<head>\n<style type="text/css">\n${combinedStyles}\n</style>\n</head>`
    );
  } else {
    // 如果连html标签都没有，创建完整的HTML结构
    const bodyContent = finalHtml;
    finalHtml = `<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="UTF-8">\n<style type="text/css">\n${combinedStyles}\n</style>\n</head>\n<body>\n${bodyContent}\n</body>\n</html>`;
  }

  // 确保内联样式属性也被保留
  finalHtml = finalHtml.replace(/<([a-z][a-z0-9]*)([^>]*?)>/gi, (match, tag, attrs) => {
    // 保留原有的style属性，并添加!important
    if (attrs.includes('style=')) {
      return match.replace(/style=(["'])(.*?)\1/gi, (styleMatch, quote, styleContent) => {
        // 为每个样式属性添加!important（如果尚未添加）
        const enhancedStyle = styleContent.replace(/([^;]+)(?=;|$)/g, prop => {
          if (!prop.includes('!important')) {
            return `${prop} !important`;
          }
          return prop;
        });
        return `style=${quote}${enhancedStyle}${quote}`;
      });
    }
    return match;
  });

  // 使用cheerio处理HTML，确保所有样式都被内联到HTML中
  let $ = cheerio.load(finalHtml, { decodeEntities: false });

  // 提取所有样式标签内容
  let allStyles = '';
  $('style').each((i, el) => {
    allStyles += $(el).html() + '\n';
  });

  // 确保所有样式规则都有!important标记
  allStyles = allStyles.replace(/([^;{}]+)(?=;|})/g, match => {
    if (!match.includes('!important')) {
      return match + ' !important';
    }
    return match;
  });

  // 移除所有现有样式标签
  $('style').remove();

  // 在<head>中添加合并后的样式
  if (!$('head').length) {
    $('html').prepend('<head></head>');
  }
  $('head').append(`<style type="text/css">${allStyles}</style>`);

  // 为内联样式属性添加!important
  $('[style]').each((i, el) => {
    const style = $(el).attr('style');
    if (style) {
      const importantStyle = style
        .split(';')
        .map(rule => {
          rule = rule.trim();
          if (rule && !rule.includes('!important')) {
            return rule + ' !important';
          }
          return rule;
        })
        .join(';');
      $(el).attr('style', importantStyle);
    }
  });

  // 更新最终HTML
  finalHtml = $.html();

  // 创建临时HTML文件并强制写入完整内容
  if (!tempHtmlPath) {
    tempHtmlPath = path.join(os.tmpdir(), `docx-conversion-${Date.now()}.html`);
    await fs.writeFile(tempHtmlPath, finalHtml, 'utf8');
    console.error(`📝 样式修复后的HTML文件已创建: ${tempHtmlPath}`);
  } else if (options.useExistingHtml && options.existingHtmlPath) {
    console.error(`📝 使用已存在的HTML文件: ${tempHtmlPath}`);
  }

  // 验证最终文件内容
  const writtenContent = await fs.readFile(tempHtmlPath, 'utf8');
  // 使用正则表达式检查样式标签，而不是简单的字符串匹配
  const styleTagsMatch = writtenContent.match(/<style[^>]*>[\s\S]*?<\/style>/gi) || [];
  const writtenHasStyle = styleTagsMatch.length > 0;
  const writtenHasDoctype = writtenContent.includes('<!DOCTYPE');
  const hasWordStyles =
    writtenContent.includes('Microsoft YaHei') || writtenContent.includes('Calibri');

  // 检查样式标签内是否有实际内容
  const hasStyleContent = styleTagsMatch.some(tag => {
    const styleContent = tag.replace(/<style[^>]*>/i, '').replace(/<\/style>/i, '');
    return styleContent.trim().length > 0;
  });

  // 检查是否有基本的CSS规则
  const hasCssRules = styleTagsMatch.some(tag => {
    const styleContent = tag.replace(/<style[^>]*>/i, '').replace(/<\/style>/i, '');
    return styleContent.includes('{') && styleContent.includes('}');
  });

  console.error('🔍 样式修复验证:', {
    filePath: tempHtmlPath,
    fileSize: writtenContent.length,
    hasStyle: writtenHasStyle,
    hasStyleContent: hasStyleContent,
    hasCssRules: hasCssRules,
    hasDoctype: writtenHasDoctype,
    hasWordStyles: hasWordStyles,
    styleTagCount: styleTagsMatch.length,
    contentPreview: writtenContent.substring(0, 500) + '...',
  });

  if (!writtenHasStyle || !hasStyleContent || !hasCssRules || !hasWordStyles) {
    console.error('❌ 样式修复失败，手动注入完整Word样式...');

    // 强制注入完整的Word样式
    // 从createPerfectWordHtml获取完整的样式定义
    const perfectHtml = createPerfectWordHtml('', options);
    const wordStyles = perfectHtml.match(/<style[^>]*>[\s\S]*?<\/style>/gi)?.[0] || '';

    // 提取文档内容，保留所有HTML结构但移除现有样式
    let bodyContent = writtenContent;

    // 如果文档已经有HTML结构，只提取body内容
    if (bodyContent.includes('<body') && bodyContent.includes('</body>')) {
      const bodyMatch = bodyContent.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
      if (bodyMatch && bodyMatch[1]) {
        bodyContent = bodyMatch[1];
      }
    } else if (bodyContent.includes('<html') || bodyContent.includes('<!DOCTYPE')) {
      // 如果有HTML标签但没有完整的body标签，移除HTML结构保留内容
      bodyContent = bodyContent
        .replace(/<!DOCTYPE[^>]*>/gi, '')
        .replace(/<html[^>]*>/gi, '')
        .replace(/<\/html>/gi, '')
        .replace(/<head[^>]*>[\s\S]*?<\/head>/gi, '')
        .replace(/<body[^>]*>/gi, '')
        .replace(/<\/body>/gi, '');
    }

    // 移除所有现有样式标签
    bodyContent = bodyContent.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

    // 创建强制注入样式的完整HTML
    const enforcedHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    ${wordStyles}
</head>
<body>
${bodyContent}
</body>
</html>`;

    await fs.writeFile(tempHtmlPath, enforcedHtml, 'utf8');
    console.error('✅ 样式强制注入完成');
  }

  // Step 5: Return explicit Playwright-MCP call instructions
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
    playwrightInstructions: [
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
              // 确保所有样式都被正确应用
              function ensureStylesApplied() {
                // 检查样式是否存在
                const styles = document.querySelectorAll('style');
                if (styles.length === 0) {
                  console.error('警告: 未找到样式标签');
                  
                  // 创建基本样式
                  const style = document.createElement('style');
                  style.textContent = 'body { font-family: "Calibri", "Microsoft YaHei", sans-serif !important; } * { -webkit-print-color-adjust: exact !important; }';
                  document.head.appendChild(style);
                }
                
                // 确保所有样式规则都有!important
                styles.forEach(style => {
                  if (style.sheet) {
                    console.log(\`样式表已加载，规则数: \${style.sheet.cssRules.length}\`);
                  }
                });
              }
              
              // 执行样式检查
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
            // 确保所有样式都被应用
            document.querySelectorAll('style').forEach(style => {
              if (style.sheet) {
                console.log(`样式表已加载，规则数: ${style.sheet.cssRules.length}`);
              }
            });
            // 强制重新计算样式
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
    ],
    outputPath: finalOutputPath,
    tempHtmlPath: tempHtmlPath,
    cleanupRequired: !options.useExistingHtml, // 如果使用的是已存在的HTML文件，则不需要清理
  };
}

// 终极版Word样式HTML生成器 - 确保100%样式还原
function createPerfectWordHtml(content: string, options: any = {}): string {
  const chineseFont = options.chineseFont || 'Microsoft YaHei';

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
  ${content}
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
        const scale = options.watermarkImageScale || 0.25;
        const opacity = options.watermarkImageOpacity || 0.6;

        let x = 0,
          y = 0;
        const position = options.watermarkImagePosition || 'top-right';

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
        const fontSize = options.watermarkFontSize || 48;
        const opacity = options.watermarkTextOpacity || 0.3;

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

// QR Code function
async function addQRCode(pdfPath: string, qrCodePath?: string, options: QRCodeOptions = {}) {
  try {
    const finalQrPath = qrCodePath || defaultResourcePaths.defaultQrCodePath;

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
      const scale = options.qrScale || 0.15;
      const opacity = options.qrOpacity || 1.0;
      const position = options.qrPosition || 'bottom-center';

      let x = 0,
        y = 0;

      switch (position) {
        case 'top-left':
          x = 20;
          y = height - qrImage.height * scale - 20;
          break;
        case 'top-right':
          x = width - qrImage.width * scale - 20;
          y = height - qrImage.height * scale - 20;
          break;
        case 'top-center':
          x = (width - qrImage.width * scale) / 2;
          y = height - qrImage.height * scale - 20;
          break;
        case 'bottom-left':
          x = 20;
          y = 20;
          break;
        case 'bottom-right':
          x = width - qrImage.width * scale - 20;
          y = 20;
          break;
        case 'bottom-center':
          x = (width - qrImage.width * scale) / 2;
          y = 20;
          break;
        case 'center':
          x = (width - qrImage.width * scale) / 2;
          y = (height - qrImage.height * scale) / 2;
          break;
      }

      page.drawImage(qrImage, {
        x,
        y,
        width: qrImage.width * scale,
        height: qrImage.height * scale,
        opacity,
      });

      if (options.addText !== false) {
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const text = options.customText || 'Scan QR code for more information';
        const textSize = options.textSize || 8;
        const textColor = options.textColor || '#000000';

        page.drawText(text, {
          x: x + (qrImage.width * scale - text.length * textSize * 0.6) / 2,
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

// Unified PDF post-processing function
async function processPdfPostConversion(
  playwrightPdfPath: string,
  targetPath?: string,
  options: any = {}
) {
  try {
    // Convert playwright-mcp temp path to target path
    const outputDir = process.env.OUTPUT_DIR || path.dirname(playwrightPdfPath);
    let finalPath = targetPath;

    if (!finalPath) {
      // Extract original filename from playwright path
      const playwrightFileName = path.basename(playwrightPdfPath);
      // Remove timestamp prefix and decode the filename
      const decodedName = playwrightFileName
        .replace(/^\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}\.\d{3}Z-/, '')
        .replace(/-/g, '/');
      finalPath = path.join(outputDir, path.basename(decodedName));
    } else if (!path.isAbsolute(finalPath)) {
      finalPath = path.join(outputDir, finalPath);
    }

    // Ensure output directory exists
    await fs.mkdir(path.dirname(finalPath), { recursive: true });

    // Copy the PDF from playwright temp location to target location
    await fs.copyFile(playwrightPdfPath, finalPath);

    const results: any[] = [];

    // Add watermark if specified
    if (finalPath && (options.addWatermark || process.env.WATERMARK_IMAGE)) {
      const watermarkOptions = {
        watermarkImage: options.watermarkImage || process.env.WATERMARK_IMAGE,
        watermarkText: options.watermarkText,
        watermarkImageScale: options.watermarkImageScale || 0.25,
        watermarkImageOpacity: options.watermarkImageOpacity || 0.6,
        watermarkImagePosition: options.watermarkImagePosition || 'top-right',
        watermarkFontSize: options.watermarkFontSize || 48,
        watermarkTextOpacity: options.watermarkTextOpacity || 0.3,
      };

      try {
        const watermarkResult = await addWatermark(finalPath, watermarkOptions);
        results.push(watermarkResult);
      } catch (error: any) {
        results.push({ success: false, error: error.message });
      }
    }

    // Add QR code if specified
    if (
      finalPath &&
      (options.addQrCode || (options.addQrCode !== false && process.env.QR_CODE_IMAGE))
    ) {
      const qrOptions = {
        qrScale: options.qrScale || 0.15,
        qrOpacity: options.qrOpacity || 1.0,
        qrPosition: options.qrPosition || 'bottom-center',
        addText: options.addText !== false,
        customText: options.customText || 'Scan QR code for more information',
        textSize: options.textSize || 8,
        textColor: options.textColor || '#000000',
      };

      try {
        const qrCodePath = options.qrCodePath || process.env.QR_CODE_IMAGE;
        if (qrCodePath) {
          const qrResult = await addQRCode(finalPath, qrCodePath, qrOptions);
          results.push(qrResult);
        } else {
          results.push({ success: false, error: 'QR code path not provided' });
        }
      } catch (error: any) {
        results.push({ success: false, error: error.message });
      }
    }

    // Clean up temporary playwright file if it's different from final path
    if (playwrightPdfPath !== finalPath && fsSync.existsSync(playwrightPdfPath)) {
      try {
        await fs.unlink(playwrightPdfPath);
      } catch (error) {
        console.warn(`Failed to clean up temporary file: ${playwrightPdfPath}`);
      }
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

// Markdown 转 PDF 函数
// 注意：此函数实际上是先将 Markdown 转换为 HTML，然后需要使用 playwright-mcp 完成最终的 PDF 转换
// 转换流程：Markdown → HTML → PDF (通过 playwright-mcp)
async function convertMarkdownToPdf(inputPath: string, outputPath?: string, options: any = {}) {
  try {
    // 使用环境变量控制的输出路径
    let finalOutputPath = outputPath;
    if (!finalOutputPath) {
      const baseName = path.basename(inputPath, path.extname(inputPath));
      finalOutputPath = path.join(defaultResourcePaths.outputDir, `${baseName}.pdf`);
    } else if (!path.isAbsolute(finalOutputPath)) {
      finalOutputPath = path.join(defaultResourcePaths.outputDir, finalOutputPath);
    }

    console.error(`🔄 Markdown 到 PDF 转换...`);
    console.error(`📄 输入: ${inputPath}`);
    console.error(`📁 输出: ${finalOutputPath}`);

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
    const defaultWatermarkPath = process.env.WATERMARK_IMAGE || null;
    const defaultQrCodePath = process.env.QR_CODE_IMAGE || null;
    const addQrCode = options.addQrCode || false;

    // 构建 playwright 命令，包含水印和二维码处理
    const playwrightCommands = [
      `browser_navigate("file://${htmlOutputPath}")`,
      `browser_wait_for({ time: 3 })`,
      `browser_pdf_save({ filename: "${finalOutputPath}" })`,
    ];

    // 如果有水印或二维码需要添加，在 playwright 命令后添加处理步骤
    const postProcessingSteps: string[] = [];
    if (defaultWatermarkPath) {
      postProcessingSteps.push(`添加水印: ${defaultWatermarkPath}`);
    }
    if (addQrCode && defaultQrCodePath) {
      postProcessingSteps.push(`添加二维码: ${defaultQrCodePath}`);
    }

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

// DOCX 转 Markdown 函数
async function convertDocxToMarkdown(inputPath: string, outputPath?: string, options: any = {}) {
  try {
    // 使用环境变量控制的输出路径
    let finalOutputPath = outputPath;
    if (!finalOutputPath) {
      const baseName = path.basename(inputPath, path.extname(inputPath));
      finalOutputPath = path.join(defaultResourcePaths.outputDir, `${baseName}.md`);
    } else if (!path.isAbsolute(finalOutputPath)) {
      finalOutputPath = path.join(defaultResourcePaths.outputDir, finalOutputPath);
    }

    console.error(`🔄 DOCX 到 Markdown 转换...`);
    console.error(`📄 输入: ${inputPath}`);
    console.error(`📁 输出: ${finalOutputPath}`);

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

// DOCX 转 TXT 函数
async function convertDocxToTxt(inputPath: string, outputPath?: string, options: any = {}) {
  try {
    // 使用环境变量控制的输出路径
    let finalOutputPath = outputPath;
    if (!finalOutputPath) {
      const baseName = path.basename(inputPath, path.extname(inputPath));
      finalOutputPath = path.join(defaultResourcePaths.outputDir, `${baseName}.txt`);
    } else if (!path.isAbsolute(finalOutputPath)) {
      finalOutputPath = path.join(defaultResourcePaths.outputDir, finalOutputPath);
    }

    console.error(`🔄 DOCX 到 TXT 转换...`);
    console.error(`📄 输入: ${inputPath}`);
    console.error(`📁 输出: ${finalOutputPath}`);

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
      .replace(/!\[([^\]]*)\]\([^\)]+\)/g, '$1') // 移除图片，保留alt文本
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
    switch (name) {
      case 'read_document':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await readDocument(args.filePath, args), null, 2),
            },
          ],
        };

      case 'write_document':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await writeDocument(args.content, args.outputPath, args),
                null,
                2
              ),
            },
          ],
        };

      case 'convert_document':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await convertDocument(args.inputPath, args.outputPath, args),
                null,
                2
              ),
            },
          ],
        };

      // case "add_watermark":  // 移除独立工具处理，避免混淆
      //   return {
      //     content: [
      //       {
      //         type: "text",
      //         text: JSON.stringify(await addWatermark(args.pdfPath, args), null, 2)
      //       }
      //     ]
      //   };

      // case "add_qrcode":     // 移除独立工具处理，避免混淆
      //   return {
      //     content: [
      //       {
      //         type: "text",
      //         text: JSON.stringify(await addQRCode(args.pdfPath, args.qrCodePath, args), null, 2)
      //       }
      //     ]
      //   };

      case 'process_pdf_post_conversion':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await processPdfPostConversion(args.playwrightPdfPath, args.targetPath, args),
                null,
                2
              ),
            },
          ],
        };

      case 'convert_docx_to_pdf':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await convertDocxToPdf(args.docxPath, args.outputPath, args),
                null,
                2
              ),
            },
          ],
        };

      case 'convert_markdown_to_html':
        const markdownResult = await convertMarkdownToHtml(args.markdownPath, {
          theme: args.theme || 'github',
          includeTableOfContents: args.includeTableOfContents || false,
          customCSS: args.customCSS,
          outputPath: args.outputPath
            ? path.isAbsolute(args.outputPath)
              ? args.outputPath
              : path.join(defaultResourcePaths.outputDir, args.outputPath)
            : path.join(
                defaultResourcePaths.outputDir,
                `${path.basename(args.markdownPath, path.extname(args.markdownPath))}.html`
              ),
          standalone: true,
          debug: true,
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(markdownResult, null, 2),
            },
          ],
        };

      case 'convert_markdown_to_docx':
        const docxResult = await convertMarkdownToDocx(args.markdownPath, {
          theme: args.theme || 'professional',
          includeTableOfContents: args.includeTableOfContents || false,
          preserveStyles: args.preserveStyles !== false,
          outputPath: args.outputPath
            ? path.isAbsolute(args.outputPath)
              ? args.outputPath
              : path.join(defaultResourcePaths.outputDir, args.outputPath)
            : path.join(
                defaultResourcePaths.outputDir,
                `${path.basename(args.markdownPath, path.extname(args.markdownPath))}.docx`
              ),
          debug: true,
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(docxResult, null, 2),
            },
          ],
        };

      case 'convert_markdown_to_pdf':
        const pdfResult = await convertMarkdownToPdf(
          args.markdownPath,
          args.outputPath
            ? path.isAbsolute(args.outputPath)
              ? args.outputPath
              : path.join(defaultResourcePaths.outputDir, args.outputPath)
            : path.join(
                defaultResourcePaths.outputDir,
                `${path.basename(args.markdownPath, path.extname(args.markdownPath))}.pdf`
              ),
          {
            theme: args.theme || 'github',
            includeTableOfContents: args.includeTableOfContents || false,
            addQrCode: args.addQrCode || false,
          }
        );
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(pdfResult, null, 2),
            },
          ],
        };

      case 'convert_html_to_markdown':
        // 使用增强的HTML到Markdown转换器
        const {
          EnhancedHtmlToMarkdownConverter,
        } = require('./tools/enhancedHtmlToMarkdownConverter');
        const enhancedConverter = new EnhancedHtmlToMarkdownConverter();

        const finalOutputPath = args.outputPath
          ? path.isAbsolute(args.outputPath)
            ? args.outputPath
            : path.join(defaultResourcePaths.outputDir, args.outputPath)
          : path.join(
              defaultResourcePaths.outputDir,
              `${path.basename(args.htmlPath, path.extname(args.htmlPath))}.md`
            );

        const htmlToMdResult = await enhancedConverter.convertHtmlToMarkdown(args.htmlPath, {
          preserveStyles: args.preserveStyles !== false,
          includeCSS: args.includeCSS || false,
          debug: args.debug || false,
          outputPath: finalOutputPath,
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(htmlToMdResult, null, 2),
            },
          ],
        };

      case 'plan_conversion':
        // 使用转换规划器
        const planner = new ConversionPlanner();
        const conversionRequest = {
          sourceFormat: args.sourceFormat,
          targetFormat: args.targetFormat,
          sourceFile: args.sourceFile,
          requirements: {
            preserveStyles: args.preserveStyles !== false,
            includeImages: args.includeImages !== false,
            theme: args.theme || 'github',
            quality: args.quality || 'balanced',
          },
        };

        const conversionPlan = await planner.planConversion(conversionRequest);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(conversionPlan, null, 2),
            },
          ],
        };

      case 'dual_parsing_docx_to_html':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await dualParsingDocxToHtml(args), null, 2),
            },
          ],
        };

      case 'dual_parsing_docx_to_pdf':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await dualParsingDocxToPdf(args), null, 2),
            },
          ],
        };

      case 'docx_style_analysis':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await analyzeDualParsingStyles(args), null, 2),
            },
          ],
        };

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
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

// Start server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('doc-ops-mcp server running on stdio');
}

main().catch(console.error);
