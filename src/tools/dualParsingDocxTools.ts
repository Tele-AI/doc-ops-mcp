/**
 * Dual Parsing Engine DOCX Tools
 * Integrate dual parsing engine into MCP tool system
 */

import { DualParsingEngine, DualParsingOptions, DualParsingResult } from './dualParsingEngine';
const fs = require('fs/promises');
const path = require('path');
const { StyleExtractor } = require('./styleExtractor');

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

// Dual parsing DOCX to HTML tool definition
export const DUAL_PARSING_DOCX_TO_HTML_TOOL = {
  name: 'dual_parsing_docx_to_html',
  description: 'Convert DOCX files to HTML using dual parsing engine, perfectly preserving styles and formatting',
  inputSchema: {
    type: 'object',
    properties: {
      docxPath: {
        type: 'string',
        description: 'DOCX file path',
      },
      outputPath: {
        type: 'string',
        description: 'Output HTML file path (optional)',
      },
      cssOutputPath: {
        type: 'string',
        description: 'Output CSS file path (optional)',
      },
      options: {
        type: 'object',
        description: 'Conversion options',
        properties: {
          extractStyles: {
            type: 'boolean',
            description: 'Whether to extract style information',
            default: true,
          },
          preserveImages: {
            type: 'boolean',
            description: 'Whether to preserve images',
            default: true,
          },

          generateCompleteHTML: {
            type: 'boolean',
            description: 'Whether to generate complete HTML document',
            default: true,
          },
          inlineCSS: {
            type: 'boolean',
            description: 'Whether to inline CSS',
            default: false,
          },
          optimizeStructure: {
            type: 'boolean',
            description: 'Whether to optimize HTML structure',
            default: true,
          },
          addAccessibility: {
            type: 'boolean',
            description: 'Whether to add accessibility support',
            default: true,
          },
          logProgress: {
            type: 'boolean',
            description: 'Whether to show conversion progress',
            default: true,
          },
        },
      },
    },
    required: ['docxPath'],
  },
};

// Dual parsing DOCX to PDF tool definition
export const DUAL_PARSING_DOCX_TO_PDF_TOOL = {
  name: 'dual_parsing_docx_to_pdf',
  description:
    'Convert DOCX files to PDF using dual parsing engine, perfectly preserving styles and formatting. Note: This tool requires playwright-mcp to complete the final PDF generation step.',
  inputSchema: {
    type: 'object',
    properties: {
      docxPath: {
        type: 'string',
        description: 'DOCX file path',
      },
      outputPath: {
        type: 'string',
        description: 'Output PDF file path (optional)',
      },
      options: {
        type: 'object',
        description: 'Conversion options',
        properties: {
          extractStyles: {
            type: 'boolean',
            description: 'Whether to extract style information',
            default: true,
          },
          preserveImages: {
            type: 'boolean',
            description: 'Whether to preserve images',
            default: true,
          },
          pdfOptions: {
            type: 'object',
            description: 'PDF generation options',
            properties: {
              format: {
                type: 'string',
                description: 'Page format',
                enum: ['A4', 'A3', 'A5', 'Letter', 'Legal'],
                default: 'A4',
              },
              margin: {
                type: 'string',
                description: 'Page margins',
                default: '1in',
              },
              printBackground: {
                type: 'boolean',
                description: 'Whether to print background',
                default: true,
              },
              landscape: {
                type: 'boolean',
                description: 'Whether to use landscape orientation',
                default: false,
              },
            },
          },
          logProgress: {
            type: 'boolean',
            description: 'Whether to show conversion progress',
            default: true,
          },
        },
      },
    },
    required: ['docxPath'],
  },
};

// Style analysis tool definition
export const DOCX_STYLE_ANALYSIS_TOOL = {
  name: 'docx_style_analysis',
  description: 'Analyze style information of DOCX files and provide detailed style reports',
  inputSchema: {
    type: 'object',
    properties: {
      docxPath: {
        type: 'string',
        description: 'DOCX file path',
      },
      outputPath: {
        type: 'string',
        description: 'Style report output path (optional)',
      },
      detailed: {
        type: 'boolean',
        description: 'Whether to generate detailed report',
        default: false,
      },
    },
    required: ['docxPath'],
  },
};

/**
 * Dual parsing DOCX to HTML (optimized version)
 */
export async function dualParsingDocxToHtml(args: any): Promise<any> {
  try {
    const { docxPath, outputPath, cssOutputPath, options = {} } = args;

    console.log('🚀 Starting dual parsing DOCX to HTML (optimized version)...');
    console.log('📁 Input file:', docxPath);

    // Validate and secure input paths
    const validatedDocxPath = validatePath(docxPath);
    const validatedOutputPath = outputPath ? validatePath(outputPath) : null;
    const validatedCssOutputPath = cssOutputPath ? validatePath(cssOutputPath) : null;

    // Validate input file
    if (
      !(await fs
        .access(validatedDocxPath)
        .then(() => true)
        .catch(() => false))
    ) {
      throw new Error(`Input file does not exist: ${validatedDocxPath}`);
    }

    // Create image output directory - 使用安全的路径处理
    const { safePathJoin } = require('../security/securityConfig');
    const imageOutputDir = safePathJoin(path.dirname(validatedDocxPath), 'images');

    // Build optimized conversion options
    const dualParsingOptions: DualParsingOptions = {
      extractStyles: options.extractStyles !== false,
      mammothOptions: {
        preserveImages: options.preserveImages !== false,
        convertImagesToBase64: false,
        includeDefaultStyles: true,
        transformDocument: true,
        customStyleMappings: [],
        imageOutputDir: imageOutputDir,
      },
      cssOptions: {
        minify: false,
        includeComments: true,
        generateResponsive: options.optimizeStructure !== false,
        generatePrintStyles: true,
        customPrefix: '',
      },
      postProcessOptions: {
        injectStyles: true,
        optimizeStructure: options.optimizeStructure !== false,
        fixFormatting: true,
        addSemanticTags: true,
        removeEmptyElements: true,
        mergeAdjacentElements: true,
        addAccessibility: options.addAccessibility !== false,
      },
      outputOptions: {
        includeCSS: true,
        inlineCSS: options.inlineCSS === true,
        generateCompleteHTML: options.generateCompleteHTML !== false,
        preserveOriginalStructure: true,
        addDocumentMetadata: true,
      },
      mediaOptions: {
        extractToFiles: true,
        outputDirectory: imageOutputDir,
        convertToBase64: false,
        optimizeImages: false,
        preserveOriginalSize: true,
      },
      debugOptions: {
        logProgress: options.logProgress !== false,
        saveIntermediateResults: false,
      },
    };

    // Create optimized dual parsing engine
    const engine = new DualParsingEngine(dualParsingOptions);

    // Execute optimized conversion
    const result: DualParsingResult = await engine.convertDocxToHtml(validatedDocxPath);

    if (!result.success) {
      throw new Error(result.error || 'Conversion failed');
    }

    console.log('✅ Dual parsing conversion completed!');
    console.log('📊 Conversion statistics:', {
      'Total styles': result.details.styleExtraction.totalStyles,
      'Paragraph styles': result.details.styleExtraction.paragraphStyles,
      'Character styles': result.details.styleExtraction.characterStyles,
      'Table styles': result.details.styleExtraction.tableStyles,
      'Media files': result.details.mediaExtraction.totalFiles,
      'CSS rules': result.details.cssGeneration.totalRules,
      'Conversion time': `${result.performance.totalTime}ms`,
      'Styles injected': result.details.htmlProcessing.stylesInjected,
    });

    // Generate output paths
    const htmlOutputPath = validatedOutputPath || validatedDocxPath.replace(/\.docx$/i, '_converted.html');
    const cssPath = validatedCssOutputPath || validatedDocxPath.replace(/\.docx$/i, '_styles.css');

    // Save HTML file
    const htmlToSave = result.completeHTML || result.html;
    await fs.writeFile(htmlOutputPath, htmlToSave, 'utf8');
    console.log('💾 HTML file saved:', htmlOutputPath);

    // Display image save information
    if (result.mediaFiles.length > 0) {
      console.log(`🖼️ Image files saved to: ${imageOutputDir}`);
      console.log(`📸 Total ${result.mediaFiles.length} image files saved`);
    }

    // Save CSS file
    if (cssOutputPath && result.css) {
      await fs.writeFile(cssPath, result.css, 'utf8');
      console.log('💾 CSS file saved:', cssPath);
    }

    // Cleanup resources
    engine.cleanup();

    return {
      success: true,
      message: 'Dual parsing conversion completed (optimized version)',
      result: {
        html: result.html,
        css: result.css,
        completeHTML: result.completeHTML,
        mediaFiles: result.mediaFiles.map(f => {
          const { safePathJoin } = require('../security/securityConfig');
          return {
            name: f.name,
            type: f.type,
            size: f.size,
            localPath: safePathJoin(imageOutputDir, f.name),
          };
        }),
        outputPath: htmlOutputPath,
        cssOutputPath: cssOutputPath,
      },
      statistics: {
        styles: result.details.styleExtraction,
        conversion: result.details.mammothConversion,
        css: result.details.cssGeneration,
        html: result.details.htmlProcessing,
        media: result.details.mediaExtraction,
        performance: result.performance,
      },
      enhancements: {
        deepParsing: true,
        completeStyleMapping: true,
        enhancedCSS: true,
        styleInjection: true,
      },
    };
  } catch (error: any) {
    console.error('❌ Dual parsing conversion failed:', error);
    return {
      success: false,
      error: error.message,
      message: 'Dual parsing conversion failed',
    };
  }
}

/**
 * Dual parsing DOCX to PDF
 */
export async function dualParsingDocxToPdf(args: any): Promise<any> {
  try {
    const { docxPath, outputPath, options = {} } = args;

    // Validate and secure input paths
    const validatedDocxPath = validatePath(docxPath);
    const validatedOutputPath = outputPath ? validatePath(outputPath) : null;

    // Validate input file
    try {
      await fs.access(validatedDocxPath);
    } catch (error) {
      throw new Error(`DOCX file does not exist: ${validatedDocxPath}`);
    }

    // First convert to HTML
    const htmlResult = await dualParsingDocxToHtml({
      docxPath: validatedDocxPath,
      options: {
        ...options,
        generateCompleteHTML: true,
        inlineCSS: true,
      },
    });

    if (!htmlResult.success) {
      throw new Error(`HTML conversion failed: ${htmlResult.error}`);
    }

    // Generate PDF output path
    const pdfOutputPath = validatedOutputPath || validatedDocxPath.replace(/\.docx$/i, '_converted.pdf');

    // Build playwright-mcp commands
    const playwrightCommands = [
      `browser_navigate("file://${htmlResult.outputPath}")`,
      `browser_wait_for({ time: 3 })`,
      `browser_pdf_save({ filename: "${pdfOutputPath}" })`,
    ];

    const instructions = `
Dual Parsing DOCX to PDF - Hybrid Mode Workflow

✅ Completed (Current MCP):
  1. DOCX file reading and parsing
  2. Dual parsing engine processing
  3. HTML file generation: ${htmlResult.outputPath}

🎯 To Execute (playwright-mcp):
  Please run the following commands to complete PDF conversion:
  
  ${playwrightCommands.map((cmd, i) => `  ${i + 1}. ${cmd}`).join('\n')}

📁 Final Output: ${pdfOutputPath}

💡 Advantages:
  - Preserves all dual parsing engine functionality
  - Uses existing playwright-mcp browser instance
  - Avoids Chromium startup issues
  - Perfectly preserves styles and formatting
`;

    return {
      success: false,
      useMcpMode: true,
      mode: 'hybrid_pdf_conversion',
      currentMcp: {
        completed: ['DOCX file reading and parsing', 'Dual parsing engine processing', 'HTML file generation'],
        htmlFile: htmlResult.outputPath,
        tempFiles: [htmlResult.outputPath],
      },
      playwrightMcp: {
        required: true,
        commands: playwrightCommands,
        reason: 'PDF conversion requires playwright-mcp browser instance',
      },
      finalOutput: pdfOutputPath,
      instructions,
      mcpCommands: playwrightCommands,
      stats: htmlResult.stats,
      pdfOptions: {
        format: options.pdfOptions?.format || 'A4',
        margin: options.pdfOptions?.margin || '1in',
        printBackground: options.pdfOptions?.printBackground !== false,
        landscape: options.pdfOptions?.landscape === true,
      },
    };
  } catch (error: any) {
    console.error('❌ Dual parsing PDF conversion failed:', error);
    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * DOCX style analysis
 */
export async function analyzeDocxStyles(args: any): Promise<any> {
  try {
    const { docxPath, outputPath, detailed = false } = args;

    // Validate and secure input paths
    const validatedDocxPath = validatePath(docxPath);
    const validatedOutputPath = outputPath ? validatePath(outputPath) : null;

    // Validate input file
    try {
      await fs.access(validatedDocxPath);
    } catch (error) {
      throw new Error(`DOCX file does not exist: ${validatedDocxPath}`);
    }

    // Create style extractor
    const extractor = new StyleExtractor();

    // Extract styles
    const result = await extractor.extractStyles(validatedDocxPath);

    // Analyze styles
    const analysis = {
      summary: {
        totalStyles: result.styles.size,
        documentStyles: result.documentStyles.length,
        timestamp: new Date().toISOString(),
      },
      styleTypes: {
        paragraph: 0,
        character: 0,
        table: 0,
        numbering: 0,
      },
      styles: detailed ? Array.from(result.styles.entries()) : [],
      documentStyles: detailed ? result.documentStyles : [],
    };

    // Count style types
    for (const [styleId, style] of result.styles) {
      analysis.styleTypes[style.type]++;
    }

    // Save analysis report
    if (validatedOutputPath) {
      await fs.writeFile(validatedOutputPath, JSON.stringify(analysis, null, 2));
    }

    return {
      success: true,
      message: 'Style analysis completed',
      analysis,
      outputPath: validatedOutputPath,
    };
  } catch (error: any) {
    console.error('❌ Style analysis failed:', error);
    return {
      success: false,
      error: error.message,
    };
  }
}

// Export tool definitions
export const DUAL_PARSING_TOOLS = [
  DUAL_PARSING_DOCX_TO_HTML_TOOL,
  DUAL_PARSING_DOCX_TO_PDF_TOOL,
  DOCX_STYLE_ANALYSIS_TOOL,
];
