/**
 * MCP 文档转换工具使用示例
 * 演示如何使用最新的 MCP 工具进行文档转换和处理
 * 
 * 注意：这些示例展示了如何调用 MCP 工具，实际使用时需要通过 MCP 客户端调用
 */

import * as path from 'path';

/**
 * 示例1：使用转换规划器进行智能转换
 * 这是所有转换操作的第一步，必须先调用此工具获取转换计划
 */
function demonstratePlanConversion() {
  console.log('🎯 示例1：智能转换规划');
  console.log('===================');
  
  // 示例：Markdown 转 PDF 的转换规划
  const markdownToPdfPlan = {
    tool: 'plan_conversion',
    params: {
      sourceFormat: 'markdown',
      targetFormat: 'pdf',
      sourceFile: '/path/to/document.md',
      preserveStyles: true,
      includeImages: true,
      theme: 'github',
      quality: 'high'
    }
  };
  
  console.log('📋 Markdown → PDF 转换规划:');
  console.log(JSON.stringify(markdownToPdfPlan, null, 2));
  
  // 示例：DOCX 转 HTML 的转换规划
  const docxToHtmlPlan = {
    tool: 'plan_conversion',
    params: {
      sourceFormat: 'docx',
      targetFormat: 'html',
      preserveStyles: true,
      theme: 'modern',
      quality: 'balanced'
    }
  };
  
  console.log('\n📋 DOCX → HTML 转换规划:');
  console.log(JSON.stringify(docxToHtmlPlan, null, 2));
}

/**
 * 示例2：基础文档读写操作
 */
function demonstrateBasicOperations() {
  console.log('\n📖 示例2：基础文档操作');
  console.log('===================');
  
  // 读取文档
  const readDocument = {
    tool: 'read_document',
    params: {
      filePath: '/path/to/document.pdf',
      extractImages: true,
      preserveFormatting: true
    }
  };
  
  console.log('📄 读取文档:');
  console.log(JSON.stringify(readDocument, null, 2));
  
  // 写入文档
  const writeDocument = {
    tool: 'write_document',
    params: {
      content: '# 我的文档\n\n这是文档内容...',
      filePath: '/output/my_document.md',
      format: 'markdown'
    }
  };
  
  console.log('\n✍️ 写入文档:');
  console.log(JSON.stringify(writeDocument, null, 2));
}

/**
 * 示例3：通用文档转换
 */
function demonstrateDocumentConversion() {
  console.log('\n🔄 示例3：通用文档转换');
  console.log('=====================');
  
  // 通用转换工具
  const convertDocument = {
    tool: 'convert_document',
    params: {
      inputPath: '/input/document.docx',
      outputPath: '/output/document.html',
      sourceFormat: 'docx',
      targetFormat: 'html',
      preserveStyles: true,
      includeImages: true
    }
  };
  
  console.log('🔄 通用文档转换:');
  console.log(JSON.stringify(convertDocument, null, 2));
}

/**
 * 示例4：Markdown 专用转换工具
 */
function demonstrateMarkdownConversions() {
  console.log('\n📝 示例4：Markdown 专用转换');
  console.log('==========================');
  
  // Markdown 转 HTML
  const mdToHtml = {
    tool: 'convert_markdown_to_html',
    params: {
      markdownPath: '/input/document.md',
      outputPath: '/output/document.html',
      theme: 'github',
      includeTableOfContents: true,
      preserveStyles: true
    }
  };
  
  console.log('📝 Markdown → HTML:');
  console.log(JSON.stringify(mdToHtml, null, 2));
  
  // Markdown 转 DOCX
  const mdToDocx = {
    tool: 'convert_markdown_to_docx',
    params: {
      markdownPath: '/input/document.md',
      outputPath: '/output/document.docx',
      theme: 'academic',
      includeTableOfContents: true,
      preserveStyles: true
    }
  };
  
  console.log('\n📝 Markdown → DOCX:');
  console.log(JSON.stringify(mdToDocx, null, 2));
  
  // Markdown 转 PDF（需要配合 playwright-mcp）
  const mdToPdf = {
    tool: 'convert_markdown_to_pdf',
    params: {
      markdownPath: '/input/document.md',
      outputPath: '/output/document.pdf',
      theme: 'modern',
      includeTableOfContents: true,
      addQrCode: true
    }
  };
  
  console.log('\n📝 Markdown → PDF:');
  console.log(JSON.stringify(mdToPdf, null, 2));
}

/**
 * 示例5：HTML 转换工具
 */
function demonstrateHtmlConversions() {
  console.log('\n🌐 示例5：HTML 转换');
  console.log('==================');
  
  // HTML 转 Markdown
  const htmlToMd = {
    tool: 'convert_html_to_markdown',
    params: {
      htmlPath: '/input/document.html',
      outputPath: '/output/document.md',
      preserveStyles: true,
      includeCSS: false,
      debug: false
    }
  };
  
  console.log('🌐 HTML → Markdown:');
  console.log(JSON.stringify(htmlToMd, null, 2));
}

/**
 * 示例6：PDF 后处理工具
 */
function demonstratePdfPostProcessing() {
  console.log('\n📄 示例6：PDF 后处理');
  console.log('===================');
  
  // PDF 后处理（添加水印和二维码）
  const pdfPostProcess = {
    tool: 'process_pdf_post_conversion',
    params: {
      playwrightPdfPath: '/tmp/playwright_generated.pdf',
      targetPath: '/output/final_document.pdf',
      addWatermark: true,
      addQrCode: true,
      watermarkImageScale: 0.25,
      watermarkImageOpacity: 0.6,
      watermarkImagePosition: 'fullscreen',
      qrScale: 0.15,
      qrOpacity: 1.0,
      qrPosition: 'bottom-center',
      customText: '扫描二维码获取更多信息'
    }
  };
  
  console.log('📄 PDF 后处理:');
  console.log(JSON.stringify(pdfPostProcess, null, 2));
}

/**
 * 示例7：文档增强功能
 */
function demonstrateDocumentEnhancements() {
  console.log('\n✨ 示例7：文档增强功能');
  console.log('======================');
  
  // 添加水印
  const addWatermark = {
    tool: 'add_watermark',
    params: {
      inputPath: '/input/document.pdf',
      outputPath: '/output/watermarked.pdf',
      watermarkText: '机密文档',
      watermarkImage: '/assets/watermark.png',
      imageScale: 0.3,
      imageOpacity: 0.5,
      imagePosition: 'center'
    }
  };
  
  console.log('🏷️ 添加水印:');
  console.log(JSON.stringify(addWatermark, null, 2));
  
  // 添加二维码
  const addQrCode = {
    tool: 'add_qrcode',
    params: {
      inputPath: '/input/document.pdf',
      outputPath: '/output/qr_document.pdf',
      qrCodePath: '/assets/qrcode.png',
      scale: 0.15,
      opacity: 1.0,
      position: 'bottom-right',
      customText: '扫码查看详情'
    }
  };
  
  console.log('\n📱 添加二维码:');
  console.log(JSON.stringify(addQrCode, null, 2));
}

/**
 * 示例8：完整的转换工作流
 */
function demonstrateCompleteWorkflow() {
  console.log('\n🔄 示例8：完整转换工作流');
  console.log('========================');
  
  console.log('步骤1: 规划转换');
  const step1 = {
    tool: 'plan_conversion',
    params: {
      sourceFormat: 'markdown',
      targetFormat: 'pdf',
      sourceFile: '/input/technical_doc.md',
      theme: 'academic',
      quality: 'high'
    }
  };
  console.log(JSON.stringify(step1, null, 2));
  
  console.log('\n步骤2: Markdown 转 PDF');
  const step2 = {
    tool: 'convert_markdown_to_pdf',
    params: {
      markdownPath: '/input/technical_doc.md',
      theme: 'academic',
      includeTableOfContents: true,
      addQrCode: true
    }
  };
  console.log(JSON.stringify(step2, null, 2));
  
  console.log('\n步骤3: PDF 后处理');
  const step3 = {
    tool: 'process_pdf_post_conversion',
    params: {
      playwrightPdfPath: '/tmp/generated.pdf',
      targetPath: '/output/final_technical_doc.pdf',
      addWatermark: true,
      addQrCode: true,
      customText: '技术文档 - 版权所有'
    }
  };
  console.log(JSON.stringify(step3, null, 2));
}

/**
 * 主函数
 */
function main() {
  console.log('🚀 MCP 文档转换工具使用示例');
  console.log('============================\n');
  
  // 运行所有示例
  demonstratePlanConversion();
  demonstrateBasicOperations();
  demonstrateDocumentConversion();
  demonstrateMarkdownConversions();
  demonstrateHtmlConversions();
  demonstratePdfPostProcessing();
  demonstrateDocumentEnhancements();
  demonstrateCompleteWorkflow();
  
  console.log('\n✅ 所有示例展示完成!');
  console.log('\n💡 使用提示:');
  console.log('1. 所有转换操作都应该先调用 plan_conversion 工具');
  console.log('2. PDF 转换需要配合 playwright-mcp 使用');
  console.log('3. 环境变量可以设置默认的水印和二维码图片');
  console.log('4. 输出路径如果不是绝对路径，会相对于 OUTPUT_DIR 环境变量');
}

// 运行示例
if (require.main === module) {
  main();
}

export {
  demonstratePlanConversion,
  demonstrateBasicOperations,
  demonstrateDocumentConversion,
  demonstrateMarkdownConversions,
  demonstrateHtmlConversions,
  demonstratePdfPostProcessing,
  demonstrateDocumentEnhancements,
  demonstrateCompleteWorkflow
};