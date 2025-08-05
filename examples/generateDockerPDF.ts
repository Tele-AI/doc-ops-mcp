import { DualParsingEngine } from '../src/tools/dualParsingEngine';
import * as path from 'path';

/**
 * PDF 生成示例
 * 包含传统 DualParsingEngine 和新 MCP 工具两种方式
 * 
 * 推荐使用 MCP 工具方式，功能更强大且更稳定
 */
async function generateDockerIntroPDF() {
  console.log('🚀 开始生成Docker简介PDF...');

  try {
    // 创建双解析引擎实例
    const engine = new DualParsingEngine();

    // 生成PDF文件路径
    const outputPath = path.join(process.cwd(), 'docker_intro_new.pdf');

    // 使用新的文档转换功能生成Docker简介PDF
    const result = await engine.generateDockerIntroDocument('pdf', outputPath);

    if (result.success) {
      console.log('✅ Docker简介PDF生成成功!');
      console.log(`📄 文件路径: ${result.outputPath}`);
      console.log(`📊 文件大小: ${(result.fileSize! / 1024).toFixed(2)} KB`);
    } else {
      console.error('❌ PDF生成失败:', result.error);
    }
  } catch (error: any) {
    console.error('❌ 发生错误:', error.message);
  }
}

/**
 * 示例：从自定义文本内容生成PDF
 */
async function generateCustomPDF() {
  console.log('🚀 开始生成自定义PDF...');

  try {
    const engine = new DualParsingEngine();

    // 自定义文本内容
    const customContent = `
这是一个自定义的技术文档示例，展示了如何使用我们的文档转换工具创建专业的技术文档。

## 概述

我们的文档转换工具支持多种格式输出，包括PDF、HTML、Markdown和Word文档。

## 特性

我们的文档转换工具具有以下特性：

- 支持多种输出格式
- 高质量的文档输出
- 灵活的样式设置
- 支持中文内容
- 简单易用的API

## 应用场景

### 技术文档
适用于API文档、用户手册、技术规范等。

### 报告生成
可用于生成项目报告、分析报告等。

### 内容发布
支持将内容发布为多种格式，满足不同需求。

## 结论

通过这个工具，您可以轻松地将文本内容转换为专业的文档格式。
    `;

    const outputPath = path.join(process.cwd(), 'custom_document_new.pdf');

    const result = await engine.generateDocumentFromText('自定义技术文档', customContent, 'pdf', {
      author: '技术团队',
      description: '展示文档转换工具功能的示例文档',
      outputPath: outputPath,
    });

    if (result.success) {
      console.log('✅ 自定义PDF生成成功!');
      console.log(`📄 文件路径: ${result.outputPath}`);
      console.log(`📊 文件大小: ${(result.fileSize! / 1024).toFixed(2)} KB`);
    } else {
      console.error('❌ PDF生成失败:', result.error);
    }
  } catch (error: any) {
    console.error('❌ 发生错误:', error.message);
  }
}

/**
 * 主函数
 */
async function main() {
  console.log('📚 PDF生成示例程序');
  console.log('==================');

  // 生成Docker简介PDF
  await generateDockerIntroPDF();

  console.log('\n');

  // 生成自定义PDF
  await generateCustomPDF();

  console.log('\n✨ 所有PDF生成完成!');
}

// 运行示例
if (require.main === module) {
  main().catch(console.error);
}

/**
 * =================================================================
 * 新增：MCP 工具 PDF 生成示例
 * =================================================================
 * 以下示例展示如何使用 MCP 工具生成高质量的 PDF 文档
 */

/**
 * MCP 示例：完整的 Markdown 转 PDF 工作流
 */
function demonstrateMcpPdfGeneration() {
  console.log('🚀 MCP PDF 生成示例');
  console.log('==================');
  
  console.log('\n📋 步骤1: 转换规划');
  const planConversion = {
    tool: 'plan_conversion',
    params: {
      sourceFormat: 'markdown',
      targetFormat: 'pdf',
      sourceFile: '/input/docker_guide.md',
      preserveStyles: true,
      theme: 'github',
      quality: 'high'
    }
  };
  console.log(JSON.stringify(planConversion, null, 2));
  
  console.log('\n🔄 步骤2: Markdown 转 PDF');
  const convertToPdf = {
    tool: 'convert_markdown_to_pdf',
    params: {
      markdownPath: '/input/docker_guide.md',
      theme: 'github',
      includeTableOfContents: true,
      addQrCode: true
    }
  };
  console.log(JSON.stringify(convertToPdf, null, 2));
  
  console.log('\n✨ 步骤3: PDF 后处理');
  const postProcess = {
    tool: 'process_pdf_post_conversion',
    params: {
      playwrightPdfPath: '/tmp/playwright_docker_guide.pdf',
      targetPath: '/output/docker_guide_final.pdf',
      addWatermark: true,
      addQrCode: true,
      watermarkImageScale: 0.25,
      watermarkImageOpacity: 0.6,
      watermarkImagePosition: 'fullscreen',
      qrPosition: 'bottom-center',
      customText: 'Docker 指南 - 扫码获取最新版本'
    }
  };
  console.log(JSON.stringify(postProcess, null, 2));
}

/**
 * MCP 示例：从文本内容生成 PDF
 */
function demonstrateMcpTextToPdf() {
  console.log('\n📝 MCP 文本转 PDF 示例');
  console.log('=====================');
  
  // 首先写入 Markdown 文件
  const writeMarkdown = {
    tool: 'write_document',
    params: {
      content: `# 容器化技术指南\n\n## 概述\n\n容器化技术是现代软件开发和部署的重要技术。\n\n### Docker 基础\n\nDocker 是最流行的容器化平台，提供了：\n\n- 轻量级虚拟化\n- 快速部署\n- 环境一致性\n- 资源隔离\n\n### 最佳实践\n\n1. **镜像优化**\n   - 使用多阶段构建\n   - 选择合适的基础镜像\n   - 清理不必要的文件\n\n2. **安全考虑**\n   - 不要以 root 用户运行\n   - 定期更新镜像\n   - 扫描安全漏洞\n\n3. **性能优化**\n   - 合理设置资源限制\n   - 使用健康检查\n   - 优化网络配置\n\n## 结论\n\n容器化技术极大地简化了应用的开发、测试和部署流程。`,
      filePath: '/tmp/containerization_guide.md',
      format: 'markdown'
    }
  };
  
  console.log('📄 步骤1 - 写入 Markdown:');
  console.log(JSON.stringify(writeMarkdown, null, 2));
  
  // 然后转换为 PDF
  const convertToPdf = {
    tool: 'convert_markdown_to_pdf',
    params: {
      markdownPath: '/tmp/containerization_guide.md',
      outputPath: '/output/containerization_guide.pdf',
      theme: 'academic',
      includeTableOfContents: true,
      addQrCode: false
    }
  };
  
  console.log('\n🔄 步骤2 - 转换为 PDF:');
  console.log(JSON.stringify(convertToPdf, null, 2));
}

/**
 * MCP 示例：批量 PDF 生成
 */
function demonstrateMcpBatchPdfGeneration() {
  console.log('\n📚 MCP 批量 PDF 生成示例');
  console.log('========================');
  
  const documents = [
    {
      name: 'api_documentation',
      theme: 'github',
      title: 'API 文档'
    },
    {
      name: 'user_manual',
      theme: 'academic',
      title: '用户手册'
    },
    {
      name: 'technical_specification',
      theme: 'professional',
      title: '技术规范'
    }
  ];
  
  documents.forEach((doc, index) => {
    console.log(`\n📄 文档 ${index + 1}: ${doc.title}`);
    
    // 规划转换
    const plan = {
      tool: 'plan_conversion',
      params: {
        sourceFormat: 'markdown',
        targetFormat: 'pdf',
        sourceFile: `/input/${doc.name}.md`,
        theme: doc.theme,
        quality: 'high'
      }
    };
    
    // 转换为 PDF
    const convert = {
      tool: 'convert_markdown_to_pdf',
      params: {
        markdownPath: `/input/${doc.name}.md`,
        outputPath: `/output/${doc.name}.pdf`,
        theme: doc.theme,
        includeTableOfContents: true,
        addQrCode: true
      }
    };
    
    // 后处理
    const postProcess = {
      tool: 'process_pdf_post_conversion',
      params: {
        playwrightPdfPath: `/tmp/${doc.name}_temp.pdf`,
        targetPath: `/output/${doc.name}_final.pdf`,
        addWatermark: true,
        addQrCode: true,
        customText: `${doc.title} - 官方文档`
      }
    };
    
    console.log('  规划:', JSON.stringify(plan, null, 2));
    console.log('  转换:', JSON.stringify(convert, null, 2));
    console.log('  后处理:', JSON.stringify(postProcess, null, 2));
  });
}

/**
 * MCP PDF 生成最佳实践指南
 */
function printMcpPdfBestPractices() {
  console.log('\n📖 MCP PDF 生成最佳实践');
  console.log('=======================');
  
  console.log('\n🎯 工作流程:');
  console.log('1. plan_conversion - 规划转换路径');
  console.log('2. convert_markdown_to_pdf - 执行转换');
  console.log('3. process_pdf_post_conversion - 后处理');
  
  console.log('\n🎨 主题选择:');
  console.log('- github: 适合技术文档和代码相关内容');
  console.log('- academic: 适合学术论文和正式报告');
  console.log('- modern: 适合现代设计风格的文档');
  console.log('- professional: 适合商务和企业文档');
  
  console.log('\n⚙️ 配置建议:');
  console.log('- 设置 OUTPUT_DIR 环境变量指定输出目录');
  console.log('- 设置 WATERMARK_IMAGE 环境变量指定默认水印');
  console.log('- 设置 QR_CODE_IMAGE 环境变量指定默认二维码');
  
  console.log('\n🔧 依赖要求:');
  console.log('- playwright-mcp 服务器必须运行');
  console.log('- 确保有足够的磁盘空间（PDF 文件较大）');
  console.log('- 建议使用 SSD 以提高转换速度');
  
  console.log('\n📏 性能优化:');
  console.log('- 单个文件建议不超过 50MB');
  console.log('- 批量处理时建议分批进行');
  console.log('- 使用 quality: "balanced" 平衡质量和速度');
}

/**
 * 运行所有 MCP PDF 示例
 */
function runAllMcpPdfExamples() {
  console.log('\n🚀 运行所有 MCP PDF 示例');
  console.log('=========================');
  
  demonstrateMcpPdfGeneration();
  demonstrateMcpTextToPdf();
  demonstrateMcpBatchPdfGeneration();
  printMcpPdfBestPractices();
  
  console.log('\n✅ 所有 MCP PDF 示例展示完成!');
  console.log('\n💡 下一步:');
  console.log('1. 启动 playwright-mcp 服务器');
  console.log('2. 配置环境变量（OUTPUT_DIR, WATERMARK_IMAGE, QR_CODE_IMAGE）');
  console.log('3. 准备 Markdown 源文件');
  console.log('4. 按照工作流程执行转换');
}

export { 
  // 原有的 DualParsingEngine 示例
  generateDockerIntroPDF, 
  generateCustomPDF,
  
  // 新增的 MCP 工具示例
  demonstrateMcpPdfGeneration,
  demonstrateMcpTextToPdf,
  demonstrateMcpBatchPdfGeneration,
  printMcpPdfBestPractices,
  runAllMcpPdfExamples
};
