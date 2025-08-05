/**
 * Enhanced Mammoth Configuration for Better Style Preservation
 * 解决 DOCX 转 HTML 时样式丢失的问题
 */

const mammoth = require('mammoth');

/**
 * 创建增强的 mammoth 配置，保留更多样式信息
 */
export function createEnhancedMammothConfig() {
  return {
    // 样式映射配置 - 这是关键！
    styleMap: [
      // 段落样式映射
      "p[style-name='Normal'] => p:fresh",
      "p[style-name='Heading 1'] => h1:fresh",
      "p[style-name='Heading 2'] => h2:fresh",
      "p[style-name='Heading 3'] => h3:fresh",
      "p[style-name='Heading 4'] => h4:fresh",
      "p[style-name='Heading 5'] => h5:fresh",
      "p[style-name='Heading 6'] => h6:fresh",
      "p[style-name='Title'] => h1.title:fresh",
      "p[style-name='Subtitle'] => h2.subtitle:fresh",

      // 字符样式映射
      "r[style-name='Strong'] => strong",
      "r[style-name='Emphasis'] => em",
      "r[style-name='Hyperlink'] => a",

      // 列表样式映射
      "p[style-name='List Paragraph'] => li:fresh",
      "p[style-name='Bullet List'] => li:fresh",
      "p[style-name='Numbered List'] => li:fresh",

      // 表格样式映射
      'table => table.docx-table',
      'tr => tr',
      'td => td',
      'th => th',

      // 通用样式映射 - 保留所有格式
      'p => p:fresh',
      'r => span',
      'b => strong',
      'i => em',
      'u => u',
      'strike => del',
      'sup => sup',
      'sub => sub',
    ],

    // 转换选项
    convertImage: mammoth.images.imgElement(function (image) {
      // 图片处理 - 转换为 base64 或保存到文件
      return image.read('base64').then(function (imageBuffer) {
        const extension = image.contentType.split('/')[1] || 'png';
        return {
          src: `data:${image.contentType};base64,${imageBuffer}`,
          alt: image.altText || 'Document Image',
        };
      });
    }),

    // 包含默认样式映射
    includeDefaultStyleMap: true,

    // 包含嵌入样式
    includeEmbeddedStyleMap: true,

    // 忽略空段落
    ignoreEmptyParagraphs: false,

    // 转换注释
    transformDocument: mammoth.transforms.paragraph(function (element) {
      // 保留段落的对齐方式
      if (element.alignment) {
        return {
          ...element,
          styleName: element.styleName || 'Normal',
        };
      }
      return element;
    }),
  };
}

/**
 * 增强的 DOCX 转 HTML 函数，保留更多样式
 */
export async function convertDocxToHtmlWithStyles(inputPath: string, options: any = {}) {
  try {
    console.error('🎨 使用增强配置转换 DOCX 到 HTML...');

    const config = createEnhancedMammothConfig();

    // 如果需要保存图片到文件而不是 base64
    if (options.saveImagesToFiles) {
      const fs = require('fs/promises');
      const path = require('path');

      const imageDir = options.imageOutputDir || path.join(process.cwd(), 'output', 'images');
      await fs.mkdir(imageDir, { recursive: true });

      config.convertImage = mammoth.images.imgElement(function (image) {
        return image.read().then(async function (imageBuffer) {
          const extension = image.contentType.split('/')[1] || 'png';
          const crypto = require('crypto');
          const randomId = crypto.randomBytes(4).toString('hex');
          const filename = `image_${Date.now()}_${randomId}.${extension}`;
          const imagePath = path.join(imageDir, filename);

          await fs.writeFile(imagePath, imageBuffer);
          console.error(`💾 图片已保存: ${imagePath}`);

          return {
            src: imagePath,
            alt: image.altText || 'Document Image',
          };
        });
      });
    }

    const result = await mammoth.convertToHtml({ path: inputPath }, config);

    console.error(`✅ 转换完成，消息数量: ${result.messages.length}`);

    // 输出转换消息以便调试
    if (result.messages.length > 0) {
      console.error('📋 转换消息:');
      result.messages.forEach((msg, index) => {
        console.error(`  ${index + 1}. ${msg.type}: ${msg.message}`);
      });
    }

    // 增强 HTML 内容，添加完整的样式
    const enhancedHtml = enhanceHtmlWithStyles(result.value, options);

    return {
      success: true,
      content: enhancedHtml,
      messages: result.messages,
      metadata: {
        format: 'html',
        originalFormat: 'docx',
        stylesPreserved: true,
        messagesCount: result.messages.length,
      },
    };
  } catch (error: any) {
    console.error('❌ 增强转换失败:', error);
    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * 增强 HTML 内容，添加完整的 CSS 样式
 */
function enhanceHtmlWithStyles(htmlContent: string, options: any = {}): string {
  const styles = `
    <style>
      /* 重置和基础样式 */
      * {
        box-sizing: border-box;
        margin: 0;
        padding: 0;
      }
      
      @page {
        size: A4;
        margin: 2.54cm;
      }
      
      body {
        font-family: "Calibri", "Microsoft YaHei", "SimSun", Arial, sans-serif;
        font-size: 11pt;
        line-height: 1.15;
        color: #000000;
        background: #ffffff;
        max-width: 21cm;
        margin: 0 auto;
        padding: 2.54cm;
        word-wrap: break-word;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      
      /* 段落样式 */
      p {
        margin: 0 0 8pt 0;
        text-align: justify;
        text-justify: inter-ideograph;
        line-height: 1.08;
        font-size: 11pt;
      }
      
      /* 标题样式 */
      h1, h2, h3, h4, h5, h6 {
        font-weight: bold;
        margin: 12pt 0 6pt 0;
        page-break-after: avoid;
      }
      
      h1 {
        font-size: 16pt;
        color: #2F5496;
      }
      
      h2 {
        font-size: 13pt;
        color: #2F5496;
      }
      
      h3 {
        font-size: 12pt;
        color: #1F3763;
      }
      
      /* 文本格式 */
      strong, b {
        font-weight: bold;
      }
      
      em, i {
        font-style: italic;
      }
      
      u {
        text-decoration: underline;
      }
      
      del, strike {
        text-decoration: line-through;
      }
      
      sup {
        vertical-align: super;
        font-size: smaller;
      }
      
      sub {
        vertical-align: sub;
        font-size: smaller;
      }
      
      /* 表格样式 */
      table.docx-table {
        width: 100%;
        border-collapse: collapse;
        margin: 12pt 0;
        font-size: 10pt;
      }
      
      table.docx-table th,
      table.docx-table td {
        border: 1pt solid #000000;
        padding: 4pt 6pt;
        vertical-align: top;
        text-align: left;
      }
      
      table.docx-table th {
        background-color: #D9E2F3;
        font-weight: bold;
      }
      
      /* 列表样式 */
      ul, ol {
        margin: 6pt 0 6pt 24pt;
        padding: 0;
      }
      
      li {
        margin: 3pt 0;
        line-height: 1.08;
      }
      
      /* 图片样式 */
      img {
        max-width: 100%;
        height: auto;
        display: block;
        margin: 6pt auto;
      }
      
      /* 链接样式 */
      a {
        color: #0563C1;
        text-decoration: underline;
      }
      
      a:visited {
        color: #954F72;
      }
      
      /* 引用样式 */
      blockquote {
        margin: 12pt 24pt;
        padding: 6pt 12pt;
        border-left: 3pt solid #CCCCCC;
        background-color: #F8F8F8;
        font-style: italic;
      }
      
      /* 代码样式 */
      code {
        font-family: "Courier New", monospace;
        background-color: #F2F2F2;
        padding: 1pt 3pt;
        border-radius: 2pt;
      }
      
      pre {
        font-family: "Courier New", monospace;
        background-color: #F8F8F8;
        padding: 6pt;
        border: 1pt solid #CCCCCC;
        border-radius: 3pt;
        overflow-x: auto;
        margin: 6pt 0;
      }
      
      /* 分页控制 */
      .page-break {
        page-break-before: always;
      }
      
      /* 中文字体优化 */
      .chinese {
        font-family: "Microsoft YaHei", "SimSun", serif;
      }
      
      /* 打印样式 */
      @media print {
        body {
          margin: 0;
          padding: 0;
        }
        
        .no-print {
          display: none;
        }
      }
    </style>
  `;

  // 如果 HTML 已经包含完整的文档结构
  if (htmlContent.includes('<html>') || htmlContent.includes('<!DOCTYPE')) {
    // 在 head 中插入样式
    return htmlContent.replace(/<\/head>/i, `${styles}</head>`);
  } else {
    // 创建完整的 HTML 文档
    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Document</title>
  ${styles}
</head>
<body>
${htmlContent}
</body>
</html>`;
  }
}

/**
 * 分析 DOCX 文档的样式信息
 */
export async function analyzeDocxStyles(inputPath: string) {
  try {
    console.error('🔍 分析 DOCX 文档样式...');

    const config = {
      ...createEnhancedMammothConfig(),
      // 添加样式分析
      transformDocument: mammoth.transforms.paragraph(function (element) {
        console.error(
          `段落样式: ${element.styleName || 'Default'}, 对齐: ${element.alignment || 'left'}`
        );
        return element;
      }),
    };

    const result = await mammoth.convertToHtml({ path: inputPath }, config);

    console.error('📊 样式分析完成');
    console.error(`消息数量: ${result.messages.length}`);

    return {
      success: true,
      messages: result.messages,
      analysis: {
        hasStyles: result.messages.some(msg => msg.message.includes('style')),
        messageCount: result.messages.length,
        content: result.value,
      },
    };
  } catch (error: any) {
    console.error('❌ 样式分析失败:', error);
    return {
      success: false,
      error: error.message,
    };
  }
}
