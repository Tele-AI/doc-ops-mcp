# PDF生成功能使用指南

本文档介绍了如何使用doc-ops-mcp中新增的PDF生成功能。

## 功能概述

新增的PDF生成功能基于Puppeteer，可以将HTML内容转换为高质量的PDF文档。主要特性包括：

- 🎨 支持完整的HTML和CSS样式
- 📄 多种页面格式支持（A4、A3、Letter等）
- 🖼️ 支持背景图片和颜色
- 📏 灵活的页边距设置
- 🌐 支持中文字体渲染
- ⚡ 高性能PDF生成

## 快速开始

### 1. 生成Docker简介PDF

```typescript
import { DualParsingEngine } from '../src/tools/dualParsingEngine';

async function generateDockerPDF() {
  const engine = new DualParsingEngine();
  
  // 生成Docker简介PDF
  const result = await engine.generateDockerIntroPDF('docker_intro.pdf');
  
  if (result.success) {
    console.log('PDF生成成功:', result.outputPath);
    console.log('文件大小:', result.fileSize, 'bytes');
  } else {
    console.error('生成失败:', result.error);
  }
}
```

### 2. 从HTML内容生成PDF

```typescript
import { DualParsingEngine } from '../src/tools/dualParsingEngine';

async function generateCustomPDF() {
  const engine = new DualParsingEngine();
  
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        h1 { color: #0066cc; }
        p { font-family: Arial, sans-serif; }
      </style>
    </head>
    <body>
      <h1>我的文档</h1>
      <p>这是一个自定义的PDF文档。</p>
    </body>
    </html>
  `;
  
  const result = await engine.generatePDFFromHTML(
    htmlContent, 
    'my_document.pdf',
    {
      pageFormat: 'A4',
      orientation: 'portrait',
      margins: {
        top: '2cm',
        bottom: '2cm',
        left: '2cm',
        right: '2cm'
      }
    }
  );
  
  console.log('PDF生成结果:', result);
}
```

### 3. 从HTML文件生成PDF

```typescript
import { DualParsingEngine } from '../src/tools/dualParsingEngine';

async function generateFromFile() {
  const engine = new DualParsingEngine();
  
  const result = await engine.generatePDFFromFile(
    'input.html',
    'output.pdf'
  );
  
  console.log('PDF生成结果:', result);
}
```

### 4. 从文档转换结果生成PDF

```typescript
import { DualParsingEngine } from '../src/tools/dualParsingEngine';

async function convertDocxToPDF() {
  const engine = new DualParsingEngine({
    outputOptions: {
      generateCompleteHTML: true,
      includeCSS: true
    }
  });
  
  // 先转换DOCX到HTML
  const conversionResult = await engine.convertDocxToHtml('document.docx');
  
  if (conversionResult.success) {
    // 然后生成PDF
    const pdfResult = await engine.generatePDFFromConversionResult(
      conversionResult,
      'document.pdf'
    );
    
    console.log('PDF生成结果:', pdfResult);
  }
}
```

## PDF生成选项

### PDFGeneratorOptions

```typescript
interface PDFGeneratorOptions {
  outputPath?: string;              // 输出路径
  pageFormat?: 'A4' | 'A3' | 'Letter';  // 页面格式
  orientation?: 'portrait' | 'landscape';  // 页面方向
  margins?: {                       // 页边距
    top?: string;
    bottom?: string;
    left?: string;
    right?: string;
  };
  includeBackground?: boolean;      // 包含背景
  printBackground?: boolean;        // 打印背景
  scale?: number;                   // 缩放比例
  headerTemplate?: string;          // 页眉模板
  footerTemplate?: string;          // 页脚模板
  displayHeaderFooter?: boolean;    // 显示页眉页脚
}
```

### 使用示例

```typescript
const pdfOptions = {
  pageFormat: 'A4' as const,
  orientation: 'portrait' as const,
  margins: {
    top: '2cm',
    bottom: '2cm',
    left: '2cm',
    right: '2cm'
  },
  printBackground: true,
  scale: 1,
  displayHeaderFooter: true,
  headerTemplate: '<div style="font-size:10px; text-align:center; width:100%;">页眉内容</div>',
  footerTemplate: '<div style="font-size:10px; text-align:center; width:100%;">第 <span class="pageNumber"></span> 页</div>'
};
```

## 运行示例

1. 确保已安装依赖：
```bash
npm install
```

2. 运行Docker PDF生成示例：
```bash
npx ts-node examples/generateDockerPDF.ts
```

3. 或者编译后运行：
```bash
npm run build
node dist/examples/generateDockerPDF.js
```

## 注意事项

1. **依赖要求**：PDF生成功能需要Puppeteer，确保已正确安装。

2. **字体支持**：系统需要支持中文字体，建议在HTML中指定字体族：
   ```css
   font-family: 'Microsoft YaHei', 'SimHei', Arial, sans-serif;
   ```

3. **性能考虑**：PDF生成是CPU密集型操作，大文档可能需要较长时间。

4. **内存使用**：Puppeteer会启动Chrome实例，注意内存使用情况。

5. **错误处理**：始终检查返回结果的`success`字段，并处理可能的错误。

## 故障排除

### 常见问题

1. **Puppeteer安装失败**
   ```bash
   npm install puppeteer --force
   ```

2. **中文字体显示问题**
   - 确保系统安装了中文字体
   - 在CSS中明确指定字体

3. **PDF生成失败**
   - 检查HTML内容是否有效
   - 确保输出目录存在且有写权限
   - 查看错误信息进行调试

4. **内存不足**
   - 减少并发PDF生成数量
   - 增加系统内存
   - 优化HTML内容大小

## API参考

### DualParsingEngine PDF方法

- `generateDockerIntroPDF(outputPath?: string): Promise<PDFGenerationResult>`
- `generatePDFFromHTML(htmlContent: string, outputPath: string, options?: PDFGeneratorOptions): Promise<PDFGenerationResult>`
- `generatePDFFromFile(htmlFilePath: string, outputPath?: string, options?: PDFGeneratorOptions): Promise<PDFGenerationResult>`
- `generatePDFFromConversionResult(result: DualParsingResult, outputPath: string, options?: PDFGeneratorOptions): Promise<PDFGenerationResult>`

### PDFGenerationResult

```typescript
interface PDFGenerationResult {
  success: boolean;     // 是否成功
  outputPath?: string;  // 输出文件路径
  error?: string;       // 错误信息
  fileSize?: number;    // 文件大小（字节）
}
```

## 更多示例

查看 `examples/generateDockerPDF.ts` 文件获取完整的使用示例。