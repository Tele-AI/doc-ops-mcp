# 文档转换功能使用指南

## 功能概述

新增的文档转换功能允许用户直接输入文字内容，然后转换为多种格式的文档，包括：

- **Markdown (.md)** - 轻量级标记语言
- **HTML (.html)** - 网页格式，支持丰富样式
- **PDF (.pdf)** - 便携式文档格式
- **Word (.docx)** - Microsoft Word文档格式

## 主要特性

✅ **多格式支持** - 一次编写，多格式输出  
✅ **样式自定义** - 支持自定义颜色、字体、边距等  
✅ **中文支持** - 完整支持中文内容（PDF格式会转换为ASCII字符）  
✅ **简单易用** - 直观的API接口  
✅ **无外部依赖** - 不依赖Puppeteer等重型库  

## 快速开始

### 1. 基本使用

```typescript
import { DualParsingEngine } from '../src/tools/dualParsingEngine';

const engine = new DualParsingEngine();

// 生成Docker简介文档（多种格式）
const result = await engine.generateDockerIntroDocument('html');
console.log('生成结果:', result);
```

### 2. 自定义内容

```typescript
// 从文本生成文档
const result = await engine.generateDocumentFromText(
  '我的文档标题',
  '这是文档的主要内容...',
  'pdf',
  {
    author: '作者姓名',
    description: '文档描述',
    styling: {
      colors: {
        primary: '#0066cc',
        secondary: '#666666',
        text: '#333333'
      }
    }
  }
);
```

## API 参考

### DualParsingEngine 方法

#### `generateDockerIntroDocument(format, outputPath?)`

生成预设的Docker简介文档。

**参数：**
- `format`: `'md' | 'pdf' | 'docx' | 'html'` - 输出格式
- `outputPath`: `string` (可选) - 输出文件路径

**返回：** `Promise<DocumentConversionResult>`

#### `generateDocumentFromText(title, content, format, options?)`

从文本内容生成文档。

**参数：**
- `title`: `string` - 文档标题
- `content`: `string` - 文档内容
- `format`: `'md' | 'pdf' | 'docx' | 'html'` - 输出格式
- `options`: `object` (可选) - 额外选项
  - `author`: `string` - 作者
  - `description`: `string` - 描述
  - `outputPath`: `string` - 输出路径
  - `styling`: `object` - 样式设置

**返回：** `Promise<DocumentConversionResult>`

#### `convertDocumentContent(content, options)`

转换文档内容到指定格式。

**参数：**
- `content`: `DocumentContent` - 文档内容对象
- `options`: `ConversionOptions` - 转换选项

**返回：** `Promise<DocumentConversionResult>`

### 数据类型

#### `DocumentContent`

```typescript
interface DocumentContent {
  title?: string;        // 文档标题
  content: string;       // 主要内容
  author?: string;       // 作者
  description?: string;  // 描述
  metadata?: {           // 元数据
    [key: string]: string;
  };
}
```

#### `ConversionOptions`

```typescript
interface ConversionOptions {
  outputPath?: string;   // 输出路径
  format: 'md' | 'pdf' | 'docx' | 'html';  // 格式
  styling?: {            // 样式选项
    fontSize?: number;
    fontFamily?: string;
    lineHeight?: number;
    margins?: {
      top?: number;
      bottom?: number;
      left?: number;
      right?: number;
    };
    colors?: {
      primary?: string;    // 主色调
      secondary?: string;  // 次色调
      text?: string;       // 文本颜色
    };
  };
}
```

#### `DocumentConversionResult`

```typescript
interface DocumentConversionResult {
  success: boolean;      // 是否成功
  outputPath?: string;   // 输出文件路径
  error?: string;        // 错误信息
  fileSize?: number;     // 文件大小（字节）
  format: string;        // 输出格式
}
```

## 使用示例

### 示例1：生成技术文档

```typescript
import { DualParsingEngine } from '../src/tools/dualParsingEngine';

async function generateTechDoc() {
  const engine = new DualParsingEngine();
  
  const content = `
人工智能技术正在快速发展，影响着各个行业。

## 主要领域

### 机器学习
机器学习是AI的核心技术，包括监督学习、无监督学习等。

### 自然语言处理
使计算机能够理解和生成人类语言。

## 应用场景

- 医疗诊断
- 金融风控
- 智能驾驶
- 教育培训

## 未来展望

AI技术将继续发展，为人类社会带来更多价值。
  `;
  
  // 生成多种格式
  const formats = ['md', 'html', 'pdf', 'docx'] as const;
  
  for (const format of formats) {
    const result = await engine.generateDocumentFromText(
      '人工智能技术概述',
      content,
      format,
      {
        author: 'AI研究员',
        description: '人工智能技术的发展现状和未来趋势'
      }
    );
    
    if (result.success) {
      console.log(`✅ ${format.toUpperCase()}: ${result.outputPath}`);
    } else {
      console.error(`❌ ${format.toUpperCase()}: ${result.error}`);
    }
  }
}

generateTechDoc();
```

### 示例2：批量生成文档

```typescript
async function batchGenerate() {
  const engine = new DualParsingEngine();
  
  const documents = [
    {
      title: '项目计划书',
      content: '项目目标、时间安排、资源分配等内容...',
      author: '项目经理'
    },
    {
      title: '技术方案',
      content: '技术架构、实现方案、风险评估等内容...',
      author: '技术负责人'
    }
  ];
  
  for (const doc of documents) {
    const result = await engine.generateDocumentFromText(
      doc.title,
      doc.content,
      'html',
      { author: doc.author }
    );
    
    console.log(`文档 "${doc.title}" 生成结果:`, result.success);
  }
}
```

### 示例3：自定义样式

```typescript
async function generateStyledDoc() {
  const engine = new DualParsingEngine();
  
  const result = await engine.generateDocumentFromText(
    '设计规范',
    '这是一份设计规范文档...',
    'html',
    {
      styling: {
        fontSize: 16,
        fontFamily: 'Georgia',
        colors: {
          primary: '#2E8B57',   // 海绿色
          secondary: '#708090', // 石板灰
          text: '#2F4F4F'       // 深石板灰
        },
        margins: {
          top: 100,
          bottom: 100,
          left: 80,
          right: 80
        }
      }
    }
  );
  
  console.log('自定义样式文档:', result);
}
```

## 运行示例

1. **安装依赖**
   ```bash
   npm install
   ```

2. **运行完整示例**
   ```bash
   npx ts-node examples/documentConversionExample.ts
   ```

3. **运行单个功能测试**
   ```bash
   npx ts-node -e "import { DualParsingEngine } from './src/tools/dualParsingEngine'; const engine = new DualParsingEngine(); engine.generateDockerIntroDocument('html').then(console.log);"
   ```

## 格式特性说明

### Markdown (.md)
- ✅ 轻量级，易于编辑
- ✅ 完整支持中文
- ✅ 支持标题、列表、链接等
- ✅ 文件体积小

### HTML (.html)
- ✅ 丰富的样式支持
- ✅ 完整支持中文
- ✅ 可在浏览器中查看
- ✅ 支持自定义CSS样式

### PDF (.pdf)
- ✅ 跨平台兼容性好
- ⚠️ 中文字符转换为ASCII（显示为?）
- ✅ 适合打印和分享
- ✅ 保持格式一致性

### Word (.docx)
- ✅ 完整支持中文
- ✅ 专业文档格式
- ✅ 支持丰富的格式设置
- ✅ 与Microsoft Office兼容

## 注意事项

1. **PDF中文支持**：由于使用标准字体，PDF格式中的中文字符会被转换为ASCII字符。如需完整中文支持，建议使用HTML或DOCX格式。

2. **文件路径**：输出文件默认保存在当前工作目录，可通过`outputPath`参数自定义。

3. **样式设置**：不同格式对样式的支持程度不同，HTML格式支持最丰富的样式设置。

4. **内容格式**：支持简单的Markdown语法，如标题（#）、列表（-）等。

5. **错误处理**：始终检查返回结果的`success`字段，并处理可能的错误。

## 故障排除

### 常见问题

1. **依赖缺失**
   ```bash
   npm install docx pdf-lib marked
   ```

2. **文件权限问题**
   - 确保输出目录有写权限
   - 检查文件是否被其他程序占用

3. **内存不足**
   - 减少大文档的处理
   - 分批处理多个文档

4. **格式问题**
   - 检查输入内容的格式
   - 验证参数的正确性

## 更多信息

- 查看 `examples/documentConversionExample.ts` 获取完整示例
- 查看源码 `src/tools/documentConverter.ts` 了解实现细节
- 参考 `src/tools/dualParsingEngine.ts` 了解集成方式