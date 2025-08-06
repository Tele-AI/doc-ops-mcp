/**
 * MediaHandler - 媒体文件处理器
 * 提取 DOCX 中的图片和其他媒体文件，转换为 base64 内嵌格式
 * 处理图片尺寸和位置信息，优化媒体文件大小
 */

const JSZip = require('jszip');
const fs = require('fs/promises');
const path = require('path');
const xml2js = require('xml2js');
const crypto = require('crypto');

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

interface MediaFile {
  id: string;
  name: string;
  type: string;
  contentType: string;
  size: number;
  data: Buffer;
  base64: string;
  width?: number;
  height?: number;
  description?: string;
}

interface MediaHandlerOptions {
  extractToFiles?: boolean;
  outputDirectory?: string;
  convertToBase64?: boolean;
  optimizeImages?: boolean;
  maxImageSize?: number;
  supportedFormats?: string[];
  preserveOriginalSize?: boolean;
}

interface ExtractionResult {
  mediaFiles: MediaFile[];
  relationships: Map<string, string>;
  success: boolean;
  error?: string;
  stats: {
    totalFiles: number;
    totalSize: number;
    imagesCount: number;
    otherMediaCount: number;
  };
}

export class MediaHandler {
  private options: MediaHandlerOptions;
  private mediaFiles: MediaFile[] = [];
  private relationships: Map<string, string> = new Map();
  private documentRelationships: Map<string, any> = new Map();

  constructor(options: MediaHandlerOptions = {}) {
    this.options = {
      extractToFiles: false,
      convertToBase64: true,
      optimizeImages: false,
      maxImageSize: 1024 * 1024, // 1MB
      supportedFormats: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp'],
      preserveOriginalSize: true,
      ...options,
    };
  }

  /**
   * 从 DOCX 文件提取媒体文件
   */
  async extractMedia(docxPath: string): Promise<ExtractionResult> {
    try {
      console.log('🖼️ 开始提取媒体文件...');

      const validatedPath = validatePath(docxPath);
      const docxBuffer = await fs.readFile(validatedPath);
      const zip = await JSZip.loadAsync(docxBuffer);

      // 重置状态
      this.mediaFiles = [];
      this.relationships.clear();
      this.documentRelationships.clear();

      // 提取关系信息
      await this.extractRelationships(zip);

      // 提取媒体文件
      await this.extractMediaFiles(zip);

      // 处理图片尺寸信息
      await this.extractImageDimensions(zip);

      // 优化媒体文件
      if (this.options.optimizeImages) {
        await this.optimizeMediaFiles();
      }

      // 保存到文件
      if (this.options.extractToFiles && this.options.outputDirectory) {
        await this.saveMediaToFiles();
      }

      const stats = this.calculateStats();

      console.log(
        `✅ 媒体提取完成：${stats.totalFiles} 个文件，总大小 ${this.formatFileSize(stats.totalSize)}`
      );

      return {
        mediaFiles: this.mediaFiles,
        relationships: this.relationships,
        success: true,
        stats,
      };
    } catch (error: any) {
      console.error('❌ 媒体提取失败:', error);
      return {
        mediaFiles: [],
        relationships: new Map(),
        success: false,
        error: error.message,
        stats: {
          totalFiles: 0,
          totalSize: 0,
          imagesCount: 0,
          otherMediaCount: 0,
        },
      };
    }
  }

  /**
   * 提取关系信息
   */
  private async extractRelationships(zip: any): Promise<void> {
    try {
      // 提取主文档关系
      const mainRelsFile = zip.file('word/_rels/document.xml.rels');
      if (mainRelsFile) {
        const relsXml = await mainRelsFile.async('text');
        const relsData = await xml2js.parseStringPromise(relsXml);

        if (relsData.Relationships && relsData.Relationships.Relationship) {
          const relationships = Array.isArray(relsData.Relationships.Relationship)
            ? relsData.Relationships.Relationship
            : [relsData.Relationships.Relationship];

          for (const rel of relationships) {
            this.relationships.set(rel.$.Id, rel.$.Target);
            this.documentRelationships.set(rel.$.Id, {
              type: rel.$.Type,
              target: rel.$.Target,
            });
          }
        }
      }

      console.log(`📋 提取了 ${this.relationships.size} 个关系`);
    } catch (error: any) {
      console.warn('⚠️ 关系提取失败:', error.message);
    }
  }

  /**
   * 提取媒体文件
   */
  private async extractMediaFiles(zip: any): Promise<void> {
    const mediaFolder = zip.folder('word/media');
    if (!mediaFolder) {
      console.log('📁 未找到媒体文件夹');
      return;
    }

    const mediaFiles = Object.keys(mediaFolder.files).filter(
      name => name.startsWith('word/media/') && !name.endsWith('/')
    );

    for (const fileName of mediaFiles) {
      try {
        const file = zip.file(fileName);
        if (file) {
          const data = await file.async('nodebuffer');
          const mediaFile = await this.createMediaFile(fileName, data);

          if (mediaFile) {
            this.mediaFiles.push(mediaFile);
          }
        }
      } catch (error: any) {
        console.warn(`⚠️ 提取媒体文件失败 ${fileName}:`, error.message);
      }
    }
  }

  /**
   * 创建媒体文件对象
   */
  private async createMediaFile(fileName: string, data: Buffer): Promise<MediaFile | null> {
    const name = path.basename(fileName);
    const ext = path.extname(name).toLowerCase().substring(1);

    // 检查是否为支持的格式
    if (!this.options.supportedFormats?.includes(ext)) {
      console.warn(`⚠️ 不支持的媒体格式: ${ext}`);
      return null;
    }

    // 检查文件大小
    if (this.options.maxImageSize && data.length > this.options.maxImageSize) {
      console.warn(`⚠️ 媒体文件过大: ${name} (${this.formatFileSize(data.length)})`);
      if (!this.options.optimizeImages) {
        return null;
      }
    }

    const contentType = this.getContentType(ext);
    const base64 = this.options.convertToBase64 ? data.toString('base64') : '';

    // 生成唯一 ID
    const id = this.generateMediaId(name);

    return {
      id,
      name,
      type: ext,
      contentType,
      size: data.length,
      data,
      base64,
    };
  }

  /**
   * 提取图片尺寸信息
   */
  private async extractImageDimensions(zip: any): Promise<void> {
    try {
      const documentFile = zip.file('word/document.xml');
      if (!documentFile) return;

      const documentXml = await documentFile.async('text');
      const documentData = await xml2js.parseStringPromise(documentXml, {
        explicitArray: false,
        mergeAttrs: true,
      });

      // 查找图片引用和尺寸信息
      this.extractImageReferences(documentData['w:document']['w:body']);
    } catch (error: any) {
      console.warn('⚠️ 图片尺寸提取失败:', error.message);
    }
  }

  /**
   * 递归提取图片引用
   */
  private extractImageReferences(element: any): void {
    if (!element) return;

    if (Array.isArray(element)) {
      for (const item of element) {
        this.extractImageReferences(item);
      }
      return;
    }

    // 查找绘图对象
    if (element['w:drawing']) {
      this.processDrawing(element['w:drawing']);
    }

    // 查找图片对象
    if (element['w:pict']) {
      this.processPicture(element['w:pict']);
    }

    // 递归处理子元素
    for (const key in element) {
      if (typeof element[key] === 'object') {
        this.extractImageReferences(element[key]);
      }
    }
  }

  /**
   * 处理绘图对象
   */
  private processDrawing(drawing: any): void {
    try {
      // 查找内联图片
      if (drawing['wp:inline']) {
        this.processInlineDrawing(drawing['wp:inline']);
      }

      // 查找锚定图片
      if (drawing['wp:anchor']) {
        this.processAnchoredDrawing(drawing['wp:anchor']);
      }
    } catch (error: any) {
      console.warn('⚠️ 绘图对象处理失败:', error.message);
    }
  }

  /**
   * 处理内联绘图
   */
  private processInlineDrawing(inline: any): void {
    const extent = inline['wp:extent'];
    if (extent && extent.cx && extent.cy) {
      const width = this.emuToPixels(parseInt(extent.cx));
      const height = this.emuToPixels(parseInt(extent.cy));

      // 查找图片引用
      const blip = this.findBlipInDrawing(inline);
      if (blip && blip['r:embed']) {
        const relationId = blip['r:embed'];
        this.updateMediaDimensions(relationId, width, height);
      }
    }
  }

  /**
   * 处理锚定绘图
   */
  private processAnchoredDrawing(anchor: any): void {
    const extent = anchor['wp:extent'];
    if (extent && extent.cx && extent.cy) {
      const width = this.emuToPixels(parseInt(extent.cx));
      const height = this.emuToPixels(parseInt(extent.cy));

      // 查找图片引用
      const blip = this.findBlipInDrawing(anchor);
      if (blip && blip['r:embed']) {
        const relationId = blip['r:embed'];
        this.updateMediaDimensions(relationId, width, height);
      }
    }
  }

  /**
   * 处理图片对象（旧格式）
   */
  private processPicture(picture: any): void {
    // 处理 VML 格式的图片
    try {
      // 这里可以添加 VML 图片处理逻辑
    } catch (error: any) {
      console.warn('⚠️ VML 图片处理失败:', error.message);
    }
  }

  /**
   * 在绘图对象中查找 blip
   */
  private findBlipInDrawing(drawing: any): any {
    // 递归查找 a:blip 元素
    const findBlip = (obj: any): any => {
      if (!obj || typeof obj !== 'object') return null;

      if (obj['a:blip']) {
        return obj['a:blip'];
      }

      for (const key in obj) {
        const result = findBlip(obj[key]);
        if (result) return result;
      }

      return null;
    };

    return findBlip(drawing);
  }

  /**
   * 更新媒体文件尺寸
   */
  private updateMediaDimensions(relationId: string, width: number, height: number): void {
    const target = this.relationships.get(relationId);
    if (target) {
      const fileName = path.basename(target);
      const mediaFile = this.mediaFiles.find(file => file.name === fileName);
      if (mediaFile) {
        mediaFile.width = width;
        mediaFile.height = height;
      }
    }
  }

  /**
   * 优化媒体文件
   */
  private async optimizeMediaFiles(): Promise<void> {
    console.log('🔧 开始优化媒体文件...');

    for (const mediaFile of this.mediaFiles) {
      if (this.isImageFile(mediaFile.type)) {
        await this.optimizeImage(mediaFile);
      }
    }
  }

  /**
   * 优化图片
   */
  private async optimizeImage(mediaFile: MediaFile): Promise<void> {
    try {
      // 这里可以集成图片优化库，如 sharp
      // 目前保持简单实现

      if (mediaFile.size > this.options.maxImageSize!) {
        console.log(`🔧 优化图片: ${mediaFile.name}`);
        // 实际的图片优化逻辑
      }
    } catch (error: any) {
      console.warn(`⚠️ 图片优化失败 ${mediaFile.name}:`, error.message);
    }
  }

  /**
   * 保存媒体文件到磁盘
   */
  private async saveMediaToFiles(): Promise<void> {
    if (!this.options.outputDirectory) return;

    try {
      // 确保输出目录存在
      await fs.mkdir(this.options.outputDirectory, { recursive: true });

      // 导入安全配置函数
      const { safePathJoin, validateAndSanitizePath } = require('../security/securityConfig');
      const allowedPaths = [this.options.outputDirectory, process.cwd()];
      
      for (const mediaFile of this.mediaFiles) {
        // 使用安全的路径处理，防止路径遍历攻击
        const sanitizedName = path.basename(mediaFile.name); // 只取文件名，防止路径遍历
        const outputPath = safePathJoin(this.options.outputDirectory, sanitizedName);
        const validatedPath = validateAndSanitizePath(outputPath, allowedPaths);
        
        await fs.writeFile(validatedPath, mediaFile.data);
        console.log(`💾 保存媒体文件: ${validatedPath}`);
      }
    } catch (error: any) {
      console.error('❌ 保存媒体文件失败:', error);
    }
  }

  /**
   * 计算统计信息
   */
  private calculateStats(): {
    totalFiles: number;
    totalSize: number;
    imagesCount: number;
    otherMediaCount: number;
  } {
    const totalFiles = this.mediaFiles.length;
    const totalSize = this.mediaFiles.reduce((sum, file) => sum + file.size, 0);
    const imagesCount = this.mediaFiles.filter(file => this.isImageFile(file.type)).length;
    const otherMediaCount = totalFiles - imagesCount;

    return {
      totalFiles,
      totalSize,
      imagesCount,
      otherMediaCount,
    };
  }

  /**
   * 生成媒体文件 ID
   */
  private generateMediaId(fileName: string): string {
    const timestamp = Date.now();
    const random = crypto.randomBytes(3).toString('hex');
    const name = path.parse(fileName).name;
    return `${name}_${timestamp}_${random}`;
  }

  /**
   * 获取内容类型
   */
  private getContentType(extension: string): string {
    const contentTypes: { [key: string]: string } = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      svg: 'image/svg+xml',
      webp: 'image/webp',
      tiff: 'image/tiff',
      ico: 'image/x-icon',
    };

    return contentTypes[extension] || 'application/octet-stream';
  }

  /**
   * 判断是否为图片文件
   */
  private isImageFile(type: string): boolean {
    const imageTypes = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp', 'tiff', 'ico'];
    return imageTypes.includes(type.toLowerCase());
  }

  /**
   * EMU 转像素
   */
  private emuToPixels(emu: number): number {
    // 1 EMU = 1/914400 inch, 1 inch = 96 pixels (默认 DPI)
    return Math.round((emu / 914400) * 96);
  }

  /**
   * 格式化文件大小
   */
  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 B';

    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));

    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  /**
   * 获取媒体文件
   */
  getMediaFiles(): MediaFile[] {
    return this.mediaFiles;
  }

  /**
   * 根据 ID 获取媒体文件
   */
  getMediaById(id: string): MediaFile | undefined {
    return this.mediaFiles.find(file => file.id === id);
  }

  /**
   * 根据名称获取媒体文件
   */
  getMediaByName(name: string): MediaFile | undefined {
    return this.mediaFiles.find(file => file.name === name);
  }

  /**
   * 获取关系映射
   */
  getRelationships(): Map<string, string> {
    return this.relationships;
  }

  /**
   * 清理资源
   */
  cleanup(): void {
    this.mediaFiles = [];
    this.relationships.clear();
    this.documentRelationships.clear();
  }
}

export { MediaFile, MediaHandlerOptions, ExtractionResult };
