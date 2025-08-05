/**
 * 文档转换规划器 - 智能转换路径规划和建议
 * 为大模型提供最佳的文档转换策略和步骤
 */

interface ConversionRequest {
  sourceFormat: string;
  targetFormat: string;
  sourceFile?: string;
  requirements?: {
    preserveStyles?: boolean;
    includeImages?: boolean;
    theme?: string;
    quality?: 'fast' | 'balanced' | 'high';
  };
}

interface ConversionStep {
  stepNumber: number;
  toolName: string;
  description: string;
  inputFormat: string;
  outputFormat: string;
  parameters: Record<string, any>;
  estimatedTime: string;
  notes?: string;
}

interface ConversionPlan {
  success: boolean;
  sourceFormat: string;
  targetFormat: string;
  conversionPath: string[];
  steps: ConversionStep[];
  totalSteps: number;
  estimatedTotalTime: string;
  recommendations: string[];
  warnings?: string[];
  error?: string;
}

class ConversionPlanner {
  // 支持的格式
  private supportedFormats = ['pdf', 'docx', 'html', 'markdown', 'md', 'txt', 'doc'];

  // 直接转换映射（一步到位）
  private directConversions: Record<string, Record<string, string>> = {
    docx: {
      pdf: 'convert_docx_to_pdf',
      html: 'convert_document',
      markdown: 'convert_document',
      md: 'convert_document',
      txt: 'convert_document',
    },
    markdown: {
      html: 'convert_markdown_to_html',
      docx: 'convert_markdown_to_docx',
      pdf: 'convert_markdown_to_html', // 需要两步：md->html->pdf
      txt: 'convert_document',
    },
    md: {
      html: 'convert_markdown_to_html',
      docx: 'convert_markdown_to_docx',
      pdf: 'convert_markdown_to_html', // 需要两步：md->html->pdf
      txt: 'convert_document',
    },
    html: {
      markdown: 'convert_html_to_markdown',
      md: 'convert_html_to_markdown',
      docx: 'convert_document',
      txt: 'convert_document',
      pdf: 'convert_document', // 需要外部工具
    },
    pdf: {
      txt: 'read_document',
      html: 'read_document', // 读取后写入
      markdown: 'read_document',
      md: 'read_document',
      docx: 'read_document',
    },
    txt: {
      html: 'write_document',
      markdown: 'write_document',
      md: 'write_document',
      docx: 'write_document',
      pdf: 'write_document',
    },
  };

  // 多步转换路径（当直接转换不可用时）
  private multiStepPaths: Record<string, Record<string, string[]>> = {
    pdf: {
      docx: ['pdf', 'html', 'docx'],
      markdown: ['pdf', 'html', 'markdown'],
      md: ['pdf', 'html', 'md'],
    },
    markdown: {
      pdf: ['markdown', 'html', 'pdf'],
    },
    md: {
      pdf: ['md', 'html', 'pdf'],
    },
    docx: {
      pdf: ['docx', 'html', 'pdf'], // 备用路径
    },
  };

  /**
   * 规划文档转换路径
   */
  async planConversion(request: ConversionRequest): Promise<ConversionPlan> {
    try {
      const sourceFormat = this.normalizeFormat(request.sourceFormat);
      const targetFormat = this.normalizeFormat(request.targetFormat);

      // 验证格式支持
      if (!this.isFormatSupported(sourceFormat)) {
        return {
          success: false,
          sourceFormat,
          targetFormat,
          conversionPath: [],
          steps: [],
          totalSteps: 0,
          estimatedTotalTime: '0分钟',
          recommendations: [],
          error: `不支持的源格式: ${sourceFormat}`,
        };
      }

      if (!this.isFormatSupported(targetFormat)) {
        return {
          success: false,
          sourceFormat,
          targetFormat,
          conversionPath: [],
          steps: [],
          totalSteps: 0,
          estimatedTotalTime: '0分钟',
          recommendations: [],
          error: `不支持的目标格式: ${targetFormat}`,
        };
      }

      // 相同格式
      if (sourceFormat === targetFormat) {
        return {
          success: true,
          sourceFormat,
          targetFormat,
          conversionPath: [sourceFormat],
          steps: [
            {
              stepNumber: 1,
              toolName: 'read_document',
              description: '文件已经是目标格式，无需转换',
              inputFormat: sourceFormat,
              outputFormat: targetFormat,
              parameters: {},
              estimatedTime: '即时',
            },
          ],
          totalSteps: 1,
          estimatedTotalTime: '即时',
          recommendations: ['文件格式已匹配，无需转换'],
        };
      }

      // 查找转换路径
      const conversionPath = this.findConversionPath(sourceFormat, targetFormat);
      if (!conversionPath || conversionPath.length === 0) {
        return {
          success: false,
          sourceFormat,
          targetFormat,
          conversionPath: [],
          steps: [],
          totalSteps: 0,
          estimatedTotalTime: '0分钟',
          recommendations: [],
          error: `无法找到从 ${sourceFormat} 到 ${targetFormat} 的转换路径`,
        };
      }

      // 生成转换步骤
      const steps = this.generateConversionSteps(conversionPath, request);
      const totalTime = this.calculateTotalTime(steps);
      const recommendations = this.generateRecommendations(sourceFormat, targetFormat, request);
      const warnings = this.generateWarnings(sourceFormat, targetFormat, conversionPath);

      return {
        success: true,
        sourceFormat,
        targetFormat,
        conversionPath,
        steps,
        totalSteps: steps.length,
        estimatedTotalTime: totalTime,
        recommendations,
        warnings,
      };
    } catch (error: any) {
      return {
        success: false,
        sourceFormat: request.sourceFormat,
        targetFormat: request.targetFormat,
        conversionPath: [],
        steps: [],
        totalSteps: 0,
        estimatedTotalTime: '0分钟',
        recommendations: [],
        error: `规划转换路径时出错: ${error.message}`,
      };
    }
  }

  /**
   * 标准化格式名称
   */
  private normalizeFormat(format: string): string {
    const normalized = format.toLowerCase().replace(/^\.|\.$/, '');
    // 统一 markdown 格式
    if (normalized === 'md') return 'markdown';
    return normalized;
  }

  /**
   * 检查格式是否支持
   */
  private isFormatSupported(format: string): boolean {
    return (
      this.supportedFormats.includes(format) ||
      this.supportedFormats.includes(format.replace('markdown', 'md'))
    );
  }

  /**
   * 查找转换路径
   */
  private findConversionPath(sourceFormat: string, targetFormat: string): string[] {
    // 首先尝试直接转换
    if (this.directConversions[sourceFormat]?.[targetFormat]) {
      return [sourceFormat, targetFormat];
    }

    // 尝试多步转换
    if (this.multiStepPaths[sourceFormat]?.[targetFormat]) {
      return this.multiStepPaths[sourceFormat][targetFormat];
    }

    // 通用转换路径（通过HTML作为中间格式）
    if (sourceFormat !== 'html' && targetFormat !== 'html') {
      // 检查是否可以通过HTML转换
      const canConvertToHtml =
        this.directConversions[sourceFormat]?.['html'] ||
        this.multiStepPaths[sourceFormat]?.['html'];
      const canConvertFromHtml =
        this.directConversions['html']?.[targetFormat] ||
        this.multiStepPaths['html']?.[targetFormat];

      if (canConvertToHtml && canConvertFromHtml) {
        return [sourceFormat, 'html', targetFormat];
      }
    }

    return [];
  }

  /**
   * 生成转换步骤
   */
  private generateConversionSteps(
    conversionPath: string[],
    request: ConversionRequest
  ): ConversionStep[] {
    const steps: ConversionStep[] = [];

    for (let i = 0; i < conversionPath.length - 1; i++) {
      const fromFormat = conversionPath[i];
      const toFormat = conversionPath[i + 1];
      const toolName = this.getToolForConversion(fromFormat, toFormat);
      const parameters = this.generateParameters(
        fromFormat,
        toFormat,
        request,
        i === conversionPath.length - 2
      );

      steps.push({
        stepNumber: i + 1,
        toolName,
        description: this.getStepDescription(fromFormat, toFormat, toolName),
        inputFormat: fromFormat,
        outputFormat: toFormat,
        parameters,
        estimatedTime: this.getEstimatedTime(fromFormat, toFormat),
        notes: this.getStepNotes(fromFormat, toFormat),
      });
    }

    return steps;
  }

  /**
   * 获取转换工具名称
   */
  private getToolForConversion(fromFormat: string, toFormat: string): string {
    // 特殊转换工具映射
    const specialMappings: Record<string, Record<string, string>> = {
      docx: {
        pdf: 'convert_docx_to_pdf',
        html: 'dual_parsing_docx_to_html',
      },
      markdown: {
        html: 'convert_markdown_to_html',
        docx: 'convert_markdown_to_docx',
      },
      html: {
        markdown: 'convert_html_to_markdown',
      },
    };

    if (specialMappings[fromFormat]?.[toFormat]) {
      return specialMappings[fromFormat][toFormat];
    }

    // 默认使用通用转换工具
    return 'convert_document';
  }

  /**
   * 生成工具参数
   */
  private generateParameters(
    fromFormat: string,
    toFormat: string,
    request: ConversionRequest,
    isFinalStep: boolean
  ): Record<string, any> {
    const params: Record<string, any> = {};

    // 根据工具类型设置参数
    const toolName = this.getToolForConversion(fromFormat, toFormat);

    switch (toolName) {
      case 'convert_markdown_to_html':
        params.theme = request.requirements?.theme || 'github';
        params.includeTableOfContents = false;
        break;

      case 'convert_markdown_to_docx':
        params.theme = request.requirements?.theme || 'professional';
        params.preserveStyles = request.requirements?.preserveStyles !== false;
        break;

      case 'convert_html_to_markdown':
        params.preserveStyles = request.requirements?.preserveStyles !== false;
        params.debug = false;
        break;

      case 'convert_docx_to_pdf':
        params.preserveFormatting = request.requirements?.preserveStyles !== false;
        break;

      case 'dual_parsing_docx_to_html':
        params.preserveStyles = request.requirements?.preserveStyles !== false;
        params.includeImages = request.requirements?.includeImages !== false;
        break;

      case 'convert_document':
        params.preserveFormatting = request.requirements?.preserveStyles !== false;
        break;
    }

    // 设置输入文件路径
    if (request.sourceFile) {
      const inputParam = this.getInputParameterName(toolName);
      params[inputParam] = request.sourceFile;
    }

    return params;
  }

  /**
   * 获取输入参数名称
   */
  private getInputParameterName(toolName: string): string {
    const mappings: Record<string, string> = {
      convert_markdown_to_html: 'markdownPath',
      convert_markdown_to_docx: 'markdownPath',
      convert_html_to_markdown: 'htmlPath',
      convert_docx_to_pdf: 'docxPath',
      dual_parsing_docx_to_html: 'docxPath',
      convert_document: 'inputPath',
    };

    return mappings[toolName] || 'inputPath';
  }

  /**
   * 获取步骤描述
   */
  private getStepDescription(fromFormat: string, toFormat: string, toolName: string): string {
    const descriptions: Record<string, string> = {
      convert_docx_to_pdf: '将DOCX文档转换为PDF格式，保持完整样式',
      dual_parsing_docx_to_html: '将DOCX文档转换为HTML格式，保留样式和图片',
      convert_markdown_to_html: '将Markdown文档转换为HTML格式，应用主题样式',
      convert_markdown_to_docx: '将Markdown文档转换为DOCX格式，应用专业样式',
      convert_html_to_markdown: '将HTML文档转换为Markdown格式，保留结构',
      convert_document: this.getConvertDocumentDescription(fromFormat, toFormat),
    };

    return descriptions[toolName] || `转换 ${fromFormat} 到 ${toFormat}`;
  }

  /**
   * 获取convert_document工具的详细描述
   */
  private getConvertDocumentDescription(fromFormat: string, toFormat: string): string {
    if (fromFormat === 'html' && toFormat === 'docx') {
      return '将HTML文档转换为DOCX格式，保留样式、结构和格式，生成完整的Word文档';
    }
    if (fromFormat === 'html' && toFormat === 'markdown') {
      return '将HTML文档转换为Markdown格式，保留文档结构和基本格式';
    }
    if (fromFormat === 'docx' && toFormat === 'markdown') {
      return '将DOCX文档转换为Markdown格式，使用优化转换器提取文本内容和结构，保留标题、段落、列表等格式';
    }
    if (fromFormat === 'docx' && toFormat === 'html') {
      return '将DOCX文档转换为HTML格式，使用增强型mammoth转换器保留样式、格式和结构，生成完整的HTML文档';
    }
    return `将${fromFormat.toUpperCase()}格式转换为${toFormat.toUpperCase()}格式，保持文档内容和基本结构`;
  }

  /**
   * 获取预估时间
   */
  private getEstimatedTime(fromFormat: string, toFormat: string): string {
    const timeMappings: Record<string, Record<string, string>> = {
      docx: {
        pdf: '30-60秒',
        html: '10-30秒',
        markdown: '10-20秒',
      },
      markdown: {
        html: '5-15秒',
        docx: '10-30秒',
      },
      html: {
        markdown: '5-15秒',
        docx: '15-30秒',
      },
      pdf: {
        html: '20-40秒',
        markdown: '20-40秒',
        txt: '10-20秒',
      },
    };

    return timeMappings[fromFormat]?.[toFormat] || '10-30秒';
  }

  /**
   * 获取步骤注意事项
   */
  private getStepNotes(fromFormat: string, toFormat: string): string | undefined {
    if (fromFormat === 'pdf') {
      return 'PDF转换可能会丢失部分格式信息';
    }

    if (toFormat === 'pdf' && fromFormat !== 'docx') {
      return '建议使用playwright-mcp完成最终的PDF生成';
    }

    return undefined;
  }

  /**
   * 计算总时间
   */
  private calculateTotalTime(steps: ConversionStep[]): string {
    if (steps.length === 0) return '0分钟';
    if (steps.length === 1) return steps[0].estimatedTime;

    const totalMinutes = steps.length * 0.5; // 平均每步30秒
    if (totalMinutes < 1) return '少于1分钟';
    if (totalMinutes < 2) return '1-2分钟';
    return `${Math.ceil(totalMinutes)}分钟左右`;
  }

  /**
   * 生成建议
   */
  private generateRecommendations(
    sourceFormat: string,
    targetFormat: string,
    request: ConversionRequest
  ): string[] {
    const recommendations: string[] = [];

    // 格式特定建议
    if (sourceFormat === 'docx' && targetFormat === 'pdf') {
      recommendations.push('建议使用convert_docx_to_pdf工具以获得最佳PDF质量');
      recommendations.push('转换过程会保持Word文档的完整样式和格式');
    }

    if (sourceFormat === 'html' && targetFormat === 'docx') {
      recommendations.push('HTML到DOCX转换将保留原始样式、结构和格式');
      recommendations.push('转换后的DOCX文件可以在Microsoft Word中正常打开和编辑');
      recommendations.push('建议启用样式保留以获得最佳转换效果');
    }

    if (sourceFormat === 'docx' && (targetFormat === 'markdown' || targetFormat === 'md')) {
      recommendations.push('DOCX到Markdown转换使用优化转换器，能够保留文档结构');
      recommendations.push('转换后的Markdown文件保留标题、段落、列表等基本格式');
      recommendations.push('复杂的样式和格式可能会简化为Markdown支持的格式');
    }

    if (sourceFormat === 'docx' && targetFormat === 'html') {
      recommendations.push('DOCX到HTML转换使用增强型mammoth转换器，完整保留样式和格式');
      recommendations.push('转换后的HTML文件包含完整的CSS样式，可在浏览器中正确显示');
      recommendations.push('建议启用样式保留和图片保存以获得最佳效果');
    }

    if (sourceFormat === 'markdown' && targetFormat === 'html') {
      recommendations.push('可以选择GitHub、学术或现代主题以获得更好的视觉效果');
      recommendations.push('建议启用样式保留以保持Markdown格式');
    }

    if (targetFormat === 'pdf') {
      recommendations.push('PDF转换可能需要额外的字体支持以正确显示中文');
    }

    if (sourceFormat === 'pdf') {
      recommendations.push('从PDF转换可能会丢失原始格式信息');
      recommendations.push('建议检查转换结果的准确性');
    }

    // 质量建议
    if (request.requirements?.quality === 'high') {
      recommendations.push('已选择高质量模式，转换时间可能较长但效果更好');
    }

    return recommendations;
  }

  /**
   * 生成警告
   */
  private generateWarnings(
    sourceFormat: string,
    targetFormat: string,
    conversionPath: string[]
  ): string[] {
    const warnings: string[] = [];

    if (conversionPath.length > 2) {
      warnings.push('此转换需要多个步骤，可能会有格式信息丢失');
    }

    if (sourceFormat === 'pdf' || targetFormat === 'pdf') {
      warnings.push('PDF格式转换可能需要额外的工具支持');
    }

    if (sourceFormat === 'html' && targetFormat === 'docx') {
      warnings.push('HTML到DOCX的转换可能无法完全保留所有CSS样式');
    }

    return warnings;
  }
}

export { ConversionPlanner, ConversionRequest, ConversionPlan, ConversionStep };
