// security/errorHandler.ts
export class SafeErrorHandler {
  private static isDevelopment = process.env.NODE_ENV === 'development';
  
  // 清理敏感信息的错误消息
  static sanitizeErrorMessage(error: any): string {
    if (this.isDevelopment) {
      return error.message; // 开发环境显示完整信息
    }
    
    // 生产环境：移除敏感信息
    let message = error.message || 'An error occurred';
    
    // 移除文件路径
    message = message.replace(/\/[^\s]+/g, '[path]');
    
    // 移除用户名
    message = message.replace(/\/home\/[^\/]+/g, '/home/[user]');
    message = message.replace(/\/Users\/[^\/]+/g, '/Users/[user]');
    
    // 移除 IP 地址
    message = message.replace(/\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b/g, '[ip]');
    
    // 移除端口号
    message = message.replace(/:\d{2,5}/g, ':[port]');
    
    return message;
  }
  
  // 记录错误但不泄露敏感信息
  static logError(context: string, error: any): void {
    const safeMessage = this.sanitizeErrorMessage(error);
    
    if (this.isDevelopment) {
      console.error(`${context}: ${safeMessage}`);
      console.error('Stack:', error.stack);
    } else {
      // 生产环境：只记录安全的信息
      console.error(`${context}: ${safeMessage}`);
      // 可以将完整错误发送到内部日志系统
      this.sendToInternalLogging(context, error);
    }
  }
  
  private static sendToInternalLogging(context: string, error: any): void {
    // 实现内部日志记录，不对外暴露
    // 例如：发送到日志服务器、写入安全的日志文件等
  }
}