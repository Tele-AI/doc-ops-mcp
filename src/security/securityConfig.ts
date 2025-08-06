/**
 * Security Configuration for doc-ops-mcp
 * Centralized security settings and utilities
 */

import * as os from 'os';
import * as path from 'path';
import * as crypto from 'crypto';

export interface SecurityConfig {
  tempFilePermissions: number;
  maxTempFileAge: number; // in milliseconds
  allowedTempDirs: string[];
  secureRandomLength: number;
}

export const defaultSecurityConfig: SecurityConfig = {
  tempFilePermissions: 0o600, // Owner read/write only
  maxTempFileAge: 3600000, // 1 hour
  allowedTempDirs: [
    process.env.TEMP_DIR || os.tmpdir(),
    '/tmp',
    path.join(process.cwd(), 'temp')
  ],
  secureRandomLength: 16
};

/**
 * Generate a secure random filename suffix
 */
export function generateSecureRandomSuffix(length: number = defaultSecurityConfig.secureRandomLength): string {
  return crypto.randomBytes(length).toString('hex');
}

/**
 * Validate if a directory is allowed for temporary files
 */
export function isAllowedTempDir(dirPath: string): boolean {
  const normalizedPath = path.resolve(dirPath);
  return defaultSecurityConfig.allowedTempDirs.some(allowedDir => 
    normalizedPath.startsWith(path.resolve(allowedDir))
  );
}

/**
 * Create a secure temporary file path
 */
export function createSecureTempPath(prefix: string, extension: string = '.tmp'): string {
  const tempDir = process.env.TEMP_DIR || os.tmpdir();
  
  if (!isAllowedTempDir(tempDir)) {
    throw new Error(`Temporary directory not allowed: ${tempDir}`);
  }
  
  const randomSuffix = generateSecureRandomSuffix();
  const filename = `${prefix}_${randomSuffix}${extension}`;
  
  return path.join(tempDir, filename);
}

/**
 * Validate and sanitize file path to prevent path traversal attacks
 */
export function validateAndSanitizePath(inputPath: string, allowedBasePaths: string[] = []): string {
  if (!inputPath || typeof inputPath !== 'string') {
    throw new Error('Invalid path: path must be a non-empty string');
  }

  // Remove null bytes and other dangerous characters
  const sanitized = inputPath.replace(/[\x00-\x1f\x7f-\x9f]/g, '');
  
  // Resolve the path to handle .. and . components
  const resolvedPath = path.resolve(sanitized);
  
  // Check for path traversal attempts
  if (sanitized.includes('..') || sanitized.includes('\x00')) {
    throw new Error('Path traversal attempt detected');
  }
  
  // If allowed base paths are specified, ensure the resolved path is within them
  if (allowedBasePaths.length > 0) {
    const isAllowed = allowedBasePaths.some(basePath => {
      const resolvedBasePath = path.resolve(basePath);
      return resolvedPath.startsWith(resolvedBasePath + path.sep) || resolvedPath === resolvedBasePath;
    });
    
    if (!isAllowed) {
      throw new Error(`Path outside allowed directories: ${resolvedPath}`);
    }
  }
  
  return resolvedPath;
}

/**
 * Safe path join that prevents path traversal
 */
export function safePathJoin(basePath: string, ...segments: string[]): string {
  const allowedBasePaths = [path.resolve(basePath)];
  
  // Validate base path
  const validatedBasePath = validateAndSanitizePath(basePath);
  
  // Join all segments
  const joinedPath = path.join(validatedBasePath, ...segments);
  
  // Validate the final path
  return validateAndSanitizePath(joinedPath, allowedBasePaths);
}

/**
 * HTML entity encoding for XSS prevention
 */
export function escapeHtml(unsafe: string): string {
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * CSS property sanitization
 */
export function sanitizeCssProperty(property: string): string {
  // Remove potentially dangerous CSS properties and values
  return property
    .replace(/javascript:/gi, '')
    .replace(/vbscript:/gi, '')
    .replace(/data:/gi, '')
    .replace(/expression\s*\(/gi, '')
    .replace(/@import/gi, '')
    .replace(/url\s*\(/gi, 'url(')
    .trim();
}

/**
 * Validate file extension against allowed types
 */
export function isAllowedFileExtension(filename: string, allowedExtensions: string[]): boolean {
  const ext = path.extname(filename).toLowerCase();
  return allowedExtensions.includes(ext);
}

/**
 * Security headers for HTTP responses
 */
export const securityHeaders = {
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'X-XSS-Protection': '1; mode=block',
  'Strict-Transport-Security': 'max-age=31536000; includeSubDomains',
  'Content-Security-Policy': "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; font-src 'self'"
};