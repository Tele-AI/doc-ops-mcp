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
  
  const timestamp = Date.now();
  const randomSuffix = generateSecureRandomSuffix();
  const filename = `${prefix}-${timestamp}-${randomSuffix}${extension}`;
  
  return path.join(tempDir, filename);
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