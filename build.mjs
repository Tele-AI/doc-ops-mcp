#!/usr/bin/env node

import { build } from 'esbuild';
import { wasmLoader } from 'esbuild-plugin-wasm';
import { readFileSync, mkdirSync, existsSync, statSync } from 'fs';
import { resolve } from 'path';

// Safely read and parse package.json with size limit
let pkg;
try {
  const packagePath = './package.json';
  const stats = statSync(packagePath);
  
  // Limit package.json size to 1MB to prevent DoS
  if (stats.size > 1024 * 1024) {
    throw new Error('package.json file is too large');
  }
  
  const packageContent = readFileSync(packagePath, 'utf8');
  pkg = JSON.parse(packageContent);
} catch (error) {
  console.error('Error reading package.json:', error.message);
  process.exit(1);
}

// Safely ensure dist directory exists
try {
  const distPath = resolve('./dist');
  
  // Validate the path to prevent directory traversal
  if (!distPath.startsWith(resolve('./'))) {
    throw new Error('Invalid directory path');
  }
  
  if (!existsSync(distPath)) {
    // Create directory without deep recursion to prevent DoS
    mkdirSync(distPath, { recursive: false });
  }
} catch (error) {
  console.error('Error creating dist directory:', error.message);
  process.exit(1);
}

console.log('Building doc-ops-mcp...');

build({
  entryPoints: ['src/index.ts'],
  bundle: true,
  platform: 'node',
  target: 'node18',
  format: 'cjs',
  outfile: 'dist/index.cjs',
  external: [
    // Node.js 内置模块
    'fs',
    'path',
    'crypto',
    'util',
    'stream',
    'os',
    'url',
    'events',
    'buffer',
    'http',
    'https',
    'net',
    'tls',
    'zlib',
    'child_process',
    'worker_threads',
    'perf_hooks',
    'async_hooks',
    'readline',
    'tty',
    'cluster',
    'dns',
    'dgram',
    'module',
    // 有问题的 jsdom 依赖
    'jsdom',
    './xhr-sync-worker.js',
    // Puppeteer 相关依赖
    'puppeteer',
    'escalade/sync',
    'esutils',
    'jsbn',
  ],
  plugins: [wasmLoader()],
  sourcemap: false, // Disable source maps for production security
  minify: true, // 生产环境启用压缩
  keepNames: true,
  banner: {
    js: '#!/usr/bin/env node',
  },
  define: {
    'process.env.NODE_ENV': '"production"',
  },
  logLevel: 'warning', // Reduced debug output for production
})
  .then(async () => {
    console.log('✅ Build completed successfully');

    // Make the binary executable
    try {
      const { chmodSync } = await import('fs');
      chmodSync('dist/index.cjs', 0o755);
      console.log('✅ Binary permissions set');
    } catch (error) {
      console.warn('⚠️  Could not set binary permissions:', error.message);
    }
  })
  .catch(error => {
    console.error('❌ Build failed:', error.message || 'Unknown error');
    process.exit(1);
  });
