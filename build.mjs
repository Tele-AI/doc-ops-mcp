#!/usr/bin/env node

import { build } from 'esbuild';
import { wasmLoader } from 'esbuild-plugin-wasm';
import { promises as fs } from 'fs';  // 使用异步 fs
import { resolve } from 'path';

// 异步读取和解析 package.json
let pkg;
try {
  const packagePath = './package.json';
  const stats = await fs.stat(packagePath);  // 改为异步
  
  // 限制文件大小防止 DoS
  if (stats.size > 1024 * 1024) {
    throw new Error('package.json file is too large');
  }
  
  const packageContent = await fs.readFile(packagePath, 'utf8');  // 改为异步
  pkg = JSON.parse(packageContent);
} catch (error) {
  process.exit(1);
}

// 异步确保 dist 目录存在
try {
  const distPath = resolve('./dist');
  
  // 验证路径安全性
  if (!distPath.startsWith(resolve('./'))) {
    throw new Error('Invalid directory path');
  }
  
  // 使用异步操作，支持递归创建
  await fs.mkdir(distPath, { recursive: true });  // 改为异步且 recursive: true
} catch (error) {
  process.exit(1);
}

// 生产环境移除调试输出

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
      // 生产环境移除调试输出
    } catch (error) {
      // 生产环境移除调试输出
    }
  })
  .catch(error => {
    // 生产环境移除调试输出
    process.exit(1);
  });
