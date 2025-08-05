#!/usr/bin/env node

import { build } from 'esbuild';
import { wasmLoader } from 'esbuild-plugin-wasm';
import { readFileSync, mkdirSync, existsSync } from 'fs';

const pkg = JSON.parse(readFileSync('./package.json', 'utf8'));

// Ensure dist directory exists
if (!existsSync('dist')) {
  mkdirSync('dist', { recursive: true });
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
  sourcemap: false,
  minify: false,
  keepNames: true,
  banner: {
    js: '#!/usr/bin/env node',
  },
  define: {
    'process.env.NODE_ENV': '"production"',
  },
  logLevel: 'info',
})
  .then(async () => {
    console.log('✅ Build completed successfully');

    // Make the binary executable
    try {
      const { execSync } = await import('child_process');
      execSync('chmod +x dist/index.cjs');
      console.log('✅ Binary permissions set');
    } catch (error) {
      console.warn('⚠️  Could not set binary permissions:', error.message);
    }
  })
  .catch(error => {
    console.error('❌ Build failed:', error);
    process.exit(1);
  });
