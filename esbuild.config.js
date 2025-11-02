// esbuild.config.js
const { build } = require('esbuild');

build({
  entryPoints: ['code.ts'],
  bundle: true,
  outfile: 'code.js',
  platform: 'node',
  format: 'iife', // Output as a global script for Figma
  sourcemap: false,
  minify: true,
  target: ['es6'],
  external: [],
}).catch(() => process.exit(1));
