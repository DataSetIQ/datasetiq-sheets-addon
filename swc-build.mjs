#!/usr/bin/env node
import { transformFileSync } from '@swc/core';
import { writeFileSync } from 'fs';

const result = transformFileSync('src/Code.ts', {
  jsc: {
    parser: {
      syntax: 'typescript',
      tsx: false,
    },
    target: 'es2019',
  },
  module: {
    type: 'es6',
  },
});

// Post-process: Fix JSDoc comments that end up on same line as function
let code = result.code;
code = code.replace(/\*\/ function /g, '*/\nfunction ');

writeFileSync('Code.gs', code, 'utf-8');
console.log('✅ Transpiled Code.ts → Code.gs with SWC');
