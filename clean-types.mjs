#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'fs';

let code = readFileSync('Code.gs', 'utf-8');

// Remove all the problematic TypeScript constructs
code = code
  // Remove type/interface declarations (multi-line)
  .replace(/^type\s+\w+\s*=[\s\S]*?;/gm, '')
  .replace(/^interface\s+\w+\s*{[\s\S]*?^}/gm, '')
  // Remove optional parameter markers FIRST
  .replace(/\?\s*:/g, ':')
  .replace(/\?(\s*[,)])/g, '$1')
  // Remove all ': Type' annotations including union types (function params, variables, return types)
  .replace(/:\s*[\w\s|<>\[\].?]+\s*([,)=;{])/g, '$1')
  .replace(/:\s*\w+(\[\])?\s*$/gm, '')
  .replace(/:\s*{\s*[^}]+}\s*([),])/g, '$1')
  .replace(/:\ GoogleAppsScript[^\s]*/g, '')
  // Remove 'as Type' assertions
  .replace(/\s+as\s+\w+/g, '')
  // Remove export statements
  .replace(/^export\s*{[^}]*};?\s*$/gm, '')
  // Clean up extra blank lines
  .replace(/\n\n\n+/g, '\n\n');

writeFileSync('Code.gs', code, 'utf-8');
console.log('✅ Cleaned TypeScript syntax from Code.gs');
