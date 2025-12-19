// Simple build script to remove TypeScript types for Google Apps Script
const fs = require('fs');

let code = fs.readFileSync('src/Code.ts', 'utf8');

// Remove type annotations
code = code
  // Remove type imports
  .replace(/^type .+$/gm, '')
  // Remove interface declarations  
  .replace(/^interface .+\{[\s\S]*?\}/gm, '')
  // Remove type aliases
  .replace(/^export type .+$/gm, '')
  // Remove function return types
  .replace(/\):\s*\w+(\[\])?(\s*\{)/g, ')$2')
  // Remove parameter types
  .replace(/(\w+):\s*[\w\[\]<>|{}?]+(\s*[,)])/g, '$1$2')
  // Remove variable types
  .replace(/:\s*[\w\[\]<>|{}?]+(\s*=)/g, '$1')
  // Remove as type assertions
  .replace(/\s+as\s+\w+/g, '')
  // Clean up exports
  .replace(/^export \{[\s\S]*?\};?$/gm, '');

fs.writeFileSync('Code.gs', code, 'utf8');
console.log('✅ Built Code.gs successfully');
