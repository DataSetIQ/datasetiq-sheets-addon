#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'fs';

// Read the bundled Code.gs file
const code = readFileSync('Code.gs', 'utf-8');

// Find all globalThis assignments (pattern: g.functionName = implementationName;)
const assignmentPattern = /g\.(\w+)\s*=\s*(\w+);/g;
const assignments = [];
let match;

while ((match = assignmentPattern.exec(code)) !== null) {
  assignments.push({ name: match[1], impl: match[2] });
}

console.log(`Found ${assignments.length} function assignments`);

// Generate top-level function declarations that delegate to globalThis
// Custom functions need JSDoc comments for Sheets to recognize them
const declarations = assignments.map(({ name }) => {
  const isCustomFunction = name.startsWith('DSIQ');
  const comment = isCustomFunction ? `/**\n * DataSetIQ custom function: ${name}\n * @customfunction\n */\n` : '';
  return `${comment}function ${name}() {
  var args = Array.prototype.slice.call(arguments);
  return globalThis.${name}.apply(this, args);
}`;
}).join('\n\n');

// Don't modify the IIFE - keep it using globalThis
// Just append the top-level declarations
const finalCode = `${code}\n\n// Top-level function declarations for Apps Script\n${declarations}\n`;

writeFileSync('Code.gs', finalCode, 'utf-8');
console.log('✅ Generated top-level function declarations');
console.log(`Functions: ${assignments.map(a => a.name).join(', ')}`);
