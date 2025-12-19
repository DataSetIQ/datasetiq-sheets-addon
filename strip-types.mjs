#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'fs';

// Read the TypeScript source
const source = readFileSync('src/Code.ts', 'utf-8');

// Split into lines for better control
const lines = source.split('\n');
const output = [];
let skipUntil = null;

for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  const trimmed = line.trim();
  
  // Skip type/interface declarations (including multi-line)
  if (trimmed.startsWith('type ') || trimmed.startsWith('interface ')) {
    // Skip until we find the closing semicolon or brace
    if (trimmed.startsWith('type ')) {
      skipUntil = ';';
    } else {
      skipUntil = '}';
    }
    if (line.includes(skipUntil)) {
      skipUntil = null; // Single line declaration
    }
    continue;
  }
  
  if (skipUntil) {
    if (line.includes(skipUntil)) {
      skipUntil = null;
    }
    continue;
  }
  
  // Skip export statements
  if (trimmed.startsWith('export {') || trimmed.startsWith('export type') || trimmed.startsWith('import type')) {
    continue;
  }
  
  // Process the line to remove inline type annotations
  let processed = line;
  
  // Remove type annotations from function parameters: (param: Type) -> (param)
  processed = processed.replace(/(\w+):\s*[A-Za-z<>[\]|{}&\s,?]+(\s*[=),])/g, '$1$2');
  
  // Remove return type annotations: ): Type { -> ) {
  processed = processed.replace(/\):\s*[A-Za-z<>[\]|{}&\s,?]+(\s*{)/g, ')$1');
  
  // Remove type assertions: as Type
  processed = processed.replace(/\s+as\s+\w+/g, '');
  
  // Remove generics: <Type>
  processed = processed.replace(/<[A-Za-z,\s]+>/g, '');
  
  // Remove variable type annotations: const x: Type = -> const x =
  processed = processed.replace(/(const|let|var)\s+(\w+):\s*[A-Za-z<>[\]|{}&\s,?]+(\s*=)/g, '$1 $2$3');
  
  output.push(processed);
}

writeFileSync('Code.gs', output.join('\n'), 'utf-8');
console.log('✅ Stripped types from Code.ts → Code.gs');
