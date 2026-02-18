import { cpSync, copyFileSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const thisFile = fileURLToPath(import.meta.url);
const root = resolve(dirname(thisFile), '..');
const dist = resolve(root, 'dist');

rmSync(dist, { recursive: true, force: true });
mkdirSync(dist, { recursive: true });

copyFileSync(resolve(root, 'index.html'), resolve(dist, 'index.html'));
copyFileSync(resolve(root, 'netlify.toml'), resolve(dist, 'netlify.toml'));
cpSync(resolve(root, 'core'), resolve(dist, 'core'), { recursive: true });
cpSync(resolve(root, 'modules'), resolve(dist, 'modules'), { recursive: true });

console.log('Build complete: dist');
