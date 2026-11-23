// Builds the static site into dist/: inlines the strata-kit palette into src/index.html and copies
// the workbooks and CNAME beside it. No dependencies, no network.
import { cpSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';

const kit = readFileSync('vendor/strata-kit/theme.css', 'utf8');
const html = readFileSync('src/index.html', 'utf8');
if (!html.includes('/*__KIT_CSS__*/')) throw new Error('src/index.html is missing the /*__KIT_CSS__*/ marker');

rmSync('dist', { recursive: true, force: true });
mkdirSync('dist', { recursive: true });
writeFileSync('dist/index.html', html.replace('/*__KIT_CSS__*/', () => kit));
cpSync('files', 'dist/files', { recursive: true });
cpSync('CNAME', 'dist/CNAME');
console.log('built dist/');
