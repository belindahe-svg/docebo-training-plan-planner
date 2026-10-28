// Optional: bundles src/worker.js + src/app.html into one dist/worker.js
// for pasting into the Cloudflare dashboard editor. Not needed for the
// GitHub deploy, where wrangler bundles the HTML itself.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
const html = readFileSync('src/app.html', 'utf8');
const src = readFileSync('src/worker.js', 'utf8');
mkdirSync('dist', { recursive: true });
const xlsx = readFileSync('src/vendor/xlsx.full.min.txt', 'utf8');
writeFileSync('dist/worker.js', src
  .replace("import APP_HTML from './app.html';", () => `const APP_HTML = ${JSON.stringify(html)};`)
  .replace("import XLSX_JS from './vendor/xlsx.full.min.txt';", () => `const XLSX_JS = ${JSON.stringify(xlsx)};`));
console.log('built dist/worker.js');
