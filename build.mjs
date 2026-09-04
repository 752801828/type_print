import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const output = path.join(root, 'dist');
const sdkDir = path.dirname(createRequire(import.meta.url).resolve('@lark-base-open/js-sdk'));
const docxPreviewDir = path.dirname(createRequire(import.meta.url).resolve('docx-preview'));
const jszipDir = path.dirname(path.dirname(createRequire(import.meta.url).resolve('jszip')));

await fs.rm(output, { recursive: true, force: true });
await fs.cp(path.join(root, 'public'), output, { recursive: true });
await fs.mkdir(path.join(output, 'vendor'), { recursive: true });
await fs.cp(sdkDir, path.join(output, 'vendor', 'lark-base'), { recursive: true });
await fs.copyFile(path.join(docxPreviewDir, 'docx-preview.min.js'), path.join(output, 'vendor', 'docx-preview.min.js'));
await fs.copyFile(path.join(jszipDir, 'dist', 'jszip.min.js'), path.join(output, 'vendor', 'jszip.min.js'));

const indexPath = path.join(output, 'index.html');
const index = (await fs.readFile(indexPath, 'utf8'))
  .replace('href="/feishu/styles.css"', 'href="./styles.css"')
  .replace('src="/feishu/vendor/jszip.min.js"', 'src="./vendor/jszip.min.js"')
  .replace('src="/feishu/vendor/docx-preview.min.js"', 'src="./vendor/docx-preview.min.js"')
  .replace('src="/feishu/app.js"', 'src="./app.js"');
await fs.writeFile(indexPath, index, 'utf8');
console.log('Built dist/');
