import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const output = path.join(root, 'dist');
const sdkDir = path.dirname(createRequire(import.meta.url).resolve('@lark-base-open/js-sdk'));

await fs.rm(output, { recursive: true, force: true });
await fs.cp(path.join(root, 'public'), output, { recursive: true });
await fs.mkdir(path.join(output, 'vendor'), { recursive: true });
await fs.cp(sdkDir, path.join(output, 'vendor', 'lark-base'), { recursive: true });

const indexPath = path.join(output, 'index.html');
const index = (await fs.readFile(indexPath, 'utf8'))
  .replace('href="/feishu/styles.css"', 'href="./styles.css"')
  .replace('src="/feishu/app.js"', 'src="./app.js"');
await fs.writeFile(indexPath, index, 'utf8');
console.log('Built dist/');
