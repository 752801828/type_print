import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { listTemplates, saveTemplate, renameTemplate, removeTemplate, renderTemplate, previewTemplate, previewRecord, templateFile, outputFile } from './lib/template-store.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(here, 'public');
const sdkEntry = createRequire(import.meta.url).resolve('@lark-base-open/js-sdk');
const sdkDir = path.dirname(sdkEntry);
const port = Number(process.env.PORT || 4318);
const host = process.env.HOST || '0.0.0.0';
const build = 'feiye-independent';
const mime = { '.html':'text/html; charset=utf-8', '.js':'text/javascript; charset=utf-8', '.mjs':'text/javascript; charset=utf-8', '.css':'text/css; charset=utf-8', '.json':'application/json; charset=utf-8', '.txt':'text/plain; charset=utf-8', '.layout':'application/json; charset=utf-8', '.svg':'image/svg+xml', '.doc':'application/msword', '.docx':'application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.xls':'application/vnd.ms-excel', '.pdf':'application/pdf', '.zip':'application/zip' };

const send = (res, status, body, type = 'application/json; charset=utf-8') => { res.writeHead(status, { 'content-type': type, 'cache-control': 'no-store', 'x-feiye-build': build, 'x-content-type-options': 'nosniff', 'referrer-policy': 'no-referrer', 'content-security-policy': "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors https://feishu.cn https://*.feishu.cn https://larksuite.com https://*.larksuite.com http://localhost:* http://127.0.0.1:*" }); res.end(body); };
const json = (res, status, value) => send(res, status, JSON.stringify(value));
const body = async (req, limit = 25 * 1024 * 1024) => {
  const chunks = []; let size = 0;
  for await (const chunk of req) { size += chunk.length; if (size > limit) throw new Error('请求体过大'); chunks.push(chunk); }
  return Buffer.concat(chunks);
};
const safe = (base, relative) => { const target = path.resolve(base, relative); return target === base || target.startsWith(`${base}${path.sep}`) ? target : null; };
const serve = async (req, res, pathname) => {
  const isVendor = pathname.startsWith('/vendor/lark-base/');
  const relative = isVendor ? pathname.slice('/vendor/lark-base/'.length) : pathname === '/' || pathname === '/feishu' ? 'index.html' : pathname.slice(1);
  const target = safe(isVendor ? sdkDir : publicDir, relative);
  if (!target || (!isVendor && relative.includes('..'))) return send(res, 404, 'Not found', 'text/plain; charset=utf-8');
  try { const data = await fs.readFile(target); send(res, 200, data, mime[path.extname(target).toLowerCase()] || 'application/octet-stream'); }
  catch { send(res, 404, 'Not found', 'text/plain; charset=utf-8'); }
};

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const pathname = url.pathname === '/feishu' ? '/' : url.pathname.startsWith('/feishu/') ? url.pathname.slice('/feishu'.length) : url.pathname;
  try {
    if (pathname === '/api/health' && req.method === 'GET') return json(res, 200, { ok: true, name: '排版打印 · 飞书扩展', port, build });
    if (pathname === '/api/templates' && req.method === 'GET') return json(res, 200, { templates: await listTemplates({ baseId: url.searchParams.get('baseId') || '', tableId: url.searchParams.get('tableId') || '' }) });
    if (pathname === '/api/templates' && req.method === 'POST') {
      let fileName = req.headers['x-file-name'] || 'template.docx';
      try { fileName = decodeURIComponent(fileName); } catch {}
      const header = name => { const value = req.headers[name] || ''; try { return decodeURIComponent(value); } catch { return value; } };
      const item = await saveTemplate(fileName, await body(req), { baseId: header('x-base-id'), tableId: header('x-table-id'), viewId: header('x-view-id'), baseName: header('x-base-name'), tableName: header('x-table-name') });
      return json(res, 201, { template: item });
    }
    const templateMatch = pathname.match(/^\/api\/templates\/([\w-]+)$/);
    if (templateMatch && req.method === 'PATCH') { const input = JSON.parse((await body(req, 64 * 1024)).toString('utf8')); return json(res, 200, { template: await renameTemplate(templateMatch[1], input.name) }); }
    if (templateMatch && req.method === 'DELETE') return json(res, 200, { deleted: await removeTemplate(templateMatch[1]) });
    const templateFileMatch = pathname.match(/^\/api\/templates\/([\w-]+)\/file$/);
    if (templateFileMatch && req.method === 'GET') {
      const templates = await listTemplates(); const item = templates.find(entry => entry.id === templateFileMatch[1]);
      if (!item) return json(res, 404, { error: 'TEMPLATE_NOT_FOUND' });
      const extension = item.extension || path.extname(item.fileName); const file = await fs.readFile(templateFile(item.id, extension)); res.writeHead(200, { 'content-type': mime[extension] || 'application/octet-stream', 'content-disposition': `inline; filename="${encodeURIComponent(item.name)}"`, 'cache-control': 'no-store' }); return res.end(file);
    }
    const templatePreviewMatch = pathname.match(/^\/api\/templates\/([\w-]+)\/preview$/);
    if (templatePreviewMatch && req.method === 'GET') return json(res, 200, await previewTemplate(templatePreviewMatch[1]));
    const recordPreviewMatch = pathname.match(/^\/api\/templates\/([\w-]+)\/record-preview$/);
    if (recordPreviewMatch && req.method === 'POST') { const input = JSON.parse((await body(req, 2 * 1024 * 1024)).toString('utf8')); return json(res, 200, await previewRecord(recordPreviewMatch[1], input.record || {})); }
    if (pathname === '/api/generate-docx' && req.method === 'POST') {
      const input = JSON.parse((await body(req, 2 * 1024 * 1024)).toString('utf8'));
      return json(res, 201, { output: await renderTemplate(input.templateId, input.records, input.outputFormat) });
    }
    const outputMatch = pathname.match(/^\/api\/outputs\/([\w-]+)\/download$/);
    if (outputMatch && req.method === 'GET') {
      let extension = 'docx'; let file;
      for (const candidate of ['doc', 'docx', 'xlsx', 'xls', 'pdf', 'zip']) { try { file = await fs.readFile(outputFile(outputMatch[1], candidate)); extension = candidate; break; } catch {} }
      if (!file) return json(res, 404, { error: 'OUTPUT_NOT_FOUND' });
      const type = mime[`.${extension}`] || 'application/octet-stream';
      res.writeHead(200, { 'content-type': type, 'content-disposition': `attachment; filename="feiye-${outputMatch[1]}.${extension}"`, 'cache-control': 'no-store' });
      return res.end(file);
    }
    if (pathname.startsWith('/api/')) return json(res, 404, { error: 'NOT_FOUND' });
    return serve(req, res, pathname);
  } catch (error) { return json(res, 400, { error: error.message || '请求失败' }); }
});

server.listen(port, host, () => console.log(`飞页已启动：http://${host}:${port}`));
