import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createFeishuOAuth, uploadFeishuDocument } from './feishu-oauth.mjs';
import { renderFeishuDocx, listTemplates } from './template-store.mjs';
import { recordUsage } from './usage-store.mjs';

const directory = path.join(process.env.FEIYE_FEISHU_OAUTH_DATA_DIR || fileURLToPath(new URL('../data/oauth', import.meta.url)), 'jobs');
const oauth = createFeishuOAuth();
const active = new Map();
const checking = new Map();
const escape = value => String(value).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const session = req => String(req.headers.authorization || '').match(/^Bearer ([A-Za-z0-9_-]{43})$/)?.[1] || '';
const jobFile = id => path.join(directory, `${id}.json`);
const fail = (message, status = 400) => Object.assign(new Error(message), { status });
const saveJob = async job => { const file = jobFile(job.id); await fs.writeFile(`${file}.tmp`, JSON.stringify(job), { mode: 0o600 }); await fs.rename(`${file}.tmp`, file); };
const publicJob = job => ({ id: job.id, stage: job.stage, name: job.name, error: job.error, url: job.url, ownerName: job.ownerName, warning: job.warning });
const validDocumentUrl = value => { try { const url = new URL(value); return url.protocol === 'https:' && /(^|\.)(feishu\.cn|larksuite\.com|larkoffice\.com)$/.test(url.hostname); } catch { return false; } };

async function runExport(job, input, login) {
  try {
    const output = await renderFeishuDocx(input.templateId, input.records);
    job.name = output.name; job.stage = 'uploading'; await saveJob(job);
    job.ticket = await uploadFeishuDocument(oauth, login, output.bytes, output.name);
    job.stage = 'importing'; await saveJob(job);
    // No record values or OAuth credentials are written to the job file.
  } catch (error) { job.stage = 'failed'; job.error = error.message; await saveJob(job); }
}

async function finishJob(job, login) {
  if (job.stage !== 'importing') return job;
  const result = (await oauth.api(`/drive/v1/import_tasks/${encodeURIComponent(job.ticket)}`, { token: await oauth.token(login) })).data?.result;
  if (!result || !Number.isInteger(result.job_status)) throw fail('飞书返回了无效导入状态', 502);
  if (result.job_status === 0) {
    if (!result.token || !validDocumentUrl(result.url)) throw fail('飞书未返回有效文档链接', 502);
    job.stage = 'complete'; job.url = result.url; job.documentId = result.token;
    job.warning = result.extra?.length ? '飞书转换对部分排版进行了调整，请打开文档检查。' : '';
    await saveJob(job);
    await recordUsage({ ...job.usage, format: 'feishu-doc', records: job.count }).catch(() => {});
    active.delete(job.owner);
  } else if (![1, 2].includes(result.job_status)) {
    job.stage = 'failed'; job.error = `飞书文档导入失败（${result.job_status}），请检查文件格式和应用权限`; await saveJob(job); active.delete(job.owner);
  }
  return job;
}

export async function handleFeishuDocs(req, res, pathname, url, { json, body }) {
  if (!pathname.startsWith('/api/oauth/feishu/') && !pathname.startsWith('/api/feishu-docs')) return false;
  try {
    if (pathname === '/api/oauth/feishu/start' && req.method === 'POST') {
      const input = JSON.parse((await body(req, 4096)).toString());
      json(res, 200, oauth.start(input.challenge)); return true;
    }
    if (pathname === '/api/oauth/feishu/authorize' && req.method === 'GET') {
      const result = oauth.authorize(url.searchParams.get('state'));
      res.writeHead(302, { location: result.url, 'set-cookie': result.cookie, 'cache-control': 'no-store', 'referrer-policy': 'no-referrer' }); res.end(); return true;
    }
    if (pathname === '/api/oauth/feishu/callback' && req.method === 'GET') {
      let message, ok = false;
      try { const name = await oauth.callback(url.searchParams.get('state'), url.searchParams.get('code'), req.headers.cookie, url.searchParams.get('error')); message = `已授权：${name}。请返回排版打印，导出会自动继续。`; ok = true; }
      catch (error) { message = error.message; }
      res.writeHead(ok ? 200 : 400, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store', 'referrer-policy': 'no-referrer', 'content-security-policy': "default-src 'none'; style-src 'unsafe-inline'; frame-ancestors 'none'" });
      res.end(`<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>飞书文档授权</title><body style="font:16px/1.7 sans-serif;max-width:540px;margin:12vh auto;padding:24px"><h1>${ok ? '授权完成' : '授权未完成'}</h1><p>${escape(message)}</p><p>可以关闭此窗口。</p></body></html>`); return true;
    }
    if (pathname === '/api/oauth/feishu/status' && req.method === 'POST') {
      const input = JSON.parse((await body(req, 4096)).toString()); json(res, 200, await oauth.claim(input.state, input.verifier)); return true;
    }
    const login = session(req);
    const identity = await oauth.identity(login);
    if (pathname === '/api/oauth/feishu/status' && req.method === 'GET') {
      await oauth.token(login); json(res, 200, { authorized: true, user: { name: identity.user.name, openId: identity.user.openId } }); return true;
    }
    if (pathname === '/api/oauth/feishu/disconnect' && req.method === 'POST') {
      await oauth.disconnect(login); json(res, 200, { disconnected: true }); return true;
    }
    if (pathname === '/api/feishu-docs' && req.method === 'POST') {
      const input = JSON.parse((await body(req, 2 * 1024 * 1024)).toString());
      if (!/^[a-f0-9-]{36}$/.test(input.requestId || '')) throw fail('缺少有效导出请求标识');
      if (!Array.isArray(input.records) || !input.records.length || input.records.length > 100) throw fail('请选择 1–100 条记录');
      if (!input.baseId || !input.tableId) throw fail('缺少当前多维表上下文');
      const template = (await listTemplates({ baseId: input.baseId, tableId: input.tableId })).find(item => item.id === input.templateId);
      if (!template) throw fail('当前数据表没有此模板', 404);
      if (!['.docx', '.txt', '.layout', '.json'].includes(template.extension)) throw fail('第一期仅支持 Word 和在线模板导出飞书文档');
      const id = crypto.createHash('sha256').update(`${identity.id}/${input.requestId}`).digest('hex');
      await fs.mkdir(directory, { recursive: true, mode: 0o700 });
      try { const existing = JSON.parse(await fs.readFile(jobFile(id))); json(res, 200, publicJob(existing)); return true; } catch (error) { if (error.code !== 'ENOENT') throw error; }
      if (active.has(identity.id)) throw fail('已有飞书文档正在生成，请等待完成后再导出', 409);
      const job = { id, owner: identity.id, ownerName: identity.user.name, stage: 'rendering', createdAt: Date.now(), count: input.records.length,
        usage: { userId: identity.user.openId, userName: identity.user.name, templateId: template.id, templateName: template.name, baseId: template.baseId, baseName: template.baseName, tableId: template.tableId, tableName: template.tableName } };
      active.set(identity.id, id);
      try { await saveJob(job); } catch (error) { active.delete(identity.id); throw error; }
      const release = setTimeout(() => { if (active.get(identity.id) === id) active.delete(identity.id); }, 10 * 60 * 1000);
      release.unref();
      void runExport(job, input, login).catch(() => {}).finally(() => { if (job.stage === 'failed') active.delete(identity.id); });
      json(res, 202, publicJob(job)); return true;
    }
    const match = pathname.match(/^\/api\/feishu-docs\/([a-f0-9]{64})$/);
    if (match && req.method === 'GET') {
      let job; try { job = JSON.parse(await fs.readFile(jobFile(match[1]))); } catch (error) { if (error.code === 'ENOENT') throw fail('导出任务不存在', 404); throw error; }
      if (job.owner !== identity.id) throw fail('无权查看此导出任务', 403);
      if (['rendering', 'uploading'].includes(job.stage) && active.get(identity.id) !== job.id) { job.stage = 'failed'; job.error = '服务重启中断了导出，请重新生成'; await saveJob(job); }
      if (Date.now() - job.createdAt > 10 * 60 * 1000 && !['complete', 'failed'].includes(job.stage)) { job.stage = 'failed'; job.error = '导入等待超时，请先检查个人云盘是否已有文档，再决定是否重试'; await saveJob(job); active.delete(identity.id); }
      if (!checking.has(job.id)) checking.set(job.id, finishJob(job, login).finally(() => checking.delete(job.id)));
      json(res, 200, publicJob(await checking.get(job.id))); return true;
    }
    json(res, 404, { error: '接口不存在' }); return true;
  } catch (error) { json(res, error.status || 400, { error: error.message || '飞书文档请求失败' }); return true; }
}
