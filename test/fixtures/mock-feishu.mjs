// Used only by the isolated HTTP integration test process.
import fs from 'node:fs/promises';
import path from 'node:path';
const appendFile = fs.appendFile;
fs.appendFile = (file, ...args) => appendFile(String(file).endsWith('usage-events.jsonl') ? path.join(process.env.FEIYE_FEISHU_OAUTH_DATA_DIR, 'test-usage.jsonl') : file, ...args);
const reply = value => new Response(JSON.stringify({ code: 0, ...value }), { headers: { 'content-type': 'application/json' } });
globalThis.fetch = async (url, options = {}) => {
  if (!String(url).startsWith('https://open.feishu.cn/open-apis/')) throw new Error('Unexpected mock origin');
  if (String(url).endsWith('/oauth/token')) {
    const input = JSON.parse(options.body);
    return reply({ access_token: `test-${input.code}`, expires_in: 7200, refresh_token: 'test-refresh', refresh_token_expires_in: 604800 });
  }
  if (String(url).endsWith('/user_info')) {
    const who = options.headers.authorization.includes('user-b') ? 'b' : 'a';
    return reply({ data: { open_id: `ou_${who}`, tenant_key: `tenant-${who}`, name: `用户${who}` } });
  }
  if (String(url).endsWith('/medias/upload_all')) {
    if (options.body.get('parent_type') !== 'ccm_import_open') throw new Error('Wrong parent_type');
    return reply({ data: { file_token: 'temporary-file' } });
  }
  if (String(url).endsWith('/import_tasks')) {
    const input = JSON.parse(options.body);
    if (input.point.mount_key !== '') throw new Error('Must target personal root');
    return reply({ data: { ticket: 'mock-ticket' } });
  }
  if (String(url).endsWith('/import_tasks/mock-ticket')) return reply({ data: { result: { job_status: 0, token: 'document', url: 'https://example.feishu.cn/docx/document' } } });
  throw new Error('Unexpected Feishu endpoint');
};
