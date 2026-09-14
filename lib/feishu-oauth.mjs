import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const digest = value => crypto.createHash('sha256').update(value).digest('hex');
const random = () => crypto.randomBytes(32).toString('base64url');
const fail = (message, status = 400) => Object.assign(new Error(message), { status });
const validSecret = value => typeof value === 'string' && /^[A-Za-z0-9_-]{43}$/.test(value);
const defaultDir = process.env.FEIYE_FEISHU_OAUTH_DATA_DIR || fileURLToPath(new URL('../data/oauth', import.meta.url));

// One Node service owns this store. Multiple service workers require a shared transactional store.
export function createFeishuOAuth({ env = process.env, directory = defaultDir, request = globalThis.fetch, now = Date.now } = {}) {
  const pending = new Map();
  const refreshing = new Map();
  let storage;
  let queue = Promise.resolve();
  const config = () => {
    const appId = env.FEIYE_FEISHU_APP_ID;
    const secret = env.FEIYE_FEISHU_APP_SECRET;
    const redirect = env.FEIYE_FEISHU_OAUTH_REDIRECT_URI;
    if (!appId || !secret || !redirect) throw fail('服务器尚未配置飞书 OAuth，请联系管理员', 503);
    const parsed = new URL(redirect);
    if (parsed.protocol !== 'https:' || parsed.pathname !== '/feishu/api/oauth/feishu/callback') throw fail('OAuth 回调地址配置无效', 503);
    return { appId, secret, redirect, origin: parsed.origin };
  };
  const init = () => storage ||= (async () => {
    await fs.mkdir(directory, { recursive: true, mode: 0o700 });
    const keyPath = path.join(directory, 'key');
    try { await fs.writeFile(keyPath, crypto.randomBytes(32), { flag: 'wx', mode: 0o600 }); } catch (error) { if (error.code !== 'EEXIST') throw error; }
    const key = await fs.readFile(keyPath);
    if (key.length !== 32) throw fail('OAuth 密钥文件损坏，请联系管理员', 503);
    let db = { users: {}, sessions: {} };
    try {
      const raw = await fs.readFile(path.join(directory, 'tokens.enc'));
      const decipher = crypto.createDecipheriv('aes-256-gcm', key, raw.subarray(0, 12));
      decipher.setAuthTag(raw.subarray(12, 28));
      db = JSON.parse(Buffer.concat([decipher.update(raw.subarray(28)), decipher.final()]).toString());
    } catch (error) { if (error.code !== 'ENOENT') throw fail('OAuth 令牌文件无法解密，请联系管理员', 503); }
    return { key, db };
  })();
  const mutate = fn => {
    const task = queue.then(async () => {
      const { key, db } = await init();
      const next = structuredClone(db);
      const result = fn(next);
      const iv = crypto.randomBytes(12);
      const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
      const data = Buffer.concat([cipher.update(JSON.stringify(next)), cipher.final()]);
      const temp = path.join(directory, 'tokens.enc.tmp');
      await fs.writeFile(temp, Buffer.concat([iv, cipher.getAuthTag(), data]), { mode: 0o600 });
      await fs.rename(temp, path.join(directory, 'tokens.enc'));
      (await init()).db = next;
      return result;
    });
    queue = task.catch(() => {});
    return task;
  };
  async function api(endpoint, { token, method = 'GET', json, form } = {}) {
    let response;
    try {
      response = await request(`https://open.feishu.cn/open-apis${endpoint}`, {
        method, headers: { ...(token ? { authorization: `Bearer ${token}` } : {}), ...(json ? { 'content-type': 'application/json' } : {}) },
        body: json ? JSON.stringify(json) : form, signal: AbortSignal.timeout(45000), redirect: 'error'
      });
    } catch { throw fail('连接飞书超时或网络异常，请稍后重试', 502); }
    let value;
    try { value = await response.json(); } catch { throw fail('飞书返回了无效响应', 502); }
    if (!response.ok || (value.code != null && value.code !== 0) || value.error) {
      const code = value.code || value.error || response.status;
      const expired = [99991663, 99991668, 99991671, 20003, 20004, 20024, 20026].includes(Number(code)) || code === 'invalid_grant';
      throw fail(expired ? '飞书授权已失效，请重新授权' : `飞书接口失败（${String(code).replace(/[^\w-]/g, '').slice(0, 30)}），请检查应用的云文档导入/上传权限是否已发布`, expired ? 401 : 502);
    }
    return value;
  }
  const tokenData = (value, previous = {}) => {
    if (!value.access_token || !Number.isFinite(value.expires_in) || value.expires_in <= 0) throw fail('飞书未返回有效访问令牌', 502);
    return { ...previous, accessToken: value.access_token, accessExpires: now() + value.expires_in * 1000,
      refreshToken: value.refresh_token || previous.refreshToken || '',
      refreshExpires: value.refresh_token_expires_in ? now() + value.refresh_token_expires_in * 1000 : previous.refreshExpires || 0,
      scope: value.scope || previous.scope || '' };
  };
  function start(challenge) {
    const { origin } = config();
    if (!/^[a-f0-9]{64}$/.test(challenge || '')) throw fail('无效授权请求');
    for (const [id, item] of pending) if (item.expires < now()) pending.delete(id);
    if (pending.size >= 128) throw fail('授权请求过多，请稍后再试', 429);
    const state = random();
    pending.set(state, { challenge, expires: now() + 10 * 60 * 1000 });
    return { state, url: `${origin}/feishu/api/oauth/feishu/authorize?state=${state}` };
  }
  function pendingItem(state) {
    const item = pending.get(state);
    if (!item || item.expires < now()) { pending.delete(state); throw fail('授权请求已过期，请重新点击导出', 401); }
    return item;
  }
  function authorize(state) {
    const item = pendingItem(state);
    if (item.cookie || item.used) throw fail('授权链接已使用，请重新点击导出', 401);
    item.cookie = random();
    const { appId, redirect } = config();
    const url = new URL('https://accounts.feishu.cn/open-apis/authen/v1/authorize');
    url.search = new URLSearchParams({ client_id: appId, response_type: 'code', redirect_uri: redirect, state,
      scope: env.FEIYE_FEISHU_OAUTH_SCOPES || 'offline_access docs:document:import docs:document.media:upload' }).toString();
    return { url: url.href, cookie: `feiye_oauth_${state}=${item.cookie}; Path=/feishu/api/oauth/feishu; HttpOnly; Secure; SameSite=Lax; Max-Age=600` };
  }
  async function callback(state, code, cookieHeader, denied) {
    const item = pendingItem(state);
    const cookies = Object.fromEntries(String(cookieHeader || '').split(';').map(part => part.trim().split('=')));
    if (!item.cookie || cookies[`feiye_oauth_${state}`] !== item.cookie || item.used) throw fail('授权会话校验失败，请回插件重新授权', 401);
    item.used = true;
    try {
      if (denied || !code || code.length > 2048) throw fail('已取消飞书授权');
      const { appId, secret, redirect } = config();
      const result = await api('/authen/v2/oauth/token', { method: 'POST', json: { grant_type: 'authorization_code', client_id: appId, client_secret: secret, code, redirect_uri: redirect } });
      const profile = (await api('/authen/v1/user_info', { token: result.access_token })).data;
      if (!profile?.open_id || !profile?.tenant_key) throw fail('飞书未返回用户及租户身份', 502);
      const user = { ...tokenData(result), openId: profile.open_id, tenantKey: profile.tenant_key, name: String(profile.name || profile.en_name || '飞书用户').slice(0, 100), appId };
      const id = digest(`${appId}/${user.tenantKey}/${user.openId}`);
      // Never bind OAuth credentials to an unverified browser-supplied user ID.
      await mutate(db => { db.users[id] = user; });
      item.userId = id;
      return user.name;
    } catch (error) { item.error = error.message; throw error; }
  }
  async function claim(state, verifier) {
    const item = pendingItem(state);
    if (!validSecret(verifier) || digest(verifier) !== item.challenge) throw fail('授权校验失败', 401);
    if (item.claimed) throw fail('授权结果已领取', 401);
    if (item.error) { pending.delete(state); throw fail(item.error, 401); }
    if (!item.userId) return { pending: true };
    item.claimed = true;
    const session = random();
    const user = (await init()).db.users[item.userId];
    try {
      await mutate(db => {
        for (const [id, entry] of Object.entries(db.sessions)) if (entry.expires < now()) delete db.sessions[id];
        db.sessions[digest(session)] = { userId: item.userId, expires: now() + 30 * 86400000 };
      });
    } catch (error) { item.claimed = false; throw error; }
    pending.delete(state);
    return { session, user: { name: user.name, openId: user.openId } };
  }
  async function identity(session) {
    config();
    if (!validSecret(session)) throw fail('请先授权飞书文档', 401);
    await queue;
    const { db } = await init();
    const login = db.sessions[digest(session)];
    const user = login && db.users[login.userId];
    if (!login || login.expires <= now() || !user || user.appId !== config().appId) throw fail('飞书授权已过期，请重新授权', 401);
    return { id: login.userId, user };
  }
  async function token(session) {
    const { id, user } = await identity(session);
    if (user.accessExpires > now() + 60000) return user.accessToken;
    if (refreshing.has(id)) return refreshing.get(id);
    const task = (async () => {
      if (!user.refreshToken || user.refreshExpires <= now()) throw fail('飞书授权已过期，请重新授权', 401);
      const { appId, secret } = config();
      const result = await api('/authen/v2/oauth/token', { method: 'POST', json: { grant_type: 'refresh_token', client_id: appId, client_secret: secret, refresh_token: user.refreshToken } });
      const updated = tokenData(result, user);
      await mutate(db => { if (!db.users[id]) throw fail('授权已解除', 401); db.users[id] = updated; });
      return updated.accessToken;
    })();
    refreshing.set(id, task);
    try { return await task; } finally { refreshing.delete(id); }
  }
  async function disconnect(session) {
    const { id } = await identity(session);
    await mutate(db => { delete db.users[id]; for (const [key, entry] of Object.entries(db.sessions)) if (entry.userId === id) delete db.sessions[key]; });
  }
  return { config, start, authorize, callback, claim, identity, token, disconnect, api };
}

export async function uploadFeishuDocument(oauth, session, bytes, name) {
  if (!bytes.length || bytes.length > 20 * 1024 * 1024) throw fail('飞书文档导入文件必须小于 20MB');
  const form = new FormData();
  form.set('file_name', `${name.replace(/\.docx$/i, '')}.docx`);
  form.set('parent_type', 'ccm_import_open');
  form.set('size', String(bytes.length));
  form.set('extra', JSON.stringify({ obj_type: 'docx', file_extension: 'docx' }));
  form.set('file', new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), `${name}.docx`);
  const token = await oauth.token(session);
  const uploaded = await oauth.api('/drive/v1/medias/upload_all', { token, method: 'POST', form });
  if (!uploaded.data?.file_token) throw fail('飞书未返回上传文件标识', 502);
  const imported = await oauth.api('/drive/v1/import_tasks', { token, method: 'POST', json: {
    file_extension: 'docx', file_token: uploaded.data.file_token, type: 'docx', file_name: name.replace(/\.docx$/i, ''), point: { mount_type: 1, mount_key: '' }
  } });
  if (!imported.data?.ticket) throw fail('飞书未返回导入任务标识', 502);
  return imported.data.ticket;
}
