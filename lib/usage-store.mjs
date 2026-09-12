import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import https from 'node:https';
import { fileURLToPath } from 'node:url';

const dataDir = path.resolve(fileURLToPath(new URL('../data', import.meta.url)));
export const usageFile = path.join(dataDir, 'usage-events.jsonl');
const clean = (value, limit = 120) => String(value || '').replace(/[\r\n\t]/g, ' ').trim().slice(0, limit);
let accessToken = '', accessTokenExpiresAt = 0;
const requestJson = (method, url, body, headers = {}) => new Promise((resolve, reject) => { const request = https.request(url, { method, headers: { accept: 'application/json', ...(body ? { 'content-type': 'application/json' } : {}), ...headers }, timeout: 2500 }, response => { let text = ''; response.setEncoding('utf8'); response.on('data', chunk => { text += chunk; }); response.on('end', () => { try { const data = JSON.parse(text); if (response.statusCode >= 200 && response.statusCode < 300) resolve(data); else reject(new Error(`Feishu API ${response.statusCode}`)); } catch (error) { reject(error); } }); }); request.on('timeout', () => request.destroy(new Error('Feishu API timeout'))); request.on('error', reject); if (body) request.write(JSON.stringify(body)); request.end(); });
const tenantAccessToken = async () => { const appId = clean(process.env.FEIYE_FEISHU_APP_ID, 200), appSecret = clean(process.env.FEIYE_FEISHU_APP_SECRET, 200); if (!appId || !appSecret) return ''; if (accessToken && accessTokenExpiresAt > Date.now() + 60000) return accessToken; const result = await requestJson('POST', 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', { app_id: appId, app_secret: appSecret }); accessToken = clean(result.tenant_access_token, 4000); accessTokenExpiresAt = Date.now() + Math.max(60000, (Number(result.expire) || 7200) * 1000); return accessToken; };
const resolveUserProfile = async userId => { if (!userId || !process.env.FEIYE_FEISHU_APP_ID || !process.env.FEIYE_FEISHU_APP_SECRET) return {}; const userIdType = userId.startsWith('ou_') ? 'open_id' : userId.startsWith('on_') ? 'union_id' : 'user_id'; try { const token = await tenantAccessToken(); if (!token) return {}; const result = await requestJson('GET', `https://open.feishu.cn/open-apis/contact/v3/users/${encodeURIComponent(userId)}?user_id_type=${userIdType}`, null, { authorization: `Bearer ${token}` }); const user = result.data?.user || {}; const avatar = user.avatar || {}; return { name: clean(user.name || user.en_name || user.enName), avatar: clean(avatar.avatar_origin || avatar.avatar_72x72 || avatar.avatar_240x240 || avatar.avatar_640x640, 1000) }; } catch { return {}; } };

export async function recordUsage(value = {}) {
  const userId = clean(value.userId, 200);
  const providedName = clean(value.userName);
  const profile = providedName ? {} : await resolveUserProfile(userId);
  const userName = providedName || profile.name || '';
  const userAvatar = clean(value.userAvatar, 1000) || profile.avatar || '';
  const event = {
    at: new Date().toISOString(),
    user: userId ? crypto.createHash('sha256').update(userId).digest('hex').slice(0, 16) : 'anonymous',
    userId,
    userName,
    userAvatar,
    baseId: clean(value.baseId, 100),
    baseName: clean(value.baseName),
    tableId: clean(value.tableId, 100),
    tableName: clean(value.tableName),
    templateId: clean(value.templateId, 100),
    templateName: clean(value.templateName),
    format: clean(value.format, 16).toLowerCase(),
    records: Math.max(0, Math.min(10000, Number(value.records) || 0))
  };
  await fs.mkdir(dataDir, { recursive: true });
  await fs.appendFile(usageFile, `${JSON.stringify(event)}\n`, 'utf8');
  return event;
}
