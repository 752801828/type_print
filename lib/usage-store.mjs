import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

const dataDir = path.resolve(fileURLToPath(new URL('../data', import.meta.url)));
export const usageFile = path.join(dataDir, 'usage-events.jsonl');
const clean = (value, limit = 120) => String(value || '').replace(/[\r\n\t]/g, ' ').trim().slice(0, limit);

export async function recordUsage(value = {}) {
  const userId = clean(value.userId, 200);
  const event = {
    at: new Date().toISOString(),
    user: userId ? crypto.createHash('sha256').update(userId).digest('hex').slice(0, 16) : 'anonymous',
    userName: clean(value.userName),
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
