let sdk;
const servedByOwnServer = location.pathname.startsWith('/feishu') || ['127.0.0.1', 'localhost'].includes(location.hostname);
const basePath = location.pathname.startsWith('/feishu') ? '/feishu' : servedByOwnServer ? '' : 'https://gzwy.online/feishu';
const appUrl = path => `${basePath}${path}`;
const sdkUrl = servedByOwnServer ? appUrl('/vendor/lark-base/index.mjs') : new URL('./vendor/lark-base/index.mjs', location.href).href;
let reading = false;
let selectionBound = false;
let selectionPoll;
let selectionTimer;
let selectionPolling = false;
const linkedSchemaCache = new Map();
const state = { fields: [], viewFields: [], records: [], templates: [], selectedTemplate: null, context: null, table: null, view: null };
const $ = id => document.getElementById(id);
const text = value => Array.isArray(value) ? value.map(text).join('、') : value && typeof value === 'object' ? (value.text || value.name || (value.value !== undefined ? text(value.value) : JSON.stringify(value))) : value == null ? '' : String(value);
const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[char]));
const stripFieldMark = value => String(value ?? '').replace(/^(?:🔵|🟡|🔴|🟢|⚪)\s*/u, '').trim();
const normalizeKey = value => stripFieldMark(value).replace(/[\u2600-\u27BF\u{1F300}-\u{1FAFF}]/gu, '').replace(/[\s_\-：:（）()·]/g, '').toLowerCase();
const setStatus = message => { const node = $('status'); if (node) node.textContent = message; };
const toast = async (message, type = 'info') => { try { await sdk?.bitable?.ui?.showToast?.({ toastType: type, message }); } catch {} };
const isLayoutTemplate = item => ['.txt', '.layout', '.json'].includes(String(item?.extension || '').toLowerCase());
const outputLabel = value => ({ pdf: 'PDF', word: 'WORD', docx: 'DOCX', xlsx: 'XLSX', xls: 'XLS' }[value] || String(value || '文件').toUpperCase());
const updateOutputOptions = item => { const current = $('outputFormat').value; const extension = String(item?.extension || '.docx').slice(1).toLowerCase(); const options = isLayoutTemplate(item) ? [['pdf','PDF'],['word','Word'],['xlsx','Excel（XLSX）']] : [[extension, outputLabel(extension)]]; $('outputFormat').innerHTML = options.map(([value, label]) => `<option value="${value}">${label}</option>`).join(''); $('outputFormat').value = options.some(([value]) => value === current) ? current : options[0][0]; };
const updateAction = () => { const enabled = Boolean(state.selectedTemplate && state.records.length); $('generate').disabled = !enabled; $('selectedSummary').textContent = enabled ? `${state.selectedTemplate.name} · ${state.records.length} 条记录` : '选择模板并读取记录后继续'; $('generate').textContent = state.selectedTemplate ? `生成 ${outputLabel($('outputFormat').value)}　→` : '生成文件　→'; };
const formatStats = item => { const stats = item?.stats || {}; if (stats.pages) return `${stats.pages} 页 · ${item?.fields?.length || 0} 个变量`; return stats.worksheets || stats.rows || stats.cells || stats.formulas ? `工作表: ${stats.worksheets || 0}　行数: ${stats.rows || 0}　单元格数: ${stats.cells || 0}　公式数: ${stats.formulas || 0}` : `${item?.fields?.length || 0} 个变量 · ${String(item?.extension || 'docx').replace('.', '').toUpperCase()}`; };

async function basicFieldName(api) { const meta = api.getMeta ? await api.getMeta().catch(() => null) : null; const relationTableId = meta?.property?.tableId || meta?.property?.table_id || meta?.property?.relationTableId || meta?.property?.relation_table_id || (api.getTableId ? await api.getTableId().catch(() => '') : ''); return { id: String(api.id || meta?.id || ''), name: String(meta?.name || await api.getName?.() || api.name || api.id || ''), api, relationTableId: String(relationTableId || '') }; }
const fallbackRecordField = (field, record) => { const fields = record?.fields || {}; if (fields[field.id] !== undefined) return fields[field.id]; if (fields[field.name] !== undefined) return fields[field.name]; const key = Object.keys(fields).find(candidate => normalizeKey(candidate) === normalizeKey(field.name)); return key === undefined ? undefined : fields[key]; };
const hasCellValue = value => value != null && value !== '' && value !== 'undefined' && !(typeof value === 'string' && !value.trim());
async function readField(field, recordId, record, table) { let value; try { value = await table?.getCellValue?.(field.id, recordId); } catch {} if (!hasCellValue(value)) { try { value = await field.api.getValue(record || recordId); } catch {} } if (!hasCellValue(value)) { try { value = await field.api.getValue(recordId); } catch {} } if (!hasCellValue(value)) { try { value = await field.api.getCellString(recordId); } catch {} } if (!hasCellValue(value)) value = fallbackRecordField(field, record); return value ?? ''; }
async function readRawField(field, recordId, record, table) { try { const value = await table?.getCellValue?.(field.id, recordId); if (hasCellValue(value) && !(Array.isArray(value) && !value.length)) return value; } catch {} for (const target of [record, recordId]) { if (!target) continue; try { const value = await field.api.getValue(target); if (value != null) return value; } catch {} } return fallbackRecordField(field, record) ?? ''; }
const relationIds = value => { const items = Array.isArray(value) ? value : value == null ? [] : [value]; return [...new Set(items.flatMap(item => { if (typeof item === 'string') return [item]; if (!item || typeof item !== 'object') return []; return item.recordIds || item.record_ids || item.linkRecordIds || item.link_record_ids || [item.recordId || item.record_id || item.id]; }).filter(Boolean).map(String))]; };
async function readLinkedRows(raw, bitable, fallbackTableId = '') {
  const values = Array.isArray(raw) ? raw : Array.isArray(raw?.value) ? raw.value : raw ? [raw] : [];
  const relationItems = values.flatMap(item => Array.isArray(item?.value) ? item.value : [item]);
  const ids = relationIds([raw, ...values, ...relationItems]);
  const tableId = raw?.tableId || raw?.table_id || raw?.property?.tableId || raw?.property?.table_id || relationItems.find(item => item && typeof item === 'object')?.tableId || relationItems.find(item => item && typeof item === 'object')?.table_id || relationItems.find(item => item && typeof item === 'object')?.property?.tableId || relationItems.find(item => item && typeof item === 'object')?.property?.table_id || fallbackTableId;
  if (!ids.length || !tableId) return [];
  try {
    let schema = linkedSchemaCache.get(tableId);
    if (!schema) { const table = await bitable.base.getTable(tableId); schema = { table, fields: await Promise.all((await table.getFieldList()).map(basicFieldName)) }; linkedSchemaCache.set(tableId, schema); }
    const requested = new Map();
    for (const templateField of state.selectedTemplate?.fields || []) { const field = templateField.marker === '#' ? null : findTemplateField(schema.fields, templateField.name, false); if (field) requested.set(field.id, field); }
    const linkedFields = requested.size ? [...requested.values()] : schema.fields;
    return Promise.all(ids.map(async id => {
      let rawRecord; try { rawRecord = await schema.table.getRecordById(id); } catch {}
      const entries = await Promise.all(linkedFields.map(async field => [field, text(await readField(field, id, rawRecord, schema.table)) || text(rawRecord?.fields?.[field.id])]));
      const row = {};
      for (const [field, value] of entries) { row[field.name] = value; row[stripFieldMark(field.name)] ??= value; }
      for (const templateField of state.selectedTemplate?.fields || []) if (templateField.marker !== '#') {
        const candidate = findTemplateField(schema.fields, templateField.name);
        if (candidate && row[templateField.name] === undefined) row[templateField.name] = row[candidate.name] ?? '';
      }
      return row;
    }));
  } catch { return []; }
}
const matchingTemplateLoops = field => (state.selectedTemplate?.fields || []).filter(item => item.marker === '#' && fieldMatchesTemplate(field.name, item.name));
const isTemplateLoopField = field => matchingTemplateLoops(field).length > 0;
const activeRecordFields = () => {
  const selected = new Map();
  const add = field => { if (field?.id) selected.set(field.id, field); };
  add(state.viewFields[0] || state.fields[0]);
  for (const templateField of state.selectedTemplate?.fields || []) {
    if (/大写/.test(templateField.name) || ['__合同创建时间', '__SKU总数'].includes(templateField.name)) continue;
    add(findTemplateField(state.fields, templateField.name));
  }
  return [...selected.values()];
};
async function readRecord(id, table, bitable) { let record; try { record = await table.getRecordById(id); } catch {} const entries = await Promise.all(activeRecordFields().map(async field => { const valueTask = readField(field, id, record, table); const linkedTableId = field.relationTableId || ''; const linkedTask = linkedTableId || isTemplateLoopField(field) ? readRawField(field, id, record, table).then(raw => readLinkedRows(raw, bitable, linkedTableId)) : Promise.resolve([]); return { field, value: await valueTask, linked: await linkedTask }; })); const fields = {}; const loops = {}; for (const { field, value, linked } of entries) { fields[field.id] = value; if (linked.length) { loops[field.name] = linked; loops[stripFieldMark(field.name)] = linked; for (const templateField of matchingTemplateLoops(field)) loops[templateField.name] = linked; } } return { id, fields, loops }; }

async function readCurrentRecord(force = false) {
  if (reading) return;
  reading = true;
  setStatus('正在读取选中记录…');
  try {
    if (!sdk) sdk = await import(sdkUrl);
    bindSelectionListener();
    const { bitable } = sdk;
    let selection = {}; try { selection = await bitable.base.getSelection(); } catch {}
    let table = null; if (selection.tableId && bitable.base.getTableById) { try { table = await bitable.base.getTableById(selection.tableId); } catch {} } if (!table) table = await bitable.base.getActiveTable();
    const meta = selection.tableId || table.id ? null : await table.getMeta();
    let view = null; try { view = selection.viewId && table.getViewById ? await table.getViewById(selection.viewId) : await table.getActiveView(); } catch {}
    const checkedRecordIds = view?.getSelectedRecordIdList ? (await view.getSelectedRecordIdList().catch(() => [])).map(String) : [];
    const selectedRecordIds = currentReadRecordIds(selection, checkedRecordIds);
    const tableId = String(selection.tableId || table.id || meta?.id || ''); const viewId = String(selection.viewId || view?.id || ''); const selectionKey = `${selection.baseId || ''}/${tableId}/${viewId}/${selectedRecordIds.join(',')}`; const previous = state.context;
    if (!force && previous?.selectionKey === selectionKey) return;
    const sameScope = previous?.tableId === tableId && previous?.viewId === viewId && state.fields.length;
    if (!sameScope) {
      linkedSchemaCache.clear();
      state.fields = await Promise.all((await table.getFieldList()).map(basicFieldName));
      let visibleFieldIds = []; try { visibleFieldIds = await view?.getVisibleFieldIdList?.() || []; } catch {}
      const fieldsById = new Map(state.fields.map(field => [field.id, field])); state.viewFields = visibleFieldIds.map(id => fieldsById.get(String(id))).filter(Boolean); if (!state.viewFields.length) state.viewFields = state.fields;
    }
    const names = sameScope ? [previous.baseName, previous.tableName, previous.viewName] : await Promise.all([bitable.base.getBaseName?.().catch(() => '') || '', table.getName().catch(() => ''), view?.getName?.().catch(() => '') || '']);
    state.table = table; state.view = view; state.context = { baseId: String(selection.baseId || ''), tableId, viewId, selectionKey, baseName: String(names[0] || ''), tableName: String(names[1] || ''), viewName: String(names[2] || '当前记录') };
    if (!previous || previous.baseId !== state.context.baseId || previous.tableId !== state.context.tableId) await loadTemplates();
    let ids = selectedRecordIds;
    if (!ids.length) { try { ids = (await table.getRecordIdList?.()).slice(0, 1); } catch {} }
    if (!ids.length) throw new Error('当前没有可读取的记录');
    state.records = await Promise.all(ids.map(id => readRecord(id, table, bitable))); drawDependencies();
    $('contextTitle').textContent = `${state.context.baseName || '当前多维表'} / ${state.context.tableName}`; $('contextMeta').textContent = checkedRecordIds.length ? `${state.context.viewName} · 已读取 ${state.records.length} 条勾选记录` : `${state.context.viewName} · 当前活动行`; drawRecord(); setStatus('已连接当前多维表');
  } catch (error) { const message = /not registered/i.test(String(error?.message || '')) ? '当前页面未注册飞书 Base 宿主，请从多维表「扩展脚本」打开' : error.message || '读取记录失败'; $('contextMeta').textContent = message; setStatus(message); await toast(message, 'error'); }
  finally { reading = false; }
}

async function loadTemplates() { const query = state.context?.tableId ? `?baseId=${encodeURIComponent(state.context.baseId)}&tableId=${encodeURIComponent(state.context.tableId)}` : ''; const response = await fetch(appUrl(`/api/templates${query}`)); const result = await response.json(); state.templates = state.context?.tableId ? (result.templates || []) : []; if (!state.templates.some(item => item.id === state.selectedTemplate?.id)) state.selectedTemplate = state.templates[0] || null; drawTemplates(); updateAction(); }
const layoutText = value => Array.isArray(value) ? value.map(layoutText).join('') : value && typeof value === 'object' ? (value.text || '') : String(value || '');
const cleanLayoutName = stripFieldMark;
const matchingValue = (object, key) => { if (!object || typeof object !== 'object') return undefined; if (Object.hasOwn(object, key)) return object[key]; const normalized = normalizeKey(key); const match = Object.keys(object).find(candidate => legacyPlaceholderName(candidate) === key) ?? Object.keys(object).find(candidate => normalizeKey(candidate) === normalized); return match === undefined ? undefined : object[match]; };
const templateFieldAliases = { '__合同编号': ['合同编号'], '__供应商': ['供应商'], '__合同明细_采购明细': ['合同明细', '采购明细', '合同明细采购明细'], '__开票品名': ['开票品名'], 'SKU__': ['SKU', 'SKU💻'], '单价_含税___': ['单价（含税）', '单价(含税)', '单价(含税)💻', '单价（含税）💻', '含税单价'], '实收数量__': ['实收数量', '实收数量💻', '数量'], '总价_含税_': ['总价（含税）', '总价(含税)', '总价(含税)💻', '含税总价'], '__产品金额': ['产品金额'], '产品金额_大写_': ['产品金额（大写）', '产品金额(大写)'], '__产品运输费用': ['产品运输费用'], '产品运输费用_大写_': ['产品运输费用（大写）', '产品运输费用(大写)'], '__合同合计金额': ['合同合计金额'], '合同合计金额_大写_': ['合同合计金额（大写）', '合同合计金额(大写)'], '__采购合同': ['采购合同'], '__SKU种类': ['SKU种类'], '__SKU总数': ['SKU总数'], '__合同创建时间': ['合同创建时间'] };
const templateNames = name => [String(name || ''), ...(templateFieldAliases[String(name || '')] || [])];
const legacyPlaceholderName = value => String(value || '').replace(/[^A-Za-z0-9_\u3400-\u9fff]/g, '_');
const fieldMatchesTemplate = (fieldName, templateName) => templateNames(templateName).some(candidate => { const left = normalizeKey(candidate); const right = normalizeKey(fieldName); return candidate === legacyPlaceholderName(fieldName) || left === right || (left.length >= 2 && right.includes(left)) || (right.length >= 2 && left.includes(right)); });
const findTemplateField = (fields, name, allowFuzzy = true) => {
  const exact = fields.find(field => field.name === name) || fields.find(field => legacyPlaceholderName(field.name) === name);
  if (exact) return exact;
  const normalizedName = normalizeKey(name); const normalizedExact = fields.find(field => normalizeKey(field.name) === normalizedName);
  if (normalizedExact) return normalizedExact;
  for (const candidate of templateNames(name)) {
    const alias = fields.find(field => field.name === candidate || legacyPlaceholderName(field.name) === candidate);
    if (alias) return alias;
  }
  for (const candidate of templateNames(name)) {
    const normalized = fields.find(field => normalizeKey(field.name) === normalizeKey(candidate));
    if (normalized) return normalized;
  }
  return allowFuzzy ? fields.find(field => fieldMatchesTemplate(field.name, name)) : undefined;
};
const numberValue = value => { if (value && typeof value === 'object' && value.value !== undefined) return numberValue(value.value); const number = Number(String(value ?? '').replace(/[,，￥¥\s]/g, '')); return Number.isFinite(number) ? number : 0; };
const chineseCurrency = value => { const amount = Math.round(numberValue(value) * 100); const digits = '零壹贰叁肆伍陆柒捌玖'; const small = ['', '拾', '佰', '仟']; const large = ['', '万', '亿', '兆']; const integer = Math.floor(amount / 100); const groupText = group => { let out = ''; let zero = false; for (let i = 3; i >= 0; i--) { const digit = Math.floor(group / (10 ** i)) % 10; if (!digit) { if (out) zero = true; continue; } if (zero) out += '零'; out += digits[digit] + small[i]; zero = false; } return out; }; let out = ''; let missing = false; for (let i = 3; i >= 0; i--) { const group = Math.floor(integer / (10000 ** i)) % 10000; if (!group) { if (out) missing = true; continue; } if (out && (missing || group < 1000)) out += '零'; out += groupText(group) + large[i]; missing = false; } if (!out) out = '零'; const jiao = Math.floor(amount / 10) % 10; const fen = amount % 10; return `${numberValue(value) < 0 ? '负' : ''}${out}元${jiao ? `${digits[jiao]}角` : ''}${fen ? `${!jiao && integer ? '零' : ''}${digits[fen]}分` : (!jiao ? '整' : '')}`; };
const procurementRows = record => Object.entries(record.loops || {}).find(([name, rows]) => Array.isArray(rows) && /合同明细|采购明细/.test(normalizeKey(name)))?.[1] || [];
const applyProcurementRules = (values, record) => { const rows = procurementRows(record); if (!rows.length && !state.selectedTemplate?.fields?.some(field => /合同明细.*采购明细/.test(normalizeKey(field.name)))) return values; values['__合同明细_采购明细'] = rows; const detailValue = (row, name) => { const fields = Object.keys(row || {}).filter(key => hasCellValue(row[key])).map(key => ({ name: key })); const field = findTemplateField(fields, name, false); return field ? row[field.name] : undefined; }; for (const row of rows) { for (const name of ['__开票品名', 'SKU__', '单位', '单价_含税___', '实收数量__', '总价_含税_']) { const value = detailValue(row, name); if (value !== undefined) row[name] = value; } } values['__SKU总数'] = rows.reduce((sum, row) => sum + numberValue(detailValue(row, '实收数量__')), 0); const total = matchingValue(values, '__合同合计金额') ?? matchingValue(values, '合同合计金额'); values['合同合计金额_大写_'] = chineseCurrency(total); values['产品金额_大写_'] = chineseCurrency(matchingValue(values, '__产品金额') ?? matchingValue(values, '产品金额')); values['产品运输费用_大写_'] = chineseCurrency(matchingValue(values, '__产品运输费用') ?? matchingValue(values, '产品运输费用')); values['__合同创建时间'] = `${new Date().getFullYear()}年${new Date().getMonth() + 1}月${new Date().getDate()}日`; return values; };
const recordScope = record => { const values = {}; for (const field of state.fields) { const value = text(record.fields[field.id]); values[field.name] = value; values[stripFieldMark(field.name)] ??= value; } for (const templateField of state.selectedTemplate?.fields || []) if (templateField.marker !== '#') { const sourceField = findTemplateField(state.fields, templateField.name); if (sourceField) values[templateField.name] = text(record.fields[sourceField.id]); } values['表格名'] = state.context?.tableName || ''; values['打印时间'] = new Date().toLocaleString('zh-CN'); for (const [name, rows] of Object.entries(record.loops || {})) { values[name] = rows; values[stripFieldMark(name)] ??= rows; } applyProcurementRules(values, record); for (const field of state.selectedTemplate?.fields || []) if (matchingValue(values, field.name) === undefined) values[field.name] = field.marker === '#' ? [] : ''; return values; };
const selectionRecordIds = selection => { const primary = selection?.recordId || selection?.record_id; const extra = selection?.recordIds || selection?.record_ids || selection?.recordIdList || selection?.record_id_list; return [...new Set([primary, ...(Array.isArray(extra) ? extra : [])].filter(Boolean).map(String))]; };
const currentReadRecordIds = (selection, checkedIds) => { const checked = [...new Set((checkedIds || []).filter(Boolean).map(String))]; return checked.length ? checked : selectionRecordIds(selection); };
const layoutVariableValue = (names, scope, rowContext) => { const path = (names || []).map(cleanLayoutName); if (scope === null) return `[${path.join(' / ')}]`; if (path[0] === '#') return String((rowContext?.index ?? 0) + 1); if (rowContext && normalizeKey(path[0]) === normalizeKey(rowContext.root)) return text(matchingValue(rowContext.item, path.at(-1))); let value = matchingValue(scope, path[0]); for (const key of path.slice(1)) value = matchingValue(value, key); return text(value); };
const layoutInline = (nodes, scope = null, rowContext = null) => (nodes || []).map(node => node.type === 'variable' ? `<span class="${scope === null ? 'layout-variable' : 'layout-value'}">${escapeHtml(layoutVariableValue(node.name, scope, rowContext))}</span>` : `<span style="font-size:${escapeHtml(node.fontSize || '9pt')};font-weight:${node.bold ? 700 : 400};font-style:${node.italic ? 'italic' : 'normal'}">${escapeHtml(node.text || layoutText(node.children))}</span>`).join('');
const renderLayoutPreview = (layout, scope = null) => {
  const root = document.createElement('div');
  const setting = layout?.pageSetting || {}; const adaptive = setting.mode === 'auto' || setting.pageSizeMode === 'auto';
  for (const page of layout?.document?.pages || []) {
    const pageEl = document.createElement('section'); pageEl.className = 'layout-page';
    if (adaptive) pageEl.classList.add('layout-page-adaptive'); else { pageEl.style.width = `${Number(setting.width || 210)}mm`; pageEl.style.minHeight = `${Number(setting.height || 297)}mm`; }
    for (const row of page.rows || []) {
      const rowEl = document.createElement('div'); rowEl.className = 'layout-row';
      for (const column of row.columns || []) {
        const colEl = document.createElement('div'); colEl.className = 'layout-column'; colEl.style.width = `${column.width || 100}%`;
        for (const block of column.blocks || []) {
          const blockEl = document.createElement('div'); blockEl.className = 'layout-block-preview';
          if (block.type === 4 && block.table) {
            const table = document.createElement('table'); const colgroup = document.createElement('colgroup');
            for (const item of block.table.columns || []) { const col = document.createElement('col'); col.style.width = `${item.width || 1}%`; colgroup.append(col); } table.append(colgroup);
            const tbody = document.createElement('tbody'); const merges = block.table.merges || []; const covered = new Set();
            for (const merge of merges) for (let rr = merge.rowIndex; rr < merge.rowIndex + merge.rowSpan; rr++) for (let cc = merge.colIndex; cc < merge.colIndex + merge.colSpan; cc++) if (rr !== merge.rowIndex || cc !== merge.colIndex) covered.add(`${rr}:${cc}`);
            const appendTableRow = (tableRow, rowIndex, rowContext = null) => {
              const tr = document.createElement('tr');
              for (const [cellIndex, cell] of (tableRow.cells || []).entries()) {
                if (covered.has(`${rowIndex}:${cellIndex}`)) continue;
                const merge = merges.find(item => item.rowIndex === rowIndex && item.colIndex === cellIndex);
                const td = document.createElement('td'); td.style.background = cell.background || ''; td.style.textAlign = cell.content?.[0]?.align || '';
                if (merge?.colSpan > 1) td.colSpan = merge.colSpan; if (merge?.rowSpan > 1) td.rowSpan = merge.rowSpan;
                for (const paragraph of cell.content || []) { const p = document.createElement('p'); p.innerHTML = layoutInline(paragraph.children, scope, rowContext); p.style.textAlign = paragraph.align || ''; td.append(p); }
                tr.append(td);
              }
              tbody.append(tr);
            };
            for (const [rowIndex, tableRow] of (block.table.rows || []).entries()) {
              const dynamic = (block.table.dynamicRows || []).find(item => item.rowIndex === rowIndex); const rootName = cleanLayoutName(dynamic?.dataSource?.rootPath?.[0]); const rows = scope && rootName && Array.isArray(matchingValue(scope, rootName)) ? matchingValue(scope, rootName) : null;
              if (rows?.length) rows.forEach((item, index) => appendTableRow(tableRow, rowIndex, { root: rootName, item, index })); else appendTableRow(tableRow, rowIndex);
            }
            table.append(tbody); blockEl.append(table);
          } else for (const paragraph of block.content || []) { const p = document.createElement('p'); p.innerHTML = layoutInline(paragraph.children, scope); p.style.textAlign = paragraph.align || ''; blockEl.append(p); }
          colEl.append(blockEl);
        }
        rowEl.append(colEl);
      }
      pageEl.append(rowEl);
    }
    root.append(pageEl);
  }
  return root;
};
const resetPreview = () => {
  $('previewDocx').replaceChildren(); $('previewDocx').hidden = true;
  $('previewFrame').hidden = true; $('previewFrame').removeAttribute('src'); $('previewFrame').removeAttribute('srcdoc');
  $('previewLayout').hidden = true; $('previewLayout').replaceChildren();
  $('previewText').hidden = true; $('previewText').textContent = '';
};
const renderDocxPreview = async blob => {
  if (!window.docx?.renderAsync) throw new Error('DOCX 布局渲染器加载失败');
  const host = $('previewDocx'); host.hidden = false;
  await window.docx.renderAsync(await blob.arrayBuffer(), host, host, { breakPages: true, ignoreWidth: false, ignoreHeight: false, ignoreFonts: false, renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true, useBase64URL: true });
};
async function openPreview(id) {
  try {
    const response = await fetch(appUrl(`/api/templates/${id}/preview`)); const data = await response.json(); if (!response.ok) throw new Error(data.error);
    const fileUrl = appUrl(data.fileUrl); resetPreview();
    $('previewTitle').textContent = data.template.name; $('previewMeta').textContent = `${String(data.template.extension || '').replace('.', '').toUpperCase()} · ${formatStats(data.template)}`; $('previewMeta').classList.remove('record-render-meta');
    $('downloadPreview').hidden = false; $('downloadPreview').href = fileUrl; $('previewFields').hidden = false; $('previewFields').innerHTML = (data.fields || []).length ? `<strong>识别到的变量</strong><div>${data.fields.map(field => `<span class="preview-chip">${escapeHtml(field.marker ? `{${field.marker}${field.name}}` : `{${field.name}}`)}</span>`).join('')}</div>` : '<span>未识别到变量</span>';
    $('previewDialog').showModal();
    if (data.kind === 'docx') { const file = await fetch(fileUrl); if (!file.ok) throw new Error('DOCX 模板读取失败'); await renderDocxPreview(await file.blob()); }
    else if (data.kind === 'pdf') { $('previewFrame').hidden = false; $('previewFrame').src = fileUrl; }
    else if (data.kind === 'html') { $('previewFrame').hidden = false; $('previewFrame').srcdoc = data.html; }
    else if (data.kind === 'layout') { $('previewLayout').hidden = false; $('previewLayout').replaceChildren(renderLayoutPreview(data.layout)); }
    else { $('previewText').hidden = false; $('previewText').textContent = data.content || '此格式无法在浏览器内直接渲染，已显示模板变量信息。'; }
  } catch (error) { await toast(error.message, 'error'); }
}
async function openRecordPreview(record) {
  if (!state.selectedTemplate) return toast('请先选择模板', 'error');
  try {
    const response = await fetch(appUrl(`/api/templates/${state.selectedTemplate.id}/record-preview`), { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ record: recordScope(record) }) });
    if (!response.ok) { const error = await response.json(); throw new Error(error.error); }
    const primary = state.viewFields[0] || state.fields[0]; const label = primary ? text(record.fields[primary.id]) : record.id; resetPreview();
    $('previewTitle').textContent = `${state.selectedTemplate.name} · 记录预览`; $('previewMeta').textContent = `${label || '当前记录'} · 已填充真实数据`; $('previewMeta').classList.add('record-render-meta'); $('previewFields').hidden = true; $('downloadPreview').hidden = true; $('previewDialog').showModal();
    if (response.headers.get('content-type')?.includes('wordprocessingml')) await renderDocxPreview(await response.blob());
    else { const data = await response.json(); $('previewFrame').hidden = false; $('previewFrame').srcdoc = data.html; }
  } catch (error) { await toast(`预览失败：${error.message}`, 'error'); }
}
const templateFieldDiagnostics = () => {
  const fields = [...(state.fields || []), ...[...linkedSchemaCache.values()].flatMap(schema => schema.fields || [])];
  return (state.selectedTemplate?.fields || []).map(templateField => {
    const special = /大写/.test(templateField.name) || templateField.name === '__合同创建时间' || templateField.name === '__SKU总数'; const matched = special ? { name: '系统计算', id: 'system' } : templateField.marker === '#' ? fields.find(field => matchingTemplateLoops(field).some(item => item.name === templateField.name)) : findTemplateField(fields, templateField.name);
    return { ...templateField, matched, fieldId: matched?.id || '', matchedName: matched?.name || '', relationTableId: matched?.relationTableId || '' };
  });
};
function drawDependencies() { const fields = state.selectedTemplate?.fields || []; $('dependencyCount').textContent = fields.length; $('dependencyList').innerHTML = fields.length ? fields.map(field => { const parts = field.name.split(/\s*[·/]\s*/).filter(Boolean); const kind = field.marker === '#' ? '循环' : field.marker === '?' ? '条件' : '字段'; return `<div class="dependency-row"><span class="dependency-kind">${kind}</span>${parts.map(part => `<span class="dependency-chip">${escapeHtml(part)}</span>`).join('<b>/</b>')}</div>`; }).join('') : '<div class="empty-template">当前模板没有识别到字段依赖</div>'; const diagnostics = templateFieldDiagnostics(); $('fieldDiagnostics').innerHTML = diagnostics.length ? diagnostics.map(item => item.matched ? `<div class="field-diagnostic matched"><span>✓</span><code>${escapeHtml(item.name)}</code><small>已匹配：${escapeHtml(item.matchedName)}　字段 ID：${escapeHtml(item.fieldId)}</small></div>` : `<div class="field-diagnostic missing"><span>!</span><code>${escapeHtml(item.name)}</code><small>未匹配；请修改模板。当前字段：${escapeHtml((state.fields || []).map(field => `${field.name}（${field.id}）`).join('、') || '无')}</small></div>`).join('') : '<div class="empty-template">暂无字段诊断</div>'; }
function drawRecord() { const list = $('recordPreviewList'); if (!state.records.length) { list.innerHTML = '<div class="empty-template">请先在飞书中选择记录</div>'; return; } const primary = state.viewFields[0] || state.fields[0]; list.innerHTML = state.records.map((record, index) => { const label = primary ? text(record.fields[primary.id]) : ''; return `<button class="record-preview-button" data-record-preview="${escapeHtml(record.id)}" ${state.selectedTemplate ? '' : 'disabled'}><span class="record-preview-copy"><strong>${escapeHtml(label || `记录 ${index + 1}`)}</strong><small>${escapeHtml(state.context?.tableName || '当前数据表')} · 点击查看完整文件渲染</small></span><span class="record-preview-action">预览</span></button>`; }).join(''); list.querySelectorAll('[data-record-preview]').forEach(button => button.onclick = () => openRecordPreview(state.records.find(record => record.id === button.dataset.recordPreview))); updateAction(); }
async function deleteTemplateItem(id) { const item = state.templates.find(template => template.id === id); if (!item || !confirm(`确定删除模板“${item.name}”吗？删除后不可恢复。`)) return; try { const response = await fetch(appUrl(`/api/templates/${id}`), { method: 'DELETE' }); const result = await response.json(); if (!response.ok || !result.deleted) throw new Error(result.error || '删除失败'); state.templates = state.templates.filter(template => template.id !== id); if (state.selectedTemplate?.id === id) state.selectedTemplate = state.templates[0] || null; drawTemplates(); await toast('模板已删除', 'success'); } catch (error) { await toast(error.message, 'error'); } }
function drawTemplates() { const current = state.selectedTemplate || state.templates[0]; updateOutputOptions(current); $('templateList').innerHTML = current ? `<div class="template selected"><div><strong>${escapeHtml(current.name)}</strong><small>${formatStats(current)}</small></div><span class="template-actions"><button class="preview-template" data-preview-id="${escapeHtml(current.id)}">预览</button><button class="delete-template" data-delete-id="${escapeHtml(current.id)}">删除</button></span></div>` : '<div class="empty">还没有模板<br><small>请上传模板文件</small></div>'; $('categoryList').innerHTML = state.templates.length ? state.templates.map(item => `<div class="category-template-row"><button class="category-template ${current?.id === item.id ? 'active' : ''}" data-id="${escapeHtml(item.id)}">▣　${escapeHtml(item.name)}</button><button class="mini-template-action danger" data-delete-id="${escapeHtml(item.id)}" aria-label="删除 ${escapeHtml(item.name)}">×</button></div>`).join('') : '<div class="side-empty">暂无模板</div>'; $('templateType').textContent = isLayoutTemplate(current) ? '在线' : String(current?.extension || 'docx').replace('.', '').toUpperCase(); $('templateInfo').textContent = current ? formatStats(current) : '支持 DOCX、XLSX、XLS、PDF、飞书在线模板'; drawDependencies(); if (state.records.length) drawRecord(); document.querySelectorAll('.category-template').forEach(button => button.onclick = async () => { if (state.selectedTemplate?.id === button.dataset.id) return; state.selectedTemplate = state.templates.find(item => item.id === button.dataset.id) || null; drawTemplates(); updateAction(); await readCurrentRecord(true); }); document.querySelectorAll('.preview-template').forEach(button => button.onclick = event => { event.stopPropagation(); openPreview(button.dataset.previewId); }); document.querySelectorAll('[data-delete-id]').forEach(button => button.onclick = event => { event.stopPropagation(); deleteTemplateItem(button.dataset.deleteId); }); updateAction(); }

async function uploadTemplate(file) { const headers = { 'content-type': file.type || 'application/octet-stream', 'x-file-name': encodeURIComponent(file.name) }; for (const [key, value] of Object.entries({ 'x-base-id': state.context?.baseId, 'x-table-id': state.context?.tableId, 'x-view-id': state.context?.viewId, 'x-base-name': state.context?.baseName, 'x-table-name': state.context?.tableName })) if (value) headers[key] = encodeURIComponent(value); const response = await fetch(appUrl('/api/templates'), { method: 'POST', headers, body: await file.arrayBuffer() }); const result = await response.json(); if (!response.ok) throw new Error(result.error || '上传失败'); state.templates.unshift(result.template); state.selectedTemplate = result.template; drawTemplates(); await readCurrentRecord(true); await toast('模板已保存并绑定当前数据表', 'success'); await openPreview(result.template.id); }
$('createTemplate').onclick = () => $('templateFile').click();
$('menuToggle').onclick = () => $('appShell').classList.toggle('menu-collapsed');
$('generateTop').onclick = () => state.selectedTemplate && state.records.length ? $('generate').click() : toast(state.selectedTemplate ? '请先在飞书表格中勾选记录' : '请先导入模板', 'error');
$('debugInfo').onclick = () => { const panel = $('dependencyPanel'); panel.hidden = !panel.hidden; $('debugInfo').textContent = panel.hidden ? '字段匹配诊断' : '收起字段诊断'; drawDependencies(); };
$('closePreview').onclick = () => { $('previewFrame').src = 'about:blank'; $('previewDialog').close(); };
$('templateFile').onchange = async event => { const file = event.target.files[0]; if (file) try { await uploadTemplate(file); } catch (error) { $('contextMeta').textContent = error.message; await toast(error.message, 'error'); } event.target.value = ''; };
$('refreshRecord').onclick = () => readCurrentRecord(true);
$('outputFormat').onchange = updateAction;
$('generate').onclick = async () => { if (!state.selectedTemplate || !state.records.length) return; const missing = templateFieldDiagnostics().filter(item => !item.matched); if (missing.length) return toast(`模板有 ${missing.length} 个字段未匹配：${missing.map(item => item.name).join('、')}，请先修改模板`, 'error'); const button = $('generate'); button.disabled = true; button.textContent = '生成中…'; try { const records = state.records.map(recordScope); const response = await fetch(appUrl('/api/generate-docx'), { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ templateId: state.selectedTemplate.id, records, outputFormat: $('outputFormat').value }) }); const result = await response.json(); if (!response.ok || !result.output) throw new Error(result.error || '生成失败'); $('result').hidden = false; $('result').innerHTML = `生成完成。<a href="${appUrl(`/api/outputs/${result.output.id}/download`)}">下载 ${escapeHtml(result.output.name)} ↗</a>`; await toast(`${records.length} 条记录已生成`, 'success'); } catch (error) { await toast(error.message, 'error'); } finally { updateAction(); } };
const bindSelectionListener = () => { if (selectionBound || !sdk?.bitable?.base) return; selectionBound = true; try { sdk.bitable.base.onSelectionChange(() => { clearTimeout(selectionTimer); selectionTimer = setTimeout(readCurrentRecord, 20); }); } catch {} selectionPoll = setInterval(async () => { if (selectionPolling) return; selectionPolling = true; try { const [current, checked] = await Promise.all([sdk.bitable.base.getSelection(), state.view?.getSelectedRecordIdList ? state.view.getSelectedRecordIdList() : []]); const ids = currentReadRecordIds(current, checked); const key = `${current.baseId || ''}/${current.tableId || ''}/${current.viewId || ''}/${ids.join(',')}`; if (state.context?.selectionKey && state.context.selectionKey !== key) { $('contextMeta').textContent = '正在读取勾选记录…'; readCurrentRecord(); } } catch {} finally { selectionPolling = false; } }, 200); };
readCurrentRecord();
