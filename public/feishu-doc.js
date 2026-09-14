const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
const hex = buffer => [...new Uint8Array(buffer)].map(byte => byte.toString(16).padStart(2, '0')).join('');
const read = key => { try { return sessionStorage.getItem(key) || ''; } catch { return ''; } };
const write = (key, value) => { try { value ? sessionStorage.setItem(key, value) : sessionStorage.removeItem(key); } catch {} };

export async function exportFeishuDocument({ appUrl, escapeHtml, result, payload, storageKey, popup }) {
  let login = read(storageKey), cancelled = false;
  const controller = new AbortController();
  const stage = (title, detail, percent) => {
    result.hidden = false; result.className = 'result generating';
    result.innerHTML = `<div class="generation-status"><strong>${escapeHtml(title)}</strong><span>${percent}%</span></div><div class="progress-track" role="progressbar" aria-label="飞书文档导出进度" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${percent}"><i style="width:${percent}%"></i></div><small>${escapeHtml(detail)}</small>`;
  };
  async function api(path, data, authenticated = true) {
    const attemptController = new AbortController();
    const abort = () => attemptController.abort();
    controller.signal.addEventListener('abort', abort, { once: true });
    const timer = setTimeout(abort, 55000);
    try {
    if (controller.signal.aborted) abort();
    const response = await fetch(appUrl(path), { method: data === undefined ? 'GET' : 'POST',
      headers: { ...(data === undefined ? {} : { 'content-type': 'application/json' }), ...(authenticated && login ? { authorization: `Bearer ${login}` } : {}) },
      body: data === undefined ? undefined : JSON.stringify(data), signal: attemptController.signal });
    const value = await response.json();
    if (!response.ok) throw Object.assign(new Error(value.error || '请求失败'), { status: response.status });
    return value;
    } finally { clearTimeout(timer); controller.signal.removeEventListener('abort', abort); }
  }
  try {
    stage('正在检查飞书授权', '仅导出飞书文档时需要授权，普通文件下载不受影响。', 5);
    let profile;
    if (login) {
      try { profile = (await api('/api/oauth/feishu/status')).user; }
      catch (error) { if (error.status !== 401) throw error; write(storageKey, ''); login = ''; }
    }
    if (!login) {
      const random = crypto.getRandomValues(new Uint8Array(32));
      const verifier = btoa(String.fromCharCode(...random)).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '');
      const challenge = hex(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
      const attempt = await api('/api/oauth/feishu/start', { challenge }, false);
      stage('等待飞书授权', '请在授权窗口确认账户。授权完成后自动继续，文档将进入该账户的个人云盘。', 10);
      const actions = document.createElement('div'); actions.className = 'download-row';
      const link = document.createElement('a'); link.href = attempt.url; link.target = '_blank'; link.rel = 'noopener noreferrer'; link.textContent = '打开飞书授权';
      const cancel = document.createElement('button'); cancel.textContent = '取消'; cancel.type = 'button'; cancel.onclick = () => { cancelled = true; controller.abort(); popup?.close(); };
      actions.append(link, cancel); result.append(actions);
      if (popup && !popup.closed) popup.location.replace(attempt.url);
      const deadline = Date.now() + 10 * 60 * 1000;
      while (Date.now() < deadline && !cancelled) {
        await pause(1500);
        const value = await api('/api/oauth/feishu/status', { state: attempt.state, verifier }, false);
        if (!value.pending) { login = value.session; profile = value.user; write(storageKey, login); break; }
      }
      if (!login) throw new Error(cancelled ? '已取消授权' : '授权等待超时，请重试');
    }
    popup?.close();
    // Display the actual OAuth account; do not assume plugin open_id belongs to this application.
    stage('正在生成文档', `授权账户：${profile.name} · ${payload.records.length} 条记录合并为一份文档`, 25);
    const saved = read(`${storageKey}:job`);
    let job;
    if (saved) {
      try { job = await api(`/api/feishu-docs/${saved}`); } catch (error) { if (![400, 404].includes(error.status)) throw error; write(`${storageKey}:job`, ''); }
      if (job && ['failed', 'complete'].includes(job.stage)) { write(`${storageKey}:job`, ''); job = null; }
    }
    if (!job) {
      const request = { ...payload, requestId: crypto.randomUUID() };
      // Retry the same idempotency key if the response is lost, never create a second document.
      try { job = await api('/api/feishu-docs', request); }
      catch (error) { if (error.status || controller.signal.aborted) throw error; job = await api('/api/feishu-docs', request); }
      write(`${storageKey}:job`, job.id);
    }
    const deadline = Date.now() + 10 * 60 * 1000;
    while (!['complete', 'failed'].includes(job.stage)) {
      if (Date.now() > deadline) throw new Error('飞书仍在处理，稍后再次点击可查询原任务；请勿重复创建');
      const stages = { rendering: ['正在填充模板', 30], uploading: ['正在上传文档', 55], importing: ['飞书正在转换文档', 80] };
      const [label, percent] = stages[job.stage] || ['正在处理', 30];
      stage(label, `归属账户：${job.ownerName || profile.name} · ${job.name || '正在整理内容'}`, percent);
      await pause(2000);
      job = await api(`/api/feishu-docs/${job.id}`);
    }
    write(`${storageKey}:job`, '');
    if (job.stage === 'failed') throw new Error(job.error || '导入失败');
    result.className = 'result generated';
    result.innerHTML = `<div class="generation-status"><strong>飞书文档已创建</strong><span>100%</span></div><div class="progress-track done"><i style="width:100%"></i></div><small>已保存到 ${escapeHtml(job.ownerName || profile.name)} 的个人云盘${job.warning ? ` · ${escapeHtml(job.warning)}` : ''}</small><div class="download-row"><span>${escapeHtml(job.name)}</span><a href="${escapeHtml(job.url)}" target="_blank" rel="noopener noreferrer">打开飞书文档</a></div>`;
    const actions = document.createElement('div'); actions.className = 'feishu-doc-actions';
    const copy = document.createElement('button'); copy.type = 'button'; copy.textContent = '复制文档链接';
    copy.onclick = async () => { try { await navigator.clipboard.writeText(job.url); copy.textContent = '已复制'; } catch { copy.textContent = '请右键打开文档链接复制'; } };
    const disconnect = document.createElement('button'); disconnect.type = 'button'; disconnect.textContent = '解除本插件授权';
    disconnect.onclick = async () => { if (!confirm('删除本插件保存的飞书文档授权？已创建的文档不会删除。')) return; try { await api('/api/oauth/feishu/disconnect', {}); write(storageKey, ''); disconnect.disabled = true; disconnect.textContent = '已解除，下次导出重新授权'; } catch (error) { disconnect.textContent = error.message; } };
    actions.append(copy, disconnect); result.append(actions);
  } catch (error) {
    if (error.status === 401) write(storageKey, '');
    result.hidden = false; result.className = 'result generation-error';
    result.innerHTML = `<strong>${cancelled ? '已取消' : '飞书文档未完成'}</strong><small>${escapeHtml(cancelled ? '没有创建文档。' : error.message || '网络异常，请重试')}</small>`;
    popup?.close();
  }
}
