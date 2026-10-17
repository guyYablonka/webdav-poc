// Page script for the demo site: Edit buttons, uploads, admin actions, and live updates over SSE.
(() => {
  const page = document.body.dataset.page;
  const fileId = document.body.dataset.fileId || null;
  if (page === 'login' || page === 'error') return;

  const formats = {
    time: new Intl.DateTimeFormat(undefined, { hour: '2-digit', minute: '2-digit' }),
    datetime: new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'medium' }),
  };

  function formatTimes(root = document) {
    for (const el of root.querySelectorAll('time[datetime]')) {
      const fmt = formats[el.dataset.fmt] || formats.datetime;
      el.textContent = fmt.format(new Date(el.getAttribute('datetime')));
    }
  }

  let toastTimer;
  function toast(message, kind = 'info') {
    const el = document.getElementById('toast');
    el.textContent = message;
    el.className = `toast ${kind}`;
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (el.hidden = true), 6000);
  }

  async function api(path, options = {}) {
    const res = await fetch(path, { ...options, headers: { 'x-requested-with': 'fetch', ...(options.headers || {}) } });
    if (res.status === 401) {
      location.assign('/login');
      throw new Error('Signed out');
    }
    if (!res.ok) {
      let message = res.statusText;
      try {
        const body = await res.json();
        message = body.message || body.error || message;
      } catch {}
      throw new Error(message);
    }
    return res.status === 204 ? null : res.json();
  }

  async function swap(id, url) {
    const el = document.getElementById(id);
    if (!el) return;
    const res = await fetch(url);
    if (res.status === 401) return location.assign('/login');
    if (!res.ok) return;
    el.innerHTML = await res.text();
    formatTimes(el);
  }

  // Refreshes are coalesced so a burst of events (LOCK + PUT + UNLOCK) costs one round trip.
  let pendingDetails = false;
  let pendingPreview = false;
  let scheduled = null;
  function scheduleRefresh({ preview = false } = {}) {
    pendingDetails = true;
    pendingPreview ||= preview;
    if (scheduled) return;
    scheduled = setTimeout(async () => {
      const withPreview = pendingPreview;
      pendingDetails = pendingPreview = false;
      scheduled = null;
      if (page === 'list') await swap('file-rows', '/partials/files');
      if (page === 'file') {
        await Promise.all([
          swap('file-details', `/partials/files/${fileId}`),
          withPreview ? swap('preview', `/files/${fileId}/preview`) : null,
        ]);
      }
    }, 150);
  }

  function flash(selector) {
    const el = document.querySelector(selector);
    if (!el) return;
    el.classList.remove('flash');
    void el.offsetWidth;
    el.classList.add('flash');
  }

  const DESCRIBE = {
    'version.created': (e) => `v${e.versionNo} saved by ${e.userId} (${e.source})`,
    'lock.acquired': (e) => `locked by ${e.userId}`,
    'lock.refreshed': (e) => `lock refreshed by ${e.userId}`,
    'lock.revived': (e) => `expired lock revived for ${e.userId}`,
    'lock.released': (e) => `lock released (${e.reason.replace('_', ' ')})`,
    'session.revoked': (e) => `DAV session of ${e.userId} revoked`,
    'file.created': (e) => `file created by ${e.userId}`,
  };

  function logActivity(e) {
    const list = document.getElementById('activity');
    if (!list) return;
    list.querySelector('.muted')?.remove();
    const item = document.createElement('li');
    const at = document.createElement('time');
    at.dateTime = e.at;
    at.dataset.fmt = 'datetime';
    item.append(at, ' ', (DESCRIBE[e.type] || (() => e.type))(e));
    item.className = e.type.replace('.', '-');
    list.prepend(item);
    formatTimes(item);
    while (list.children.length > 30) list.lastElementChild.remove();
  }

  function onEvent(e) {
    if (page === 'list') {
      scheduleRefresh();
      if (e.type === 'version.created') setTimeout(() => flash(`tr[data-file-id="${e.fileId}"]`), 400);
      return;
    }
    if (page === 'file' && e.fileId === fileId) {
      logActivity(e);
      scheduleRefresh({ preview: e.type === 'version.created' });
      if (e.type === 'version.created') setTimeout(() => flash(`tr[data-version-id="${e.versionId}"]`), 400);
    }
  }

  // Live updates --------------------------------------------------------------

  const live = document.getElementById('live');
  const events = new EventSource('/events');
  events.onopen = () => live?.classList.add('on');
  events.onerror = () => live?.classList.remove('on');
  for (const type of Object.keys(DESCRIBE)) {
    events.addEventListener(type, (message) => onEvent(JSON.parse(message.data)));
  }
  // Locks lapse without an event; re-render now and then so badges stay honest.
  setInterval(() => scheduleRefresh(), 30_000);

  // Actions -----------------------------------------------------------------

  document.addEventListener('click', async (ev) => {
    const button = ev.target.closest('button[data-action]');
    if (!button) return;
    ev.preventDefault();
    try {
      switch (button.dataset.action) {
        case 'edit': {
          const lockedBy = button.dataset.lockedBy;
          if (lockedBy && !confirm(`${lockedBy} is editing this file. Office will offer to open it read-only. Continue?`)) return;
          const session = await api(`/api/files/${button.dataset.fileId}/edit-session`, { method: 'POST' });
          toast(`Opening in ${session.app}… If nothing happens, check that Office desktop is installed.`);
          window.location.href = session.uri;
          break;
        }
        case 'force-unlock':
          await api(`/api/files/${button.dataset.fileId}/force-unlock`, { method: 'POST' });
          toast('Lock removed. The editor’s next save will fail.');
          scheduleRefresh();
          break;
        case 'revoke':
          await api(`/api/sessions/${button.dataset.sessionId}/revoke`, { method: 'POST' });
          toast('DAV session revoked. Its next request gets 401.');
          scheduleRefresh();
          break;
      }
    } catch (err) {
      toast(err.message, 'error');
    }
  });

  document.addEventListener('change', async (ev) => {
    const input = ev.target;
    if (!(input instanceof HTMLInputElement) || input.type !== 'file' || !input.files?.length) return;
    const form = new FormData();
    form.append('file', input.files[0]);
    const url = input.dataset.uploadVersion ? `/api/files/${input.dataset.uploadVersion}/versions` : '/api/files';
    try {
      await api(url, { method: 'POST', body: form });
      toast('Uploaded.');
    } catch (err) {
      toast(`Upload failed: ${err.message}`, 'error');
    } finally {
      input.value = '';
    }
  });

  formatTimes();
})();
