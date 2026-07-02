const STORAGE_KEY = 'raceTrackerLocalEntries';
const ENDPOINT_KEY = 'raceTrackerEndpoint';
const SYNC_INTERVAL_MS = 1500;
let syncRunning = false;

function qs(name) {
  return new URLSearchParams(window.location.search).get(name);
}

function endpoint() {
  return localStorage.getItem(ENDPOINT_KEY) || '';
}

function getEntries() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); }
  catch { return []; }
}

function setEntries(entries) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(entries.slice(0, 500)));
}

function makeEntry(payload) {
  return {
    id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    status: endpoint() ? 'pending' : 'demo',
    attempts: 0,
    createdAt: new Date().toISOString(),
    ...payload
  };
}

function queueEntry(payload) {
  const entry = makeEntry(payload);
  const entries = getEntries();
  entries.unshift(entry);
  setEntries(entries);
  renderSyncStatus();
  // Start upload in the background. Do not await this.
  setTimeout(syncQueue, 0);
  return entry;
}

function updateEntry(id, patch) {
  const entries = getEntries();
  const updated = entries.map(e => e.id === id ? { ...e, ...patch } : e);
  setEntries(updated);
}

function formatClock(date = new Date()) {
  return date.toLocaleTimeString('en-GB', { hour12: false });
}

function setStatus() {
  const pill = document.getElementById('connectionStatus');
  if (!pill) return;
  if (endpoint()) {
    pill.textContent = 'Live sheet';
    pill.classList.add('live');
  } else {
    pill.textContent = 'Demo';
    pill.classList.remove('live');
  }
}

async function submitToSheet(payload) {
  const url = endpoint();
  if (!url) return { ok: true, demo: true };

  await fetch(url, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload)
  });

  // no-cors means we cannot read the Google response, so successful fetch completion = queued to Google.
  return { ok: true, live: true };
}

async function syncQueue() {
  if (syncRunning || !endpoint()) return;
  syncRunning = true;

  try {
    let entries = getEntries();
    const pending = entries
      .filter(e => e.status === 'pending' || e.status === 'failed')
      .reverse(); // upload oldest first

    for (const entry of pending) {
      updateEntry(entry.id, { status: 'uploading', attempts: (entry.attempts || 0) + 1 });
      renderRecentForCurrentPage();
      renderSyncStatus();

      try {
        await submitToSheet(entry);
        updateEntry(entry.id, { status: 'uploaded', uploadedAt: new Date().toISOString() });
      } catch (err) {
        updateEntry(entry.id, { status: 'failed', lastError: String(err) });
        break; // signal may be down; retry later
      }
    }
  } finally {
    syncRunning = false;
    renderRecentForCurrentPage();
    renderSyncStatus();
  }
}

function showMessage(text, type = '') {
  const el = document.getElementById('message');
  if (!el) return;
  el.className = 'message ' + type;
  el.textContent = text;
}

function statusLabel(entry) {
  if (entry.status === 'uploaded') return 'uploaded';
  if (entry.status === 'uploading') return 'syncing';
  if (entry.status === 'failed') return 'retrying';
  if (entry.status === 'pending') return 'queued';
  return 'demo';
}

function renderSyncStatus() {
  const pill = document.getElementById('connectionStatus');
  if (!pill) return;

  if (!endpoint()) {
    pill.textContent = 'Demo';
    pill.classList.remove('live');
    return;
  }

  const entries = getEntries();
  const pending = entries.filter(e => ['pending', 'uploading', 'failed'].includes(e.status)).length;
  pill.classList.add('live');
  pill.textContent = pending ? `Syncing ${pending}` : 'Live sheet';
}

function renderRecent(filterType = null, limit = 8) {
  const list = document.getElementById('recentList');
  if (!list) return;

  let entries = getEntries();
  if (filterType) entries = entries.filter(e => e.type === filterType);
  entries = entries.slice(0, limit);

  if (!entries.length) {
    list.innerHTML = '<p class="muted">No entries yet.</p>';
    return;
  }

  list.innerHTML = entries.map(e => {
    const main = e.bib ? `#${e.bib}` : e.label;
    const meta = [e.location, e.type, e.clockTime, statusLabel(e)].filter(Boolean).join(' · ');
    return `<div class="recent-item"><strong>${main}</strong><span class="recent-meta">${meta}</span></div>`;
  }).join('');
}

function currentMode() {
  const path = window.location.pathname.replace(/\/$/, '');
  if (path.includes('checkpoint.html') || path.endsWith('/checkpoint')) return 'checkpoint';
  if (path.includes('finish.html') || path.endsWith('/finish')) return 'finish-recorder';
  if (path.includes('timer.html') || path.endsWith('/timer')) return 'finish-timer';
  return null;
}

function renderRecentForCurrentPage() {
  const mode = currentMode();
  if (mode) renderRecent(mode);
}

function initKeypad(mode, locationName = '') {
  let value = '';
  const display = document.getElementById('display');

  function updateDisplay() {
    display.innerHTML = value || '&nbsp;';
  }

  function submit() {
    if (!value) return;

    // Capture immediately, then clear input immediately.
    const capturedAt = new Date();
    const bib = value;
    value = '';
    updateDisplay();

    const payload = {
      type: mode,
      location: locationName,
      bib,
      timestamp: capturedAt.toISOString(),
      clockTime: formatClock(capturedAt)
    };

    queueEntry(payload);
    renderRecent(mode);
    showMessage(`#${bib} queued`, 'good');
  }

  document.querySelectorAll('[data-key]').forEach(btn => {
    btn.addEventListener('click', () => {
      if (value.length >= 5) return;
      value += btn.dataset.key;
      updateDisplay();
    });
  });

  document.querySelectorAll('[data-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      if (btn.dataset.action === 'backspace') {
        value = value.slice(0, -1);
        updateDisplay();
      }
      if (btn.dataset.action === 'submit') submit();
    });
  });

  window.addEventListener('keydown', (e) => {
    if (/^[0-9]$/.test(e.key)) {
      if (value.length < 5) value += e.key;
      updateDisplay();
    }
    if (e.key === 'Backspace') {
      value = value.slice(0, -1);
      updateDisplay();
    }
    if (e.key === 'Enter') submit();
  });

  updateDisplay();
}

function initCheckpoint() {
  const locationName = qs('location') || 'Checkpoint';
  document.getElementById('pageTitle').textContent = locationName;
  initKeypad('checkpoint', locationName);
  renderRecent('checkpoint');
}

function initFinishRecorder() {
  initKeypad('finish-recorder', 'Finish');
  renderRecent('finish-recorder');
}

function initTimer() {
  const btn = document.getElementById('finishButton');
  btn.addEventListener('click', () => {
    // Capture immediately at tap time. Do not wait for Google Sheets.
    const capturedAt = new Date();
    const payload = {
      type: 'finish-timer',
      location: 'Finish',
      label: 'Finish time',
      timestamp: capturedAt.toISOString(),
      clockTime: formatClock(capturedAt)
    };

    queueEntry(payload);
    renderRecent('finish-timer');
    showMessage(`Finish queued: ${payload.clockTime}`, 'good');
  });
  renderRecent('finish-timer');
}

function initSettings() {
  const input = document.getElementById('endpoint');
  input.value = endpoint();

  document.getElementById('saveSettings').addEventListener('click', () => {
    localStorage.setItem(ENDPOINT_KEY, input.value.trim());
    showMessage('Settings saved', 'good');
    setStatus();
    syncQueue();
  });

  document.getElementById('clearSettings').addEventListener('click', () => {
    localStorage.removeItem(ENDPOINT_KEY);
    input.value = '';
    showMessage('Demo mode restored', 'good');
    setStatus();
  });
}

function initRecent() {
  renderRecent(null, 100);
  document.getElementById('clearLocal').addEventListener('click', () => {
    localStorage.removeItem(STORAGE_KEY);
    renderRecent(null, 100);
    renderSyncStatus();
  });
}

window.addEventListener('online', syncQueue);
setInterval(syncQueue, SYNC_INTERVAL_MS);

setStatus();
renderSyncStatus();

const path = window.location.pathname.replace(/\/$/, '');
if (path.includes('checkpoint.html') || path.endsWith('/checkpoint')) initCheckpoint();
if (path.includes('finish.html') || path.endsWith('/finish')) initFinishRecorder();
if (path.includes('timer.html') || path.endsWith('/timer')) initTimer();
if (path.includes('settings.html') || path.endsWith('/settings')) initSettings();
if (path.includes('recent.html') || path.endsWith('/recent')) initRecent();

// Try to flush anything left from earlier use.
syncQueue();
