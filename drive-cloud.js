/* NECD Submission — Google Drive connector via Google Apps Script.
   Uses a hidden HTML form to avoid cross-origin fetch restrictions. */
'use strict';
(() => {
  const ENDPOINT = 'https://script.google.com/macros/s/AKfycbzIauQtfs0ScLQqCR62-sOloFr77o_RBUFp8PsPfMkzWgoZiBMcOOcH75Exu73RA2ziTw/exec';
  const $ = id => document.getElementById(id);
  const pending = new Map();
  let bridge = null;
  let online = false;
  let sequence = Promise.resolve();

  function status(message, ok = false) {
    const el = $('cloudState');
    if (el) {
      el.textContent = message;
      el.dataset.ok = String(ok);
    }
  }

  function toBase64(bytes) {
    let binary = '';
    for (let i = 0; i < bytes.length; i += 8192) {
      binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
    }
    return btoa(binary);
  }

  function fromBase64(encoded) {
    const binary = atob(encoded);
    return Uint8Array.from(binary, character => character.charCodeAt(0));
  }

  function makeId() {
    return globalThis.crypto?.randomUUID?.() ||
      ('req-' + Date.now() + '-' + Math.random().toString(36).slice(2));
  }

  // The Apps Script response runs in a Google-hosted iframe and posts to this page.
  window.addEventListener('message', event => {
    if (!/^https:\/\/(?:[a-z0-9-]+\.)*(?:googleusercontent\.com|google\.com)$/.test(event.origin)) return;
    const data = event.data;
    if (!data || data.channel !== 'necd-drive-response') return;
    const id = String(data.id || '');
    const job = pending.get(id);
    if (!job) return;
    pending.delete(id);
    clearTimeout(job.timeout);
    if (data.error) job.reject(new Error(String(data.error)));
    else job.resolve(data.result);
  });

  function send(action, data = {}) {
    if (!bridge) return Promise.reject(new Error('Conexiunea Google Drive nu este inițializată.'));
    return new Promise((resolve, reject) => {
      const id = makeId();
      const timeout = setTimeout(() => {
        if (!pending.has(id)) return;
        pending.delete(id);
        reject(new Error('Google Drive nu a răspuns. Verifică publicarea noii versiuni Apps Script.'));
      }, 90000);
      pending.set(id, { resolve, reject, timeout });

      const form = document.createElement('form');
      form.method = 'POST';
      form.action = ENDPOINT;
      form.target = bridge.name;
      form.style.display = 'none';
      const input = document.createElement('input');
      input.type = 'hidden';
      input.name = 'payload';
      input.value = JSON.stringify({ id, action, data });
      form.appendChild(input);
      document.body.appendChild(form);
      try {
        form.submit();
      } catch (error) {
        clearTimeout(timeout);
        pending.delete(id);
        reject(error);
      } finally {
        setTimeout(() => form.remove(), 1000);
      }
    });
  }

  // A single target iframe is used, so requests must be serialized.
  function rpc(action, data = {}) {
    const task = sequence.then(() => send(action, data));
    sequence = task.catch(() => {});
    return task;
  }

  async function loadWorkspace() {
    const response = await rpc('load');
    if (!response || !response.template) throw new Error('Nu am primit șablonul Excel NECD din Google Drive.');
    await window.NECD_APP.loadPrivateTemplate(new Blob([fromBase64(response.template)]));
    const empty = { schema: 3, year: 2025, records: [], files: [], submissions: [], audit: [], manualSites: {} };
    window.NECD_APP.restoreWorkspace(response.workspace ? JSON.parse(response.workspace) : empty);
    window.NECD_APP.render();
  }

  async function connect() {
    status('Conectare la Google Drive…');
    bridge = document.createElement('iframe');
    bridge.name = 'necd_drive_' + makeId().replace(/[^a-zA-Z0-9_]/g, '_');
    bridge.title = 'NECD Google Drive';
    bridge.hidden = true;
    document.body.appendChild(bridge);
    try {
      await loadWorkspace();
      online = true;
      status('Sincronizat · Google Drive', true);
    } catch (error) {
      online = false;
      status('Conexiune nereușită');
      window.NECD_APP.notice(error.message || String(error), 'error');
    }
  }

  async function saveWorkspace() {
    if (!online) throw new Error('Google Drive nu este conectat.');
    status('Salvez progresul…');
    try {
      const response = await rpc('saveWorkspace', { json: JSON.stringify(window.NECD_APP.serializeWorkspace()) });
      // Merge server-side to preserve submissions from different computers.
      if (response?.workspace) window.NECD_APP.restoreWorkspace(JSON.parse(response.workspace));
      status('Salvat · Google Drive', true);
    } catch (error) {
      status('Salvarea a eșuat');
      throw error;
    }
  }

  async function archiveSource(file, id) {
    if (!online) throw new Error('Google Drive nu este conectat.');
    const extension = /\.csv$/i.test(file.name) ? 'csv' : 'xlsx';
    const bytes = new Uint8Array(await file.arrayBuffer());
    return rpc('saveSource', {
      id: String(id).replace(/[^a-zA-Z0-9_-]/g, '_'),
      extension,
      base64: toBase64(bytes)
    });
  }

  async function saveReport(blob, audit, year) {
    if (!online) throw new Error('Google Drive nu este conectat.');
    status('Salvez raportul…');
    try {
      const [bytes, auditText] = await Promise.all([
        blob.arrayBuffer().then(buffer => new Uint8Array(buffer)),
        typeof audit === 'string' ? Promise.resolve(audit) : audit.text()
      ]);
      await rpc('saveReport', { year, base64: toBase64(bytes), audit: auditText });
      status('Raport salvat · Google Drive', true);
    } catch (error) {
      status('Salvarea raportului a eșuat');
      throw error;
    }
  }

  async function refresh() {
    if (!online) return;
    if (window.NECD_APP.state.busy) throw new Error('O depunere este în curs.');
    await loadWorkspace();
    status('Sincronizat · Google Drive', true);
  }

  window.NECD_CLOUD = { isConnected: () => online, saveWorkspace, archiveSource, saveReport, refresh };
  document.addEventListener('DOMContentLoaded', connect);
})();
