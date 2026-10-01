'use strict';

/* Classer dans Drive : complément Outlook.
   Tout se passe dans le navigateur : Outlook -> OneDrive (Microsoft Graph), sans serveur intermédiaire. */

const GRAPH = 'https://graph.microsoft.com/v1.0';
const SIMPLE_MAX = 4 * 1024 * 1024;
const CHUNK = 320 * 1024 * 16;
const ROOT = { id: 'root', name: 'OneDrive', driveId: null };
const LS = { token: 'cd.mstoken', root: 'cd.msroot', mem: 'cd.msmemory', opts: 'cd.options', cid: 'cd.msClientId', manual: 'cd.manualAccount' };
const PROP = 'cdFiled';
const CAT_NAME = 'Classé';
const OPTS = ['optEml', 'optPdf', 'optPrefix', 'optSub', 'optTag'];

const state = {
  item: null, token: null, tokenExp: 0, email: '',
  root: null, path: [], dest: null, currentList: [],
  pca: null, naa: false, searchTimer: null, busy: false
};

const $ = (id) => document.getElementById(id);
const show = (id, on = true) => $(id).classList.toggle('hidden', !on);
const escHtml = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const escQ = (s) => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");

function lsGet(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* stockage indisponible */ } }

class AuthError extends Error {
  constructor() { super('Session Microsoft expirée : reconnecte-toi.'); this.name = 'AuthError'; }
}

function officeCall(fn) {
  return new Promise((resolve, reject) => {
    fn((r) => r.status === Office.AsyncResultStatus.Succeeded
      ? resolve(r.value)
      : reject(new Error((r.error && r.error.message) || 'Erreur Outlook')));
  });
}

/* ---------- Démarrage ---------- */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    document.body.textContent = 'Ce complément fonctionne uniquement dans Outlook.';
    return;
  }
  const t = lsGet(LS.token, null);
  if (t && t.exp > Date.now() + 60000) { state.token = t.token; state.tokenExp = t.exp; }
  state.root = lsGet(LS.root, null) || ROOT;
  state.path = [state.root];

  restoreOptions();
  bindUI();

  if (Office.context.requirements.isSetSupported('Mailbox', '1.5')) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      if (state.busy) { state.reloadAfterSave = true; return; }
      loadItem();
    });
  }
  if (!Office.context.requirements.isSetSupported('Mailbox', '1.14')) {
    const o = $('optEml'); o.checked = false; o.disabled = true;
    o.nextElementSibling.textContent = 'Mail original (.eml) : non disponible dans cette version d\'Outlook';
  }

  if (!canTag()) { const o = $('optTag'); o.checked = false; o.disabled = true; }
  $('redirect1').textContent = 'brk-multihub://' + location.host;
  $('redirect2').textContent = new URL('auth.html', location.href).href.split('?')[0];
  if (!clientId()) show('cidRow', true);

  loadItem();
  startAuth();
});

function clientId() { return CONFIG.CLIENT_ID || lsGet(LS.cid, ''); }

function canTag() { return Office.context.requirements.isSetSupported('Mailbox', '1.8'); }

function bindUI() {
  $('btnConnect').addEventListener('click', connect);
  $('btnLogout').addEventListener('click', logout);
  $('btnAuto').addEventListener('click', useOutlookAccount);
  $('btnSave').addEventListener('click', save);
  $('btnNewFolder').addEventListener('click', () => { show('newFolderRow', true); $('newFolderName').focus(); });
  $('btnCreate').addEventListener('click', createFolderHere);
  $('newFolderName').addEventListener('keydown', (e) => { if (e.key === 'Enter') createFolderHere(); });
  $('btnSetRoot').addEventListener('click', setRoot);
  $('search').addEventListener('input', onSearchInput);
  $('folders').addEventListener('click', onFolderClick);
  $('crumbs').addEventListener('click', onCrumbClick);
  OPTS.forEach((id) => $(id).addEventListener('change', saveOptions));
}

function restoreOptions() {
  const o = lsGet(LS.opts, {});
  OPTS.forEach((id) => { if (typeof o[id] === 'boolean') $(id).checked = o[id]; });
}
function saveOptions() {
  const o = {};
  OPTS.forEach((id) => { o[id] = $(id).checked; });
  lsSet(LS.opts, o);
}

/* ---------- Mail courant ---------- */

function loadItem() {
  state.item = Office.context.mailbox.item;
  selectDest(null);
  setStatus('');
  if (!state.item) {
    show('main', false); show('footer', false); show('empty', true);
    return;
  }
  show('empty', false); show('main', true); show('footer', true);
  renderMail();
  renderFiled();
}

function correspondent(item) {
  const me = ((Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || '').toLowerCase();
  const from = item.from || item.sender;
  if (from && from.emailAddress && from.emailAddress.toLowerCase() !== me) return from;
  return (item.to && item.to[0]) || from || { displayName: '', emailAddress: '' };
}

function cleanSubject(s) {
  return String(s || '').replace(/^\s*((re|tr|fw|fwd|wg|aw)\s*:\s*)+/i, '').trim();
}

function isoDate(d) {
  d = d instanceof Date ? d : new Date(d || Date.now());
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function sanitize(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '')
    .slice(0, 150);
}

function fmtSize(n) {
  if (!n && n !== 0) return '';
  if (n < 1024) return n + ' o';
  if (n < 1048576) return Math.round(n / 1024) + ' Ko';
  return (n / 1048576).toFixed(1).replace('.', ',') + ' Mo';
}

function renderMail() {
  const item = state.item;
  const who = correspondent(item);
  const whoName = who.displayName || (who.emailAddress || '').split('@')[0] || 'Inconnu';
  const date = item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date();

  $('mSubject').textContent = item.subject || '(sans objet)';
  $('mMeta').textContent = `${whoName}${who.emailAddress ? ' <' + who.emailAddress + '>' : ''}, ${date.toLocaleString('fr-FR', { dateStyle: 'medium', timeStyle: 'short' })}`;
  $('baseName').value = sanitize(`${isoDate(date)} - ${whoName} - ${cleanSubject(item.subject) || 'sans objet'}`);

  const atts = item.attachments || [];
  const cloud = Office.MailboxEnums && Office.MailboxEnums.AttachmentType ? Office.MailboxEnums.AttachmentType.Cloud : 'cloud';
  $('attList').innerHTML = atts.length
    ? atts.map((a, i) => {
        const isCloud = a.attachmentType === cloud;
        const note = isCloud ? ' (lien cloud, non copiable)' : a.isInline ? ' (image intégrée)' : '';
        return `<label class="check"><input type="checkbox" data-i="${i}" ${a.isInline || isCloud ? '' : 'checked'} ${isCloud ? 'disabled' : ''}>
          <span class="name">${escHtml(a.name)}${escHtml(note)}</span><span class="size">${fmtSize(a.size)}</span></label>`;
      }).join('')
    : '<p class="muted small">Aucune pièce jointe.</p>';
}

function selectedAttachments() {
  const atts = state.item.attachments || [];
  return [...$('attList').querySelectorAll('input[type=checkbox]:checked')].map((c) => atts[Number(c.dataset.i)]);
}

/* ---------- Connexion Microsoft ---------- */

function hasToken(marginMs = 60000) { return !!state.token && state.tokenExp > Date.now() + marginMs; }

// Authentification imbriquée (NAA) : Outlook fournit le jeton sans fenêtre quand c'est possible.
async function initMsal() {
  if (state.pca || !clientId()) return;
  if (lsGet(LS.manual, false)) return;
  if (Office.context.requirements.isSetSupported('NestedAppAuth', '1.1') && msal.createNestablePublicClientApplication) {
    try {
      state.pca = await msal.createNestablePublicClientApplication({ auth: { clientId: clientId(), authority: CONFIG.AUTHORITY } });
      state.naa = true;
    } catch (e) { console.warn('NAA indisponible', e); state.pca = null; state.naa = false; }
  }
}

async function startAuth() {
  await initMsal();
  try { await getToken(false); onConnected(); }
  catch (e) { show('auth', true); show('autoRow', lsGet(LS.manual, false)); }
}

async function getToken(interactive) {
  if (state.naa) {
    const req = { scopes: CONFIG.SCOPES };
    try { return (await state.pca.acquireTokenSilent(req)).accessToken; }
    catch (e) {
      if (!interactive) throw new AuthError();
      return (await state.pca.acquireTokenPopup(req)).accessToken;
    }
  }
  if (hasToken()) return state.token;
  if (!interactive) throw new AuthError();
  await authorizeDialog();
  return state.token;
}

// Solution de secours : fenêtre de connexion Office (versions d'Outlook sans NAA).
function authorizeDialog() {
  return new Promise((resolve, reject) => {
    const url = new URL('auth.html', location.href);
    url.search = '?cid=' + encodeURIComponent(clientId());
    Office.context.ui.displayDialogAsync(url.href, { height: 60, width: 35 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error('Impossible d\'ouvrir la fenêtre de connexion : ' + res.error.message));
        return;
      }
      const dlg = res.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dlg.close();
        let m = {};
        try { m = JSON.parse(arg.message); } catch (e) { /* message illisible */ }
        if (!m.token) { reject(new Error('Connexion refusée' + (m.error ? ' : ' + m.error : ''))); return; }
        state.token = m.token;
        state.tokenExp = m.exp || Date.now() + 3300000;
        lsSet(LS.token, { token: state.token, exp: state.tokenExp });
        resolve(m.token);
      });
      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => reject(new Error('Fenêtre de connexion fermée.')));
    });
  });
}

async function connect() {
  if (!clientId()) {
    const v = $('clientId').value.trim();
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v)) {
      setStatus('Colle l\'ID d\'application (client) : 36 caractères, du type 1a2b3c4d-....', 'warn');
      return;
    }
    lsSet(LS.cid, v);
    show('cidRow', false);
  }
  try {
    setStatus('Connexion à Microsoft...');
    await initMsal();
    await getToken(true);
    setStatus('');
    onConnected();
  } catch (e) {
    setStatus(e.message || String(e), 'err');
  }
}

// Changer de compte : on quitte la connexion automatique d'Outlook et on passe par la fenêtre
// de connexion, qui propose le choix du compte. Le dossier racine appartenait à l'ancien
// OneDrive : il est réinitialisé.
function logout() {
  state.token = null; state.tokenExp = 0; state.email = ''; state.connected = false;
  lsSet(LS.token, null);
  lsSet(LS.manual, true);
  state.pca = null; state.naa = false;
  state.root = ROOT; state.path = [ROOT];
  lsSet(LS.root, null);
  show('autoRow', true);
  show('account', false); show('pick', false); show('auth', true);
  selectDest(null);
}

async function useOutlookAccount() {
  lsSet(LS.manual, false);
  show('autoRow', false);
  await connect();
}

async function onConnected() {
  state.connected = true;
  show('auth', false); show('pick', true);
  try {
    const r = await (await graph('/me?$select=userPrincipalName,mail')).json();
    state.email = r.mail || r.userPrincipalName || '';
    $('accountEmail').textContent = state.email;
    show('account', !!state.email);
  } catch (e) { handleError(e); return; }
  openPath(state.path.length ? state.path : [state.root]);
}

/* ---------- Appels Microsoft Graph (OneDrive) ---------- */

async function graph(path, opts = {}) {
  const token = await getToken(false);
  const url = path.startsWith('http') ? path : GRAPH + path;
  const r = await fetch(url, { ...opts, headers: { Authorization: 'Bearer ' + token, ...(opts.headers || {}) } });
  if (r.status === 401) { state.token = null; lsSet(LS.token, null); throw new AuthError(); }
  if (!r.ok) {
    let msg = r.status + '';
    try { const j = await r.json(); msg = (j.error && j.error.message) || msg; } catch (e) { /* corps vide */ }
    throw new Error('OneDrive : ' + msg);
  }
  return r;
}

const itemPath = (f) => (f.id === 'root' ? '/me/drive/root' : `/drives/${f.driveId}/items/${f.id}`);
const SELECT = '$select=id,name,folder,parentReference,webUrl';

function parentFromPath(p) {
  if (!p) return '';
  const after = p.split(':').slice(1).join(':');
  const segs = decodeURIComponent(after || '').split('/').filter(Boolean);
  return segs.length ? segs[segs.length - 1] : 'OneDrive';
}

function toFolder(it) {
  const pr = it.parentReference || {};
  return { id: it.id, name: it.name, driveId: pr.driveId || null, parentName: parentFromPath(pr.path), webUrl: it.webUrl || '' };
}

async function listFolders(parent) {
  const out = [];
  let url = `${itemPath(parent)}/children?${SELECT}&$top=999`;
  while (url) {
    const r = await (await graph(url)).json();
    (r.value || []).filter((x) => x.folder).forEach((x) => out.push(toFolder(x)));
    url = r['@odata.nextLink'] || '';
  }
  return out.sort((a, b) => a.name.localeCompare(b.name, 'fr', { numeric: true, sensitivity: 'base' }));
}

async function listSharedFolders() {
  try {
    const r = await (await graph('/me/drive/sharedWithMe?$top=200')).json();
    return (r.value || []).filter((x) => x.remoteItem && x.remoteItem.folder).map((x) => {
      const f = toFolder(x.remoteItem);
      f.name = x.name || f.name; f.shared = true; f.webUrl = f.webUrl || x.webUrl || '';
      return f;
    });
  } catch (e) {
    if (e instanceof AuthError) throw e;
    return [];
  }
}

async function searchFolders(text, limit = 30) {
  const q = String(text).replace(/'/g, "''");
  const r = await (await graph(`/me/drive/root/search(q='${encodeURIComponent(q)}')?${SELECT}&$top=${Math.max(limit * 3, 50)}`)).json();
  const tl = text.toLowerCase();
  return (r.value || []).filter((x) => x.folder && x.name.toLowerCase().includes(tl)).slice(0, limit).map(toFolder);
}

async function parentName(folder) { return folder.parentName || ''; }

async function createFolder(name, parent) {
  const r = await graph(`${itemPath(parent)}/children`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' })
  });
  return toFolder(await r.json());
}

async function upload(name, blob, parent) {
  const target = `${itemPath(parent)}:/${encodeURIComponent(name)}:`;
  if (blob.size <= SIMPLE_MAX) {
    return (await graph(`${target}/content?@microsoft.graph.conflictBehavior=rename`, {
      method: 'PUT', headers: { 'Content-Type': blob.type || 'application/octet-stream' }, body: blob
    })).json();
  }
  // Gros fichiers : session d'envoi par morceaux
  const s = await (await graph(`${target}/createUploadSession`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'rename' } })
  })).json();
  let r = null;
  for (let start = 0; start < blob.size; start += CHUNK) {
    const end = Math.min(start + CHUNK, blob.size);
    r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${start}-${end - 1}/${blob.size}` }, body: blob.slice(start, end) });
    if (!r.ok) throw new Error('Envoi interrompu (' + r.status + ')');
  }
  return r.json();
}

async function convertToPdf(it) {
  const driveId = it.parentReference && it.parentReference.driveId;
  const r = await graph(`/drives/${driveId}/items/${it.id}/content?format=pdf`);
  const b = await r.blob();
  if (!b.size) throw new Error('conversion PDF vide');
  return new Blob([b], { type: 'application/pdf' });
}

async function deleteItem(it) {
  const driveId = it.parentReference && it.parentReference.driveId;
  await graph(`/drives/${driveId}/items/${it.id}`, { method: 'DELETE' });
}

/* ---------- Navigation dans les dossiers ---------- */

async function openPath(path) {
  state.path = path;
  $('search').value = '';
  renderCrumbs();
  $('folders').innerHTML = '<li class="info">Chargement...</li>';
  try {
    const cur = path[path.length - 1];
    let list = await listFolders(cur);
    if (path.length === 1 && cur.id === 'root') list = list.concat(await listSharedFolders());
    if (state.path !== path) return;
    renderFolderList(list, false);
  } catch (e) { handleError(e); }
}

function renderCrumbs() {
  $('crumbs').innerHTML = state.path.map((f, i) =>
    (i ? '<span class="sep">/</span>' : '') + `<button data-i="${i}">${escHtml(f.name)}</button>`).join('');
}

function onCrumbClick(e) {
  const b = e.target.closest('button[data-i]');
  if (!b) return;
  openPath(state.path.slice(0, Number(b.dataset.i) + 1));
}

async function renderFolderList(list, withParents) {
  state.currentList = list;
  const ul = $('folders');
  if (!list.length) {
    ul.innerHTML = `<li class="info">${withParents ? 'Aucun dossier trouvé.' : 'Aucun sous-dossier. Tu peux choisir le dossier courant ou en créer un.'}</li>`;
    if (!withParents) prependCurrentFolderRow();
    return;
  }
  ul.innerHTML = list.map((f, i) => `
    <li data-i="${i}" class="${state.dest && state.dest.id === f.id ? 'sel' : ''}">
      <button class="f-name" data-act="pick" title="Choisir ce dossier">${escHtml(f.name)}${f.shared ? '<span class="f-sub">Partagé avec moi</span>' : ''}<span class="f-sub" data-parent="${i}"></span></button>
      <button class="f-open" data-act="open" title="Ouvrir" aria-label="Ouvrir ${escHtml(f.name)}">›</button>
    </li>`).join('');
  if (!withParents) prependCurrentFolderRow();
  if (withParents) {
    for (let i = 0; i < list.length; i++) {
      const name = await parentName(list[i]);
      const el = ul.querySelector(`[data-parent="${i}"]`);
      if (el && name) el.textContent = 'dans ' + name;
    }
  }
}

function prependCurrentFolderRow() {
  const cur = state.path[state.path.length - 1];
  const li = document.createElement('li');
  li.dataset.current = '1';
  if (state.dest && state.dest.id === cur.id) li.className = 'sel';
  li.innerHTML = `<button class="f-name" data-act="pick-current" title="Choisir ce dossier"><span class="f-sub">Choisir le dossier courant</span>${escHtml(cur.name)}</button>`;
  $('folders').prepend(li);
}

function onFolderClick(e) {
  const btn = e.target.closest('button[data-act]');
  if (!btn) return;
  const act = btn.dataset.act;
  if (act === 'pick-current') { const cur = state.path[state.path.length - 1]; selectDest(cur); return; }
  const f = state.currentList[Number(btn.closest('li').dataset.i)];
  if (!f) return;
  if (act === 'pick') selectDest(f);
  if (act === 'open') {
    if ($('search').value.trim()) openPath([state.root, f]);
    else openPath(state.path.concat([f]));
  }
}

function onSearchInput() {
  clearTimeout(state.searchTimer);
  const text = $('search').value.trim();
  state.searchTimer = setTimeout(async () => {
    if (!text) { openPath(state.path); return; }
    if (/^https?:\/\//i.test(text)) { openLink(text); return; }
    $('folders').innerHTML = '<li class="info">Recherche...</li>';
    try {
      const res = await searchFolders(text);
      if ($('search').value.trim() !== text) return;
      renderFolderList(res, true);
    } catch (e) { handleError(e); }
  }, 300);
}

async function createFolderHere() {
  const name = sanitize($('newFolderName').value);
  if (!name) return;
  const cur = state.path[state.path.length - 1];
  try {
    const f = await createFolder(name, cur);
    $('newFolderName').value = '';
    show('newFolderRow', false);
    await openPath(state.path);
    selectDest(f);
  } catch (e) { handleError(e); }
}

function setRoot() {
  const cur = state.path[state.path.length - 1];
  state.root = { ...cur };
  lsSet(LS.root, state.root);
  openPath([state.root]);
  setStatus(`Racine définie : ${cur.name}`, 'ok');
}

function selectDest(f) {
  state.dest = f ? { ...f } : null;
  const d = $('dest');
  if (state.dest) { d.className = 'tab'; d.innerHTML = 'Destination : <b>' + escHtml(state.dest.name) + '</b>'; }
  else { d.className = 'tab empty'; d.textContent = 'Aucun dossier choisi'; }
  $('btnSave').disabled = !state.dest || state.busy;
  document.querySelectorAll('#folders li').forEach((li) => {
    const lf = li.dataset.current ? state.path[state.path.length - 1] : state.currentList[Number(li.dataset.i)];
    li.classList.toggle('sel', !!(state.dest && lf && lf.id === state.dest.id));
  });
}


/* ---------- Dossier désigné par un lien OneDrive collé ---------- */

function shareId(url) {
  const b64 = btoa(unescape(encodeURIComponent(url))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  return 'u!' + b64;
}

async function resolveLink(raw) {
  const attempts = [];
  let u = null;
  try { u = new URL(raw); } catch (e) { throw new Error('Lien illisible.'); }
  const id = u.searchParams.get('id');
  const cid = u.searchParams.get('cid');
  // OneDrive personnel : ...onedrive.live.com/?id=XXX!123&cid=YYY
  if (id && cid && !id.startsWith('/')) attempts.push(`/drives/${encodeURIComponent(cid)}/items/${encodeURIComponent(id)}`);
  // OneDrive / SharePoint pro, barre d'adresse : .../my?id=%2Fpersonal%2F...%2FDocuments%2FDossier
  if (id && id.startsWith('/')) attempts.push(`/shares/${shareId(u.origin + id)}/driveItem`);
  // Lien de partage ("Copier le lien") ou adresse directe du dossier
  attempts.push(`/shares/${shareId(raw)}/driveItem`);
  let lastErr = null;
  for (const a of attempts) {
    try {
      const it = await (await graph(a + '?' + SELECT)).json();
      return it;
    } catch (e) {
      if (e instanceof AuthError) throw e;
      lastErr = e;
    }
  }
  throw new Error('Dossier introuvable à partir de ce lien' + (lastErr ? ' (' + lastErr.message + ')' : '') + '.');
}

async function openLink(raw) {
  $('folders').innerHTML = '<li class="info">Ouverture du lien...</li>';
  try {
    const it = await resolveLink(raw);
    if (!it.folder) throw new Error('Ce lien désigne un fichier, pas un dossier.');
    const f = toFolder(it);
    if ($('search').value.trim() !== raw) return;
    $('search').value = '';
    await openPath([state.root, f]);
    selectDest(f);
  } catch (e) {
    handleError(e);
    $('folders').innerHTML = '<li class="info">' + escHtml(e.message) + '</li>';
  }
}

/* ---------- Enregistrement ---------- */

function b64ToBlob(b64, type) {
  const bin = atob(String(b64).replace(/\s/g, ''));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type });
}

function fmtAddr(list) {
  return (list || []).filter(Boolean).map((x) =>
    x.displayName && x.displayName !== x.emailAddress ? `${x.displayName} <${x.emailAddress}>` : x.emailAddress).join(', ');
}

function buildMailHtml(item, bodyHtml) {
  const doc = new DOMParser().parseFromString(bodyHtml || '', 'text/html');
  doc.querySelectorAll('script, img[src^="cid:"]').forEach((n) => n.remove());
  const date = item.dateTimeCreated ? new Date(item.dateTimeCreated) : null;
  const rows = [
    ['De', fmtAddr([item.from || item.sender])],
    ['À', fmtAddr(item.to)],
    ['Cc', fmtAddr(item.cc)],
    ['Date', date ? date.toLocaleString('fr-FR', { dateStyle: 'full', timeStyle: 'short' }) : ''],
    ['Objet', item.subject || ''],
    ['Pièces jointes', (item.attachments || []).filter((a) => !a.isInline).map((a) => a.name).join(', ')]
  ].filter((r) => r[1]);
  const head = rows.map(([k, v]) =>
    `<tr><td style="padding:2pt 8pt 2pt 0;vertical-align:top;font-weight:bold;white-space:nowrap">${k} :</td><td style="padding:2pt 0">${escHtml(v)}</td></tr>`).join('');
  return `<!doctype html><html><head><meta charset="utf-8"></head><body style="font-family:Arial,sans-serif;font-size:10pt">
<table style="border-collapse:collapse;margin-bottom:10pt">${head}</table><hr>${doc.body ? doc.body.innerHTML : ''}</body></html>`;
}

async function savePdf(item, base, target, emlItem) {
  let pdf = null;
  let tmp = null;
  // 1) conversion directe du .eml par OneDrive (garde les images intégrées)
  if (emlItem) { try { pdf = await convertToPdf(emlItem); } catch (e) { if (e instanceof AuthError) throw e; } }
  // 2) sinon : page HTML temporaire avec en-tête, convertie puis supprimée
  if (!pdf) {
    const body = await officeCall((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
    tmp = await upload(base + ' (conversion).html', new Blob([buildMailHtml(item, body)], { type: 'text/html' }), target);
    try { pdf = await convertToPdf(tmp); }
    finally { await deleteItem(tmp).catch(() => {}); }
  }
  await upload(base + '.pdf', pdf, target);
}

async function saveAttachment(item, a, target, datePrefix) {
  const c = await officeCall((cb) => item.getAttachmentContentAsync(a.id, cb));
  const F = Office.MailboxEnums.AttachmentContentFormat;
  let name = sanitize(a.name) || 'piece-jointe';
  let blob;
  switch (c.format) {
    case F.Base64: blob = b64ToBlob(c.content, a.contentType || 'application/octet-stream'); break;
    case F.Eml: blob = new Blob([c.content], { type: 'message/rfc822' }); if (!/\.eml$/i.test(name)) name += '.eml'; break;
    case F.ICalendar: blob = new Blob([c.content], { type: 'text/calendar' }); if (!/\.ics$/i.test(name)) name += '.ics'; break;
    default: throw new Error('pièce jointe cloud (lien), non copiée');
  }
  if (datePrefix) name = `${datePrefix} - ${name}`;
  await upload(name, blob, target);
  return name;
}

async function save() {
  if (state.busy || !state.dest || !state.item) return;
  try { await getToken(true); } catch (e) { setStatus(e.message, 'err'); return; }
  const item = state.item;
  const dest = state.dest;
  const base = sanitize($('baseName').value) || 'Mail';
  const date = isoDate(item.dateTimeCreated);
  const atts = selectedAttachments();
  const wantEml = $('optEml').checked && !$('optEml').disabled;
  const wantPdf = $('optPdf').checked;
  if (!wantEml && !wantPdf && !atts.length) { setStatus('Coche au moins un élément à enregistrer.', 'warn'); return; }

  state.busy = true;
  $('btnSave').disabled = true;
  let done = 0;
  const failures = [];
  const step = (m) => setStatus(m + '... (ne change pas de mail pendant l\'envoi)');

  try {
    let target = dest;
    let emlItem = null;
    if ($('optSub').checked) {
      step('Création du sous-dossier');
      target = await createFolder(base, dest);
    }
    if (wantEml) {
      step('Mail original');
      try {
        const b64 = await officeCall((cb) => item.getAsFileAsync(cb));
        emlItem = await upload(base + '.eml', b64ToBlob(b64, 'message/rfc822'), target);
        done++;
      } catch (e) { if (e instanceof AuthError) throw e; failures.push('.eml (' + e.message + ')'); }
    }
    if (wantPdf) {
      step('Mail en PDF');
      try { await savePdf(item, base, target, emlItem); done++; }
      catch (e) { if (e instanceof AuthError) throw e; failures.push('PDF (' + e.message + ')'); }
    }
    for (let i = 0; i < atts.length; i++) {
      step(`Pièce jointe ${i + 1}/${atts.length}`);
      try { await saveAttachment(item, atts[i], target, $('optPrefix').checked ? date : ''); done++; }
      catch (e) { if (e instanceof AuthError) throw e; failures.push(atts[i].name + ' (' + e.message + ')'); }
    }

    let tagNote = '';
    if (done) {
      try { await markFiled(item, dest, target); }
      catch (e) { tagNote = '<br>Fichiers enregistrés, mais le marquage dans Outlook a échoué (' + escHtml(e.message) + ').'; }
    }
    const url = target.webUrl || dest.webUrl || '';
    const msg = `${done} fichier(s) enregistré(s) dans <b>${escHtml(dest.name)}</b>.` + (url ? ` <a href="${escHtml(url)}" target="_blank" rel="noopener">Ouvrir le dossier</a>` : '') +
      (failures.length ? `<br>Échec : ${escHtml(failures.join(' ; '))}` : '') + tagNote;
    setStatus(msg, failures.length || tagNote ? 'warn' : 'ok', true);
    if (state.item === item) renderFiled();
  } catch (e) {
    handleError(e);
  } finally {
    state.busy = false;
    $('btnSave').disabled = !state.dest;
    if (state.reloadAfterSave) {
      state.reloadAfterSave = false;
      const keep = $('status').innerHTML; const cls = $('status').className;
      loadItem();
      $('status').innerHTML = keep; $('status').className = cls;
    }
  }
}

/* ---------- Marquage du mail classé (v2) ---------- */

async function readFiled(item) {
  try {
    const props = await officeCall((cb) => item.loadCustomPropertiesAsync(cb));
    const v = props.get(PROP);
    return { props, list: v ? JSON.parse(v) : [] };
  } catch (e) {
    return { props: null, list: [] };
  }
}

async function markFiled(item, dest, target) {
  const label = dest.parentName ? `${dest.parentName} / ${dest.name}` : dest.name;
  const { props, list } = await readFiled(item);
  if (props) {
    list.push({ id: dest.id, url: target.webUrl || dest.webUrl || '', name: label, date: new Date().toISOString() });
    props.set(PROP, JSON.stringify(list.slice(-10)));
    await officeCall((cb) => props.saveAsync(cb));
  }
  if ($('optTag').checked && canTag()) {
    const name = CAT_NAME;
    const master = await officeCall((cb) => Office.context.mailbox.masterCategories.getAsync(cb)).catch(() => []);
    if (!(master || []).some((c) => c.displayName === name)) {
      await officeCall((cb) => Office.context.mailbox.masterCategories.addAsync([{ displayName: name, color: Office.MailboxEnums.CategoryColor.Preset4 }], cb))
        .catch(() => { /* déjà créée ailleurs */ });
    }
    await officeCall((cb) => item.categories.addAsync([name], cb));
  }
}

async function renderFiled() {
  const item = state.item;
  const box = $('filed');
  if (!item) { show('filed', false); return; }
  let { list } = await readFiled(item);
  if (!list.length && canTag()) {
    try {
      const cats = await officeCall((cb) => item.categories.getAsync(cb));
      if ((cats || []).some((c) => c.displayName === CAT_NAME)) list = [{ name: '' }];
    } catch (e) { /* catégories illisibles */ }
  }
  if (state.item !== item) return;
  if (!list.length) { show('filed', false); return; }
  const seen = new Set();
  const parts = list.filter((e) => { const k = e.url || e.name; if (seen.has(k)) return false; seen.add(k); return true; })
    .map((e) => {
      const when = e.date ? ' le ' + new Date(e.date).toLocaleDateString('fr-FR') : '';
      return e.url
        ? `<a href="${escHtml(e.url)}" target="_blank" rel="noopener">${escHtml(e.name)}</a>${when}`
        : `<b>${escHtml(e.name)}</b>${when}`;
    });
  const named = parts.filter((p) => p && p !== '<b></b>');
  box.innerHTML = named.length ? 'Déjà classé dans ' + named.join(', ') + '.' : 'Déjà classé (dossier non précisé).';
  show('filed', true);
}

/* ---------- Divers ---------- */

function setStatus(msg, kind = '', isHtml = false) {
  const s = $('status');
  s.className = 'status small' + (kind ? ' ' + kind : '');
  if (isHtml) s.innerHTML = msg; else s.textContent = msg;
}

function handleError(e) {
  console.error(e);
  if (e instanceof AuthError) {
    show('pick', false); show('auth', true); show('account', false);
  }
  setStatus(e.message || String(e), 'err');
}
