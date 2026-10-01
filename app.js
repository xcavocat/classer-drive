'use strict';

/* Classer dans Drive : complément Outlook.
   Tout se passe dans le navigateur : Outlook -> Google Drive, sans serveur intermédiaire. */

const API = 'https://www.googleapis.com/drive/v3';
const UPLOAD_API = 'https://www.googleapis.com/upload/drive/v3';
const FOLDER = 'application/vnd.google-apps.folder';
const GDOC = 'application/vnd.google-apps.document';
const MULTIPART_MAX = 5 * 1024 * 1024;
const LS = { token: 'cd.token', root: 'cd.root', mem: 'cd.memory', opts: 'cd.options', cid: 'cd.clientId' };
const PROP = 'cdFiled';
const CAT_PREFIX = 'Drive - ';
const GENERIC_FOLDERS = new Set(['correspondance', 'correspondances', 'courrier', 'courriers', 'mails', 'emails', 'e-mails',
  'pieces', 'pièces', 'procedure', 'procédure', 'factures', 'facturation', 'divers', 'echanges', 'échanges', 'notes',
  'docs', 'documents', 'admin', 'administratif', 'client', 'adversaire', 'juridiction', 'actes', 'ecritures', 'écritures']);
const OPTS = ['optEml', 'optPdf', 'optPrefix', 'optSub', 'optTag'];

const GENERIC_DOMAINS = new Set([
  'gmail.com', 'googlemail.com', 'outlook.com', 'outlook.fr', 'hotmail.com', 'hotmail.fr', 'live.com', 'live.fr',
  'msn.com', 'yahoo.com', 'yahoo.fr', 'icloud.com', 'me.com', 'orange.fr', 'wanadoo.fr', 'free.fr', 'sfr.fr',
  'laposte.net', 'neuf.fr', 'bbox.fr', 'aol.com', 'protonmail.com', 'proton.me'
]);

const STOPWORDS = new Set((
  'objet dossier dossiers affaire suite votre votres notre notres vous nous pour avec dans sans sous chez entre ' +
  'demande concernant relatif relative relatifs madame monsieur maitre maître cher chère bonjour merci cordialement ' +
  'document documents piece pièce pieces pièces jointe jointes copie envoi information informations ' +
  'mail email courriel message reponse réponse urgent important rappel transfert fwd forward ' +
  'cette celle celui leurs leur mais donc ainsi plus moins tres très bien être avoir fait faire ' +
  'janvier fevrier février mars avril juin juillet aout août septembre octobre novembre decembre décembre ' +
  'lundi mardi mercredi jeudi vendredi samedi dimanche the and for your with from this that'
).split(/\s+/));

const state = {
  item: null, token: null, tokenExp: 0, email: '',
  root: null, path: [], dest: null, currentList: [],
  parentNames: new Map(), searchTimer: null, busy: false, suggestRun: 0
};

const $ = (id) => document.getElementById(id);
const show = (id, on = true) => $(id).classList.toggle('hidden', !on);
const escHtml = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const escQ = (s) => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");

function lsGet(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* stockage indisponible */ } }

class AuthError extends Error {
  constructor() { super('Session Google expirée : reconnecte-toi.'); this.name = 'AuthError'; }
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
  if (t && t.exp > Date.now() + 60000) { state.token = t.token; state.tokenExp = t.exp; state.email = t.email || ''; }
  state.root = lsGet(LS.root, null) || { id: CONFIG.ROOT_FOLDER_ID || 'root', name: CONFIG.ROOT_FOLDER_NAME || 'Mon Drive' };
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
  if (!clientId()) {
    show('cidRow', true);
    $('redirectUri').textContent = new URL('auth.html', location.href).href.split('?')[0];
  }

  loadItem();
  if (hasToken()) onConnected();
});

function clientId() {
  const c = CONFIG.GOOGLE_CLIENT_ID;
  if (c && !c.startsWith('__')) return c;
  return lsGet(LS.cid, '');
}

function canTag() { return Office.context.requirements.isSetSupported('Mailbox', '1.8'); }

function bindUI() {
  $('btnConnect').addEventListener('click', connect);
  $('btnLogout').addEventListener('click', logout);
  $('btnSave').addEventListener('click', save);
  $('btnNewFolder').addEventListener('click', () => { show('newFolderRow', true); $('newFolderName').focus(); });
  $('btnCreate').addEventListener('click', createFolderHere);
  $('newFolderName').addEventListener('keydown', (e) => { if (e.key === 'Enter') createFolderHere(); });
  $('btnSetRoot').addEventListener('click', setRoot);
  $('search').addEventListener('input', onSearchInput);
  $('folders').addEventListener('click', onFolderClick);
  $('suggest').addEventListener('click', onSuggestClick);
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
  if (hasToken()) computeSuggestions();
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

/* ---------- Connexion Google ---------- */

function hasToken(marginMs = 60000) { return !!state.token && state.tokenExp > Date.now() + marginMs; }

function authorize() {
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
        if (!m.token) { reject(new Error('Connexion refusée' + (m.error ? ' (' + m.error + ')' : ''))); return; }
        state.token = m.token;
        state.tokenExp = Date.now() + (m.expiresIn || 3600) * 1000;
        lsSet(LS.token, { token: state.token, exp: state.tokenExp, email: state.email });
        resolve(m.token);
      });
      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => reject(new Error('Fenêtre de connexion fermée.')));
    });
  });
}

async function connect() {
  if (!clientId()) {
    const v = $('clientId').value.trim();
    if (!v.endsWith('.apps.googleusercontent.com')) { setStatus('Colle l\'ID client Google (il finit par .apps.googleusercontent.com).', 'warn'); return; }
    lsSet(LS.cid, v);
    show('cidRow', false);
  }
  try {
    setStatus('Connexion à Google...');
    await authorize();
    setStatus('');
    onConnected();
  } catch (e) {
    setStatus(e.message, 'err');
  }
}

function logout() {
  state.token = null; state.tokenExp = 0; state.email = '';
  lsSet(LS.token, null);
  show('account', false); show('pick', false); show('auth', true);
  selectDest(null);
}

async function onConnected() {
  show('auth', false); show('pick', true);
  try {
    const r = await (await drive('/about?fields=user(emailAddress)')).json();
    state.email = r.user && r.user.emailAddress || '';
    lsSet(LS.token, { token: state.token, exp: state.tokenExp, email: state.email });
    $('accountEmail').textContent = state.email;
    show('account', !!state.email);
  } catch (e) { handleError(e); return; }
  openPath(state.path.length ? state.path : [state.root]);
  if (state.item) computeSuggestions();
}

/* ---------- Appels Drive ---------- */

async function drive(path, opts = {}) {
  if (!hasToken(0)) throw new AuthError();
  const url = path.startsWith('http') ? path : API + path;
  const r = await fetch(url, { ...opts, headers: { Authorization: 'Bearer ' + state.token, ...(opts.headers || {}) } });
  if (r.status === 401) { state.token = null; lsSet(LS.token, null); throw new AuthError(); }
  if (!r.ok) {
    let msg = r.status + '';
    try { const j = await r.json(); msg = (j.error && j.error.message) || msg; } catch (e) { /* corps vide */ }
    throw new Error('Drive : ' + msg);
  }
  return r;
}

const ALL_DRIVES = { supportsAllDrives: 'true', includeItemsFromAllDrives: 'true' };

async function listFolders(parentId) {
  const out = [];
  let pageToken = '';
  do {
    const p = new URLSearchParams({
      q: `'${escQ(parentId)}' in parents and mimeType='${FOLDER}' and trashed=false`,
      fields: 'nextPageToken,files(id,name,parents)', orderBy: 'name_natural', pageSize: '1000', ...ALL_DRIVES
    });
    if (pageToken) p.set('pageToken', pageToken);
    const r = await (await drive('/files?' + p)).json();
    out.push(...(r.files || []));
    pageToken = r.nextPageToken;
  } while (pageToken);
  return out;
}

async function listSharedDrives() {
  try {
    const r = await (await drive('/drives?pageSize=100&fields=drives(id,name)')).json();
    return (r.drives || []).map((d) => ({ id: d.id, name: d.name, shared: true }));
  } catch (e) {
    if (e instanceof AuthError) throw e;
    return [];
  }
}

async function searchFolders(text, limit = 30) {
  const p = new URLSearchParams({
    q: `mimeType='${FOLDER}' and trashed=false and name contains '${escQ(text)}'`,
    fields: 'files(id,name,parents)', pageSize: String(limit), corpora: 'allDrives', ...ALL_DRIVES
  });
  const r = await (await drive('/files?' + p)).json();
  return r.files || [];
}

async function parentName(folder) {
  const pid = folder.parents && folder.parents[0];
  if (!pid) return '';
  if (state.parentNames.has(pid)) return state.parentNames.get(pid);
  let name = '';
  try {
    const r = await (await drive(`/files/${encodeURIComponent(pid)}?fields=name&supportsAllDrives=true`)).json();
    name = r.name || '';
  } catch (e) { if (e instanceof AuthError) throw e; }
  state.parentNames.set(pid, name);
  return name;
}

async function createFolder(name, parentId) {
  const r = await drive('/files?supportsAllDrives=true&fields=id,name,parents', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=UTF-8' },
    body: JSON.stringify({ name, mimeType: FOLDER, parents: [parentId] })
  });
  return r.json();
}

async function upload(name, blob, parentId, targetMime) {
  const meta = { name, parents: [parentId] };
  if (targetMime) meta.mimeType = targetMime;
  const type = blob.type || 'application/octet-stream';

  if (blob.size <= MULTIPART_MAX) {
    const fd = new FormData();
    fd.append('metadata', new Blob([JSON.stringify(meta)], { type: 'application/json' }));
    fd.append('file', blob);
    return (await drive(UPLOAD_API + '/files?uploadType=multipart&supportsAllDrives=true&fields=id,name', { method: 'POST', body: fd })).json();
  }

  // Gros fichiers : envoi en reprise (resumable)
  const init = await drive(UPLOAD_API + '/files?uploadType=resumable&supportsAllDrives=true&fields=id,name', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Upload-Content-Type': type },
    body: JSON.stringify(meta)
  });
  const loc = init.headers.get('Location');
  if (!loc) throw new Error('Drive n\'a pas fourni d\'adresse d\'envoi pour ce gros fichier.');
  const r = await fetch(loc, { method: 'PUT', headers: { 'Content-Type': type }, body: blob });
  if (!r.ok) throw new Error('Envoi interrompu (' + r.status + ')');
  return r.json();
}

/* ---------- Navigation dans les dossiers ---------- */

async function openPath(path) {
  state.path = path;
  $('search').value = '';
  renderCrumbs();
  $('folders').innerHTML = '<li class="info">Chargement...</li>';
  try {
    const cur = path[path.length - 1];
    let list = await listFolders(cur.id);
    if (path.length === 1 && cur.id === 'root') list = list.concat(await listSharedDrives());
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
      <button class="f-name" data-act="pick" title="Choisir ce dossier">${escHtml(f.name)}${f.shared ? '<span class="f-sub">Drive partagé</span>' : ''}<span class="f-sub" data-parent="${i}"></span></button>
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
    if ($('search').value.trim()) openPath([state.root, { id: f.id, name: f.name, parents: f.parents }]);
    else openPath(state.path.concat([{ id: f.id, name: f.name, parents: f.parents }]));
  }
}

function onSearchInput() {
  clearTimeout(state.searchTimer);
  const text = $('search').value.trim();
  state.searchTimer = setTimeout(async () => {
    if (!text) { openPath(state.path); return; }
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
    const f = await createFolder(name, cur.id);
    $('newFolderName').value = '';
    show('newFolderRow', false);
    await openPath(state.path);
    selectDest(f);
  } catch (e) { handleError(e); }
}

function setRoot() {
  const cur = state.path[state.path.length - 1];
  state.root = { id: cur.id, name: cur.name };
  lsSet(LS.root, state.root);
  openPath([state.root]);
  setStatus(`Racine définie : ${cur.name}`, 'ok');
}

function selectDest(f) {
  state.dest = f ? { id: f.id, name: f.name, parents: f.parents } : null;
  const d = $('dest');
  if (state.dest) { d.className = 'tab'; d.innerHTML = 'Destination : <b>' + escHtml(state.dest.name) + '</b>'; }
  else { d.className = 'tab empty'; d.textContent = 'Aucun dossier choisi'; }
  $('btnSave').disabled = !state.dest || state.busy;
  document.querySelectorAll('#folders li').forEach((li) => {
    const lf = li.dataset.current ? state.path[state.path.length - 1] : state.currentList[Number(li.dataset.i)];
    li.classList.toggle('sel', !!(state.dest && lf && lf.id === state.dest.id));
  });
  document.querySelectorAll('#suggest .sugg').forEach((b) => b.classList.toggle('sel', !!(state.dest && b.dataset.id === state.dest.id)));
}

/* ---------- Suggestions et mémoire ---------- */

function memoryKeys(item) {
  const keys = [];
  if (item.conversationId) keys.push(['conv:' + item.conversationId, 10, 'même conversation']);
  const who = correspondent(item);
  const email = (who.emailAddress || '').toLowerCase();
  if (email) {
    keys.push(['mail:' + email, 4, 'même correspondant']);
    const dom = email.split('@')[1];
    const myDom = ((Office.context.mailbox.userProfile.emailAddress || '').split('@')[1] || '').toLowerCase();
    if (dom && !GENERIC_DOMAINS.has(dom) && dom !== myDom) keys.push(['dom:' + dom, 2, 'même organisation']);
  }
  return keys;
}

function remember(item, dest) {
  const mem = lsGet(LS.mem, {});
  const now = Date.now();
  for (const [k] of memoryKeys(item)) {
    const list = mem[k] || [];
    const e = list.find((x) => x.id === dest.id);
    if (e) { e.n++; e.t = now; e.name = dest.name; } else list.unshift({ id: dest.id, name: dest.name, parents: dest.parents, n: 1, t: now });
    list.sort((a, b) => b.t - a.t);
    mem[k] = list.slice(0, 5);
  }
  const keys = Object.keys(mem);
  if (keys.length > 1500) {
    keys.sort((a, b) => Math.max(...mem[a].map((x) => x.t)) - Math.max(...mem[b].map((x) => x.t)));
    keys.slice(0, 300).forEach((k) => delete mem[k]);
  }
  lsSet(LS.mem, mem);
}

function keywords(item) {
  const who = correspondent(item);
  const text = cleanSubject(item.subject) + ' ' + (who.displayName || '');
  const seen = new Set();
  const out = [];
  for (const raw of text.split(/[^\p{L}\p{N}]+/u)) {
    const t = raw.trim();
    const low = t.toLowerCase();
    if (!t || seen.has(low) || STOPWORDS.has(low)) continue;
    const isAcronym = /^\p{Lu}{2,}$/u.test(t);
    if (!isAcronym && t.length < 4) continue;
    if (/^\d+$/.test(t) && t.length < 4) continue;
    seen.add(low);
    const weight = isAcronym ? 3 : /^\p{Lu}/u.test(t) ? 2 : 1;
    out.push({ t, weight });
  }
  return out.sort((a, b) => b.weight - a.weight).map((x) => x.t);
}

async function computeSuggestions() {
  const item = state.item;
  const run = ++state.suggestRun;
  const box = $('suggest');
  box.innerHTML = '<p class="muted small">Recherche de suggestions...</p>';

  const scores = new Map();
  const add = (f, pts, why) => {
    const e = scores.get(f.id) || { f, score: 0, why: new Set() };
    e.score += pts; e.why.add(why);
    scores.set(f.id, e);
  };

  const mem = lsGet(LS.mem, {});
  for (const [k, pts, why] of memoryKeys(item)) {
    for (const m of (mem[k] || [])) add({ id: m.id, name: m.name, parents: m.parents }, pts * Math.min(m.n, 5), why);
  }

  try {
    const toks = keywords(item).slice(0, 6);
    const results = await Promise.all(toks.map((t) => searchFolders(t, 15).catch((e) => { if (e instanceof AuthError) throw e; return []; })));
    results.forEach((fs, i) => fs.forEach((f) => add(f, 3, 'mot « ' + toks[i] + ' »')));
  } catch (e) {
    if (run === state.suggestRun) handleError(e);
    return;
  }
  if (run !== state.suggestRun) return;

  const top = [...scores.values()].sort((a, b) => b.score - a.score).slice(0, 5);
  if (!top.length) { box.innerHTML = '<p class="muted small">Pas de suggestion : cherche ou parcours tes dossiers.</p>'; return; }

  state.suggestions = top.map((e) => e.f);
  box.innerHTML = top.map((e, i) =>
    `<button class="sugg" data-i="${i}" data-id="${escHtml(e.f.id)}">${escHtml(e.f.name)}<span class="why" data-why="${i}">${escHtml([...e.why].join(', '))}</span></button>`).join('');

  const best = top[0];
  if (!state.dest && (best.why.has('même conversation') || best.why.has('même correspondant'))) selectDest(best.f);
  else selectDest(state.dest);

  for (let i = 0; i < top.length; i++) {
    try {
      const pn = await parentName(top[i].f);
      if (run !== state.suggestRun) return;
      const el = box.querySelector(`[data-why="${i}"]`);
      if (el && pn) el.textContent = 'dans ' + pn + ', ' + el.textContent;
    } catch (e) { return; }
  }
}

function onSuggestClick(e) {
  const b = e.target.closest('.sugg');
  if (!b) return;
  selectDest(state.suggestions[Number(b.dataset.i)]);
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

async function savePdf(item, base, target) {
  const body = await officeCall((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
  const html = buildMailHtml(item, body);
  // Conversion via Google Docs, export PDF, puis suppression du document temporaire.
  const tmp = await upload(base + ' (conversion)', new Blob([html], { type: 'text/html' }), target, GDOC);
  try {
    const pdf = await (await drive(`/files/${tmp.id}/export?mimeType=application/pdf`)).blob();
    await upload(base + '.pdf', new Blob([pdf], { type: 'application/pdf' }), target);
  } finally {
    await drive(`/files/${tmp.id}?supportsAllDrives=true`, { method: 'DELETE' }).catch(() => {});
  }
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
  if (!hasToken(300000)) {
    try { await authorize(); } catch (e) { setStatus(e.message, 'err'); return; }
  }
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
    let target = dest.id;
    if ($('optSub').checked) {
      step('Création du sous-dossier');
      target = (await createFolder(base, target)).id;
    }
    if (wantEml) {
      step('Mail original');
      try {
        const b64 = await officeCall((cb) => item.getAsFileAsync(cb));
        await upload(base + '.eml', b64ToBlob(b64, 'message/rfc822'), target);
        done++;
      } catch (e) { if (e instanceof AuthError) throw e; failures.push('.eml (' + e.message + ')'); }
    }
    if (wantPdf) {
      step('Mail en PDF');
      try { await savePdf(item, base, target); done++; }
      catch (e) { if (e instanceof AuthError) throw e; failures.push('PDF (' + e.message + ')'); }
    }
    for (let i = 0; i < atts.length; i++) {
      step(`Pièce jointe ${i + 1}/${atts.length}`);
      try { await saveAttachment(item, atts[i], target, $('optPrefix').checked ? date : ''); done++; }
      catch (e) { if (e instanceof AuthError) throw e; failures.push(atts[i].name + ' (' + e.message + ')'); }
    }

    let tagNote = '';
    if (done) {
      remember(item, dest);
      try { await markFiled(item, dest, target); }
      catch (e) { tagNote = '<br>Fichiers enregistrés, mais le marquage dans Outlook a échoué (' + escHtml(e.message) + ').'; }
    }
    const url = 'https://drive.google.com/drive/folders/' + encodeURIComponent(target);
    const msg = `${done} fichier(s) enregistré(s) dans <b>${escHtml(dest.name)}</b>. <a href="${url}" target="_blank" rel="noopener">Ouvrir le dossier</a>` +
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

async function folderLabel(dest) {
  if (!GENERIC_FOLDERS.has(String(dest.name).toLowerCase().trim())) return dest.name;
  let pn = '';
  try { pn = await parentName(dest); } catch (e) { /* sans parent */ }
  return pn ? `${pn} / ${dest.name}` : dest.name;
}

function categoryName(label) {
  return (CAT_PREFIX + String(label).replace(/[,;]/g, ' ').replace(/\s+/g, ' ').trim()).slice(0, 80);
}

function colorFor(label) {
  let h = 0;
  for (const c of String(label)) h = (h * 31 + c.charCodeAt(0)) >>> 0;
  return Office.MailboxEnums.CategoryColor['Preset' + (h % 25)];
}

async function markFiled(item, dest, target) {
  const label = await folderLabel(dest);
  const { props, list } = await readFiled(item);
  if (props) {
    list.push({ id: dest.id, target, name: label, date: new Date().toISOString() });
    props.set(PROP, JSON.stringify(list.slice(-10)));
    await officeCall((cb) => props.saveAsync(cb));
  }
  if ($('optTag').checked && canTag()) {
    const name = categoryName(label);
    const master = await officeCall((cb) => Office.context.mailbox.masterCategories.getAsync(cb)).catch(() => []);
    if (!(master || []).some((c) => c.displayName === name)) {
      await officeCall((cb) => Office.context.mailbox.masterCategories.addAsync([{ displayName: name, color: colorFor(label) }], cb))
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
      list = (cats || []).filter((c) => c.displayName.startsWith(CAT_PREFIX)).map((c) => ({ name: c.displayName.slice(CAT_PREFIX.length) }));
    } catch (e) { /* catégories illisibles */ }
  }
  if (state.item !== item) return;
  if (!list.length) { show('filed', false); return; }
  const seen = new Set();
  const parts = list.filter((e) => { const k = e.target || e.name; if (seen.has(k)) return false; seen.add(k); return true; })
    .map((e) => {
      const when = e.date ? ' le ' + new Date(e.date).toLocaleDateString('fr-FR') : '';
      return e.target
        ? `<a href="https://drive.google.com/drive/folders/${encodeURIComponent(e.target)}" target="_blank" rel="noopener">${escHtml(e.name)}</a>${when}`
        : `<b>${escHtml(e.name)}</b>`;
    });
  box.innerHTML = 'Déjà classé dans ' + parts.join(', ') + '.';
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
