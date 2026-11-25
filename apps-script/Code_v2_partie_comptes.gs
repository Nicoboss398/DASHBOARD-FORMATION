// =====================================================================
// v2 — COMPTES, RÔLES ET DROITS PAR PROJET (25/11/2026)
// ---------------------------------------------------------------------
// Chaque personne se connecte avec son e-mail + son mot de passe.
// Le script ne renvoie à chacun QUE les projets qui lui sont attribués,
// et à un CLIENT ni les actions « Interne Opus » ni les notes internes.
// Rôles : admin (tout + gestion des accès) · equipe (ses projets, tout voir) ·
// partenaire (ses projets, modifie, sans interne) · client (lecture seule, sans interne).
// Les comptes sont dans l'onglet « Utilisateurs » du même classeur.
// Première mise en route : exécuter une fois initialiserAdmin() depuis
// l'éditeur → un code d'invitation est envoyé par e-mail à l'administrateur.
// L'ancien jeton partagé (TOKEN) n'ouvre plus l'accès (LEGACY_TOKEN_ENABLED).
// =====================================================================
const USERS_SHEET = 'Utilisateurs';
const USERS_HEADERS = ['email', 'nom', 'role', 'projets', 'hash', 'sel', 'invitation', 'invitation_expire', 'actif', 'derniere_connexion'];
const ROLES = ['admin', 'equipe', 'partenaire', 'client'];
const ADMIN_EMAIL = 'nicolas.bossard@opusformation.fr';
const DASHBOARD_URL = 'https://nicoboss398.github.io/DASHBOARD-FORMATION/';
const SUBTASK_DELIM = '␞';
const META_DELIM = '␟';
const SESSION_DAYS = 30;
const INVITE_DAYS = 7;
const LEGACY_TOKEN_ENABLED = false;

function checkToken_(t) {
  return TOKEN && t === TOKEN;
}

// --- GET : plus de lecture sans connexion ---
function doGet(e) {
  return json_({ success: false, error: 'Utilisez le tableau de bord pour vous connecter.' });
}

// --- POST : body = {"action":"...","session":"...", ...} ---
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const a = body.action;
    if (a === 'login') return json_(login_(body));
    if (a === 'activate') return json_(activate_(body));
    const ctx = userCtx_(body);
    if (!ctx) return json_({ success: false, error: 'unauthorized' });
    if (a === 'me') return json_({ success: true, user: publicUser_(ctx) });
    if (a === 'get') return json_(readTasksFor_(ctx));
    if (a === 'save_all') {
      if (ctx.role === 'client') return json_({ success: false, error: 'forbidden' });
      return json_(saveScoped_(ctx, body.tasks || []));
    }
    if (a === 'notify') {
      if (ctx.role === 'client') return json_({ success: false, error: 'forbidden' });
      return json_(sendNotify_(body));
    }
    if (a === 'users_list' || a === 'user_save' || a === 'user_invite') {
      if (ctx.role !== 'admin') return json_({ success: false, error: 'forbidden' });
      return json_(adminAction_(a, body, ctx));
    }
    return json_({ success: false, error: 'Action inconnue: ' + a });
  } catch (err) {
    return json_({ success: false, error: String(err) });
  }
}

// ---------- Sessions signées (aucun stockage côté serveur) ----------
function props_() { return PropertiesService.getScriptProperties(); }
function secret_() {
  let s = props_().getProperty('SESSION_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props_().setProperty('SESSION_SECRET', s); }
  return s;
}
function b64_(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function unb64_(s) { while (s.length % 4) s += '='; return Utilities.newBlob(Utilities.base64DecodeWebSafe(s)).getDataAsString(); }
function hmac_(txt) { return b64_(Utilities.computeHmacSha256Signature(txt, secret_())); }
function makeSession_(email) {
  const payload = email + '|' + (Date.now() + SESSION_DAYS * 86400000);
  return b64_(Utilities.newBlob(payload).getBytes()) + '.' + hmac_(payload);
}
function readSession_(tok) {
  if (!tok || typeof tok !== 'string' || tok.indexOf('.') < 0) return null;
  const parts = tok.split('.');
  let payload;
  try { payload = unb64_(parts[0]); } catch (e) { return null; }
  if (hmac_(payload) !== parts[1]) return null;
  const i = payload.lastIndexOf('|');
  const exp = Number(payload.slice(i + 1));
  if (!exp || Date.now() > exp) return null;
  return payload.slice(0, i);
}

// ---------- Mots de passe ----------
function hash_(secret, salt) {
  let h = salt + '|' + secret;
  for (let i = 0; i < 300; i++) {
    h = b64_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
  }
  return h;
}
function randomCode_() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const raw = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  let out = '';
  for (let i = 0; i < 10; i++) out += alphabet[parseInt(raw.substr(i * 2, 2), 16) % alphabet.length];
  return out.slice(0, 5) + '-' + out.slice(5);
}
function tooManyFailures_(email) {
  const c = CacheService.getScriptCache();
  return Number(c.get('fail_' + email) || 0) >= 8;
}
function noteFailure_(email) {
  const c = CacheService.getScriptCache();
  c.put('fail_' + email, String(Number(c.get('fail_' + email) || 0) + 1), 900);
}

// ---------- Onglet « Utilisateurs » ----------
function usersSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.getRange(1, 1, 1, USERS_HEADERS.length).setValues([USERS_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1000, USERS_HEADERS.length).setNumberFormat('@');
  }
  return sh;
}
function readUsers_() {
  const sh = usersSheet_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, USERS_HEADERS.length).getValues().map(function (r, i) {
    const u = { _row: i + 2 };
    USERS_HEADERS.forEach(function (h, j) { u[h] = (r[j] === null || r[j] === undefined) ? '' : String(r[j]); });
    u.email = u.email.trim().toLowerCase();
    return u;
  }).filter(function (u) { return u.email; });
}
function writeUser_(u) {
  const sh = usersSheet_();
  const row = USERS_HEADERS.map(function (h) { return (u[h] === null || u[h] === undefined) ? '' : String(u[h]); });
  if (u._row) sh.getRange(u._row, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  else { sh.appendRow(row); u._row = sh.getLastRow(); }
}
function isActive_(u) { return ['non', 'false', '0'].indexOf(String(u.actif).trim().toLowerCase()) < 0; }
function parseProjects_(s) { return String(s || '').split(/\s*[;\n]\s*/).map(function (x) { return x.trim(); }).filter(Boolean); }
function nowTxt_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function publicUser_(ctx) { return { email: ctx.email, name: ctx.name, role: ctx.role, projects: ctx.projects }; }
function ctxFromUser_(u) {
  return { email: u.email, name: u.nom || u.email, role: ROLES.indexOf(u.role) >= 0 ? u.role : 'client', projects: parseProjects_(u.projets), _u: u };
}
function userCtx_(body) {
  if (LEGACY_TOKEN_ENABLED && body.token && checkToken_(body.token)) {
    return { email: 'legacy', name: 'legacy', role: 'admin', projects: ['*'] };
  }
  const email = readSession_(body.session);
  if (!email) return null;
  const u = readUsers_().filter(function (x) { return x.email === email; })[0];
  if (!u || !isActive_(u) || !u.hash) return null;
  return ctxFromUser_(u);
}
function canSee_(ctx, proj) {
  return ctx.role === 'admin' || ctx.projects.indexOf('*') >= 0 || ctx.projects.indexOf(proj) >= 0;
}

// ---------- Connexion / activation ----------
function login_(b) {
  const email = String(b.email || '').trim().toLowerCase();
  const pwd = String(b.password || '');
  Utilities.sleep(300);
  if (tooManyFailures_(email)) return { success: false, error: 'locked' };
  const u = readUsers_().filter(function (x) { return x.email === email; })[0];
  if (!u || !isActive_(u) || !u.hash || hash_(pwd, u.sel) !== u.hash) {
    noteFailure_(email);
    return { success: false, error: 'invalid' };
  }
  u.derniere_connexion = nowTxt_();
  writeUser_(u);
  const ctx = ctxFromUser_(u);
  return { success: true, session: makeSession_(email), user: publicUser_(ctx) };
}
function activate_(b) {
  const email = String(b.email || '').trim().toLowerCase();
  const code = String(b.code || '').trim().toUpperCase();
  const pwd = String(b.password || '');
  Utilities.sleep(300);
  if (tooManyFailures_(email)) return { success: false, error: 'locked' };
  if (pwd.length < 8) return { success: false, error: 'weak' };
  const u = readUsers_().filter(function (x) { return x.email === email; })[0];
  if (!u || !isActive_(u) || !u.invitation || u.invitation !== 'inv:' + hash_(code, 'invitation')) {
    noteFailure_(email);
    return { success: false, error: 'invalid_code' };
  }
  if (u.invitation_expire && Number(u.invitation_expire) < Date.now()) return { success: false, error: 'expired' };
  u.sel = Utilities.getUuid();
  u.hash = hash_(pwd, u.sel);
  u.invitation = '';
  u.invitation_expire = '';
  u.derniere_connexion = nowTxt_();
  writeUser_(u);
  const ctx = ctxFromUser_(u);
  return { success: true, session: makeSession_(email), user: publicUser_(ctx) };
}

// ---------- Lecture filtrée ----------
function clientView_(t) {
  let c = String(t.comments || '');
  let subs = '';
  const si = c.indexOf(SUBTASK_DELIM);
  if (si >= 0) { subs = c.slice(si); c = c.slice(0, si); }
  let meta = null;
  const mi = c.indexOf(META_DELIM);
  if (mi >= 0) {
    try { meta = JSON.parse(c.slice(mi + 1)); } catch (e) { meta = {}; }
    c = c.slice(0, mi);
  }
  if (meta && meta.visibility === 'Interne Opus') return null;
  if (meta) { delete meta.internalNote; delete meta.source; }
  const out = {};
  Object.keys(t).forEach(function (k) { out[k] = t[k]; });
  out.comments = c + ((meta && Object.keys(meta).length) ? META_DELIM + JSON.stringify(meta) : '') + subs;
  return out;
}
// Partenaires et clients : jamais d'actions « Interne Opus » ni de notes internes
function restricted_(ctx) { return ctx.role === 'client' || ctx.role === 'partenaire'; }
function readTasksFor_(ctx) {
  const mine = readTasks_().filter(function (t) { return canSee_(ctx, t.project); });
  if (!restricted_(ctx)) return mine;
  return mine.map(clientView_).filter(function (t) { return t; });
}
function splitComments_(c) {
  c = String(c || '');
  let subs = '', meta = null;
  const si = c.indexOf(SUBTASK_DELIM);
  if (si >= 0) { subs = c.slice(si); c = c.slice(0, si); }
  const mi = c.indexOf(META_DELIM);
  if (mi >= 0) { try { meta = JSON.parse(c.slice(mi + 1)); } catch (e) { meta = null; } c = c.slice(0, mi); }
  return { human: c, meta: meta, subs: subs };
}
// Un partenaire enregistre une vue filtrée : on restaure ce qu'il ne voit pas
// (actions internes de ses projets, notes internes et sources des actions).
function restoreHidden_(existing, incoming, ctx) {
  const byId = {};
  existing.forEach(function (t) { byId[String(t.id)] = t; });
  const hiddenTasks = existing.filter(function (t) { return canSee_(ctx, t.project) && clientView_(t) === null; });
  const merged = incoming.map(function (t) {
    const old = byId[String(t.id)];
    if (!old) return t;
    const o = splitComments_(old.comments);
    if (!o.meta || (!o.meta.internalNote && !o.meta.source)) return t;
    const n = splitComments_(t.comments);
    const meta = n.meta || {};
    if (o.meta.internalNote) meta.internalNote = o.meta.internalNote;
    if (o.meta.source) meta.source = o.meta.source;
    const out = {};
    Object.keys(t).forEach(function (k) { out[k] = t[k]; });
    out.comments = n.human + META_DELIM + JSON.stringify(meta) + n.subs;
    return out;
  });
  const incomingIds = {};
  merged.forEach(function (t) { incomingIds[String(t.id)] = true; });
  return merged.concat(hiddenTasks.filter(function (t) { return !incomingIds[String(t.id)]; }));
}

// ---------- Écriture limitée aux projets autorisés ----------
function saveScoped_(ctx, incoming) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (ctx.role === 'admin' || ctx.projects.indexOf('*') >= 0) return writeAll_(incoming);
    const existing = readTasks_();
    const known = {};
    existing.forEach(function (t) { known[t.project] = true; });
    const created = [];
    incoming.forEach(function (t) {
      if (t.project && !known[t.project] && !canSee_(ctx, t.project) && created.indexOf(t.project) < 0) created.push(t.project);
    });
    if (created.length && ctx._u) {
      ctx._u.projets = parseProjects_(ctx._u.projets).concat(created).join(' ; ');
      writeUser_(ctx._u);
      ctx.projects = ctx.projects.concat(created);
    }
    const kept = existing.filter(function (t) { return !canSee_(ctx, t.project); });
    let mine = incoming.filter(function (t) { return canSee_(ctx, t.project); });
    if (restricted_(ctx)) mine = restoreHidden_(existing, mine, ctx);
    return writeAll_(kept.concat(mine));
  } finally {
    lock.releaseLock();
  }
}
// Compatibilité : ancienne fonction (écriture complète avec verrou)
function saveAll_(tasks) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return writeAll_(tasks); } finally { lock.releaseLock(); }
}
function writeAll_(tasks) {
  const sh = getSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  if (tasks.length > 0) {
    const rows = tasks.map(function (t) {
      return HEADERS.map(function (h) {
        const v = t[h];
        return (v === null || v === undefined) ? '' : String(v);
      });
    });
    sh.getRange(2, 1, rows.length, HEADERS.length).setNumberFormat('@').setValues(rows);
  }
  return { success: true, count: tasks.length };
}

// ---------- Administration des accès ----------
function adminAction_(a, b, ctx) {
  if (a === 'users_list') {
    const projs = {};
    readTasks_().forEach(function (t) { if (t.project) projs[t.project] = true; });
    return {
      success: true,
      projects: Object.keys(projs).sort(),
      users: readUsers_().map(function (u) {
        return {
          email: u.email, name: u.nom, role: u.role, projects: parseProjects_(u.projets),
          active: isActive_(u), hasPassword: !!u.hash, pendingInvite: !!u.invitation, lastLogin: u.derniere_connexion
        };
      })
    };
  }
  const email = String(b.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { success: false, error: 'E-mail invalide' };
  const users = readUsers_();
  let u = users.filter(function (x) { return x.email === email; })[0];
  if (a === 'user_save') {
    if (!u) u = { email: email };
    const role = ROLES.indexOf(b.role) >= 0 ? b.role : 'client';
    const active = b.active !== false;
    // Garde-fou : il doit toujours rester au moins un administrateur actif
    const otherAdmins = users.filter(function (x) { return x.email !== email && x.role === 'admin' && isActive_(x); }).length;
    if (u.role === 'admin' && (role !== 'admin' || !active) && otherAdmins === 0) {
      return { success: false, error: 'Impossible : ce compte est le dernier administrateur actif.' };
    }
    u.nom = String(b.name || u.nom || '').trim();
    u.role = role;
    u.projets = (b.projects || []).map(function (p) { return String(p).trim(); }).filter(Boolean).join(' ; ');
    u.actif = active ? 'oui' : 'non';
    writeUser_(u);
    return { success: true };
  }
  if (a === 'user_invite') {
    if (!u) return { success: false, error: 'Compte inconnu : enregistrez-le d\'abord.' };
    const code = randomCode_();
    u.invitation = 'inv:' + hash_(code, 'invitation');
    u.invitation_expire = String(Date.now() + INVITE_DAYS * 86400000);
    writeUser_(u);
    let sent = false;
    if (b.sendEmail) { sendInvitationMail_(u, code); sent = true; }
    return { success: true, code: code, sent: sent };
  }
  return { success: false, error: 'Action inconnue' };
}
function sendInvitationMail_(u, code) {
  MailApp.sendEmail(u.email, 'Votre accès au tableau de bord OPUS',
    'Bonjour' + (u.nom ? ' ' + u.nom : '') + ',\n\n'
    + 'Un accès au tableau de bord de suivi OPUS Formation vous a été ouvert.\n\n'
    + '1. Ouvrez : ' + DASHBOARD_URL + '\n'
    + '2. Cliquez sur « Première connexion »\n'
    + '3. Saisissez votre e-mail (' + u.email + ') et ce code d\'invitation : ' + code + '\n'
    + '4. Choisissez votre mot de passe (8 caractères minimum).\n\n'
    + 'Ce code est valable ' + INVITE_DAYS + ' jours et ne sert qu\'une fois.\n\n'
    + '— OPUS Formation');
}

// À exécuter UNE fois depuis l'éditeur : crée le compte administrateur et
// envoie son code d'invitation par e-mail à ADMIN_EMAIL.
function initialiserAdmin() {
  const users = readUsers_();
  let u = users.filter(function (x) { return x.email === ADMIN_EMAIL; })[0];
  if (!u) u = { email: ADMIN_EMAIL, nom: 'Nicolas BOSSARD' };
  u.role = 'admin';
  u.projets = '*';
  u.actif = 'oui';
  const code = randomCode_();
  u.invitation = 'inv:' + hash_(code, 'invitation');
  u.invitation_expire = String(Date.now() + INVITE_DAYS * 86400000);
  writeUser_(u);
  sendInvitationMail_(u, code);
  Logger.log('Compte administrateur prêt. Code d\'invitation envoyé par e-mail à ' + ADMIN_EMAIL);
}

// --- Notification e-mail (chat OPUS) ---
function sendNotify_(body) {
  var recipients = (body.recipients || [])
    .map(function (e) { return String(e || '').trim(); })
    .filter(function (e) { return e && e.indexOf('@') > 0; });
  recipients = recipients.filter(function (e, i) { return recipients.indexOf(e) === i; });
  if (recipients.length === 0) return { success: false, error: 'no recipients' };
  var subject = body.subject || 'Nouveau message dans le chat OPUS';
  var message = body.message || 'Vous avez un nouveau message dans le chat OPUS';
  var sent = 0;
  recipients.forEach(function (to) {
    try { MailApp.sendEmail(to, subject, message); sent++; } catch (e) { /* on continue */ }
  });
  return { success: true, sent: sent };
}

// --- Lecture de toutes les tâches ---
function readTasks_() {
  const sh = getSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  return values
    .filter(function (row) { return row[3] !== '' && row[3] !== null; })
    .map(function (row) {
      const task = {};
      HEADERS.forEach(function (h, i) {
        let v = row[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        task[h] = (v === null || v === undefined) ? '' : v;
      });
      return task;
    });
}

// --- Accès à l'onglet des tâches (créé avec en-têtes si absent) ---
function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  return sh;
}

// --- Réponse JSON ---
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// --- Test manuel (exécuter depuis l'éditeur pour vérifier les droits) ---
function testAcces() {
  const sh = getSheet_();
  Logger.log('OK — onglet: ' + sh.getName() + ', lignes: ' + sh.getLastRow() + ', comptes: ' + readUsers_().length);
}
