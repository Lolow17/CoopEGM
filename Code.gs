/**
 * Coop'EGM — Caisse de bar · Base de données Google Sheet
 *
 * INSTALLATION
 * 1. Créer un Google Sheet vide, puis Extensions > Apps Script.
 * 2. Remplacer le contenu de Code.gs par ce fichier, enregistrer.
 * 3. Sélectionner la fonction « setup » puis Exécuter (autoriser l'accès).
 * 4. Déployer > Nouveau déploiement > Type : Application Web
 *      Exécuter en tant que : Moi
 *      Qui a accès : Tout le monde
 * 5. Copier l'URL qui finit par /exec et la coller dans l'appli
 *    (pastille de synchro en haut à droite > Adresse du script).
 *
 * Après chaque modification de ce script : Déployer > Gérer les déploiements >
 * crayon > Version : Nouvelle version > Déployer (l'URL reste la même).
 *
 * CODES D'ACCÈS (onglet Config)
 *   PIN         = code ADMIN : accès à tout (Gestion, Réception, Stock,
 *                 inventaire, articles, annulations anciennes).
 *   PIN_CAISSE  = code CAISSE : Caisse, Ardoises, Clôture ; annulation d'un
 *                 ticket ou d'un règlement seulement dans les 5 minutes.
 *   Les deux codes doivent être différents. Tapez-les avec une apostrophe
 *   devant ('0123) pour garder les zéros du début.
 *   PIN vide        → aucun code, tout est ouvert.
 *   PIN_CAISSE vide → la caisse est ouverte sans code, seul l'admin en demande un.
 *   Après un changement de code, les appareils redemandent le code.
 *
 * ARDOISES
 * Une vente payée « Ardoise » est notée au nom d'une personne (colonne client
 * de l'onglet Ventes). Quand la personne paie, un règlement est ajouté dans
 * l'onglet Reglements avec son vrai moyen de paiement (Espèces, Carte,
 * Virement) : c'est ce règlement qui entre dans la caisse du jour.
 * Le solde de chaque ardoise = ventes « Ardoise » valides − règlements valides,
 * calculé sur tout l'historique.
 */

const PERIODE_JOURS = 120;          // historique renvoyé à l'appli
const METHODES = ['Espèces', 'Carte', 'Virement', 'Autre', 'Ardoise'];
const REGLEMENT = ['Espèces', 'Carte', 'Virement'];   // moyens acceptés pour solder une ardoise
const HEAD = {
  Config:      ['cle', 'valeur', 'aide'],
  Produits:    ['id', 'nom', 'categorie', 'mode', 'contenance', 'seuil', 'couleur', 'ordre', 'actif', 'formats', 'stock', 'stock_lisible'],
  Ventes:      ['id', 'date', 'jour', 'paiement', 'statut', 'total', 'detail', 'poste', 'annule_le', 'motif_annulation', 'articles', 'client'],
  Reglements:  ['id', 'date', 'jour', 'client', 'paiement', 'montant', 'statut', 'poste', 'annule_le', 'motif_annulation'],
  Mouvements:  ['id', 'date', 'article_id', 'article', 'delta', 'type', 'ref', 'note', 'poste'],
  Receptions:  ['id', 'date', 'fournisseur', 'reference', 'montant', 'detail', 'justificatif', 'lien_justificatif', 'poste', 'lignes'],
  Inventaires: ['id', 'date', 'motif', 'ecarts', 'detail', 'poste', 'lignes'],
  Clotures:    ['id', 'date', 'jour', 'fond', 'especes_comptees', 'especes_attendues', 'ecart', 'total', 'especes', 'carte', 'virement', 'autre', 'note', 'poste', 'ardoise', 'reglements_ardoise'],
  Journal:     ['date', 'type', 'detail', 'ref', 'poste'],
  Ops:         ['opId', 'date', 'type']
};

/* ---------------- Points d'entrée ---------------- */

function doGet(e) {
  try {
    // Sans code configuré : comportement historique. Avec des codes : rien n'est lu sans code (POST action « state »).
    if (!roleOf('', config())) return json({ ok: true, locked: true });
    return json({ ok: true, state: state() });
  }
  catch (err) { return json({ ok: false, error: String(err.message || err) }); }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json({ ok: false, error: 'Requête illisible' }); }
  try {
    const conf = config(), role = roleOf(body.pin, conf);
    if (body.action === 'pin') return json({ ok: !!role, role: role });
    if (!role) return json({ ok: false, auth: true, error: 'Code requis' });
    if (body.action === 'state') return json({ ok: true, role: role, state: state() });
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      if (body.action === 'upload') {
        if (role !== 'admin') return json({ ok: false, error: 'Code admin requis' });
        return json(upload(body));
      }
      if (body.action === 'ops') {
        const results = applyOps(body.ops || [], body.device, conf);
        SpreadsheetApp.flush();
        return json({ ok: true, results: results, state: state() });
      }
    } finally { lock.releaseLock(); }
    return json({ ok: false, error: 'Action inconnue' });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

/* ---------------- Utilitaires ---------------- */

function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
const SHEETS = {};
function sh(name) {
  if (SHEETS[name]) return SHEETS[name];
  const h = HEAD[name];
  let s = ss().getSheetByName(name);
  if (!s) {
    s = ss().insertSheet(name);
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');
    s.setFrozenRows(1);
  } else if (s.getLastColumn() < h.length) {
    // Onglet créé par une version précédente : on ajoute les nouvelles colonnes d'en-tête
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');
  }
  SHEETS[name] = s;
  return s;
}
function rows(name) {
  const s = sh(name), n = s.getLastRow(), h = HEAD[name];
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, h.length).getValues().map(function (r) {
    const o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o;
  });
}
function append(name, list) {
  if (!list.length) return;
  const s = sh(name);
  s.getRange(s.getLastRow() + 1, 1, list.length, list[0].length).setValues(list);
}
function S(v) { const s = String(v == null ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; } // anti-formule
function num(v) { const n = Number(String(v).replace(',', '.')); return isFinite(n) ? n : 0; }
function money(n) { return Math.round(num(n) * 100) / 100; }
function parseJ(v, d) { try { return v ? JSON.parse(v) : d; } catch (e) { return d; } }
function iso(v) { if (!v) return ''; const d = v instanceof Date ? v : new Date(v); return isNaN(d) ? '' : d.toISOString(); }
function toDate(v) { const d = new Date(v); return isNaN(d) ? new Date() : d; }
function uid() { return Utilities.getUuid(); }
function ymd(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function dayStr(v, d) { if (v instanceof Date) return ymd(v); if (v) return String(v); return ymd(toDate(d)); }
function eur(n) { return money(n).toFixed(2).replace('.', ',') + ' €'; }
function cleanName(v) { return String(v == null ? '' : v).trim().replace(/\s+/g, ' ').slice(0, 60); }
function tabKey(v) { return cleanName(v).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); } // « Élodie » = « elodie »
/* Rôle donné par un code : 'admin', 'caisse' ou '' (refusé) */
function roleOf(code, conf) {
  const a = String(conf.PIN || '').trim(), k = String(conf.PIN_CAISSE || '').trim(), p = String(code || '').trim();
  if (!a) return 'admin';
  if (p && p === a) return 'admin';
  if (!k) return 'caisse';
  if (p && p === k) return 'caisse';
  return '';
}
function config() { const o = {}; rows('Config').forEach(function (r) { o[String(r.cle)] = r.valeur; }); return o; }
function setConfig(k, v) {
  const s = sh('Config'), last = s.getLastRow();
  if (last > 1) {
    const keys = s.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) if (String(keys[i][0]) === k) { s.getRange(i + 2, 2).setValue(v); return; }
  }
  s.appendRow([k, v, '']);
}

/* ---------------- Produits ---------------- */

function prodFromRow(r) {
  const mode = String(r.mode) === 'unit' ? 'unit' : 'ml';
  return {
    id: String(r.id), name: String(r.nom || ''), category: String(r.categorie || ''), mode: mode,
    packageSize: mode === 'ml' ? (num(r.contenance) || 1000) : 1, low: num(r.seuil),
    color: String(r.couleur || ''), order: num(r.ordre),
    active: !(r.actif === false || /^(false|faux|non|0)$/i.test(String(r.actif))),
    formats: parseJ(r.formats, [])
  };
}
function lisible(p, st) { return p.mode === 'ml' ? (st / 1000).toFixed(2).replace('.', ',') + ' L' : st + ' u'; }
function prodToRow(x) {
  const p = x.p;
  return [S(p.id), S(p.name), S(p.category), p.mode, p.packageSize, p.low, S(p.color), p.order, p.active, JSON.stringify(p.formats), x.stock, lisible(p, x.stock)];
}
function loadProducts() {
  const list = rows('Produits').filter(function (r) { return r.id !== ''; }).map(function (r) { return { p: prodFromRow(r), stock: num(r.stock) }; });
  const by = {}; list.forEach(function (x) { by[x.p.id] = x; });
  return { list: list, by: by, dirty: false };
}
function saveProducts(P) {
  if (!P.dirty) return;
  const s = sh('Produits'), w = HEAD.Produits.length, last = s.getLastRow();
  if (last > 1) s.getRange(2, 1, last - 1, w).clearContent();
  if (P.list.length) s.getRange(2, 1, P.list.length, w).setValues(P.list.map(prodToRow));
}

/* ---------------- Opérations ---------------- */

function applyOps(ops, device, conf) {
  const os = sh('Ops'), n = os.getLastRow();
  const done = new Set(n > 1 ? os.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); }) : []);
  const ctx = {
    P: loadProducts(), device: S(device || ''), now: new Date(), admin: false,
    out: { Ventes: [], Reglements: [], Mouvements: [], Receptions: [], Inventaires: [], Clotures: [], Journal: [], Ops: [] }
  };
  const results = [];
  ops.forEach(function (op) {
    if (!op || !op.opId) return;
    const id = String(op.opId);
    if (done.has(id)) { results.push({ opId: id, ok: true, dup: true }); return; }
    try {
      const h = HANDLERS[op.type];
      if (!h) throw new Error('Type inconnu : ' + op.type);
      ctx.admin = roleOf(op.pin, conf) === 'admin';
      h(op.data || {}, ctx);
      ctx.out.Ops.push([id, ctx.now, String(op.type)]);
      done.add(id);
      results.push({ opId: id, ok: true });
    } catch (err) {
      results.push({ opId: id, ok: false, error: String(err.message || err) });
    }
  });
  saveProducts(ctx.P);
  Object.keys(ctx.out).forEach(function (k) { append(k, ctx.out[k]); });
  return results;
}

function move(ctx, pid, delta, type, ref, note, at) {
  const x = ctx.P.by[pid];
  if (x) { x.stock = Math.round((x.stock + delta) * 1000) / 1000; ctx.P.dirty = true; }
  ctx.out.Mouvements.push([uid(), at, S(pid), S(x ? x.p.name : ''), delta, type, S(ref), S(note), ctx.device]);
}
function journal(ctx, type, detail, ref) { ctx.out.Journal.push([ctx.now, type, S(detail), S(ref), ctx.device]); }

/* Retrouve une ligne (en attente d'écriture ou déjà dans l'onglet) pour l'annuler */
function findRow(ctx, name, id) {
  const H = HEAD[name], iSt = H.indexOf('statut'), iAnn = H.indexOf('annule_le');
  const pending = ctx.out[name].filter(function (r) { return r[0] === id; })[0];
  if (pending) {
    const o = {}; H.forEach(function (k, i) { o[k] = pending[i]; });
    o.set = function (st, at, m) { pending[iSt] = st; pending[iAnn] = at; pending[iAnn + 1] = m; };
    return o;
  }
  const s = sh(name), last = s.getLastRow();
  if (last < 2) return null;
  const cell = s.getRange(2, 1, last - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  if (!cell) return null;
  const row = cell.getRow(), v = s.getRange(row, 1, 1, H.length).getValues()[0];
  const o = {}; H.forEach(function (k, i) { o[k] = v[i]; });
  o.set = function (st, at, m) { s.getRange(row, iSt + 1).setValue(st); s.getRange(row, iAnn + 1, 1, 2).setValues([[at, m]]); };
  return o;
}

function checkVoid(d, ctx, row, label) {
  const reason = String(d.reason || '').trim();
  if (!reason) throw new Error('Motif obligatoire');
  if (!row) throw new Error(label + ' introuvable');
  if (String(row.statut) === 'annulee') return '';
  const recent = (ctx.now.getTime() - toDate(row.date).getTime()) < 5 * 60000;
  if (!ctx.admin && !(d.quick && recent)) throw new Error('Code admin requis pour annuler ce ' + label.toLowerCase());
  return reason;
}

const HANDLERS = {
  sale: function (d, ctx) {
    if (!d.id) throw new Error('Vente sans identifiant');
    const items = (d.items || []).map(function (i) {
      return { pid: String(i.pid || ''), fid: String(i.fid || ''), article: String(i.article || ''), format: String(i.format || ''),
               qty: Math.max(1, Math.round(num(i.qty))), consumed: num(i.consumed), unitPrice: money(i.unitPrice) };
    });
    if (!items.length) throw new Error('Vente vide');
    const method = METHODES.indexOf(d.method) >= 0 ? d.method : 'Autre';
    const client = cleanName(d.client);
    if (method === 'Ardoise' && !client) throw new Error('Nom obligatoire pour une ardoise');
    const total = money(items.reduce(function (a, i) { return a + i.qty * i.unitPrice; }, 0));
    const at = toDate(d.at);
    const detail = items.map(function (i) { return i.qty + ' × ' + i.article + ' (' + i.format + ')'; }).join(', ');
    ctx.out.Ventes.push([S(d.id), at, "'" + String(d.day || ymd(at)), method, 'valide', total, S(detail), S(d.device || ctx.device), '', '', JSON.stringify(items), S(client)]);
    items.forEach(function (i) { move(ctx, i.pid, -i.consumed, 'vente', d.id, i.format, at); });
    journal(ctx, 'Vente', eur(total) + ' · ' + method + (client ? ' · ' + client : ''), d.id);
  },

  void: function (d, ctx) {
    const sale = findRow(ctx, 'Ventes', String(d.id));
    const reason = checkVoid(d, ctx, sale, 'Ticket');
    if (!reason) return;
    sale.set('annulee', ctx.now, S(reason));
    parseJ(sale.articles, []).forEach(function (i) { move(ctx, String(i.pid), num(i.consumed), 'annulation', d.id, reason, ctx.now); });
    journal(ctx, 'Annulation', eur(sale.total) + ' · ' + reason, d.id);
  },

  settle: function (d, ctx) {
    if (!d.id) throw new Error('Règlement sans identifiant');
    const client = cleanName(d.client);
    if (!client) throw new Error('Nom manquant');
    if (REGLEMENT.indexOf(d.method) < 0) throw new Error('Moyen de paiement invalide');
    const amount = money(d.amount);
    if (amount <= 0) throw new Error('Montant invalide');
    const at = toDate(d.at);
    ctx.out.Reglements.push([S(d.id), at, "'" + String(d.day || ymd(at)), S(client), d.method, amount, 'valide', S(d.device || ctx.device), '', '']);
    journal(ctx, 'Règlement ardoise', client + ' · ' + eur(amount) + ' · ' + d.method, d.id);
  },

  settlevoid: function (d, ctx) {
    const r = findRow(ctx, 'Reglements', String(d.id));
    const reason = checkVoid(d, ctx, r, 'Règlement');
    if (!reason) return;
    r.set('annulee', ctx.now, S(reason));
    journal(ctx, 'Annulation règlement', String(r.client) + ' · ' + eur(r.montant) + ' · ' + reason, d.id);
  },

  reception: function (d, ctx) {
    if (!ctx.admin) throw new Error('Code admin requis pour la réception');
    const supplier = String(d.supplier || '').trim();
    if (!supplier) throw new Error('Fournisseur manquant');
    const lines = (d.lines || []).map(function (l) {
      const x = ctx.P.by[String(l.pid)], pk = Math.round(num(l.packages));
      return { pid: String(l.pid), article: x ? x.p.name : String(l.article || ''), packages: pk, added: pk * (x && x.p.mode === 'ml' ? x.p.packageSize : 1) };
    }).filter(function (l) { return l.packages > 0; });
    if (!lines.length) throw new Error('Aucune quantité reçue');
    const at = toDate(d.at);
    ctx.out.Receptions.push([S(d.id), at, S(supplier), S(d.invoiceRef || ''), money(d.amount),
      S(lines.map(function (l) { return l.packages + ' × ' + l.article; }).join(', ')),
      S(d.fileName || ''), S(d.fileUrl || ''), ctx.device, JSON.stringify(lines)]);
    lines.forEach(function (l) { move(ctx, l.pid, l.added, 'reception', d.id, supplier, at); });
    journal(ctx, 'Réception', supplier + ' · ' + eur(d.amount) + ' · ' + lines.length + ' réf.', d.id);
  },

  inventory: function (d, ctx) {
    if (!ctx.admin) throw new Error("Code admin requis pour l'inventaire");
    const reason = String(d.reason || '').trim() || 'Inventaire';
    const at = toDate(d.at);
    const lines = (d.lines || []).map(function (l) {
      const x = ctx.P.by[String(l.pid)], th = x ? x.stock : 0, ph = Math.max(0, Math.round(num(l.physical)));
      return { pid: String(l.pid), article: x ? x.p.name : '', theoretical: th, physical: ph, diff: Math.round((ph - th) * 1000) / 1000 };
    });
    const ecarts = lines.filter(function (l) { return l.diff !== 0; });
    ctx.out.Inventaires.push([S(d.id), at, S(reason), ecarts.length,
      S(ecarts.map(function (l) { return l.article + ' ' + (l.diff > 0 ? '+' : '') + l.diff; }).join(', ')), ctx.device, JSON.stringify(lines)]);
    ecarts.forEach(function (l) { move(ctx, l.pid, l.diff, 'inventaire', d.id, reason, at); });
    journal(ctx, 'Inventaire', reason + ' · ' + ecarts.length + ' écart(s)', d.id);
  },

  closing: function (d, ctx) {
    const day = String(d.day || ''), fl = money(d.float), ct = money(d.count);
    const sales = rows('Ventes').map(function (r) { return { jour: dayStr(r.jour, r.date), statut: String(r.statut), paiement: String(r.paiement), total: num(r.total) }; });
    ctx.out.Ventes.forEach(function (r) { sales.push({ jour: String(r[2]).replace(/^'/, ''), statut: r[4], paiement: r[3], total: r[5] }); });
    const regs = rows('Reglements').map(function (r) { return { jour: dayStr(r.jour, r.date), statut: String(r.statut), paiement: String(r.paiement), montant: num(r.montant) }; });
    ctx.out.Reglements.forEach(function (r) { regs.push({ jour: String(r[2]).replace(/^'/, ''), statut: r[6], paiement: r[4], montant: r[5] }); });
    // pay[] = encaissements du jour par moyen (ventes + règlements d'ardoise) ; pay.Ardoise = crédit accordé ce jour
    const pay = { 'Espèces': 0, 'Carte': 0, 'Virement': 0, 'Autre': 0, 'Ardoise': 0 };
    let total = 0, settled = 0;
    sales.forEach(function (r) { if (r.jour === day && r.statut === 'valide') { pay[r.paiement] = money((pay[r.paiement] || 0) + r.total); total = money(total + r.total); } });
    regs.forEach(function (r) { if (r.jour === day && r.statut === 'valide') { pay[r.paiement] = money((pay[r.paiement] || 0) + r.montant); settled = money(settled + r.montant); } });
    const exp = money(fl + pay['Espèces']), diff = money(ct - exp);
    ctx.out.Clotures.push([S(d.id), toDate(d.at), "'" + day, fl, ct, exp, diff, total, pay['Espèces'], pay['Carte'], pay['Virement'], pay['Autre'], S(d.note || ''), ctx.device, pay['Ardoise'], settled]);
    journal(ctx, 'Clôture', day + ' · écart ' + eur(diff), d.id);
  },

  product: function (d, ctx) {
    if (!ctx.admin) throw new Error('Code admin requis');
    const name = String(d.name || '').trim();
    if (!name) throw new Error('Nom manquant');
    const mode = d.mode === 'unit' ? 'unit' : 'ml';
    const formats = (d.formats || []).map(function (f) {
      return { id: String(f.id || uid()), name: String(f.name || '').trim(), qty: Math.max(1, Math.round(num(f.qty))), price: money(f.price) };
    }).filter(function (f) { return f.name; });
    if (!formats.length) throw new Error('Au moins un format requis');
    const p = { id: String(d.id || uid()), name: name, category: String(d.category || '').trim(), mode: mode,
      packageSize: mode === 'ml' ? Math.max(1, Math.round(num(d.packageSize))) : 1, low: Math.max(0, num(d.low)),
      color: String(d.color || ''), order: num(d.order), active: d.active !== false, formats: formats };
    const x = ctx.P.by[p.id];
    if (x) {
      if (x.p.mode !== p.mode && x.stock !== 0) throw new Error("Changement d'unité impossible : stock non nul");
      x.p = p;
    } else {
      const nx = { p: p, stock: 0 }; ctx.P.list.push(nx); ctx.P.by[p.id] = nx;
    }
    ctx.P.dirty = true;
    journal(ctx, x ? 'Article modifié' : 'Article créé', name + (p.active ? '' : ' (archivé)'), p.id);
  }
};

/* ---------------- État renvoyé à l'appli ---------------- */

function state() {
  const cutoff = Date.now() - PERIODE_JOURS * 864e5;
  const keep = function (r) { const t = toDate(r.date).getTime(); return t >= cutoff; };
  const P = rows('Produits').filter(function (r) { return r.id !== ''; });
  const stock = {}; P.forEach(function (r) { stock[String(r.id)] = num(r.stock); });
  const V = rows('Ventes').filter(function (r) { return r.id !== ''; });
  const R = rows('Reglements').filter(function (r) { return r.id !== ''; });

  // Soldes des ardoises sur tout l'historique (pas seulement 120 jours)
  const tabs = {};
  const tab = function (name, date, delta) {
    const k = tabKey(name); if (!k) return;
    const t = tabs[k] || (tabs[k] = { key: k, name: cleanName(name), balance: 0, last: '' });
    t.balance = money(t.balance + delta);
    const at = iso(date); if (at > t.last) t.last = at;
  };
  V.forEach(function (r) { if (String(r.paiement) === 'Ardoise' && String(r.statut) === 'valide') tab(r.client, r.date, num(r.total)); });
  R.forEach(function (r) { if (String(r.statut) === 'valide') tab(r.client, r.date, -num(r.montant)); });

  return {
    products: P.map(prodFromRow),
    stock: stock,
    sales: V.filter(keep).map(function (r) {
      return { id: String(r.id), at: iso(r.date), day: dayStr(r.jour, r.date), method: String(r.paiement), status: String(r.statut),
               total: num(r.total), items: parseJ(r.articles, []), device: String(r.poste || ''), client: String(r.client || ''),
               voidAt: iso(r.annule_le), voidReason: String(r.motif_annulation || '') };
    }),
    settlements: R.filter(keep).map(function (r) {
      return { id: String(r.id), at: iso(r.date), day: dayStr(r.jour, r.date), client: String(r.client || ''), method: String(r.paiement),
               amount: num(r.montant), status: String(r.statut), device: String(r.poste || ''),
               voidAt: iso(r.annule_le), voidReason: String(r.motif_annulation || '') };
    }),
    tabs: tabs,
    receptions: rows('Receptions').filter(function (r) { return r.id !== '' && keep(r); }).map(function (r) {
      return { id: String(r.id), at: iso(r.date), supplier: String(r.fournisseur), invoiceRef: String(r.reference || ''), amount: num(r.montant),
               lines: parseJ(r.lignes, []), fileName: String(r.justificatif || ''), fileUrl: String(r.lien_justificatif || ''), device: String(r.poste || '') };
    }),
    inventories: rows('Inventaires').filter(function (r) { return r.id !== '' && keep(r); }).map(function (r) {
      return { id: String(r.id), at: iso(r.date), reason: String(r.motif || ''), lines: parseJ(r.lignes, []) };
    }),
    closings: rows('Clotures').filter(function (r) { return r.id !== '' && keep(r); }).map(function (r) {
      return { id: String(r.id), at: iso(r.date), day: dayStr(r.jour, r.date), float: num(r.fond), count: num(r.especes_comptees),
               expected: num(r.especes_attendues), difference: num(r.ecart), total: num(r.total), note: String(r.note || ''),
               payments: { 'Espèces': num(r.especes), 'Carte': num(r.carte), 'Virement': num(r.virement), 'Autre': num(r.autre), 'Ardoise': num(r.ardoise) },
               settled: num(r.reglements_ardoise) };
    }),
    audit: rows('Journal').slice(-200).reverse().map(function (r) {
      return { at: iso(r.date), type: String(r.type), description: String(r.detail), reference: String(r.ref || ''), device: String(r.poste || '') };
    }),
    pinSet: !!String(config().PIN || '').trim(),
    cashierPinSet: !!String(config().PIN_CAISSE || '').trim(),
    serverTime: new Date().toISOString()
  };
}

/* ---------------- Justificatifs (Drive) ---------------- */

function upload(body) {
  const data = String(body.data || '');
  if (!data) throw new Error('Fichier vide');
  if (data.length > 7.5e6) throw new Error('Fichier trop volumineux (5 Mo max)');
  const type = String(body.type || '');
  if (!(type === 'application/pdf' || type.indexOf('image/') === 0)) throw new Error('Format accepté : photo ou PDF');
  const blob = Utilities.newBlob(Utilities.base64Decode(data), type, String(body.name || 'justificatif').slice(0, 120));
  const f = folder().createFile(blob);
  return { ok: true, url: f.getUrl(), name: f.getName() };
}
function folder() {
  const id = String(config().DOSSIER_ID || '');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder("Coop'EGM - Justificatifs");
  setConfig('DOSSIER_ID', f.getId());
  return f;
}

/* ---------------- Installation et maintenance ---------------- */

function exemples() {
  function F(name, qty, price) { return { id: uid(), name: name, qty: qty, price: price }; }
  return [
    { name: 'Bière 33 cl', category: 'Bières', mode: 'unit', packageSize: 1, low: 24, color: '#f2b544', formats: [F('Bouteille', 1, 2.5)] },
    { name: 'Pression (fût 30 L)', category: 'Bières', mode: 'ml', packageSize: 30000, low: 5000, color: '#e8743b', formats: [F('Demi 25 cl', 250, 3), F('Pinte 50 cl', 500, 5)] },
    { name: 'Rhum blanc 1 L', category: 'Spiritueux', mode: 'ml', packageSize: 1000, low: 300, color: '#c94f9b', formats: [F('2 cl', 20, 2), F('4 cl', 40, 3)] },
    { name: 'Whisky 70 cl', category: 'Spiritueux', mode: 'ml', packageSize: 700, low: 200, color: '#7a5cff', formats: [F('2 cl', 20, 2.5), F('4 cl', 40, 4)] },
    { name: 'Coca 33 cl', category: 'Softs', mode: 'unit', packageSize: 1, low: 24, color: '#e0484f', formats: [F('Canette', 1, 1.5)] },
    { name: 'Eau 50 cl', category: 'Softs', mode: 'unit', packageSize: 1, low: 12, color: '#2bb3c0', formats: [F('Bouteille', 1, 1)] },
    { name: 'Café', category: 'Chaud', mode: 'unit', packageSize: 1, low: 50, color: '#9aa3b2', formats: [F('Tasse', 1, 1)] }
  ].map(function (p, i) { p.id = uid(); p.order = i; p.active = true; return p; });
}

function setup() {
  Object.keys(HEAD).forEach(sh);
  const c = sh('Config');
  if (c.getLastRow() < 2) c.getRange(2, 1, 3, 3).setValues([
    ['PIN', "'0000", 'Code ADMIN : accès à tout. Vide = aucun code.'],
    ['NOM', "Coop'EGM", 'Nom de la caisse'],
    ['DOSSIER_ID', '', 'Rempli automatiquement : dossier Drive des justificatifs']
  ]);
  if (!('PIN_CAISSE' in config())) c.appendRow(['PIN_CAISSE', '', 'Code CAISSE : Caisse, Ardoises, Clôture. Vide = caisse sans code. Différent du PIN admin.']);
  if (sh('Produits').getLastRow() < 2) {
    const P = { list: exemples().map(function (p) { return { p: p, stock: 0 }; }), by: {}, dirty: true };
    saveProducts(P);
  }
  [['Ventes', 2], ['Ventes', 9], ['Reglements', 2], ['Reglements', 9], ['Mouvements', 2], ['Receptions', 2], ['Inventaires', 2], ['Clotures', 2], ['Journal', 1], ['Ops', 2]].forEach(function (x) {
    const s = sh(x[0]); s.getRange(2, x[1], s.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  });
  ['Feuille 1', 'Sheet1', 'Feuille1'].forEach(function (n) {
    const s = ss().getSheetByName(n);
    if (s && ss().getSheets().length > 1 && s.getLastRow() === 0) ss().deleteSheet(s);
  });
  Logger.log('Installation terminée. Déployez maintenant en Application Web.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Coop'EGM")
    .addItem('Initialiser les onglets', 'setup')
    .addItem('Recalculer le stock depuis les mouvements', 'recalculerStock')
    .addItem("Purger les identifiants d'envoi de plus de 60 jours", 'purgerOps')
    .addToUi();
}

function recalculerStock() {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const tot = {};
    rows('Mouvements').forEach(function (m) { const k = String(m.article_id); tot[k] = (tot[k] || 0) + num(m.delta); });
    const P = loadProducts();
    P.list.forEach(function (x) { x.stock = Math.round((tot[x.p.id] || 0) * 1000) / 1000; });
    P.dirty = true; saveProducts(P);
  } finally { lock.releaseLock(); }
}

function purgerOps() {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const s = sh('Ops'), n = s.getLastRow();
    if (n < 2) return;
    const lim = Date.now() - 60 * 864e5;
    const keep = s.getRange(2, 1, n - 1, 3).getValues().filter(function (r) { return toDate(r[1]).getTime() >= lim; });
    s.getRange(2, 1, n - 1, 3).clearContent();
    if (keep.length) s.getRange(2, 1, keep.length, 3).setValues(keep);
  } finally { lock.releaseLock(); }
}
