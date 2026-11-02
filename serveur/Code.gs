/**
 * Plantes mères — Amaltis
 * Script Google Apps Script autonome, appelé par l'écran hébergé sur GitHub Pages (doPost).
 *
 * Données : Google Sheet « Saisie boutures et végétation »
 *   Mères        : Numéro | Surnom | Lot | Cultivar | Date de création | Source de la date | Mère d'origine |
 *                  Emplacement | Statut | Date de destruction | Notes | Créée par | Créée le | Modifiée par | Modifiée le
 *   Destructions : ID | Date | Numéro | Lot | Cultivar | Poids (g) | Méthode | Raison | Témoin 1 | Témoin 2 |
 *                  RSDC complété | Saisi par | Saisi le | Statut | Annulé par | Annulé le | Motif d'annulation
 *   Boutures     : (onglet existant, lu par registre_bv.py) + colonne I « ID »
 *   Végétation   : (onglet existant, lu par registre_bv.py) — tenu à jour automatiquement, une ligne par mère
 *   Témoins      : Nom | Habilitation | Actif      (personnes habilitées à témoigner d'une destruction)
 *   Paramètres   : Clé | Valeur
 *   Journal      : Horodatage | Employé | Action | Objet | ID | Détails   (ajout seulement, jamais modifié)
 *
 * Connexion : mêmes noms et codes que le calendrier (onglet « Employés » du Sheet du calendrier).
 * Dates stockées en texte AAAA-MM-JJ.
 */

const SHEET_ID = '1nxHQQEp1_xw7AuD7OveNNnWZRjncxtjURGZElklWSCw';      // Saisie boutures et végétation
const EMP_SHEET_ID = '1Ur9Kd_wEPs-ERF8cXhqZFm292eFSiot5wp7wrU7zQGk';  // Calendrier de culture — Amaltis (Employés)
const TZ = 'America/Toronto';

const HEAD = {
  'Mères': ['Numéro', 'Surnom', 'Lot', 'Cultivar', 'Date de création', 'Source de la date', "Mère d'origine",
    'Emplacement', 'Statut', 'Date de destruction', 'Notes', 'Créée par', 'Créée le', 'Modifiée par', 'Modifiée le'],
  'Destructions': ['ID', 'Date', 'Numéro', 'Lot', 'Cultivar', 'Poids (g)', 'Méthode', 'Raison', 'Témoin 1', 'Témoin 2',
    'RSDC complété', 'Saisi par', 'Saisi le', 'Statut', 'Annulé par', 'Annulé le', "Motif d'annulation"],
  'Boutures': ['Date de création', 'No. unique du plant mère', 'Nombre de boutures', 'Cultivar', 'Salle', 'No. de lot',
    'Responsable', 'Notes', 'ID'],
  'Végétation': ['Date de création', 'Date de destruction', 'No. unique du plant mère', 'Nombre de plants', 'Cultivar',
    'Responsable', 'Notes'],
  'Témoins': ['Nom', 'Habilitation', 'Actif'],
  'Paramètres': ['Clé', 'Valeur'],
  'Journal': ['Horodatage', 'Employé', 'Action', 'Objet', 'ID', 'Détails'],
};
// Index des colonnes de l'onglet Mères
const M = { num: 0, surnom: 1, lot: 2, cult: 3, crea: 4, src: 5, parent: 6, empl: 7, statut: 8, destr: 9, notes: 10,
  creePar: 11, creeLe: 12, modPar: 13, modLe: 14 };

/* ====================== Installation (facultatif : les onglets existent déjà) ====================== */

function installer() {
  const ss = SS_();
  Object.keys(HEAD).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const n = HEAD[name].length;
    if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
    const h = sh.getRange(1, 1, 1, n).getDisplayValues()[0];
    HEAD[name].forEach((t, i) => { if (!h[i]) sh.getRange(1, i + 1).setValue(t).setFontWeight('bold'); });
    sh.setFrozenRows(1);
  });
}

/* ====================== Outils ====================== */

let ss_;
const sheets_ = {};
function SS_() { return ss_ || (ss_ = SpreadsheetApp.openById(SHEET_ID)); }
function sheet_(name) {
  const sh = sheets_[name] || (sheets_[name] = SS_().getSheetByName(name));
  if (!sh) throw new Error('Onglet « ' + name + ' » introuvable dans le Sheet.');
  return sh;
}
function rows_(name) {
  const sh = sheet_(name), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEAD[name].length).getDisplayValues()
    .map((r, i) => ({ r: r.map(v => String(v).trim()), row: i + 2 }))
    .filter(x => x.r.some(v => v !== ''));
}
const fmt_ = v => typeof v !== 'number' ? '@' : Number.isInteger(v) ? '0' : '0.0##';
function append_(name, values) {
  const sh = sheet_(name);
  const r = sh.getLastRow() + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
  const rng = sh.getRange(r, 1, 1, values.length);
  rng.setNumberFormats([values.map(fmt_)]);
  rng.setValues([values]);
  return r;
}
function setCells_(name, row, col0, values) {
  const rng = sheet_(name).getRange(row, col0 + 1, 1, values.length);
  rng.setNumberFormats([values.map(fmt_)]);
  rng.setValues([values]);
}
const yes_ = v => /^(oui|o|yes|y|1|vrai|true|x)$/i.test(String(v).trim());
const now_ = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
const today_ = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const id_ = () => Utilities.getUuid().replace(/-/g, '').slice(0, 10);
const txt_ = (v, max) => String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim().slice(0, max || 200);
function isDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s))) return false;
  const [y, m, d] = s.split('-').map(Number), x = new Date(Date.UTC(y, m - 1, d));
  return x.getUTCFullYear() === y && x.getUTCMonth() === m - 1 && x.getUTCDate() === d;
}
/** Date lue dans les 8 premiers chiffres du numéro complet (AAAAMMJJ…), sinon ''. */
function dateDuNumero_(num) {
  const m = /^(\d{4})(\d{2})(\d{2})\d{4}_/.exec(String(num));
  if (!m) return '';
  const s = m[1] + '-' + m[2] + '-' + m[3];
  return isDate_(s) && +m[1] >= 2015 && s <= today_() ? s : '';
}
/** Lot = suffixe après le dernier « _ » du numéro. */
const lotDuNumero_ = num => { const i = String(num).lastIndexOf('_'); return i > 0 ? String(num).slice(i + 1) : ''; };
const FORMAT_NUM_ = /^\d{12}_[0-9A-Za-z]+$/;
const normNum_ = s => txt_(s, 60).replace(/\s+/g, '').toUpperCase();

function journal_(nom, action, objet, id, details) {
  append_('Journal', [now_().slice(0, 16), nom, action, objet, String(id || ''), txt_(details, 1000)]);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
/** Un même envoi (requestId) relancé après une coupure réseau n'est enregistré qu'une fois. */
function unique_(requestId, fn) {
  const c = CacheService.getScriptCache(), k = requestId ? 'req|' + txt_(requestId, 60) : null;
  if (k) { const hit = c.get(k); if (hit) return JSON.parse(hit); }
  const res = fn();
  if (k) try { c.put(k, JSON.stringify(res == null ? null : res), 3600); } catch (e) {}
  return res;
}

/* ====================== Connexion (Employés du calendrier) ====================== */

function employes_() {
  const c = CacheService.getScriptCache(), hit = c.get('emp');
  if (hit) return JSON.parse(hit);
  const sh = SpreadsheetApp.openById(EMP_SHEET_ID).getSheetByName('Employés');
  const n = sh.getLastRow() - 1;
  const v = n < 1 ? [] : sh.getRange(2, 1, n, 4).getDisplayValues()
    .filter(r => String(r[0]).trim() && yes_(r[3]))
    .map(r => ({ nom: String(r[0]).trim(), code: String(r[1]).trim(), role: String(r[2]).trim().toLowerCase() === 'admin' ? 'admin' : 'opérateur' }));
  try { c.put('emp', JSON.stringify(v), 60); } catch (e) {}
  return v;
}
const ESSAIS_MAX_ = 5, BLOCAGE_S_ = 900;  // 5 codes erronés pour un même nom → connexion suspendue 15 minutes
function auth_(nom, code) {
  const cache = CacheService.getScriptCache();
  const cle = 'echecs_' + String(nom || '').trim().toLowerCase().replace(/[^a-z0-9àâäçéèêëîïôöùûüÿ]/g, '_').slice(0, 60);
  const echecs = parseInt(cache.get(cle) || '0', 10);
  if (echecs >= ESSAIS_MAX_) throw new Error("Trop d'essais. Réessayez dans 15 minutes.");
  const e = employes_().find(x => x.nom === String(nom || '').trim() && x.code === String(code || '').trim());
  if (!e) { cache.put(cle, String(echecs + 1), BLOCAGE_S_); throw new Error('Nom ou code invalide.'); }
  if (echecs) cache.remove(cle);
  return { nom: e.nom, role: e.role };
}
function admin_(nom, code) {
  const u = auth_(nom, code);
  if (u.role !== 'admin') throw new Error('Action réservée à un administrateur.');
  return u;
}

/* ====================== Lecture ====================== */

function listerNoms() { return employes_().map(e => e.nom); }

function mereDto_(x) {
  const r = x.r;
  return { numero: r[M.num], surnom: r[M.surnom], lot: r[M.lot], cultivar: r[M.cult], creation: r[M.crea],
    source: r[M.src], parent: r[M.parent], emplacement: r[M.empl], statut: r[M.statut] || 'active',
    destruction: r[M.destr], notes: r[M.notes], creePar: r[M.creePar], creeLe: r[M.creeLe],
    modPar: r[M.modPar], modLe: r[M.modLe] };
}
function boutureDto_(x) {
  const r = x.r;
  return { id: r[8], date: normDate_(r[0]), numero: r[1], nb: r[2], cultivar: r[3], salle: r[4], lot: r[5], par: r[6], notes: r[7] };
}
function destructionDto_(x) {
  const r = x.r;
  return { id: r[0], date: r[1], numero: r[2], lot: r[3], cultivar: r[4], poids: r[5], methode: r[6], raison: r[7],
    temoin1: r[8], temoin2: r[9], rsdc: r[10], par: r[11], le: r[12], statut: r[13] || 'valide',
    annulePar: r[14], annuleLe: r[15], motif: r[16] };
}
/** Les anciennes saisies de Boutures sont parfois en JJ/MM/AAAA. */
function normDate_(s) {
  s = String(s || '').trim();
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s;
}
function parametres_() {
  const p = {};
  rows_('Paramètres').forEach(x => { if (x.r[0]) p[x.r[0]] = x.r[1]; });
  const liste = s => String(s || '').split(/[,;\n]/).map(t => t.trim()).filter(Boolean);
  return {
    methodes: liste(p['Méthodes de destruction']),
    salles: liste(p['Salles de destination']).length ? liste(p['Salles de destination']) : ['Salle 3', 'Salle 4', 'Salle 5', 'Salle 6', 'Salle 7'],
  };
}

function charger(nom, code) {
  const u = auth_(nom, code);
  const p = parametres_();
  const depuis = Utilities.formatDate(new Date(Date.now() - 400 * 864e5), TZ, 'yyyy-MM-dd');
  return {
    moi: u.nom, role: u.role, today: today_(),
    meres: rows_('Mères').filter(x => x.r[M.num]).map(mereDto_),
    boutures: rows_('Boutures').map(boutureDto_).filter(b => b.date >= depuis),
    destructions: rows_('Destructions').map(destructionDto_),
    temoins: rows_('Témoins').filter(x => x.r[0] && (x.r[2] === '' || yes_(x.r[2]))).map(x => ({ nom: x.r[0], habilitation: x.r[1] })),
    employes: employes_().map(e => e.nom),
    methodes: p.methodes, salles: p.salles,
  };
}

function journal(nom, code, n) {
  admin_(nom, code);
  const all = rows_('Journal');
  return all.slice(-Math.min(+n || 300, 2000)).reverse()
    .map(x => ({ quand: x.r[0], qui: x.r[1], action: x.r[2], objet: x.r[3], id: x.r[4], details: x.r[5] }));
}

/* ====================== Mères ====================== */

function trouverMere_(numero) {
  const n = normNum_(numero);
  return rows_('Mères').find(x => normNum_(x.r[M.num]) === n) || null;
}

/**
 * Création d'une mère. m = { numero, surnom, lot, cultivar, creation, parent, emplacement, notes, formatLibre }
 */
function creerMere(nom, code, m, requestId) {
  const u = auth_(nom, code);
  return unique_(requestId, () => withLock_(() => {
    m = m || {};
    const numero = normNum_(m.numero);
    if (!numero) throw new Error('Le numéro unique est obligatoire.');
    if (!FORMAT_NUM_.test(numero) && !m.formatLibre)
      throw new Error('Le numéro doit être au format complet (12 chiffres, « _ », lot), par exemple 202411020026_036BG.');
    if (trouverMere_(numero)) throw new Error('Ce numéro existe déjà : ' + numero + '.');
    const cultivar = txt_(m.cultivar, 60);
    if (!cultivar) throw new Error('Le cultivar est obligatoire.');
    let creation = txt_(m.creation, 10), source = 'saisie';
    if (!creation) { creation = dateDuNumero_(numero); source = creation ? 'numéro' : 'inconnue'; }
    else if (!isDate_(creation)) throw new Error('Date de création invalide.');
    else if (creation > today_()) throw new Error('La date de création ne peut pas être dans le futur.');
    const lot = txt_(m.lot, 30).toUpperCase() || lotDuNumero_(numero);
    const parent = m.parent ? normNum_(m.parent) : '';
    const row = [numero, txt_(m.surnom, 30), lot, cultivar, creation, source, parent, txt_(m.emplacement, 60),
      'active', '', txt_(m.notes, 500), u.nom, now_(), '', ''];
    append_('Mères', row);
    // Registre de végétation : une ligne par mère (date de création ; destruction remplie plus tard)
    append_('Végétation', [creation, '', numero, 1, cultivar, u.nom, txt_(m.notes, 200)]);
    journal_(u.nom, 'création', 'mère', numero,
      [m.surnom && 'surnom ' + txt_(m.surnom, 30), 'lot ' + lot, cultivar, 'créée ' + (creation || 'date inconnue'),
        parent && 'origine ' + parent, m.formatLibre && 'format non standard'].filter(Boolean).join(' · '));
    return mereDto_({ r: row });
  }));
}

/** Modification d'une fiche. Tout employé : surnom, emplacement, notes, date de création si inconnue. Admin : tout sauf le numéro et le statut. */
function modifierMere(nom, code, numero, champs) {
  const u = auth_(nom, code);
  return withLock_(() => {
    const x = trouverMere_(numero);
    if (!x) throw new Error('Mère introuvable.');
    const avant = x.r.slice(), apres = x.r.slice();
    champs = champs || {};
    const set = (k, i, max) => { if (Object.prototype.hasOwnProperty.call(champs, k)) apres[i] = txt_(champs[k], max); };
    set('surnom', M.surnom, 30); set('emplacement', M.empl, 60); set('notes', M.notes, 500);
    const dateInconnue = !avant[M.crea];
    if (Object.prototype.hasOwnProperty.call(champs, 'creation') && (u.role === 'admin' || dateInconnue)) {
      const d = txt_(champs.creation, 10);
      if (d && !isDate_(d)) throw new Error('Date de création invalide.');
      if (d > today_()) throw new Error('La date de création ne peut pas être dans le futur.');
      if (d !== avant[M.crea]) { apres[M.crea] = d; apres[M.src] = d ? 'saisie' : 'inconnue'; }
    }
    if (u.role === 'admin') {
      set('cultivar', M.cult, 60); set('lot', M.lot, 30); set('parent', M.parent, 60);
      if (!apres[M.cult]) throw new Error('Le cultivar est obligatoire.');
    }
    const libelles = { [M.surnom]: 'surnom', [M.empl]: 'emplacement', [M.notes]: 'notes', [M.crea]: 'création',
      [M.cult]: 'cultivar', [M.lot]: 'lot', [M.parent]: 'origine' };
    const diff = Object.keys(libelles).map(Number).filter(i => apres[i] !== avant[i])
      .map(i => libelles[i] + ' : ' + (avant[i] || '∅') + ' → ' + (apres[i] || '∅'));
    if (!diff.length) return mereDto_(x);
    apres[M.modPar] = u.nom; apres[M.modLe] = now_();
    setCells_('Mères', x.row, 0, apres);
    // Végétation : on reporte la date de création et le cultivar sur la ligne de la mère
    const v = rows_('Végétation').find(y => normNum_(y.r[2]) === normNum_(apres[M.num]));
    if (v) { setCells_('Végétation', v.row, 0, [apres[M.crea]]); setCells_('Végétation', v.row, 4, [apres[M.cult]]); }
    journal_(u.nom, 'modification', 'mère', apres[M.num], diff.join(' · '));
    return mereDto_({ r: apres });
  });
}

/* ====================== Prises de boutures ====================== */

/**
 * p = { date, salle, lot, notes, lignes: [{ numero, nb }] } — une ligne par mère dans l'onglet Boutures.
 */
function prendreBoutures(nom, code, p, requestId) {
  const u = auth_(nom, code);
  return unique_(requestId, () => withLock_(() => {
    p = p || {};
    const date = txt_(p.date, 10);
    if (!isDate_(date)) throw new Error('Date invalide.');
    if (date > today_()) throw new Error('La date ne peut pas être dans le futur.');
    const salle = txt_(p.salle, 30), lot = txt_(p.lot, 40);
    if (!salle) throw new Error('Choisissez la salle de destination.');
    if (!lot) throw new Error('Le numéro de lot est obligatoire.');
    const lignes = (p.lignes || []).map(l => ({ numero: normNum_(l.numero), nb: Math.round(+l.nb) })).filter(l => l.numero && l.nb > 0);
    if (!lignes.length) throw new Error('Indiquez au moins une mère et un nombre de boutures.');
    const meres = rows_('Mères');
    const vus = {};
    const out = lignes.map(l => {
      if (vus[l.numero]) throw new Error('Mère inscrite deux fois : ' + l.numero + '.');
      vus[l.numero] = 1;
      if (l.nb > 5000) throw new Error('Nombre de boutures invraisemblable pour ' + l.numero + '.');
      const x = meres.find(y => normNum_(y.r[M.num]) === l.numero);
      if (!x) throw new Error('Mère inconnue : ' + l.numero + '.');
      if ((x.r[M.statut] || 'active') !== 'active') throw new Error('Mère détruite : ' + l.numero + '.');
      if (x.r[M.crea] && x.r[M.crea] > date) throw new Error('La mère ' + l.numero + " n'existait pas encore à cette date.");
      return { x, nb: l.nb };
    });
    const res = out.map(o => {
      const id = id_();
      append_('Boutures', [date, o.x.r[M.num], o.nb, o.x.r[M.cult], salle, lot, u.nom, txt_(p.notes, 200), id]);
      return { id, date, numero: o.x.r[M.num], nb: String(o.nb), cultivar: o.x.r[M.cult], salle, lot, par: u.nom, notes: txt_(p.notes, 200) };
    });
    const total = res.reduce((s, b) => s + +b.nb, 0);
    journal_(u.nom, 'création', 'boutures', lot, date + ' · ' + salle + ' · ' + total + ' boutures · ' +
      res.map(b => b.numero + ' ×' + b.nb).join(', '));
    return res;
  }));
}

/** Suppression d'une ligne de boutures saisie par erreur (admin). L'ancienne valeur reste au journal. */
function supprimerBouture(nom, code, id, motif) {
  const u = admin_(nom, code);
  return withLock_(() => {
    if (!txt_(motif)) throw new Error('Indiquez le motif de la suppression.');
    const x = rows_('Boutures').find(y => y.r[8] && y.r[8] === String(id));
    if (!x) throw new Error('Ligne introuvable.');
    sheet_('Boutures').deleteRow(x.row);
    journal_(u.nom, 'suppression', 'boutures', id,
      'Motif : ' + txt_(motif) + ' · ligne supprimée : ' + x.r.slice(0, 8).join(' | '));
    return true;
  });
}

/* ====================== Destructions ====================== */

/**
 * d = { numero, date, poids, methode, raison, temoin1, temoin2, rsdc }
 * Registre officiel : poids net, méthode, deux témoins habilités.
 */
function detruire(nom, code, d, requestId) {
  const u = auth_(nom, code);
  return unique_(requestId, () => withLock_(() => {
    d = d || {};
    const x = trouverMere_(d.numero);
    if (!x) throw new Error('Mère introuvable.');
    if ((x.r[M.statut] || 'active') !== 'active') throw new Error('Cette mère est déjà détruite.');
    const date = txt_(d.date, 10);
    if (!isDate_(date)) throw new Error('Date invalide.');
    if (date > today_()) throw new Error('La date ne peut pas être dans le futur.');
    if (x.r[M.crea] && date < x.r[M.crea]) throw new Error('La date de destruction précède la date de création.');
    const poids = parseFloat(String(d.poids).replace(',', '.'));
    if (!(poids > 0) || poids > 100000) throw new Error('Indiquez le poids net en grammes.');
    const methode = txt_(d.methode, 100), raison = txt_(d.raison, 300);
    if (!methode) throw new Error('Indiquez la méthode de destruction.');
    if (!raison) throw new Error('Indiquez la raison de la destruction.');
    const t1 = txt_(d.temoin1, 60), t2 = txt_(d.temoin2, 60);
    const habilites = rows_('Témoins').filter(y => y.r[0] && (y.r[2] === '' || yes_(y.r[2]))).map(y => y.r[0]);
    if (!habilites.length) throw new Error("Aucun témoin habilité n'est inscrit. Demandez à un administrateur de remplir l'onglet « Témoins ».");
    if (!t1 || !t2) throw new Error('Deux témoins habilités sont obligatoires.');
    if (t1 === t2) throw new Error('Les deux témoins doivent être des personnes différentes.');
    [t1, t2].forEach(t => { if (habilites.indexOf(t) < 0) throw new Error(t + " n'est pas inscrit comme témoin habilité."); });
    const id = id_();
    const row = [id, date, x.r[M.num], x.r[M.lot], x.r[M.cult], poids, methode, raison, t1, t2, d.rsdc ? 'oui' : 'non',
      u.nom, now_(), 'valide', '', '', ''];
    append_('Destructions', row);
    setCells_('Mères', x.row, M.statut, ['détruite', date]);
    setCells_('Mères', x.row, M.modPar, [u.nom, now_()]);
    const v = rows_('Végétation').find(y => normNum_(y.r[2]) === normNum_(x.r[M.num]) && !y.r[1]);
    if (v) setCells_('Végétation', v.row, 1, [date]);
    journal_(u.nom, 'destruction', 'mère', x.r[M.num],
      date + ' · ' + poids + ' g · ' + methode + ' · ' + raison + ' · témoins ' + t1 + ', ' + t2 + ' · RSDC ' + (d.rsdc ? 'oui' : 'non') + ' · ID ' + id);
    return { destruction: destructionDto_({ r: row.map(String) }), mere: Object.assign(mereDto_(x), { statut: 'détruite', destruction: date }) };
  }));
}

/** Indique après coup que la destruction a été inscrite au RSDC. */
function marquerRsdc(nom, code, id) {
  const u = auth_(nom, code);
  return withLock_(() => {
    const x = rows_('Destructions').find(y => y.r[0] === String(id));
    if (!x) throw new Error('Destruction introuvable.');
    if (x.r[10] === 'oui') return true;
    setCells_('Destructions', x.row, 10, ['oui']);
    journal_(u.nom, 'RSDC complété', 'destruction', id, x.r[2]);
    return true;
  });
}

/** Annulation d'une destruction saisie par erreur (admin) : la mère redevient active. */
function annulerDestruction(nom, code, id, motif) {
  const u = admin_(nom, code);
  return withLock_(() => {
    if (!txt_(motif)) throw new Error("Indiquez le motif de l'annulation.");
    const x = rows_('Destructions').find(y => y.r[0] === String(id));
    if (!x) throw new Error('Destruction introuvable.');
    if (x.r[13] === 'annulée') throw new Error('Déjà annulée.');
    setCells_('Destructions', x.row, 13, ['annulée', u.nom, now_(), txt_(motif, 300)]);
    const m = trouverMere_(x.r[2]);
    if (m) {
      setCells_('Mères', m.row, M.statut, ['active', '']);
      setCells_('Mères', m.row, M.modPar, [u.nom, now_()]);
      const v = rows_('Végétation').find(y => normNum_(y.r[2]) === normNum_(x.r[2]) && y.r[1] === x.r[1]);
      if (v) setCells_('Végétation', v.row, 1, ['']);
    }
    journal_(u.nom, 'annulation', 'destruction', id, x.r[2] + ' · détruite le ' + x.r[1] + ' · motif : ' + txt_(motif, 300));
    return true;
  });
}

/* ====================== Porte d'entrée (GitHub Pages) ====================== */

const API_ = {
  listerNoms: () => listerNoms(),
  charger: (n, c) => charger(n, c),
  journal: (n, c, k) => journal(n, c, k),
  creerMere: (n, c, m, r) => creerMere(n, c, m, r),
  modifierMere: (n, c, num, ch) => modifierMere(n, c, num, ch),
  prendreBoutures: (n, c, p, r) => prendreBoutures(n, c, p, r),
  supprimerBouture: (n, c, id, m) => supprimerBouture(n, c, id, m),
  detruire: (n, c, d, r) => detruire(n, c, d, r),
  marquerRsdc: (n, c, id) => marquerRsdc(n, c, id),
  annulerDestruction: (n, c, id, m) => annulerDestruction(n, c, id, m),
};
function doPost(e) {
  let rep;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const f = Object.prototype.hasOwnProperty.call(API_, req.fn) ? API_[req.fn] : null;
    if (!f) throw new Error('Action inconnue.');
    rep = { ok: true, data: f.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    rep = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(rep)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return ContentService.createTextOutput('Plantes mères — Amaltis : service actif.');
}
