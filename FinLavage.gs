/* =====================================================================
   FIN DE PRODUCTION LAVAGE + TRAIN DE REPASSAGE
   Le Plan de travail envoie régulièrement son heure de fin estimée (et les galettes du Repassage) ;
   l'appli bureau (facturation) l'affiche avec son évolution heure par heure.
   BdL Admin envoie l'heure de départ du référent de la partie haute (fin du train de repassage).
   Onglets « Fin lavage » et « Fin lavage — historique » du tableur (créés automatiquement).
   ===================================================================== */

function feuilleFinLavage_() {
  const ss = SpreadsheetApp.openByUrl(URL_TABLEUR);
  let sh = ss.getSheetByName('Fin lavage');
  if (!sh) {
    sh = ss.insertSheet('Fin lavage');
    sh.appendRow(['Journée', 'Relevé à', 'Fin prévue', 'Reste (galettes)', 'Total (galettes)', 'Terminé']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// POST {action:'finLavage', date:'2026-10-22', fin:'17:45', reste:23, total:96, termine:false, rep:140}
// POST {action:'finLavage', date:'2026-10-22', repFin:'23:36'}   <- envoyé par BdL Admin (départ du référent partie haute)
function finLavageEnregistrer_(b) {
  if (!b.date) return { ok: false, error: 'Données manquantes' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (b.repFin) { bilanMaj_(String(b.date), { repFin: String(b.repFin) }); return { ok: true }; }
    if (!b.fin) return { ok: false, error: 'Données manquantes' };
    feuilleFinLavage_().appendRow(["'" + String(b.date), new Date(), "'" + String(b.fin),
      Number(b.reste) || 0, Number(b.total) || 0, b.termine ? 'oui' : '']);
    bilanMaj_(String(b.date), { termine: !!b.termine, fin: String(b.fin), total: Number(b.total) || 0, rep: b.rep });
    alarmeMemoriser_(b);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

// GET ?action=finLavage&date=2026-10-22
function finLavageListe_(p) {
  const date = String((p && p.date) || Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM-dd'));
  const sh = feuilleFinLavage_();
  const der = sh.getLastRow();
  const out = [];
  if (der >= 2) {
    const debut = Math.max(2, der - 600);            // les ~600 derniers relevés suffisent
    const v = sh.getRange(debut, 1, der - debut + 1, 6).getValues();
    v.forEach(function (r) {
      const j = (r[0] instanceof Date) ? Utilities.formatDate(r[0], 'Europe/Paris', 'yyyy-MM-dd') : String(r[0]);
      if (j !== date) return;
      const le = (r[1] instanceof Date) ? Utilities.formatDate(r[1], 'Europe/Paris', 'HH:mm') : String(r[1]);
      const fin = (r[2] instanceof Date) ? Utilities.formatDate(r[2], 'Europe/Paris', 'HH:mm') : String(r[2]);
      out.push({ le: le, fin: fin, reste: Number(r[3]) || 0, total: Number(r[4]) || 0, termine: String(r[5]) === 'oui' });
    });
  }
  const H = bilanTout_(), h = H.find(function (x) { return x.date === date; });
  // Journée trop ancienne pour les derniers relevés : on lit l'historique (une ligne par jour)
  if (!out.length && h && h.fin) out.push({ le: h.fin, fin: h.fin, reste: 0, total: h.total, termine: true });
  return { ok: true, date: date, releves: out, repassage: repassageInfo_(H, date, h) };
}

/* HISTORIQUE : une ligne par journée.
   A Journée · B Fin du lavage (figée quand tout est pointé) · C Galettes lavées · D Mis à jour le
   E Galettes Repassage (du Plan) · F Fin du train de repassage (départ du référent partie haute) */
function feuilleBilanLavage_() {
  const ss = SpreadsheetApp.openByUrl(URL_TABLEUR);
  let sh = ss.getSheetByName('Fin lavage — historique');
  if (!sh) {
    sh = ss.insertSheet('Fin lavage — historique');
    sh.appendRow(['Journée', 'Fin du lavage', 'Galettes lavées', 'Mis à jour le', 'Galettes repassage', 'Fin train repassage']);
    sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, 5).getValue()) === '') {
    sh.getRange(1, 4, 1, 3).setValues([['Mis à jour le', 'Galettes repassage', 'Fin train repassage']]);
  }
  return sh;
}
function txtHeure_(x) { return (x instanceof Date) ? Utilities.formatDate(x, 'Europe/Paris', 'HH:mm') : String(x || ''); }
function txtJour_(x) { return (x instanceof Date) ? Utilities.formatDate(x, 'Europe/Paris', 'yyyy-MM-dd') : String(x || ''); }
function bilanTout_() {
  const sh = feuilleBilanLavage_(), der = sh.getLastRow();
  if (der < 2) return [];
  return sh.getRange(2, 1, der - 1, 6).getValues().map(function (r, i) {
    return { ligne: i + 2, date: txtJour_(r[0]), fin: txtHeure_(r[1]), total: Number(r[2]) || 0, rep: Number(r[4]) || 0, repFin: txtHeure_(r[5]) };
  });
}
function bilanMaj_(date, m) {
  const sh = feuilleBilanLavage_(), H = bilanTout_();
  let x = H.find(function (y) { return y.date === date; });
  if (!x) { sh.appendRow(["'" + date, '', '', new Date(), '', '']); x = { ligne: sh.getLastRow() }; }
  if (m.termine === true) sh.getRange(x.ligne, 2, 1, 2).setValues([["'" + m.fin, m.total]]);
  else if (m.termine === false) sh.getRange(x.ligne, 2, 1, 2).setValues([['', '']]);   // la journée n'est plus terminée
  if (m.rep != null && m.rep !== '' && !isNaN(Number(m.rep))) sh.getRange(x.ligne, 5).setValue(Number(m.rep));
  if (m.repFin) sh.getRange(x.ligne, 6).setValue("'" + m.repFin);
  sh.getRange(x.ligne, 4).setValue(new Date());
}

/* ESTIMATION DE LA FIN DU TRAIN DE REPASSAGE
   On prend les journées précédentes (au plus les 10 dernières, sur 3 semaines) où l'on connaît
   les galettes Repassage ET l'heure de fin réelle, et on en tire la relation « galettes -> heure de fin »
   (droite ajustée). Elle se réajuste toute seule chaque jour. Moins de 3 journées : pas d'estimation. */
function minutes_(t) { const m = String(t || '').match(/(\d+):(\d+)/); if (!m) return null; let v = (+m[1]) * 60 + (+m[2]); if (v < 4 * 60) v += 1440; return v; }
function repassageInfo_(H, date, h) {
  const info = { fin: h && h.repFin ? h.repFin : null, galettes: h ? h.rep || 0 : 0, estimation: null, jours: 0 };
  const lim = new Date(date + 'T12:00:00'); lim.setDate(lim.getDate() - 21);
  const limite = Utilities.formatDate(lim, 'Europe/Paris', 'yyyy-MM-dd');
  const pts = H.filter(function (x) { return x.date < date && x.date >= limite && x.rep > 0 && minutes_(x.repFin) != null; })
    .sort(function (a, b) { return a.date < b.date ? -1 : 1; }).slice(-10)
    .map(function (x) { return { g: x.rep, m: minutes_(x.repFin) }; });
  info.jours = pts.length;
  if (pts.length < 3 || !info.galettes) return info;
  const n = pts.length, mg = pts.reduce(function (s, p) { return s + p.g; }, 0) / n, mm = pts.reduce(function (s, p) { return s + p.m; }, 0) / n;
  let sxy = 0, sxx = 0; pts.forEach(function (p) { sxy += (p.g - mg) * (p.m - mm); sxx += (p.g - mg) * (p.g - mg); });
  let pente = sxx > 0 ? sxy / sxx : 0;
  if (!(pente > 0)) pente = 0;                       // pas de lien clair : on prend la moyenne des fins
  let est = Math.round(mm + pente * (info.galettes - mg));
  est = Math.max(est, 12 * 60);
  info.estimation = String(Math.floor(est / 60) % 24).padStart(2, '0') + ':' + String(est % 60).padStart(2, '0');
  info.minutesParGalette = Math.round(pente * 10) / 10;
  return info;
}


/* =====================================================================
   ALARME « PAS DE POINTAGE DEPUIS 30 MIN » (laveur)
   Le Plan de travail envoie, avec l'heure de fin, l'heure du dernier « fait » du tunnel.
   Un déclencheur toutes les 5 min vérifie : s'il reste des galettes et que le dernier « fait »
   (ou le démarrage, si rien n'est pointé) date de 30 min ou plus, on prévient le responsable,
   UNE fois par trou de pointage.
   Réglages (Paramètres du projet → Propriétés du script) — à saisir vous-même, jamais dans le chat :
     ALERTE_EMAIL  adresse(s) mail du responsable, séparées par des virgules
     FREE_USER / FREE_PASS  identifiant + clé « Notifications par SMS » Free Mobile du responsable
                            (SMS gratuit, uniquement vers le téléphone du titulaire de la ligne Free)
   Installation : exécuter UNE fois installerAlarmePointage().  Essai : testerAlerte().
   ===================================================================== */
const ALARME_MIN = 30;

function alarmeMemoriser_(b) {
  const P = PropertiesService.getScriptProperties();
  let e = {}; try { e = JSON.parse(P.getProperty('ALARME_ETAT') || '{}'); } catch (x) {}
  if (e.date !== String(b.date)) e = { date: String(b.date) };
  e.reste = Number(b.reste) || 0; e.termine = !!b.termine;
  if (b.dernierFait !== undefined) e.dernierFait = String(b.dernierFait || '');
  if (b.demarrage) e.demarrage = String(b.demarrage);
  e.recu = new Date().toISOString();
  P.setProperty('ALARME_ETAT', JSON.stringify(e));
}

function verifierPointageLavage() {
  const P = PropertiesService.getScriptProperties();
  let e; try { e = JSON.parse(P.getProperty('ALARME_ETAT') || 'null'); } catch (x) { e = null; }
  if (!e || !e.date || e.termine || !(e.reste > 0)) return;
  const maint = new Date();
  const auj = Utilities.formatDate(maint, 'Europe/Paris', 'yyyy-MM-dd');
  let nMin = Number(Utilities.formatDate(maint, 'Europe/Paris', 'H')) * 60 + Number(Utilities.formatDate(maint, 'Europe/Paris', 'm'));
  const hier = Utilities.formatDate(new Date(maint.getTime() - 86400000), 'Europe/Paris', 'yyyy-MM-dd');
  if (e.date === hier && nMin < 4 * 60) nMin += 1440;          // journée qui déborde après minuit
  else if (e.date !== auj) return;
  const ref = e.dernierFait || e.demarrage || '06:00';
  const m = String(ref).match(/(\d+):(\d+)/); if (!m) return;
  let rMin = (+m[1]) * 60 + (+m[2]); if (nMin >= 1440 && rMin < 4 * 60) rMin += 1440;
  const ecart = nMin - rMin;
  if (ecart < ALARME_MIN) return;
  const cle = e.date + ' ' + ref;
  if (e.alerte === cle) return;                                 // déjà prévenu pour ce trou
  const texte = 'BdL lavage : aucune galette pointée ' + (e.dernierFait ? 'depuis ' + e.dernierFait : 'depuis le démarrage (' + ref + ')') +
    ' (' + ecart + ' min), ' + e.reste + ' galettes restantes.';
  envoyerAlerte_(texte);
  e.alerte = cle; P.setProperty('ALARME_ETAT', JSON.stringify(e));
}

function envoyerAlerte_(texte) {
  const P = PropertiesService.getScriptProperties();
  const mail = P.getProperty('ALERTE_EMAIL'), fu = P.getProperty('FREE_USER'), fp = P.getProperty('FREE_PASS');
  const res = [];
  if (fu && fp) {
    try {
      const r = UrlFetchApp.fetch('https://smsapi.free-mobile.fr/sendmsg?user=' + encodeURIComponent(fu) + '&pass=' + encodeURIComponent(fp) + '&msg=' + encodeURIComponent(texte), { muteHttpExceptions: true });
      res.push('SMS ' + r.getResponseCode());
    } catch (x) { res.push('SMS erreur ' + x); }
  }
  if (mail) { try { MailApp.sendEmail(mail, '⚠ Alarme pointage lavage', texte); res.push('mail ok'); } catch (x) { res.push('mail erreur ' + x); } }
  console.log('Alerte : ' + texte + ' → ' + (res.join(', ') || 'aucun destinataire réglé'));
  return res;
}

function installerAlarmePointage() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'verifierPointageLavage') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('verifierPointageLavage').timeBased().everyMinutes(5).create();
  console.log('Alarme installée : vérification toutes les 5 minutes.');
}

function testerAlerte() { console.log(envoyerAlerte_('BdL lavage : ceci est un essai de l\'alarme de pointage.').join(', ')); }
