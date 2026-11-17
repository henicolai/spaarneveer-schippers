// RESERVEKOPIE van het Google Apps Script 'Spaarneveer schippersrooster'.
// De echte codes zijn hier weggelaten (deze repository is openbaar).
// Bij herinstallatie: vul SCHIPPERSCODE en BEHEERCODE in, en het Sheet-ID bij openById.

/**
 * Spaarneveer schippersrooster — opslag in de Google Sheet 'Spaarneveer wijzigingen'.
 * Na een wijziging: Implementeren > Implementaties beheren > potlood > Nieuwe versie > Implementeren.
 */

const SCHIPPERSCODE = 'VUL_HIER_DE_SCHIPPERSCODE_IN';   // code die alle schippers krijgen
const BEHEERCODE    = 'VUL_HIER_DE_BEHEERCODE_IN';  // alleen voor de beheerder

const PAGINA_URL = 'https://henicolai.github.io/spaarneveer-schippers/';
const BEHEERDER_EMAIL = 'hnicolai@gmail.com';   // afzender/ontvanger van de mails (schippers in BCC)
const TAB_WIJZ = 'Wijzigingen';
const TAB_BER  = 'Mededelingen';

function sheet_(naam, kop) {
  const ss = SpreadsheetApp.openById('10Yq7jPnUFWJdhryzoRmuHbk5UqX-b92Ylvf4fbk-esg');  // Sheet 'Spaarneveer wijzigingen'
  let sh = ss.getSheetByName(naam);
  if (!sh) { sh = ss.insertSheet(naam); sh.appendRow(kop); sh.setFrozenRows(1); }
  return sh;
}
const schippersSheet_ = () => sheet_('Schippers', ['naam', 'email', 'mail (ja/nee)']);
const roosterSheet_ = () => sheet_('Rooster', ['datum', '1a', '1b', '2a', '2b', '3a', '3b', 'opmerking']);
const wijzSheet_ = () => sheet_(TAB_WIJZ, ['sleutel', 'naam', 'aangeboden', 'vorige', 'tijd', 'aan']);
const berSheet_  = () => sheet_(TAB_BER,  ['id', 'tekst', 'tijd']);
const passSheet_ = () => sheet_('Passagiers', ['dienst', 'aantal', 'ingevuld door', 'tijd']);

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function lees_() {
  const wijz = {};
  wijzSheet_().getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) wijz[String(r[0])] = { naam: String(r[1] || ''), aangeboden: r[2] === true || r[2] === 'TRUE', vorige: String(r[3] || ''), ts: Number(r[4]) || 0, aan: String(r[5] || '') };
  });
  const berichten = berSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0]).map(r => ({ id: String(r[0]), text: String(r[1]), ts: Number(r[2]) || 0 }));
  const rooster = {};
  const tz = Session.getScriptTimeZone();
  roosterSheet_().getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    const ds = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]).replace(/^'/, '');
    rooster[ds] = { p: r.slice(1, 7).map(x => String(x || '')), opm: String(r[7] || '') };
  });
  const passagiers = {};
  passSheet_().getDataRange().getValues().slice(1).forEach(r => { if (r[0] !== '' && r[1] !== '') passagiers[String(r[0])] = Number(r[1]); });
  const namen = schippersSheet_().getDataRange().getValues().slice(1).map(r => String(r[0] || '').trim()).filter(Boolean);
  return { ok: true, wijzigingen: wijz, berichten, rooster, passagiers, namen };
}


const DAGEN_ = ['zondag','maandag','dinsdag','woensdag','donderdag','vrijdag','zaterdag'];
const MAANDEN_ = ['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december'];
function dienstTekst_(key) {
  const [ds, pos] = key.split('_');
  const [y, m, d] = ds.split('-').map(Number);
  const dt = new Date(y, m - 1, d);
  const weekend = dt.getDay() === 0 || dt.getDay() === 6;
  const tijden = weekend ? ['9.00–12.00', '12.00–15.00', '15.00–18.00'] : ['7.30–11.00', '11.00–14.30', '14.30–18.00'];
  const n = Number(pos[0]);
  return DAGEN_[dt.getDay()] + ' ' + d + ' ' + MAANDEN_[m - 1] + ', dienst ' + n + ' (' + tijden[n - 1] + ')';
}
function adresVan_(naam) {
  const norm = v => String(v || '').trim().toLowerCase();
  const rijen = schippersSheet_().getDataRange().getValues().slice(1);
  const r = rijen.find(x => norm(x[0]) === norm(naam) && String(x[1] || '').includes('@'));
  if (!r) console.warn('Geen e-mailadres gevonden voor ' + naam + '. Namen in tab Schippers: ' + rijen.map(x => '[' + x[0] + '|' + (String(x[1] || '').includes('@') ? 'mail' : 'geen') + ']').join(' '));
  return r ? [String(r[1]).trim()] : [];
}
function mailOntvangers_(behalveNaam) {
  return schippersSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[1] && String(r[1]).includes('@') && String(r[2] || 'ja').toLowerCase() !== 'nee' && String(r[0]) !== behalveNaam)
    .map(r => String(r[1]).trim());
}
function stuurAanbodMail_(key, naam, aan) {
  try {
    const lijst = aan ? adresVan_(aan) : mailOntvangers_(naam);
    if (!lijst.length) return 0;
    if (MailApp.getRemainingDailyQuota() < lijst.length) { console.warn('Mailquotum op, geen mail voor ' + key); return -1; }
    const wat = dienstTekst_(key);
    MailApp.sendEmail({
      to: BEHEERDER_EMAIL,
      bcc: lijst.join(','),
      subject: 'Spaarneveer: dienst ter overname – ' + wat,
      body: (aan ? 'Hoi ' + aan + ',\n\n' + naam + ' kan niet varen en vraagt of jij deze dienst wilt overnemen:\n\n' : naam + ' kan niet varen en biedt deze dienst ter overname aan:\n\n') + wat + '\n\n' +
            'Kun jij? Tik op deze link en dan op Overnemen:\n' + PAGINA_URL + '#d-' + key + '\n\n' +
            '(Wie het eerst overneemt, vaart. Geen mail meer ontvangen? Laat het de beheerder weten.)',
      name: 'Spaarneveer schippersrooster'
    });
    return lijst.length;
  } catch (err) { console.error('Mail mislukt: ' + err); return -2; }
}
// Mail naar iedereen zodra een aangeboden dienst is overgenomen.
function stuurOvernameMail_(key, naam, vorigeNaam, alleen) {
  try {
    const lijst = alleen ? adresVan_(alleen) : mailOntvangers_(naam);
    if (!lijst.length) return;
    if (MailApp.getRemainingDailyQuota() < lijst.length) { console.warn('Mailquotum op, geen overnamemail voor ' + key); return; }
    const wat = dienstTekst_(key);
    MailApp.sendEmail({
      to: BEHEERDER_EMAIL,
      bcc: lijst.join(','),
      subject: 'Spaarneveer: dienst overgenomen – ' + wat,
      body: naam + ' heeft deze dienst overgenomen' + (vorigeNaam ? ' van ' + vorigeNaam : '') + ':\n\n' + wat + '\n\n' +
            'De dienst is niet meer beschikbaar; je hoeft niets te doen.\n' + PAGINA_URL + '#d-' + key + '\n\n' +
            '(Geen mail meer ontvangen? Laat het de beheerder weten.)',
      name: 'Spaarneveer schippersrooster'
    });
  } catch (err) { console.error('Overnamemail mislukt: ' + err); }
}
// ---- Wachtrij: aanbiedingen bundelen (elke 5 minuten één mail per aanbieder) ----
const wachtSheet_ = () => sheet_('Wachtrij', ['sleutel', 'naam', 'aan', 'tijd']);
function inWachtrij_(key, naam, aan) {
  wachtSheet_().appendRow([key, naam, aan || '', new Date()]);
  const lijst = aan ? adresVan_(aan) : mailOntvangers_(naam);
  return lijst.length;
}
function uitWachtrij_(key) {
  const sh = wachtSheet_(); const v = sh.getDataRange().getValues();
  if (v.length < 2) return false;
  const rest = v.slice(1).filter(r => String(r[0]) !== key);
  if (rest.length === v.length - 1) return false;
  sh.getRange(2, 1, v.length - 1, v[0].length).clearContent();
  if (rest.length) sh.getRange(2, 1, rest.length, v[0].length).setValues(rest);
  return true;
}
function verwerkWachtrij() {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = wachtSheet_(); const v = sh.getDataRange().getValues();
    if (v.length < 2) return;
    sh.getRange(2, 1, v.length - 1, v[0].length).clearContent();
    const wijz = {};
    wijzSheet_().getDataRange().getValues().slice(1).forEach(r => { if (r[0]) wijz[String(r[0])] = { naam: String(r[1] || ''), aangeboden: r[2] === true || r[2] === 'TRUE' }; });
    const groepen = {};
    v.slice(1).forEach(r => {
      const key = String(r[0] || ''), naam = String(r[1] || ''), aan = String(r[2] || '');
      const w = wijz[key];
      if (!key || !w || !w.aangeboden || w.naam !== naam) return;   // intussen ingetrokken of overgenomen
      const g = naam + '|' + aan;
      (groepen[g] = groepen[g] || { naam, aan, keys: [] });
      if (groepen[g].keys.indexOf(key) < 0) groepen[g].keys.push(key);
    });
    Object.keys(groepen).forEach(g => stuurAanbodBundel_(groepen[g].naam, groepen[g].aan, groepen[g].keys.sort()));
  } finally { lock.releaseLock(); }
}
function stuurAanbodBundel_(naam, aan, keys) {
  if (keys.length === 1) return stuurAanbodMail_(keys[0], naam, aan);
  try {
    const lijst = aan ? adresVan_(aan) : mailOntvangers_(naam);
    if (!lijst.length) return 0;
    if (MailApp.getRemainingDailyQuota() < lijst.length) { console.warn('Mailquotum op, geen bundelmail voor ' + naam); return -1; }
    const regels = keys.map(k => '- ' + dienstTekst_(k) + '\n  ' + PAGINA_URL + '#d-' + k).join('\n\n');
    MailApp.sendEmail({
      to: BEHEERDER_EMAIL,
      bcc: lijst.join(','),
      subject: 'Spaarneveer: ' + keys.length + ' diensten ter overname – ' + naam,
      body: (aan ? 'Hoi ' + aan + ',\n\n' + naam + ' kan niet varen en vraagt of jij deze diensten wilt overnemen:\n\n' : naam + ' kan niet varen en biedt deze diensten ter overname aan:\n\n') + regels + '\n\n' +
            'Kun jij er een of meer? Tik op de link en dan op Overnemen.\n\n' +
            '(Wie het eerst overneemt, vaart. Geen mail meer ontvangen? Laat het de beheerder weten.)',
      name: 'Spaarneveer schippersrooster'
    });
    return lijst.length;
  } catch (err) { console.error('Bundelmail mislukt: ' + err); return -2; }
}
// Eenmalig uitvoeren vanuit de editor: zet de 5-minuten-trigger voor de wachtrij aan.
function zetWachtrijTriggerAan() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'verwerkWachtrij').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('verwerkWachtrij').timeBased().everyMinutes(5).create();
}
// Eenmalig uitvoeren vanuit de editor om toestemming voor e-mail te geven.
function testMail() {
  MailApp.sendEmail(BEHEERDER_EMAIL, 'Spaarneveer: test e-mail', 'Het schippersrooster kan nu e-mails versturen.\n' + PAGINA_URL);
}

function doGet() {
  return json_(lees_());
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, fout: 'Ongeldig verzoek' }); }
  const beheer = req.code === BEHEERCODE;
  if (req.actie === 'check') return json_({ ok: req.code === SCHIPPERSCODE || beheer, beheer });
  if (req.code !== SCHIPPERSCODE && !beheer) return json_({ ok: false, fout: 'Onjuiste schipperscode' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (req.actie === 'zet' || req.actie === 'wis') {
      const key = String(req.sleutel || '');
      if (!/^\d{4}-\d{2}-\d{2}_[123][ab]$/.test(key)) return json_({ ok: false, fout: 'Ongeldige dienst' });
      const sh = wijzSheet_();
      const rows = sh.getDataRange().getValues();
      let idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === key);
      // controle: is de dienst intussen door iemand anders veranderd?
      const huidig = idx > 0 ? String(rows[idx][1] || '') : null;
      if ('verwacht' in req && huidig !== req.verwacht) {
        return json_({ ok: false, fout: 'Deze dienst is net door iemand anders gewijzigd. Ververs de pagina.', ...lees_() });
      }
      let gemaild = null;
      if (req.actie === 'wis') {
        if (idx > 0) sh.deleteRow(idx + 1);
      } else {
        const naam = String(req.naam || '').slice(0, 40);
        const wasAangeboden = idx > 0 && (rows[idx][2] === true || rows[idx][2] === 'TRUE');
        const vorigAan = idx > 0 ? String(rows[idx][5] || '') : '';
        const aan = req.aangeboden ? String(req.aan || '').slice(0, 40) : '';
        const rij = [key, naam, !!req.aangeboden, String(req.vorige || '').slice(0, 40), Date.now(), aan];
        if (idx > 0) sh.getRange(idx + 1, 1, 1, 6).setValues([rij]); else sh.appendRow(rij);
        const stil = beheer && req.stil === true;   // beheerwijziging zonder e-mail
        if (!stil && req.aangeboden && !wasAangeboden) gemaild = inWachtrij_(key, naam, aan);
        const vorigeNaam = idx > 0 ? String(rows[idx][1] || '') : '';
        const nogInWachtrij = wasAangeboden && !req.aangeboden ? uitWachtrij_(key) : false;   // aanbod was nog niet gemaild
        if (!stil && !nogInWachtrij && wasAangeboden && !req.aangeboden && naam && naam !== vorigeNaam) stuurOvernameMail_(key, naam, vorigeNaam, vorigAan ? vorigeNaam : '');
      }
      return json_(Object.assign(lees_(), { gemaild }));
    }
    if (req.actie === 'rooster' && beheer) {
      const dagen = req.dagen || {};
      const sh = roosterSheet_();
      sh.getRange('A:A').setNumberFormat('@');
      const rows = sh.getDataRange().getValues();
      const idx = {};
      rows.forEach((r, i) => { if (i > 0 && r[0]) idx[r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[0])] = i + 1; });
      const nieuw = [];
      Object.keys(dagen).sort().forEach(ds => {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) return;
        const d = dagen[ds] || {};
        const p = (Array.isArray(d.p) ? d.p : []).slice(0, 6).map(x => String(x || '').slice(0, 40));
        while (p.length < 6) p.push('');
        const rij = [ds].concat(p, [String(d.opm || '').slice(0, 80)]);
        if (idx[ds]) sh.getRange(idx[ds], 1, 1, 8).setValues([rij]); else nieuw.push(rij);
      });
      if (nieuw.length) sh.getRange(sh.getLastRow() + 1, 1, nieuw.length, 8).setValues(nieuw);
      return json_(lees_());
    }
    if (req.actie === 'schippers' && beheer) {
      const sh = schippersSheet_();
      const bekend = sh.getDataRange().getValues().slice(1).map(r => String(r[0]));
      const nieuw = (req.namen || []).map(String).filter(n => n && bekend.indexOf(n) < 0).map(n => [n.slice(0, 40), '', 'ja']);
      if (nieuw.length) sh.getRange(sh.getLastRow() + 1, 1, nieuw.length, 3).setValues(nieuw);
      return json_({ ok: true, toegevoegd: nieuw.length });
    }
    if (req.actie === 'passagiers') {
      const key = String(req.sleutel || '');
      if (!/^\d{4}-\d{2}-\d{2}_[123]$/.test(key)) return json_({ ok: false, fout: 'Ongeldige dienst' });
      const v = String(req.aantal === undefined ? '' : req.aantal).trim();
      if (v !== '' && !/^\d{1,4}$/.test(v)) return json_({ ok: false, fout: 'Ongeldig aantal' });
      const sh = passSheet_();
      sh.getRange('A:A').setNumberFormat('@');
      const rows = sh.getDataRange().getValues();
      const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === key);
      if (v === '') { if (idx > 0) sh.deleteRow(idx + 1); }
      else {
        const rij = [key, Number(v), String(req.naam || '').slice(0, 40), new Date()];
        if (idx > 0) sh.getRange(idx + 1, 1, 1, 4).setValues([rij]); else sh.appendRow(rij);
      }
      return json_(lees_());
    }
    if (req.actie === 'hernoem' && beheer) {
      // naam van een schipper overal in de Sheet vervangen (Rooster, Wijzigingen, Schippers)
      const oud = String(req.oud || '').trim(), nieuw = String(req.nieuw || '').trim().slice(0, 40);
      if (!oud || !nieuw) return json_({ ok: false, fout: 'Oude en nieuwe naam nodig' });
      const aantal = {};
      [['Rooster', [1, 2, 3, 4, 5, 6]], [TAB_WIJZ, [1, 3, 5]], ['Schippers', [0]]].forEach(([tab, kol]) => {
        const sh = sheet_(tab, []); const v = sh.getDataRange().getValues(); let n = 0;
        for (let i = 1; i < v.length; i++) kol.forEach(k => { if (k < v[i].length && String(v[i][k]).trim() === oud) { sh.getRange(i + 1, k + 1).setValue(nieuw); n++; } });
        aantal[tab] = n;
      });
      return json_(Object.assign(lees_(), { hernoemd: aantal }));
    }
    if (req.actie === 'verwijderschipper' && beheer) {
      // schipper uit tabblad Schippers halen (rooster blijft ongewijzigd)
      const naam = String(req.naam || '').trim(); if (!naam) return json_({ ok: false, fout: 'Naam nodig' });
      const sh = schippersSheet_(); const v = sh.getDataRange().getValues(); let n = 0;
      for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]).trim() === naam) { sh.deleteRow(i + 1); n++; }
      return json_(Object.assign(lees_(), { verwijderd: n }));
    }
    if (req.actie === 'bericht' && beheer) {
      const t = String(req.tekst || '').trim().slice(0, 2000);
      if (t) berSheet_().appendRow(['b' + Date.now(), t, Date.now()]);
      return json_(lees_());
    }
    if (req.actie === 'wisbericht' && beheer) {
      const sh = berSheet_(); const rows = sh.getDataRange().getValues();
      const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === String(req.id));
      if (idx > 0) sh.deleteRow(idx + 1);
      return json_(lees_());
    }
    return json_({ ok: false, fout: 'Niet toegestaan' });
  } finally {
    lock.releaseLock();
  }
}
