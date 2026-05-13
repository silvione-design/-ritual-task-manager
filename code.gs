// ═══════════════════════════════════════════════════════════════
// RITUAL TASK MANAGER — Apps Script Backend
// Account: ritualbrasileiro@gmail.com
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME_TASKS   = 'Task';
const SHEET_NAME_PEOPLE  = 'Persone';
const SHEET_NAME_LOG     = 'Log';
const SHEET_NAME_CONFIG  = 'Config';

// ── SETUP INIZIALE ──────────────────────────────────────────────
// Esegui questa funzione UNA SOLA VOLTA per creare la struttura del foglio

function setupFoglio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── FOGLIO TASK ──
  let shTask = ss.getSheetByName(SHEET_NAME_TASKS);
  if (!shTask) shTask = ss.insertSheet(SHEET_NAME_TASKS);
  shTask.clearContents();
  shTask.getRange(1,1,1,13).setValues([[
    'ID','Nome','Categoria','FreqType','FreqValue','FreqGiornoMese','FreqGiornoSett',
    'OraInvio','PrimaryId','BackupId','Note','Attivo','UltimoInvio'
  ]]);
  shTask.getRange(1,1,1,13).setFontWeight('bold').setBackground('#1e2330').setFontColor('#f5a623');
  shTask.setFrozenRows(1);

  // Task di default
  const defaultTasks = [
    ['t1','Decalcificazione lavabicchieri','Manutenzione','giorno-settimana-n',2,'',2,'09:00','p2','p1','Usare decalcificante apposito. Ciclo completo senza bicchieri.',true,''],
    ['t2','Controllo bottiglie acqua 1,5L','Scorte','giorno-settimana','','',1,'09:00','p1','p2','Verificare scorte. Riordinare se sotto 24 bottiglie.',true,''],
    ['t3','Pulizia mantovana cappa','Pulizie','mesi',6,'','','08:00','p2','p3','Smontare filtri e sgrassare completamente.',true,''],
    ['t4','Pulizia filtri cappa aspirante','Pulizie','mesi',1,'','','08:00','p1','p2','Smontare e lavare in lavastoviglie. Asciugare prima di rimontare.',true,''],
    ['t5','Controllo temperatura frigoriferi','HACCP','giorni',1,'','','07:30','p2','p4','Celle: 0-4°C. Freezer: -18°C. Registrare su registro HACCP.',true,''],
    ['t6','Pulizia fondo frigoriferi','Pulizie','mesi',1,'','','08:00','p3','p1','Svuotare, pulire con igienizzante, asciugare.',true,''],
    ['t7','Controllo scadenze prodotti magazzino','HACCP','giorno-settimana','','',1,'08:00','p2','p1','Verificare date scadenza. Rimuovere prodotti scaduti.',true,''],
    ['t8','Pulizia pietre laviche','Manutenzione','giorno-settimana','','',5,'10:00','p1','p4','Spazzolare a freddo. Riposizionare nel trespolo.',true,''],
    ['t9','Verifica scorte prodotti pulizia','Scorte','giorno-settimana-n',2,'',1,'09:00','p3','p2','Detergenti, igienizzanti, guanti, sacchi. Avvisare se sotto soglia.',true,''],
    ['t10','Pulizia vetrine e specchi','Pulizie','giorni',3,'','','09:00','p1','p3','Spray per vetri. Controllare anche specchi bagni.',true,''],
    ['t11','Invio ore al consulente paghe','Amministrazione','giorno-fisso-mese','',4,'','08:00','p2','','Inviare prospetto ore entro il 4 del mese.',true,''],
    ['t12','Sanificazione superfici cucina','HACCP','giorni',1,'','','07:30','p4','p1','Piani lavoro, taglieri, attrezzature. Igienizzante CE.',true,''],
    ['t13','Cambio olio friggitrice','Manutenzione','giorno-settimana','','',5,'16:00','p1','p4','Svuotare, pulire vasca, olio nuovo. Smaltire il vecchio.',true,''],
    ['t14','Stampa stock vini per ordine','Scorte','giorno-settimana','','',1,'09:00','p2','p1','Stampare giacenze vini. Consegnare a Silvio per ordine.',true,''],
  ];
  if (shTask.getLastRow() < 2) {
    shTask.getRange(2,1,defaultTasks.length,13).setValues(defaultTasks);
  }

  // ── FOGLIO PERSONE ──
  let shPeople = ss.getSheetByName(SHEET_NAME_PEOPLE);
  if (!shPeople) shPeople = ss.insertSheet(SHEET_NAME_PEOPLE);
  shPeople.clearContents();
  shPeople.getRange(1,1,1,4).setValues([['ID','Nome','WhatsApp','Email']]);
  shPeople.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1e2330').setFontColor('#f5a623');
  shPeople.setFrozenRows(1);
  const defaultPeople = [
    ['p1','Marco','+39 333 0000001','marco@ritual.it'],
    ['p2','María','+39 333 0000002','maria@ritual.it'],
    ['p3','Rosario','+39 333 0000003','rosario@ritual.it'],
    ['p4','Gaetano','+39 333 0000004','gaetano@ritual.it'],
  ];
  if (shPeople.getLastRow() < 2) {
    shPeople.getRange(2,1,defaultPeople.length,4).setValues(defaultPeople);
  }

  // ── FOGLIO LOG ──
  let shLog = ss.getSheetByName(SHEET_NAME_LOG);
  if (!shLog) shLog = ss.insertSheet(SHEET_NAME_LOG);
  shLog.clearContents();
  shLog.getRange(1,1,1,6).setValues([['Timestamp','TaskID','NomeTask','Destinatario','Email','Esito']]);
  shLog.getRange(1,1,1,6).setFontWeight('bold').setBackground('#1e2330').setFontColor('#f5a623');
  shLog.setFrozenRows(1);

  // ── FOGLIO CONFIG ──
  let shConfig = ss.getSheetByName(SHEET_NAME_CONFIG);
  if (!shConfig) shConfig = ss.insertSheet(SHEET_NAME_CONFIG);
  shConfig.clearContents();
  shConfig.getRange(1,1,5,2).setValues([
    ['Chiave','Valore'],
    ['email_admin','ritualbrasileiro@gmail.com'],
    ['nome_ristorante','Ritual Brasileiro'],
    ['password_silvio','ritual2024'],
    ['password_catia','ritual2024'],
  ]);
  shConfig.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1e2330').setFontColor('#f5a623');

  SpreadsheetApp.getUi().alert('✅ Setup completato! Struttura foglio creata correttamente.');
}

// ── TRIGGER SETUP ───────────────────────────────────────────────
// Esegui questa funzione UNA SOLA VOLTA per attivare i reminder automatici

function installaTrigger() {
  // Rimuovi trigger esistenti per evitare duplicati
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'controllaEInviaReminder') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Trigger ogni ora — lo script controlla l'orario di ogni task
  ScriptApp.newTrigger('controllaEInviaReminder')
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert('✅ Trigger installato! I reminder partiranno automaticamente ogni ora.');
}

// ── CORE: CONTROLLA E INVIA REMINDER ───────────────────────────

function controllaEInviaReminder() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const tasks    = leggiTask(ss);
  const people   = leggiPersone(ss);
  const config   = leggiConfig(ss);
  const now      = new Date();
  const oraOra   = now.getHours();
  const minOra   = now.getMinutes();

  tasks.forEach(task => {
    if (!task.attivo) return;

    // Controlla se oggi è il giorno giusto per questo task
    if (!eTocca(task, now)) return;

    // Controlla se l'orario corrisponde (finestra di 59 minuti)
    const [hTask, mTask] = (task.oraInvio || '09:00').split(':').map(Number);
    if (oraOra !== hTask) return;

    // Controlla se già inviato oggi
    if (giaInviatoOggi(task, now)) return;

    // Trova responsabile
    const primary = people.find(p => p.id === task.primaryId);
    const backup  = people.find(p => p.id === task.backupId);
    const dest    = primary || backup;
    if (!dest || !dest.email) return;

    // Invia email
    const esito = inviaEmailReminder(dest, task, config, now);

    // Aggiorna UltimoInvio nel foglio
    aggiornaUltimoInvio(ss, task.id, now);

    // Log
    scriviLog(ss, task, dest, esito);
  });
}

// ── LOGICA FREQUENZE ────────────────────────────────────────────

function eTocca(task, now) {
  const oggi     = new Date(now); oggi.setHours(0,0,0,0);
  const dowOggi  = oggi.getDay() === 0 ? 7 : oggi.getDay(); // 1=lun…7=dom
  const domMese  = oggi.getDate();

  switch (task.freqType) {

    case 'giorni': {
      if (!task.ultimoInvio) return true;
      const last = new Date(task.ultimoInvio); last.setHours(0,0,0,0);
      const diff = Math.round((oggi - last) / 86400000);
      return diff >= task.freqValue;
    }

    case 'settimane': {
      if (!task.ultimoInvio) return true;
      const last = new Date(task.ultimoInvio); last.setHours(0,0,0,0);
      const diff = Math.round((oggi - last) / 86400000);
      return diff >= task.freqValue * 7;
    }

    case 'mesi': {
      if (!task.ultimoInvio) return true;
      const last = new Date(task.ultimoInvio);
      const mesiPassati = (oggi.getFullYear() - last.getFullYear()) * 12 + (oggi.getMonth() - last.getMonth());
      return mesiPassati >= task.freqValue;
    }

    case 'giorno-fisso-mese': {
      if (domMese !== task.freqGiornoMese) return false;
      if (!task.ultimoInvio) return true;
      const last = new Date(task.ultimoInvio);
      return !(last.getMonth() === oggi.getMonth() && last.getFullYear() === oggi.getFullYear());
    }

    case 'giorno-settimana': {
      if (dowOggi !== task.freqGiornoSett) return false;
      if (!task.ultimoInvio) return true;
      const last = new Date(task.ultimoInvio); last.setHours(0,0,0,0);
      return Math.round((oggi - last) / 86400000) >= 7;
    }

    case 'giorno-settimana-n': {
      if (dowOggi !== task.freqGiornoSett) return false;
      if (!task.ultimoInvio) return true;
      const last = new Date(task.ultimoInvio); last.setHours(0,0,0,0);
      return Math.round((oggi - last) / 86400000) >= task.freqValue * 7;
    }

    default: return false;
  }
}

function giaInviatoOggi(task, now) {
  if (!task.ultimoInvio) return false;
  const last = new Date(task.ultimoInvio);
  const oggi = new Date(now);
  return last.getDate()===oggi.getDate() && last.getMonth()===oggi.getMonth() && last.getFullYear()===oggi.getFullYear();
}

// ── EMAIL ───────────────────────────────────────────────────────

function inviaEmailReminder(dest, task, config, now) {
  const nomeRistorante = config['nome_ristorante'] || 'Ritual Brasileiro';
  const freqTesto = formattaFreq(task);
  const subject = `🔔 Reminder ${nomeRistorante} — ${task.nome}`;
  const body = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:'Helvetica Neue',Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:30px 0;">
    <tr><td align="center">
      <table width="580" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.1);">
        <!-- Header -->
        <tr><td style="background:#0f1117;padding:28px 32px;">
          <p style="margin:0;font-size:11px;letter-spacing:3px;color:#8891a8;text-transform:uppercase;">${nomeRistorante}</p>
          <p style="margin:6px 0 0;font-size:22px;font-weight:700;color:#f5a623;">📋 Reminder Procedura</p>
        </td></tr>
        <!-- Body -->
        <tr><td style="padding:32px;">
          <p style="margin:0 0 8px;font-size:15px;color:#444;">Ciao <strong>${dest.nome}</strong>,</p>
          <p style="margin:0 0 24px;font-size:15px;color:#444;">Oggi tocca a te eseguire questa procedura:</p>

          <div style="background:#f8f9fa;border-left:4px solid #f5a623;border-radius:6px;padding:20px 24px;margin-bottom:24px;">
            <p style="margin:0 0 6px;font-size:19px;font-weight:700;color:#0f1117;">${task.nome}</p>
            <p style="margin:0;font-size:13px;color:#8891a8;">Categoria: ${task.categoria} &nbsp;·&nbsp; Frequenza: ${freqTesto}</p>
          </div>

          ${task.note ? `
          <div style="background:#fffbf0;border:1px solid #f5e6c3;border-radius:6px;padding:16px 20px;margin-bottom:24px;">
            <p style="margin:0 0 4px;font-size:12px;font-weight:600;color:#c4831a;text-transform:uppercase;letter-spacing:1px;">📌 Note</p>
            <p style="margin:0;font-size:14px;color:#555;">${task.note}</p>
          </div>` : ''}

          <p style="margin:24px 0 0;font-size:13px;color:#aaa;">Questo è un reminder automatico — ${now.toLocaleDateString('it-IT')}</p>
        </td></tr>
        <!-- Footer -->
        <tr><td style="background:#0f1117;padding:16px 32px;">
          <p style="margin:0;font-size:12px;color:#4a5068;">© ${nomeRistorante} · Gestito da Ritual Task Manager</p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;

  try {
    GmailApp.sendEmail(dest.email, subject, '', { htmlBody: body });
    return 'OK';
  } catch(e) {
    Logger.log('Errore invio email a ' + dest.email + ': ' + e.message);
    return 'ERRORE: ' + e.message;
  }
}

// ── LETTURA DATI ────────────────────────────────────────────────

function leggiTask(ss) {
  const sh   = ss.getSheetByName(SHEET_NAME_TASKS);
  const data = sh.getDataRange().getValues();
  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    tasks.push({
      id:            r[0],
      nome:          r[1],
      categoria:     r[2],
      freqType:      r[3],
      freqValue:     Number(r[4]) || 0,
      freqGiornoMese:Number(r[5]) || 0,
      freqGiornoSett:Number(r[6]) || 0,
      oraInvio:      r[7] || '09:00',
      primaryId:     r[8],
      backupId:      r[9],
      note:          r[10],
      attivo:        r[11] === true || r[11] === 'TRUE' || r[11] === 1,
      ultimoInvio:   r[12] ? new Date(r[12]) : null,
      riga:          i + 1,
    });
  }
  return tasks;
}

function leggiPersone(ss) {
  const sh   = ss.getSheetByName(SHEET_NAME_PEOPLE);
  const data = sh.getDataRange().getValues();
  const people = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    people.push({ id:r[0], nome:r[1], whatsapp:r[2], email:r[3] });
  }
  return people;
}

function leggiConfig(ss) {
  const sh   = ss.getSheetByName(SHEET_NAME_CONFIG);
  const data = sh.getDataRange().getValues();
  const cfg  = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) cfg[data[i][0]] = data[i][1];
  }
  return cfg;
}

// ── SCRITTURA DATI ──────────────────────────────────────────────

function aggiornaUltimoInvio(ss, taskId, now) {
  const sh   = ss.getSheetByName(SHEET_NAME_TASKS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) {
      sh.getRange(i+1, 13).setValue(now.toISOString());
      break;
    }
  }
}

function scriviLog(ss, task, dest, esito) {
  const sh  = ss.getSheetByName(SHEET_NAME_LOG);
  const now = new Date();
  sh.appendRow([now, task.id, task.nome, dest.nome, dest.email, esito]);
}

// ── FORMATO FREQUENZA ───────────────────────────────────────────

function formattaFreq(task) {
  const DSET = ['','Lunedì','Martedì','Mercoledì','Giovedì','Venerdì','Sabato','Domenica'];
  switch(task.freqType) {
    case 'giorni':            return task.freqValue===1 ? 'ogni giorno' : `ogni ${task.freqValue} giorni`;
    case 'settimane':         return task.freqValue===1 ? 'ogni settimana' : `ogni ${task.freqValue} settimane`;
    case 'mesi':              return task.freqValue===1 ? 'ogni mese' : `ogni ${task.freqValue} mesi`;
    case 'giorno-fisso-mese': return `il ${task.freqGiornoMese} di ogni mese`;
    case 'giorno-settimana':  return `ogni ${DSET[task.freqGiornoSett]}`;
    case 'giorno-settimana-n':return `ogni ${task.freqValue} settimane — ${DSET[task.freqGiornoSett]}`;
    default: return '';
  }
}

// ── WEB APP ─────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ritual Task Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// API chiamate dall'interfaccia HTML

function apiGetAll() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const tasks  = leggiTask(ss);
  const people = leggiPersone(ss);
  const config = leggiConfig(ss);
  return JSON.stringify({ tasks, people, config });
}

function apiSaveTask(taskJson) {
  const task = JSON.parse(taskJson);
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(SHEET_NAME_TASKS);
  const data = sh.getDataRange().getValues();
  const row  = [
    task.id, task.nome, task.categoria, task.freqType,
    task.freqValue||'', task.freqGiornoMese||'', task.freqGiornoSett||'',
    task.oraInvio||'09:00', task.primaryId||'', task.backupId||'',
    task.note||'', task.attivo, task.ultimoInvio||''
  ];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === task.id) {
      sh.getRange(i+1, 1, 1, 13).setValues([row]);
      return 'updated';
    }
  }
  sh.appendRow(row);
  return 'inserted';
}

function apiDeleteTask(taskId) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(SHEET_NAME_TASKS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) { sh.deleteRow(i+1); return 'deleted'; }
  }
  return 'not found';
}

function apiSavePerson(personJson) {
  const person = JSON.parse(personJson);
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sh     = ss.getSheetByName(SHEET_NAME_PEOPLE);
  const data   = sh.getDataRange().getValues();
  const row    = [person.id, person.nome, person.whatsapp||'', person.email||''];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === person.id) { sh.getRange(i+1,1,1,4).setValues([row]); return 'updated'; }
  }
  sh.appendRow(row);
  return 'inserted';
}

function apiDeletePerson(personId) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(SHEET_NAME_PEOPLE);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === personId) { sh.deleteRow(i+1); return 'deleted'; }
  }
  return 'not found';
}

function apiMarkSent(taskId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  aggiornaUltimoInvio(ss, taskId, new Date());
  return 'ok';
}

function apiLogin(user, pwd) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const config = leggiConfig(ss);
  if (user === 'silvio' && pwd === config['password_silvio']) return 'ok';
  if (user === 'catia'  && pwd === config['password_catia'])  return 'ok';
  return 'error';
}

// ── TEST MANUALE ─────────────────────────────────────────────────
// Esegui per testare l'invio senza aspettare il trigger automatico

function testInvioManuale() {
  controllaEInviaReminder();
  SpreadsheetApp.getUi().alert('✅ Test eseguito. Controlla il foglio Log e le email.');
}
