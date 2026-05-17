import { google } from 'googleapis';

const SHEET_ID = process.env.GOOGLE_SHEET_ID;

async function getAuth() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
}

async function getSheets() {
  const auth = await getAuth();
  return google.sheets({ version: 'v4', auth });
}

async function readSheet(sheets, range) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  return res.data.values || [];
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { action } = req.query;

  try {
    const sheets = await getSheets();

    // ── LOGIN ──
    if (action === 'login') {
      const { user, pwd } = req.body;
      const rows = await readSheet(sheets, 'Config!A:B');
      const cfg = {};
      rows.slice(1).forEach(r => { if (r[0]) cfg[r[0]] = r[1]; });
      if (user === 'silvio' && pwd === cfg['password_silvio']) return res.json({ ok: true });
      if (user === 'catia'  && pwd === cfg['password_catia'])  return res.json({ ok: true });
      return res.json({ ok: false });
    }

    // ── SAVE TASK ──
    if (action === 'saveTask') {
      const task = req.body;
      const rows = await readSheet(sheets, 'Task!A:M');
      const row = [
        task.id, task.nome, task.categoria, task.freqType,
        task.freqValue || '', task.freqGiornoMese || '', task.freqGiornoSett || '',
        task.oraInvio || '09:00', task.primaryId || '', task.backupId || '',
        task.note || '', task.attivo ? 'TRUE' : 'FALSE', task.ultimoInvio || ''
      ];
      const idx = rows.findIndex(r => r[0] === task.id);
      if (idx >= 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `Task!A${idx + 1}:M${idx + 1}`,
          valueInputOption: 'RAW',
          requestBody: { values: [row] },
        });
        return res.json({ result: 'updated' });
      } else {
        await sheets.spreadsheets.values.append({
          spreadsheetId: SHEET_ID,
          range: 'Task!A:M',
          valueInputOption: 'RAW',
          requestBody: { values: [row] },
        });
        return res.json({ result: 'inserted' });
      }
    }

    // ── DELETE TASK ──
    if (action === 'deleteTask') {
      const { id } = req.body;
      const rows = await readSheet(sheets, 'Task!A:A');
      const idx = rows.findIndex(r => r[0] === id);
      if (idx >= 0) {
        // Get spreadsheet to find sheet ID
        const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
        const sheet = meta.data.sheets.find(s => s.properties.title === 'Task');
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: {
            requests: [{
              deleteDimension: {
                range: { sheetId: sheet.properties.sheetId, dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 }
              }
            }]
          }
        });
      }
      return res.json({ result: 'deleted' });
    }

    // ── SAVE PERSON ──
    if (action === 'savePerson') {
      const person = req.body;
      const rows = await readSheet(sheets, 'Persone!A:D');
      const row = [person.id, person.nome, person.whatsapp || '', person.email || ''];
      const idx = rows.findIndex(r => r[0] === person.id);
      if (idx >= 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `Persone!A${idx + 1}:D${idx + 1}`,
          valueInputOption: 'RAW',
          requestBody: { values: [row] },
        });
        return res.json({ result: 'updated' });
      } else {
        await sheets.spreadsheets.values.append({
          spreadsheetId: SHEET_ID,
          range: 'Persone!A:D',
          valueInputOption: 'RAW',
          requestBody: { values: [row] },
        });
        return res.json({ result: 'inserted' });
      }
    }

    // ── DELETE PERSON ──
    if (action === 'deletePerson') {
      const { id } = req.body;
      const rows = await readSheet(sheets, 'Persone!A:A');
      const idx = rows.findIndex(r => r[0] === id);
      if (idx >= 0) {
        const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
        const sheet = meta.data.sheets.find(s => s.properties.title === 'Persone');
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: {
            requests: [{
              deleteDimension: {
                range: { sheetId: sheet.properties.sheetId, dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 }
              }
            }]
          }
        });
      }
      return res.json({ result: 'deleted' });
    }

    // ── MARK SENT ──
    if (action === 'markSent') {
      const { id } = req.body;
      const rows = await readSheet(sheets, 'Task!A:A');
      const idx = rows.findIndex(r => r[0] === id);
      if (idx >= 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `Task!M${idx + 1}`,
          valueInputOption: 'RAW',
          requestBody: { values: [[new Date().toISOString()]] },
        });
      }
      return res.json({ result: 'ok' });
    }

    res.status(400).json({ error: 'Unknown action' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
