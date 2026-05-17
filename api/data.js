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

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const auth = await getAuth();
    const sheets = google.sheets({ version: 'v4', auth });

    // Read all sheets
    const [taskRes, peopleRes, configRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Task!A:M' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Persone!A:D' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Config!A:B' }),
    ]);

    const taskRows    = taskRes.data.values || [];
    const peopleRows  = peopleRes.data.values || [];
    const configRows  = configRes.data.values || [];

    // Parse tasks (skip header row)
    const tasks = taskRows.slice(1).filter(r => r[0]).map(r => ({
      id:             r[0]  || '',
      nome:           r[1]  || '',
      categoria:      r[2]  || '',
      freqType:       r[3]  || '',
      freqValue:      Number(r[4]) || 0,
      freqGiornoMese: Number(r[5]) || 0,
      freqGiornoSett: Number(r[6]) || 0,
      oraInvio:       r[7]  || '09:00',
      primaryId:      r[8]  || '',
      backupId:       r[9]  || '',
      note:           r[10] || '',
      attivo:         r[11] === 'TRUE' || r[11] === true || r[11] === '1',
      ultimoInvio:    r[12] || null,
    }));

    // Parse people (skip header row)
    const people = peopleRows.slice(1).filter(r => r[0]).map(r => ({
      id:       r[0] || '',
      nome:     r[1] || '',
      whatsapp: r[2] || '',
      email:    r[3] || '',
    }));

    // Parse config (skip header row)
    const config = {};
    configRows.slice(1).forEach(r => { if (r[0]) config[r[0]] = r[1]; });

    res.status(200).json({ tasks, people, config });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
