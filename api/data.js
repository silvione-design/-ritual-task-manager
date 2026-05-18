const { Redis } = require('@upstash/redis');

const redis = new Redis({
  url: process.env.STORAGE_KV_REST_API_URL || process.env.KV_REST_API_URL,
  token: process.env.STORAGE_KV_REST_API_TOKEN || process.env.KV_REST_API_TOKEN,
});

const DEFAULT_TASKS = [
  { id:'t1', nome:'Decalcificazione lavabicchieri', categoria:'Manutenzione', freqType:'giorno-settimana-n', freqValue:2, freqGiornoSett:2, freqGiornoMese:0, oraInvio:'09:00', primaryId:'p2', backupId:'p1', note:'Usare decalcificante apposito. Ciclo completo senza bicchieri.', attivo:true, ultimoInvio:null },
  { id:'t2', nome:'Controllo bottiglie acqua 1,5L', categoria:'Scorte', freqType:'giorno-settimana', freqValue:0, freqGiornoSett:1, freqGiornoMese:0, oraInvio:'09:00', primaryId:'p1', backupId:'p2', note:'Verificare scorte. Riordinare se sotto 24 bottiglie.', attivo:true, ultimoInvio:null },
  { id:'t3', nome:'Pulizia mantovana cappa', categoria:'Pulizie', freqType:'mesi', freqValue:6, freqGiornoSett:0, freqGiornoMese:0, oraInvio:'08:00', primaryId:'p2', backupId:'p3', note:'Smontare filtri e sgrassare completamente.', attivo:true, ultimoInvio:null },
  { id:'t4', nome:'Pulizia filtri cappa aspirante', categoria:'Pulizie', freqType:'mesi', freqValue:1, freqGiornoSett:0, freqGiornoMese:0, oraInvio:'08:00', primaryId:'p1', backupId:'p2', note:'Smontare e lavare in lavastoviglie. Asciugare prima di rimontare.', attivo:true, ultimoInvio:null },
  { id:'t5', nome:'Controllo temperatura frigoriferi', categoria:'HACCP', freqType:'giorni', freqValue:1, freqGiornoSett:0, freqGiornoMese:0, oraInvio:'07:30', primaryId:'p2', backupId:'p4', note:'Celle: 0-4°C. Freezer: -18°C. Registrare su registro HACCP.', attivo:true, ultimoInvio:null },
  { id:'t6', nome:'Pulizia fondo frigoriferi', categoria:'Pulizie', freqType:'mesi', freqValue:1, freqGiornoSett:0, freqGiornoMese:0, oraInvio:'08:00', primaryId:'p3', backupId:'p1', note:'Svuotare, pulire con igienizzante, asciugare.', attivo:true, ultimoInvio:null },
  { id:'t7', nome:'Controllo scadenze prodotti magazzino', categoria:'HACCP', freqType:'giorno-settimana', freqValue:0, freqGiornoSett:1, freqGiornoMese:0, oraInvio:'08:00', primaryId:'p2', backupId:'p1', note:'Verificare date scadenza. Rimuovere prodotti scaduti.', attivo:true, ultimoInvio:null },
  { id:'t8', nome:'Pulizia pietre laviche', categoria:'Manutenzione', freqType:'giorno-settimana', freqValue:0, freqGiornoSett:5, freqGiornoMese:0, oraInvio:'10:00', primaryId:'p1', backupId:'p4', note:'Spazzolare a freddo. Riposizionare nel trespolo.', attivo:true, ultimoInvio:null },
  { id:'t9', nome:'Verifica scorte prodotti pulizia', categoria:'Scorte', freqType:'giorno-settimana-n', freqValue:2, freqGiornoSett:1, freqGiornoMese:0, oraInvio:'09:00', primaryId:'p3', backupId:'p2', note:'Detergenti, igienizzanti, guanti, sacchi. Avvisare se sotto soglia.', attivo:true, ultimoInvio:null },
  { id:'t10', nome:'Pulizia vetrine e specchi', categoria:'Pulizie', freqType:'giorni', freqValue:3, freqGiornoSett:0, freqGiornoMese:0, oraInvio:'09:00', primaryId:'p1', backupId:'p3', note:'Spray per vetri. Controllare anche specchi bagni.', attivo:true, ultimoInvio:null },
  { id:'t11', nome:'Invio ore al consulente paghe', categoria:'Amministrazione', freqType:'giorno-fisso-mese', freqValue:0, freqGiornoSett:0, freqGiornoMese:4, oraInvio:'08:00', primaryId:'p2', backupId:'', note:'Inviare prospetto ore entro il 4 del mese.', attivo:true, ultimoInvio:null },
  { id:'t12', nome:'Sanificazione superfici cucina', categoria:'HACCP', freqType:'giorni', freqValue:1, freqGiornoSett:0, freqGiornoMese:0, oraInvio:'07:30', primaryId:'p4', backupId:'p1', note:'Piani lavoro, taglieri, attrezzature. Igienizzante CE.', attivo:true, ultimoInvio:null },
  { id:'t13', nome:'Cambio olio friggitrice', categoria:'Manutenzione', freqType:'giorno-settimana', freqValue:0, freqGiornoSett:5, freqGiornoMese:0, oraInvio:'16:00', primaryId:'p1', backupId:'p4', note:'Svuotare, pulire vasca, olio nuovo. Smaltire il vecchio.', attivo:true, ultimoInvio:null },
  { id:'t14', nome:'Stampa stock vini per ordine', categoria:'Scorte', freqType:'giorno-settimana', freqValue:0, freqGiornoSett:1, freqGiornoMese:0, oraInvio:'09:00', primaryId:'p2', backupId:'p1', note:'Stampare giacenze vini. Consegnare a Silvio per ordine.', attivo:true, ultimoInvio:null },
];

const DEFAULT_PEOPLE = [
  { id:'p1', nome:'Marco', whatsapp:'+39 333 0000001', email:'marco@ritual.it' },
  { id:'p2', nome:'María', whatsapp:'+39 333 0000002', email:'maria@ritual.it' },
  { id:'p3', nome:'Rosario', whatsapp:'+39 333 0000003', email:'rosario@ritual.it' },
  { id:'p4', nome:'Gaetano', whatsapp:'+39 333 0000004', email:'gaetano@ritual.it' },
];

const DEFAULT_CONFIG = {
  password_silvio: 'ritual2024',
  password_catia: 'ritual2024',
  nome_ristorante: 'Ritual Brasileiro',
};

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    // Load or initialize data
    let tasks  = await redis.get('tasks');
    let people = await redis.get('people');
    let config = await redis.get('config');

    if (!tasks)  { tasks  = DEFAULT_TASKS;   await redis.set('tasks',  tasks); }
    if (!people) { people = DEFAULT_PEOPLE;  await redis.set('people', people); }
    if (!config) { config = DEFAULT_CONFIG;  await redis.set('config', config); }

    res.status(200).json({ tasks, people, config });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
