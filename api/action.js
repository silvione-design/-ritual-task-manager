const { Redis } = require('@upstash/redis');

const redis = new Redis({
  url: process.env.STORAGE_KV_REST_API_URL || process.env.KV_REST_API_URL,
  token: process.env.STORAGE_KV_REST_API_TOKEN || process.env.KV_REST_API_TOKEN,
});

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const action = req.query.action;

  try {
    // LOGIN
    if (action === 'login') {
      const { user, pwd } = req.body;
      const config = await redis.get('config') || {};
      if (user === 'silvio' && pwd === config.password_silvio) return res.json({ ok: true });
      if (user === 'catia'  && pwd === config.password_catia)  return res.json({ ok: true });
      return res.json({ ok: false });
    }

    // SAVE TASK
    if (action === 'saveTask') {
      const task = req.body;
      let tasks = await redis.get('tasks') || [];
      const idx = tasks.findIndex(t => t.id === task.id);
      if (idx >= 0) tasks[idx] = task;
      else tasks.push(task);
      await redis.set('tasks', tasks);
      return res.json({ result: 'ok' });
    }

    // DELETE TASK
    if (action === 'deleteTask') {
      const { id } = req.body;
      let tasks = await redis.get('tasks') || [];
      tasks = tasks.filter(t => t.id !== id);
      await redis.set('tasks', tasks);
      return res.json({ result: 'ok' });
    }

    // SAVE PERSON
    if (action === 'savePerson') {
      const person = req.body;
      let people = await redis.get('people') || [];
      const idx = people.findIndex(p => p.id === person.id);
      if (idx >= 0) people[idx] = person;
      else people.push(person);
      await redis.set('people', people);
      return res.json({ result: 'ok' });
    }

    // DELETE PERSON
    if (action === 'deletePerson') {
      const { id } = req.body;
      let people = await redis.get('people') || [];
      people = people.filter(p => p.id !== id);
      await redis.set('people', people);
      return res.json({ result: 'ok' });
    }

    // MARK SENT
    if (action === 'markSent') {
      const { id } = req.body;
      let tasks = await redis.get('tasks') || [];
      tasks = tasks.map(t => t.id === id ? { ...t, ultimoInvio: new Date().toISOString() } : t);
      await redis.set('tasks', tasks);
      return res.json({ result: 'ok' });
    }

    res.status(400).json({ error: 'Unknown action: ' + action });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
