// בדיקת סיסמה בלבד - לא קוראת לקלוד, לא עולה כלום להפעיל.
// משמשת את מסך הכניסה בלבד. ההגנה האמיתית על העלות היא בתוך generate.js עצמו.

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  if (!process.env.APP_PASSWORD) {
    res.status(500).json({ error: 'server not configured' });
    return;
  }

  const body = req.body || {};
  if (body.password !== process.env.APP_PASSWORD) {
    res.status(401).json({ ok: false });
    return;
  }

  res.status(200).json({ ok: true });
};
