// פונקציית שרת (Vercel Serverless Function) - היחידה שמכירה את מפתח ה-API.
// הדפדפן של הפורש קורא ל-/api/generate בלבד; המפתח לעולם לא מגיע ללקוח.

const SYSTEM_PROMPT = `אתה כותב קורות חיים מקצועי, המתמחה בתרגום ניסיון צבאי מצה"ל (קצינים ונגדים) לשפה אזרחית עבור פורשים בתחילת דרכם האזרחית.
קיבלת מידע גולמי על מועמד. כתוב עבורו קורות חיים בעברית, מוכנים לשימוש, לפי הכללים הבאים:
- מבנה: פרטים אישיים בראש (שם, טלפון, אימייל), תקציר מקצועי קצר (2-3 משפטים), ניסיון תעסוקתי, כישורים, השכלה והכשרות, שפות (רק אם צוינו).
- תרגם דרגות, יחידות וראשי תיבות צבאיים לתפקידים ומונחים אזרחיים מקבילים ומובנים למגייס שלא מכיר את הצבא (למשל: מפקד פלוגה -> מנהל צוות של כ-100 עובדים; קב"ן -> יועץ/ת רווחה; מש"א -> מנהל/ת משאבי אנוש). הבן את משמעות התפקיד לפי ההקשר המלא שניתן, גם אם זה צירוף לא נפוץ - אל תישען על מילון קבוע.
- כתוב בגוף פעיל ובלשון הישגים (הובלתי, ניהלתי, בניתי) ולא בתיאור תפקיד סביל.
- אם ניתנו נתונים מדידים (כמויות, אחוזים, תקציבים) - שלב אותם. אל תמציא נתונים שלא ניתנו.
- שדות שלא סופקו - פשוט דלג עליהם, בלי להעיר על החיסרון.
- אורך: עמוד אחד בערך.
- החזר טקסט קורות חיים בלבד, מוכן לקריאה - עם כותרות סעיפים ברורות, ללא Markdown (בלי כוכביות או סולמיות), ללא הערות נוספות לפני או אחרי.`;

function buildUserPrompt(data) {
  const lines = [
    `שם מלא: ${data.fullName || ''}`,
    `טלפון: ${data.phone || ''}`,
    `אימייל: ${data.email || ''}`,
    `דרגה סופית: ${data.rank || ''}`,
    `שנות שירות: ${data.years || ''}`,
    `תפקיד אחרון/עיקרי: ${data.role || ''}`
  ];
  if (data.rolesHistory) lines.push(`תפקידים מרכזיים נוספים: ${data.rolesHistory}`);
  if (data.managedPeople) lines.push(`ניהול אנשים: ${data.managedPeople}`);
  if (data.managedBudget) lines.push(`ניהול תקציב/ציוד: ${data.managedBudget}`);
  if (data.training) lines.push(`הדרכה/הכשרה: ${data.training}`);
  if (data.achievements) lines.push(`הישג בולט: ${data.achievements}`);
  if (data.skills) lines.push(`כישורים טכניים/מקצועיים: ${data.skills}`);
  if (data.education) lines.push(`השכלה אזרחית: ${data.education}`);
  if (data.courses) lines.push(`קורסים והכשרות צבאיים: ${data.courses}`);
  if (data.languages) lines.push(`שפות: ${data.languages}`);
  return lines.join('\n');
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  const data = req.body || {};

  if (!data.fullName || !data.role) {
    res.status(400).json({ error: 'missing required fields' });
    return;
  }

  if (!process.env.ANTHROPIC_API_KEY) {
    res.status(500).json({ error: 'server not configured' });
    return;
  }

  try {
    const apiResponse = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-5',
        max_tokens: 1500,
        system: SYSTEM_PROMPT,
        messages: [
          { role: 'user', content: buildUserPrompt(data) }
        ]
      })
    });

    if (!apiResponse.ok) {
      const errText = await apiResponse.text();
      console.error('Anthropic API error:', apiResponse.status, errText);
      res.status(502).json({ error: 'upstream error' });
      return;
    }

    const result = await apiResponse.json();
    const text = (result.content || [])
      .filter(block => block.type === 'text')
      .map(block => block.text)
      .join('\n');

    if (!text) {
      res.status(502).json({ error: 'empty response' });
      return;
    }

    res.status(200).json({ resume: text });
  } catch (err) {
    console.error('Proxy error:', err);
    res.status(500).json({ error: 'generation failed' });
  }
};
