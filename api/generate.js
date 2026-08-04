// פונקציית שרת (Vercel Serverless Function) - היחידה שמכירה את מפתח ה-API.
// הדפדפן של הפורש קורא ל-/api/generate בלבד; המפתח לעולם לא מגיע ללקוח.
// שני מצבים: generate (קורות חיים גנריים) ו-tailor (התאמה למשרה/תפקיד).

const GENERATE_SYSTEM_PROMPT = `אתה כותב קורות חיים מקצועי, המתמחה בתרגום ניסיון צבאי מצה"ל (קצינים ונגדים) לשפה אזרחית עבור פורשים בתחילת דרכם האזרחית.
קיבלת מידע גולמי על מועמד. כתוב עבורו קורות חיים בעברית, מוכנים לשימוש, לפי הכללים הבאים:
- מבנה: פרטים אישיים בראש (שם, טלפון, אימייל), תקציר מקצועי קצר (2-3 משפטים), ניסיון תעסוקתי, כישורים, השכלה והכשרות, שפות (רק אם צוינו).
- תרגם דרגות, יחידות וראשי תיבות צבאיים לתפקידים ומונחים אזרחיים מקבילים ומובנים למגייס שלא מכיר את הצבא. הבן את משמעות התפקיד לפי ההקשר המלא שניתן, גם אם זה צירוף לא נפוץ - אל תישען על מילון קבוע.
- הראה הבנה אמיתית של התפקיד: התבסס על הידע הכללי שלך לגבי מה תפקיד כזה כולל בפועל (סוג האחריות, סביבת העבודה, הכישורים הנדרשים) ושלב את זה בניסוח - אל תסתפק בהעתקה/תרגום מילולי של מה שהמועמד כתב. עם זאת, נתונים מדידים ספציפיים (כמויות, אחוזים, תקציבים, הישגים קונקרטיים) - רק אם ניתנו, אל תמציא כאלה.
- כתוב בגוף פעיל ובלשון הישגים (הובלתי, ניהלתי, בניתי) ולא בתיאור תפקיד סביל.
- שדות שלא סופקו - פשוט דלג עליהם, בלי להעיר על החיסרון.
- אורך: עמוד אחד בערך.
- החזר טקסט קורות חיים בלבד, מוכן לקריאה - עם כותרות סעיפים ברורות, ללא Markdown (בלי כוכביות או סולמיות), ללא הערות נוספות לפני או אחרי.`;

const TAILOR_SYSTEM_PROMPT = `אתה עורך קורות חיים מקצועי. קיבלת קורות חיים קיימים של מועמד, ותיאור של משרה או תפקיד שהמועמד מעוניין בו (זה יכול להיות מודעת דרושים מלאה, או תיאור קצר של סוג התפקיד הרצוי).
המשימה שלך:
- להתאים ולמקד מחדש את קורות החיים למשרה/לתפקיד הזה - להדגיש את הניסיון והכישורים הרלוונטיים ביותר, ולסדר מחדש לפי רלוונטיות.
- להתאים את הניסוח כך שיתקשר בבירור עם השפה והדרישות של המשרה, מבלי להמציא ניסיון, כישורים או נתונים שלא היו בגרסה המקורית.
- לשמור על אותו מבנה כללי ואותו אורך בקירוב לגרסה המקורית.
- החזר את קורות החיים המותאמים בלבד, באותו פורמט (כותרות סעיפים ברורות, ללא Markdown), ללא הערות נוספות לפני או אחרי.`;

function buildGeneratePrompt(data) {
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

async function callClaude(system, userContent) {
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
      system,
      messages: [{ role: 'user', content: userContent }]
    })
  });

  if (!apiResponse.ok) {
    const errText = await apiResponse.text();
    console.error('Anthropic API error:', apiResponse.status, errText);
    throw new Error('upstream error');
  }

  const result = await apiResponse.json();
  const text = (result.content || [])
    .filter(block => block.type === 'text')
    .map(block => block.text)
    .join('\n');

  if (!text) throw new Error('empty response');
  return text;
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  if (!process.env.ANTHROPIC_API_KEY) {
    res.status(500).json({ error: 'server not configured' });
    return;
  }

  const body = req.body || {};
  const mode = body.mode === 'tailor' ? 'tailor' : 'generate';

  try {
    if (mode === 'tailor') {
      if (!body.baseResume || !body.targetJob) {
        res.status(400).json({ error: 'missing required fields' });
        return;
      }
      const userContent = `קורות החיים הגנריים:\n${body.baseResume}\n\nהמשרה/תפקיד המבוקש:\n${body.targetJob}`;
      const text = await callClaude(TAILOR_SYSTEM_PROMPT, userContent);
      res.status(200).json({ resume: text });
      return;
    }

    // מצב generate
    if (!body.fullName || !body.role) {
      res.status(400).json({ error: 'missing required fields' });
      return;
    }
    const text = await callClaude(GENERATE_SYSTEM_PROMPT, buildGeneratePrompt(body));
    res.status(200).json({ resume: text });
  } catch (err) {
    console.error('Proxy error:', err);
    res.status(502).json({ error: 'generation failed' });
  }
};
