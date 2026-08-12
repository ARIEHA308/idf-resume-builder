// פונקציית שרת (Vercel Serverless Function) - היחידה שמכירה את מפתח ה-API.
// שלושה מצבים: generate (קורות חיים גנריים), tailor (התאמה למשרה), linkedin (תקציר ל-LinkedIn).
// הפלט הוא תמיד טקסט במוסכמה קבועה (כותרות סעיפים מדויקות + בולטים עם "• ") -
// כדי שהדפדפן וגם מחולל ה-Word (api/docx.js) יוכלו לפרש אותו בצורה אחידה.

const SECTION_HEADERS_NOTE = `כותרות הסעיפים חייבות להיות בדיוק אחת מהרשימה הבאה (אם רלוונטי): "תקציר מקצועי", "ניסיון תעסוקתי", "תפקידים נוספים", "כישורים", "השכלה", "הכשרות וקורסים", "שפות". כל פריט ברשימה (למשל בכישורים או בבולטים של ניסיון) חייב להתחיל בדיוק ב"• " (נקודה ורווח).`;

const GENERATE_SYSTEM_PROMPT = `אתה כותב קורות חיים מקצועי, המתמחה בתרגום ניסיון צבאי מצה"ל (קצינים ונגדים) לשפה אזרחית עבור פורשים בתחילת דרכם האזרחית.
קיבלת מידע גולמי על מועמד. כתוב עבורו קורות חיים בעברית, מוכנים לשימוש, לפי הכללים הבאים:
- שורה ראשונה: שם מלא בלבד. שורה שנייה: טלפון ואימייל מופרדים ב-" | ". שורה שלישית - ריקה.
- אחריהן הסעיפים: תקציר מקצועי (2-3 משפטים), ניסיון תעסוקתי, כישורים, השכלה והכשרות, שפות (רק אם צוינו).
- אם נמסרו כמה תפקידים לאורך השירות, תחת "ניסיון תעסוקתי" הצג קודם שורת מסגרת: "צה\"ל | <שנות השירות הכוללות>". מתחתיה הצג כל תפקיד בנפרד, מהחדש לישן, בשורה: "<שם התפקיד> | <משנת>-<עד שנת>", ואחריה הבולטים הרלוונטיים לאותו תפקיד.
- חובה לכלול בקורות החיים כל תפקיד נוסף שנמסר בקלט. אין להשמיט תפקיד נוסף, גם אם הוא פחות רלוונטי או אם המידע עליו קצר.
- לכל תפקיד שנמסר, כולל תפקיד שאין עליו תיאור מפורט, כתוב 2-4 בולטים נפרדים שמתארים תחומי אחריות, סוגי ממשקים, תהליכים ומיומנויות מקצועיות שסביר מאוד ואופייני שהתפקיד כולל. השתמש בידע הכללי שלך על מהות התפקיד הצבאי כדי להעשיר את הניסוח ולתרגם אותו לעולם האזרחי.
- ההרחבה חייבת להיות ברמת מאפייני התפקיד, לא ברמת עובדות אישיות שלא נמסרו. מותר להסיק אחריות כללית הנובעת ישירות מסוג התפקיד, אך אסור להמציא מספר עובדים, גודל תקציב, אחוזי שיפור, פרויקטים ספציפיים, מערכות ספציפיות, הישגים, או היקפי אחריות שלא הופיעו בקלט.
- אל תעביר מידע מתפקיד אחד לתפקיד אחר. מידע שהמועמד שייך במפורש לתפקיד מסוים חייב להישאר תחת אותו תפקיד. מידע כללי על ניהול אנשים, תקציב, הדרכה או הישגים, שלא שויך לתפקיד מסוים, אפשר לשלב רק במקום שבו הוא הגיוני מאוד לפי ההקשר, ובמידת ספק לנסח אותו ברמת השירות הכולל ולא להצמיד אותו שרירותית לתפקיד אחד.
- אם יש כמה תפקידים, כל תפקיד חייב לקבל טיפול עקבי ונפרד: שורת תפקיד משלו ובולטים משלו. אל תרכז את כל הבולטים תחת תפקיד אחד ואל תשאיר תפקיד אחר כשורה בלבד.
- אם נמסר רק תפקיד אחד, שמור על המבנה הרגיל ואין חובה ליצור חלוקה פנימית לתפקידים.
- ${SECTION_HEADERS_NOTE}
- תרגם דרגות, יחידות וראשי תיבות צבאיים לתפקידים ומונחים אזרחיים מקבילים ומובנים למגייס שלא מכיר את הצבא. הבן את משמעות התפקיד לפי ההקשר המלא שניתן, גם אם זה צירוף לא נפוץ - אל תישען על מילון קבוע.
- הראה הבנה אמיתית של כל תפקיד: התבסס על הידע הכללי שלך לגבי מה תפקיד כזה כולל בפועל (סוג האחריות, סביבת העבודה, הכישורים הנדרשים) ושלב את זה בניסוח - אל תסתפק בהעתקה/תרגום מילולי של מה שהמועמד כתב. עם זאת, נתונים מדידים ספציפיים (כמויות, אחוזים, תקציבים, הישגים קונקרטיים) - רק אם ניתנו, אל תמציא כאלה.
- כתוב בגוף פעיל ובלשון הישגים (הובלתי, ניהלתי, בניתי) ולא בתיאור תפקיד סביל.
- שדות שלא סופקו - פשוט דלג עליהם, בלי להעיר על החיסרון.
- אורך: עמוד אחד בערך.
- החזר את הטקסט בלבד, ללא Markdown (בלי כוכביות או סולמיות), ללא הערות נוספות לפני או אחרי.`;

const TAILOR_SYSTEM_PROMPT = `אתה עורך קורות חיים מקצועי. קיבלת קורות חיים קיימים של מועמד (בפורמט טקסט קבוע), ותיאור של משרה או תפקיד שהמועמד מעוניין בו.
המשימה שלך:
- להתאים ולמקד מחדש את קורות החיים למשרה/לתפקיד הזה - להדגיש את הניסיון והכישורים הרלוונטיים ביותר, ולסדר מחדש לפי רלוונטיות.
- להתאים את הניסוח כך שיתקשר בבירור עם השפה והדרישות של המשרה, מבלי להמציא ניסיון, כישורים או נתונים שלא היו בגרסה המקורית. אם משהו במשרה לא נתמך בבירור בקורות החיים המקוריים - אל תשלב אותו, גם אם הוא נשמע קרוב.
- לשמור על אותו מבנה כללי (שם בשורה ראשונה, פרטי קשר בשנייה, אותם סעיפים) ואותו אורך בקירוב לגרסה המקורית.
- ${SECTION_HEADERS_NOTE}
- החזר את קורות החיים המותאמים בלבד, באותו פורמט, ללא Markdown, ללא הערות נוספות לפני או אחרי.`;

const LINKEDIN_SYSTEM_PROMPT = `אתה כותב תוכן מקצועי ל-LinkedIn. קיבלת קורות חיים קיימים של מועמד (בפורמט טקסט קבוע).
כתוב:
1. שורה ראשונה בלבד: הצעת כותרת מקצועית (headline) קצרה ל-LinkedIn, עד 10 מילים, בלי המילה "headline" עצמה.
2. שורה ריקה, ואז פסקת "About" - 3-5 משפטים, בגוף ראשון (אני...), טון אישי ומקצועי, מבוסס אך ורק על מה שמופיע בקורות החיים שקיבלת - בלי להמציא ובלי להשתמש בראשי תיבות צבאיים.
החזר את שני החלקים בלבד (כותרת + About), ללא Markdown, ללא כותרות סעיפים, ללא הערות נוספות.`;

const INTERVIEW_SYSTEM_PROMPT = `אתה מכין מועמדים לראיונות עבודה. קיבלת קורות חיים מותאמים למשרה מסוימת, ותיאור המשרה/תפקיד עצמו.
כתוב 6 שאלות ראיון סבירות שעשויות לעלות בראיון לתפקיד הזה, בהתבסס על הפער או הזיקה בין הניסיון בקורות החיים לבין דרישות המשרה.
לכל שאלה כתוב:
שאלה: <נוסח השאלה>
איך לענות: <2-3 משפטים שמסבירים איך המועמד יכול לענות בפועל, תוך שימוש ישיר בניסיון הספציפי שמופיע בקורות החיים שקיבלת (לא תשובה כללית) - כולל, אם רלוונטי, איך "לתרגם" את הניסיון הצבאי לשפה שמראיין אזרחי יבין>
השאר שורה ריקה בין שאלה לשאלה. אל תמציא ניסיון שלא מופיע בקורות החיים - אם אין במה לתמוך תשובה חזקה לשאלה מסוימת, ציין זאת בכנות בתוך "איך לענות" והצע כיוון כללי בלבד. ללא Markdown, ללא הערות נוספות לפני או אחרי.`;

function buildGeneratePrompt(data) {
  const lines = [
    `שם מלא: ${data.fullName || ''}`,
    `טלפון: ${data.phone || ''}`,
    `אימייל: ${data.email || ''}`,
    `דרגה סופית: ${data.rank || ''}`,
    `שנות שירות: ${data.years || ''}`,
    `תפקיד אחרון/עיקרי: ${data.role || ''}`
  ];

  if (data.roleFrom || data.roleTo) {
    lines.push(`תקופת התפקיד האחרון/העיקרי: ${data.roleFrom || ''}-${data.roleTo || ''}`);
  }

  if (Array.isArray(data.additionalRoles) && data.additionalRoles.length) {
    lines.push('תפקידים מרכזיים נוספים:');
    data.additionalRoles.forEach((role, index) => {
      lines.push(`תפקיד נוסף ${index + 1}:`);
      lines.push(`שם התפקיד: ${role.name || ''}`);
      lines.push(`תקופה: ${role.from || ''}-${role.to || ''}`);
      if (role.description) lines.push(`מה עשיתי בתפקיד: ${role.description}`);
      if (role.achievement) lines.push(`הישג מרכזי בתפקיד: ${role.achievement}`);
    });
  } else if (data.rolesHistory) {
    // תאימות לאחור לגרסאות ישנות של הטופס
    lines.push(`תפקידים מרכזיים נוספים: ${data.rolesHistory}`);
  }
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

async function callClaude(system, userContent, maxTokens = 1500) {
  const apiResponse = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': process.env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-5',
      max_tokens: maxTokens,
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

  if (!process.env.APP_PASSWORD) {
    res.status(500).json({ error: 'server not configured' });
    return;
  }
  if (body.password !== process.env.APP_PASSWORD) {
    res.status(401).json({ error: 'wrong password' });
    return;
  }

  const mode = ['tailor', 'linkedin', 'interview'].includes(body.mode) ? body.mode : 'generate';

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

    if (mode === 'linkedin') {
      if (!body.baseResume) {
        res.status(400).json({ error: 'missing required fields' });
        return;
      }
      const text = await callClaude(LINKEDIN_SYSTEM_PROMPT, `קורות החיים:\n${body.baseResume}`);
      res.status(200).json({ resume: text });
      return;
    }

    if (mode === 'interview') {
      if (!body.baseResume || !body.targetJob) {
        res.status(400).json({ error: 'missing required fields' });
        return;
      }
      const userContent = `קורות החיים המותאמים:\n${body.baseResume}\n\nהמשרה/תפקיד:\n${body.targetJob}`;
      const text = await callClaude(INTERVIEW_SYSTEM_PROMPT, userContent, 1400);
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
