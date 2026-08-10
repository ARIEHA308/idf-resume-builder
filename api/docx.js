// בונה קובץ Word (.docx) אמיתי מהטקסט המוצג בדף (כולל עריכות ידניות) -
// מפרש את אותה מוסכמת פורמט (כותרות סעיפים קבועות + בולטים "• ") שמ-generate.js.
//
// הערה טכנית: לא משתמשים בסגנונות המובנים של הספרייה (Title/Heading) - יש להם
// באג ידוע (docx.js #3419) שבו כיווניות RTL לא "מושכת" איתה יישור לימין כשמשתמשים
// בסגנון מוגדר-מראש. פותרים את זה עם עיצוב ישיר (bold/size) על כל פסקה בנפרד.

const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');

const SECTION_HEADERS = ['תקציר מקצועי', 'ניסיון תעסוקתי', 'תפקידים נוספים', 'כישורים', 'השכלה', 'הכשרות וקורסים', 'שפות'];

function parseResumeText(raw) {
  const nonEmpty = raw.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const name = nonEmpty[0] || '';
  const contact = nonEmpty[1] || '';
  const sections = [];
  let current = null;

  for (let i = 2; i < nonEmpty.length; i++) {
    const line = nonEmpty[i];
    if (SECTION_HEADERS.includes(line)) {
      current = { header: line, bullets: [], paragraphs: [] };
      sections.push(current);
    } else if (current) {
      if (line.startsWith('•')) {
        current.bullets.push(line.replace(/^•\s*/, ''));
      } else {
        current.paragraphs.push(line);
      }
    }
  }
  return { name, contact, sections };
}

function rtlParagraph(text, runOpts = {}, paraOpts = {}) {
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    bidirectional: true,
    ...paraOpts,
    children: [new TextRun({ text, rightToLeft: true, ...runOpts })]
  });
}

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
    res.status(401).json({ error: 'wrong password' });
    return;
  }
  if (!body.text) {
    res.status(400).json({ error: 'missing text' });
    return;
  }

  try {
    const parsed = parseResumeText(body.text);
    const children = [];

    children.push(rtlParagraph(parsed.name, { bold: true, size: 36 }, { spacing: { after: 80 } }));
    children.push(rtlParagraph(parsed.contact, { size: 20, color: '5B6270' }, { spacing: { after: 280 } }));

    parsed.sections.forEach(sec => {
      children.push(rtlParagraph(sec.header, { bold: true, size: 24, color: '57623F' }, { spacing: { before: 260, after: 120 } }));
      sec.paragraphs.forEach(p => {
        children.push(rtlParagraph(p, {}, { spacing: { after: 120 } }));
      });
      sec.bullets.forEach(b => {
        children.push(rtlParagraph(b, {}, { bullet: { level: 0 } }));
      });
    });

    const doc = new Document({ sections: [{ children }] });
    const buffer = await Packer.toBuffer(doc);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="resume.docx"');
    res.status(200).send(buffer);
  } catch (err) {
    console.error('docx error:', err);
    res.status(500).json({ error: 'docx generation failed' });
  }
};
