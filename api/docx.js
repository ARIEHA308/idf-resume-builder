// בונה קובץ Word (.docx) ידנית, ברמת ה-XML הגולמי - בלי להשתמש בספריית docx
// ובלי docxtemplater. הסיבה: שתי הספריות הראו התנהגות RTL לא עקבית/לא צפויה
// אחרי חקירה מעמיקה (כולל פירוק קבצים אמיתיים ובדיקת ה-XML הפנימי). כתיבה ידנית
// נותנת שליטה מלאה - בלי שום שכבת "פרשנות" של ספרייה שיכולה להוסיף/להשמיט משהו.
//
// RTL מוגדר בשלוש רמות בו-זמנית (הכי מקיף שאפשר לפי התקן):
// 1. docDefaults (ברירת מחדל לכל המסמך) ב-styles.xml
// 2. כל פסקה בנפרד (w:bidi + w:jc="right") + כל run (w:rtl + w:lang + w:rFonts)
// 3. ה-section עצמו (w:bidi בתוך w:sectPr) - משפיע על כיוון העימוד הכללי

const JSZip = require('jszip');

const SECTION_HEADERS = [
  'תקציר מקצועי',
  'ניסיון תעסוקתי',
  'תפקידים נוספים',
  'כישורים',
  'השכלה',
  'הכשרות וקורסים',
  'שפות'
];

function parseResumeText(raw) {
  const nonEmpty = raw
    .split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 0);

  const name = nonEmpty[0] || '';
  const contact = nonEmpty[1] || '';
  const sections = [];
  let current = null;

  for (let i = 2; i < nonEmpty.length; i++) {
    const line = nonEmpty[i];

    if (SECTION_HEADERS.includes(line)) {
      current = {
        header: line,
        bullets: [],
        paragraphs: []
      };

      sections.push(current);
    } else if (current) {
      if (line.startsWith('•')) {
        current.bullets.push(
          line.replace(/^•\s*/, '')
        );
      } else {
        current.paragraphs.push(line);
      }
    }
  }

  return {
    name,
    contact,
    sections
  };
}

function escapeXml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// בונה פסקה בודדת - כל התכונות הרלוונטיות ל-RTL מוגדרות במפורש, בלי להסתמך
// על שום ברירת מחדל (גם אם כבר הגדרנו אותה למעלה - עדיפה כפילות על חוסר).
function buildParagraph(
  text,
  {
    bold = false,
    size = null,
    color = null,
    bullet = false
  } = {}
) {
  const rPr = [
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/>',
    bold ? '<w:b/><w:bCs/>' : '',
    size
      ? `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`
      : '',
    color
      ? `<w:color w:val="${color}"/>`
      : '',
    '<w:rtl/>',
    '<w:lang w:val="he-IL" w:eastAsia="he-IL" w:bidi="he-IL"/>'
  ].join('');

  const numPr = bullet
    ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'
    : '';

  return `<w:p><w:pPr>${numPr}<w:bidi/><w:jc w:val="start"/><w:rPr>${rPr}</w:rPr></w:pPr><w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function buildDocumentXml(parsed) {
  const paragraphs = [];

  paragraphs.push(
    buildParagraph(
      parsed.name,
      {
        bold: true,
        size: 36
      }
    )
  );

  paragraphs.push(
    buildParagraph(
      parsed.contact,
      {
        size: 20,
        color: '5B6270'
      }
    )
  );

  parsed.sections.forEach(sec => {
    paragraphs.push(
      buildParagraph(
        sec.header,
        {
          bold: true,
          size: 24,
          color: '57623F'
        }
      )
    );

    sec.paragraphs.forEach(p => {
      paragraphs.push(
        buildParagraph(p)
      );
    });

    sec.bullets.forEach(b => {
      paragraphs.push(
        buildParagraph(`•  ${b}`)
      );
    });
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
${paragraphs.join('\n')}
<w:sectPr>
<w:bidi/>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar
  w:top="1417"
  w:right="1417"
  w:bottom="1417"
  w:left="1417"
  w:header="708"
  w:footer="708"
  w:gutter="0"
/>
</w:sectPr>
</w:body>
</w:document>`;
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>

<w:rPrDefault>
<w:rPr>
<w:rFonts
  w:ascii="Arial"
  w:hAnsi="Arial"
  w:cs="Arial"
  w:eastAsia="Arial"
/>
<w:rtl/>
<w:lang
  w:val="he-IL"
  w:eastAsia="he-IL"
  w:bidi="he-IL"
/>
</w:rPr>
</w:rPrDefault>

<w:pPrDefault>
<w:pPr>
<w:bidi/>
<w:jc w:val="start"/>
</w:pPr>
</w:pPrDefault>

</w:docDefaults>

<w:style
  w:type="paragraph"
  w:default="1"
  w:styleId="Normal"
>
<w:name w:val="Normal"/>
<w:qFormat/>
</w:style>

</w:styles>`;

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">

<Default
  Extension="rels"
  ContentType="application/vnd.openxmlformats-package.relationships+xml"
/>

<Default
  Extension="xml"
  ContentType="application/xml"
/>

<Override
  PartName="/word/document.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
/>

<Override
  PartName="/word/styles.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
/>

</Types>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

<Relationship
  Id="rId1"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  Target="word/document.xml"
/>

</Relationships>`;

const DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

<Relationship
  Id="rId1"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  Target="styles.xml"
/>

</Relationships>`;

async function buildDocx(parsed) {
  const zip = new JSZip();

  zip.file(
    '[Content_Types].xml',
    CONTENT_TYPES_XML
  );

  zip
    .folder('_rels')
    .file(
      '.rels',
      RELS_XML
    );

  const wordFolder = zip.folder('word');

  wordFolder.file(
    'document.xml',
    buildDocumentXml(parsed)
  );

  wordFolder.file(
    'styles.xml',
    STYLES_XML
  );

  wordFolder
    .folder('_rels')
    .file(
      'document.xml.rels',
      DOCUMENT_RELS_XML
    );

  return zip.generateAsync({
    type: 'nodebuffer'
  });
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({
      error: 'Method not allowed'
    });
    return;
  }

  if (!process.env.APP_PASSWORD) {
    res.status(500).json({
      error: 'server not configured'
    });
    return;
  }

  const body = req.body || {};

  if (body.password !== process.env.APP_PASSWORD) {
    res.status(401).json({
      error: 'wrong password'
    });
    return;
  }

  if (!body.text) {
    res.status(400).json({
      error: 'missing text'
    });
    return;
  }

  try {
    const parsed = parseResumeText(
      body.text
    );

    const buffer = await buildDocx(
      parsed
    );

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );

    res.setHeader(
      'Content-Disposition',
      'attachment; filename="resume.docx"'
    );

    res.status(200).send(buffer);

  } catch (err) {
    console.error(
      'docx error:',
      err
    );

    res.status(500).json({
      error: 'docx generation failed'
    });
  }
};
