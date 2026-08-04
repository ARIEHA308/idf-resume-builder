# מחולל קורות חיים לפורשי צה"ל

## מבנה הפרויקט
- `index.html` - הטופס והדף שרואה הפורש
- `api/generate.js` - פונקציית שרת (רצה ב-Vercel) שמכילה את הקריאה לקלוד ואת מפתח ה-API
- `package.json` - קובץ טכני קטן, לא צריך לגעת בו

## איך להעלות (בלי שורת פקודה, הכל דרך הדפדפן)

### שלב 1 - יצירת ריפו ב-GitHub
1. היכנס ל-github.com → **New repository**
2. תן שם, למשל `idf-resume-builder` → **Create repository**
3. בעמוד הריפו: **Add file → Upload files**
4. גרור לתוכו את שלושת הקבצים (`index.html`, `package.json`, ואת התיקייה `api` עם `generate.js` שבתוכה)
5. **Commit changes**

### שלב 2 - חיבור ל-Vercel
1. היכנס ל-vercel.com (אפשר עם אותו חשבון GitHub)
2. **Add New → Project**
3. בחר את הריפו `idf-resume-builder` → **Import**
4. השאר את ההגדרות כברירת מחדל (Vercel מזהה לבד את `index.html` ואת `api/generate.js`)
5. **לפני שלוחצים Deploy** - פתח **Environment Variables**:
   - Name: `ANTHROPIC_API_KEY`
   - Value: המפתח שיצרת ב-console.anthropic.com (מתחיל ב-`sk-ant-`)
6. **Deploy**

בסיום תקבל כתובת כמו `idf-resume-builder.vercel.app` - זה הלינק שאפשר לשלוח לכל הפורשים.

## עדכון בעתיד
כל שינוי בקבצים ב-GitHub (למשל עריכת `index.html`) מפעיל אוטומטית פריסה מחדש ב-Vercel - בדיוק כמו ב-tedarim-app.
