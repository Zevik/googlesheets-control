# צ'אטבוט לניהול Google Sheets

אפליקציית Next.js המאפשרת ניהול Google Sheets באמצעות בינה מלאכותית ושפה טבעית.

## תכונות

### ניהול נתונים
- קריאה והצגת נתונים מטווחים
- עדכון וכתיבת נתונים חדשים
- הוספת שורות חדשות
- מחיקת תוכן תאים
- חיפוש והחלפה
- מחיקת שורות לפי מילת מפתח
- יצירת נתונים אוטומטית באמצעות AI

### ניהול גיליונות
- יצירת גיליונות חדשים
- מחיקת גיליונות
- שינוי שם גיליונות
- שכפול גיליונות
- מיון נתונים

### עיצוב מתקדם
- צביעת תאים (14 צבעים שונים)
- עיצוב טקסט (מודגש, נטוי, קו תחתון)
- שינוי גודל גופן
- יישור טקסט

### תרשימים
- תרשימי עמודות
- תרשימי פאי
- תרשימי קו
- תרשימי פיזור
- תרשימי בר

## התקנה ופיתוח מקומי

### דרישות מוקדמות
- Node.js 18+
- חשבון Google Cloud עם Google Sheets API
- מפתח OpenAI API

### שלבי התקנה

1. **שכפל את הפרויקט:**
   ```bash
   git clone https://github.com/Zevik/googlesheets-control.git
   cd googlesheets-control
   ```

2. **התקן תלויות:**
   ```bash
   npm install
   ```

3. **הגדר משתני סביבה:**
   צור קובץ `.env.local` עם:
   ```
   OPENAI_API_KEY=your_openai_api_key
   GOOGLE_SERVICE_ACCOUNT_EMAIL=your-service-account@project.iam.gserviceaccount.com
   GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\nYOUR_KEY\n-----END PRIVATE KEY-----"
   ```

4. **הפעל את השרת:**
   ```bash
   npm run dev
   ```

5. **פתח בדפדפן:**
   http://localhost:3000

## פריסה ל-Vercel

### שלב 1: חיבור לגיטהאב
1. העלה את הפרויקט לגיטהאב
2. התחבר ל-[Vercel](https://vercel.com)
3. ייבא את הפרויקט מגיטהאב

### שלב 2: הגדרת משתני סביבה
בדשבורד של Vercel, עבור ל-Settings → Environment Variables והוסף:

- `OPENAI_API_KEY`: המפתח של OpenAI
- `GOOGLE_SERVICE_ACCOUNT_EMAIL`: כתובת המייל של Service Account
- `GOOGLE_PRIVATE_KEY`: המפתח הפרטי (עם \\n במקום שורות חדשות)

### שלב 3: פריסה
Vercel יפרוס אוטומטית את האפליקציה עם כל push לגיטהאב.

## הגדרת Google Sheets API

1. **צור פרויקט ב-Google Cloud Console**
2. **הפעל את Google Sheets API**
3. **צור Service Account:**
   - עבור ל-IAM & Admin → Service Accounts
   - צור חשבון שירות חדש
   - הורד את קובץ ה-JSON
4. **שתף את הגיליון:**
   - פתח את הגיליון ב-Google Sheets
   - לחץ על "שתף"
   - הוסף את כתובת המייל של Service Account עם הרשאות עורך

## שימוש

1. הדבק את מזהה הגיליון (Spreadsheet ID) בשדה העליון
2. כתוב בקשות בעברית, למשל:
   - "הצג לי את הנתונים מהגיליון"
   - "צבע את התאים A1:B2 בצהוב"
   - "צור גרף עמודות מהנתונים בעמודות A ו-B"
   - "מיין את הגיליון לפי עמודה 1"

## מבנה הפרויקט

```
├── components/          # קומפוננטים של React
│   └── ChatInterface.js # ממשק הצ'אט הראשי
├── lib/                 # ספריות עזר
│   └── sheets.js       # לוגיקת Google Sheets ו-OpenAI
├── pages/              # עמודים של Next.js
│   ├── api/           # API routes
│   │   └── chat.js    # endpoint לצ'אט
│   ├── _app.js        # הגדרות האפליקציה
│   └── index.js       # העמוד הראשי
├── styles/            # קבצי עיצוב
│   ├── Chat.module.css # עיצוב הצ'אט
│   └── globals.css    # עיצוב גלובלי
├── next.config.js     # הגדרות Next.js
├── package.json       # תלויות ופקודות
└── vercel.json        # הגדרות Vercel
```

## טכנולוגיות

- **Frontend**: Next.js, React, CSS Modules
- **Backend**: Next.js API Routes (Serverless)
- **AI**: OpenAI GPT-4
- **Google API**: Google Sheets API v4
- **Deployment**: Vercel

## רישיון

MIT License

## תמיכה

לבעיות או שאלות, פתח issue בגיטהאב או צור קשר עם המפתח. 