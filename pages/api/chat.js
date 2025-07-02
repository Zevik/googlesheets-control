const { translateToApiCall, executeSheetOperation } = require('../../lib/sheets');

export default async function handler(req, res) {
  // Set CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  // Handle preflight request
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { message, spreadsheet_id } = req.body;

    if (!message || !spreadsheet_id) {
      return res.status(400).json({ 
        reply: 'שגיאה: הודעה או מזהה גיליון חסרים.' 
      });
    }

    console.log(`\n[לוג] התקבלה בקשה מהמשתמש: '${message}'`);

    // תרגם לפקודת API
    const apiCall = await translateToApiCall(message, spreadsheet_id);
    console.log(`[לוג] פקודה שנוצרה:`, apiCall);
    
    // בצע את הפעולה
    const result = await executeSheetOperation(apiCall, spreadsheet_id);
    console.log(`[לוג] תוצאת הפעולה: ${result}`);
    
    res.json({ reply: result });
  } catch (error) {
    console.error('Chat API Error:', error);
    res.json({ 
      reply: `אופס, משהו השתבש. השגיאה: ${error.message}` 
    });
  }
} 