const { google } = require('googleapis');
const OpenAI = require('openai');

// Initialize OpenAI
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// Initialize Google Sheets API
async function getGoogleSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      type: 'service_account',
      project_id: process.env.GOOGLE_PROJECT_ID,
      private_key_id: process.env.GOOGLE_PRIVATE_KEY_ID,
      private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      client_id: process.env.GOOGLE_CLIENT_ID,
      auth_uri: 'https://accounts.google.com/o/oauth2/auth',
      token_uri: 'https://oauth2.googleapis.com/token',
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  return google.sheets({ version: 'v4', auth });
}

// System prompt for AI
const SYSTEM_PROMPT = `
אתה עוזר AI שמתרגם בקשות בשפה טבעית לקריאות API של Google Sheets.
תמיד החזר JSON תקין עם השדות הבאים:
- function_name: שם הפונקציה לביצוע
- parameters: פרמטרים לפונקציה

הפונקציות הזמינות:
1. get_data(range) - קריאת נתונים מטווח
2. update_data(range, values) - עדכון נתונים
3. append_data(range, values) - הוספת נתונים
4. clear_data(range) - מחיקת נתונים
5. create_sheet(sheet_name) - יצירת גיליון חדש
6. delete_sheet(sheet_name) - מחיקת גיליון
7. rename_sheet(old_name, new_name) - שינוי שם גיליון

דוגמאות:
- "הצג לי את הנתונים מהעמודות A ו-B" → {"function_name": "get_data", "parameters": {"range": "A:B"}}
- "עדכן את התא A1 לערך 100" → {"function_name": "update_data", "parameters": {"range": "A1", "values": [["100"]]}}
- "צור גיליון חדש בשם 'נתונים'" → {"function_name": "create_sheet", "parameters": {"sheet_name": "נתונים"}}
`;

// Translate user message to API call
async function translateToApiCall(userMessage, spreadsheetId) {
  try {
    const response = await openai.chat.completions.create({
      model: 'gpt-4',
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content: userMessage }
      ],
      temperature: 0.1
    });

    const content = response.choices[0].message.content.trim();
    console.log('AI Response:', content);
    
    return JSON.parse(content);
  } catch (error) {
    console.error('Translation error:', error);
    throw new Error('לא הצלחתי להבין את הבקשה');
  }
}

// Execute sheet operations
async function executeSheetOperation(apiCall, spreadsheetId) {
  const sheets = await getGoogleSheetsClient();
  const { function_name, parameters } = apiCall;

  try {
    switch (function_name) {
      case 'get_data':
        return await getData(sheets, spreadsheetId, parameters.range);
      
      case 'update_data':
        return await updateData(sheets, spreadsheetId, parameters.range, parameters.values);
      
      case 'append_data':
        return await appendData(sheets, spreadsheetId, parameters.range, parameters.values);
      
      case 'clear_data':
        return await clearData(sheets, spreadsheetId, parameters.range);
      
      case 'create_sheet':
        return await createSheet(sheets, spreadsheetId, parameters.sheet_name);
      
      case 'delete_sheet':
        return await deleteSheet(sheets, spreadsheetId, parameters.sheet_name);
      
      case 'rename_sheet':
        return await renameSheet(sheets, spreadsheetId, parameters.old_name, parameters.new_name);
      
      default:
        throw new Error(`פונקציה לא נתמכת: ${function_name}`);
    }
  } catch (error) {
    console.error('Execution error:', error);
    throw new Error(`שגיאה בביצוע הפעולה: ${error.message}`);
  }
}

// Helper functions
async function getData(sheets, spreadsheetId, range) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range
  });
  
  const values = response.data.values || [];
  if (values.length === 0) {
    return 'אין נתונים בטווח המבוקש.';
  }
  
  // Format as HTML table
  let html = '<table border="1" style="border-collapse: collapse; width: 100%;">';
  values.forEach((row, index) => {
    html += '<tr>';
    row.forEach(cell => {
      html += `<td style="padding: 8px; border: 1px solid #ddd;">${cell || ''}</td>`;
    });
    html += '</tr>';
  });
  html += '</table>';
  
  return `נתונים מהטווח ${range}:<br>${html}`;
}

async function updateData(sheets, spreadsheetId, range, values) {
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'RAW',
    resource: { values }
  });
  
  return `עודכן בהצלחה הטווח ${range}`;
}

async function appendData(sheets, spreadsheetId, range, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: 'RAW',
    resource: { values }
  });
  
  return `נוספו נתונים חדשים לטווח ${range}`;
}

async function clearData(sheets, spreadsheetId, range) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range
  });
  
  return `נוקה הטווח ${range}`;
}

async function createSheet(sheets, spreadsheetId, sheetName) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    resource: {
      requests: [{
        addSheet: {
          properties: {
            title: sheetName
          }
        }
      }]
    }
  });
  
  return `נוצר גיליון חדש: ${sheetName}`;
}

async function deleteSheet(sheets, spreadsheetId, sheetName) {
  // First get sheet ID
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = spreadsheet.data.sheets.find(s => s.properties.title === sheetName);
  
  if (!sheet) {
    throw new Error(`גיליון '${sheetName}' לא נמצא`);
  }
  
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    resource: {
      requests: [{
        deleteSheet: {
          sheetId: sheet.properties.sheetId
        }
      }]
    }
  });
  
  return `נמחק הגיליון: ${sheetName}`;
}

async function renameSheet(sheets, spreadsheetId, oldName, newName) {
  // First get sheet ID
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = spreadsheet.data.sheets.find(s => s.properties.title === oldName);
  
  if (!sheet) {
    throw new Error(`גיליון '${oldName}' לא נמצא`);
  }
  
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    resource: {
      requests: [{
        updateSheetProperties: {
          properties: {
            sheetId: sheet.properties.sheetId,
            title: newName
          },
          fields: 'title'
        }
      }]
    }
  });
  
  return `שם הגיליון שונה מ-'${oldName}' ל-'${newName}'`;
}

module.exports = {
  translateToApiCall,
  executeSheetOperation
};