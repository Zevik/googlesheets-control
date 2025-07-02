import { google } from 'googleapis';
import OpenAI from 'openai';

// הגדרת Google Sheets
const getSheetService = () => {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  return google.sheets({ version: 'v4', auth });
};

// פונקציית עזר לקבלת נתונים גולמיים
async function getRawData(spreadsheetId, rangeName) {
  try {
    const sheets = getSheetService();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: rangeName,
    });
    return result.data.values || [];
  } catch (error) {
    throw error;
  }
}

// פונקציית עזר לקבלת מזהה גיליון לפי שם
async function getSheetIdByName(spreadsheetId, sheetName) {
  try {
    const sheets = getSheetService();
    const metadata = await sheets.spreadsheets.get({
      spreadsheetId,
    });
    
    const sheet = metadata.data.sheets?.find(
      s => s.properties?.title === sheetName
    );
    
    return sheet?.properties?.sheetId || null;
  } catch (error) {
    return null;
  }
}

// === פונקציות עיקריות ===

export async function getData(spreadsheetId, rangeName) {
  try {
    const values = await getRawData(spreadsheetId, rangeName);
    if (!values.length) return "לא נמצאו נתונים.";
    
    // המר לטבלת HTML
    let tableHtml = '<table border="1" style="border-collapse: collapse; width: 100%; text-align: right;">';
    values.forEach(row => {
      tableHtml += '<tr>' + row.map(cell => `<td style="padding: 5px;">${cell || ''}</td>`).join('') + '</tr>';
    });
    tableHtml += '</table>';
    
    return tableHtml;
  } catch (error) {
    return `אירעה שגיאת API: ${error.message}`;
  }
}

export async function updateData(spreadsheetId, rangeName, values) {
  try {
    const sheets = getSheetService();
    const result = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: rangeName,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values },
    });
    
    return `בוצע בהצלחה. עודכנו ${result.data.updatedCells} תאים.`;
  } catch (error) {
    return `אירעה שגיאת API: ${error.message}`;
  }
}

export async function appendData(spreadsheetId, rangeName, values) {
  try {
    const sheets = getSheetService();
    const result = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: rangeName,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values },
    });
    
    return "השורות נוספו בהצלחה לסוף הגיליון.";
  } catch (error) {
    return `אירעה שגיאת API: ${error.message}`;
  }
}

export async function clearData(spreadsheetId, rangeName) {
  try {
    const sheets = getSheetService();
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: rangeName,
    });
    
    return `הטווח '${rangeName}' נוקה בהצלחה.`;
  } catch (error) {
    return `אירעה שגיאת API: ${error.message}`;
  }
}

export async function deleteRowsByKeyword(spreadsheetId, sheetName, keyword) {
  try {
    const fullRange = `${sheetName}!A:Z`;
    const allData = await getRawData(spreadsheetId, fullRange);
    
    if (!allData.length) return "הגיליון ריק, אין מה למחוק.";
    
    const dataToKeep = allData.filter(row => 
      !row.some(cell => String(cell).toLowerCase().includes(keyword.toLowerCase()))
    );
    
    const rowsDeletedCount = allData.length - dataToKeep.length;
    
    if (rowsDeletedCount === 0) {
      return `לא נמצאו שורות המכילות את המילה '${keyword}'.`;
    }
    
    // נקה את הגיליון
    await clearData(spreadsheetId, fullRange);
    
    // הכנס בחזרה את הנתונים שנשארו
    if (dataToKeep.length > 0) {
      await updateData(spreadsheetId, `${sheetName}!A1`, dataToKeep);
    }
    
    return `הושלם. נמחקו ${rowsDeletedCount} שורות שהכילו את המילה '${keyword}'.`;
  } catch (error) {
    return `אירעה שגיאת API: ${error.message}`;
  }
}

export async function findAndReplace(spreadsheetId, sheetName, findText, replaceText) {
  try {
    const fullRange = `${sheetName}!A:Z`;
    const allData = await getRawData(spreadsheetId, fullRange);
    
    if (!allData.length) return "הגיליון ריק, אין מה להחליף.";
    
    const newData = allData.map(row => 
      row.map(cell => String(cell).replace(new RegExp(findText, 'g'), replaceText))
    );
    
    // ספירת החלפות
    const originalText = allData.flat().join('');
    const newText = newData.flat().join('');
    const replacementsCount = (originalText.match(new RegExp(findText, 'g')) || []).length;
    
    if (replacementsCount === 0) {
      return `לא נמצאו מופעים של המילה '${findText}'.`;
    }
    
    await updateData(spreadsheetId, `${sheetName}!A1`, newData);
    
    return `הושלם. בוצעו ${replacementsCount} החלפות.`;
  } catch (error) {
    return `אירעה שגיאת API במהלך החלפת הנתונים: ${error.message}`;
  }
}

export async function generateAndAddData(spreadsheetId, sheetName, headerRange, numRows, userPrompt) {
  try {
    const headers = await getRawData(spreadsheetId, headerRange);
    if (!headers.length) return "שגיאה: לא הצלחתי לקרוא את הכותרות.";
    
    const headersList = headers[0];
    
    const openai = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });
    
    const dataGenerationPrompt = `Based on user request '${userPrompt}' and headers ${JSON.stringify(headersList)}, generate ${numRows} realistic data rows. Response MUST be ONLY a valid JSON array of arrays.`;
    
    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [{ role: "system", content: dataGenerationPrompt }],
      temperature: 0.7,
    });
    
    const generatedDataStr = response.choices[0].message.content;
    const valuesToAdd = JSON.parse(generatedDataStr);
    
    // קבל את מספר השורה האחרונה מהכותרות
    const lastHeaderRow = parseInt(headerRange.match(/\d+/g)?.slice(-1)[0] || '1');
    const appendRange = `${sheetName}!A${lastHeaderRow + 1}`;
    
    return await appendData(spreadsheetId, appendRange, valuesToAdd);
  } catch (error) {
    return `אירעה שגיאה כללית במהלך יצירת הנתונים: ${error.message}`;
  }
}

// === ניהול גיליונות ===

export async function createSheet(spreadsheetId, name) {
  try {
    const sheets = getSheetService();
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: name } } }]
      }
    });
    
    return `הגיליון '${name}' נוצר בהצלחה.`;
  } catch (error) {
    return `אירעה שגיאה ביצירת הגיליון: ${error.message}`;
  }
}

export async function deleteSheet(spreadsheetId, name) {
  try {
    const sheetId = await getSheetIdByName(spreadsheetId, name);
    if (sheetId === null) return `שגיאה: לא נמצא גיליון בשם '${name}'.`;
    
    const sheets = getSheetService();
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ deleteSheet: { sheetId } }]
      }
    });
    
    return `הגיליון '${name}' נמחק בהצלחה.`;
  } catch (error) {
    return `אירעה שגיאה במחיקת הגיליון: ${error.message}`;
  }
}

export async function renameSheet(spreadsheetId, oldName, newName) {
  try {
    const sheetId = await getSheetIdByName(spreadsheetId, oldName);
    if (sheetId === null) return `שגיאה: לא נמצא גיליון בשם '${oldName}'.`;
    
    const sheets = getSheetService();
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          updateSheetProperties: {
            properties: { sheetId, title: newName },
            fields: 'title'
          }
        }]
      }
    });
    
    return `שם הגיליון שונה מ-'${oldName}' ל-'${newName}'.`;
  } catch (error) {
    return `אירעה שגיאה בשינוי השם: ${error.message}`;
  }
}

export async function duplicateSheet(spreadsheetId, sourceSheetName, newSheetName) {
  try {
    const sheetId = await getSheetIdByName(spreadsheetId, sourceSheetName);
    if (sheetId === null) return `שגיאה: לא נמצא גיליון בשם '${sourceSheetName}' לשכפול.`;
    
    const sheets = getSheetService();
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          duplicateSheet: {
            sourceSheetId: sheetId,
            newSheetName: newSheetName
          }
        }]
      }
    });
    
    return `הגיליון '${sourceSheetName}' שוכפל בהצלחה לגיליון חדש בשם '${newSheetName}'.`;
  } catch (error) {
    return `אירעה שגיאה בשכפול הגיליון: ${error.message}`;
  }
}

// === מיון ===

export async function sortSheet(spreadsheetId, sheetName, columnIndex, order) {
  try {
    const sheetId = await getSheetIdByName(spreadsheetId, sheetName);
    if (sheetId === null) return `שגיאה: לא נמצא גיליון בשם '${sheetName}'.`;
    
    const sheets = getSheetService();
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          sortRange: {
            range: { sheetId },
            sortSpecs: [{
              dimensionIndex: columnIndex - 1,
              sortOrder: order
            }]
          }
        }]
      }
    });
    
    return `הגיליון '${sheetName}' מוין לפי עמודה ${columnIndex} בסדר ${order}.`;
  } catch (error) {
    return `אירעה שגיאה במיון: ${error.message}`;
  }
}

// === עיצוב ===

// מילון צבעים
const COLOR_MAP = {
  'red': { red: 1.0, green: 0.0, blue: 0.0 },
  'green': { red: 0.0, green: 1.0, blue: 0.0 },
  'blue': { red: 0.0, green: 0.0, blue: 1.0 },
  'yellow': { red: 1.0, green: 1.0, blue: 0.0 },
  'orange': { red: 1.0, green: 0.65, blue: 0.0 },
  'purple': { red: 0.5, green: 0.0, blue: 0.5 },
  'pink': { red: 1.0, green: 0.75, blue: 0.8 },
  'brown': { red: 0.65, green: 0.16, blue: 0.16 },
  'gray': { red: 0.5, green: 0.5, blue: 0.5 },
  'light_blue': { red: 0.8, green: 0.9, blue: 1.0 },
  'light_green': { red: 0.8, green: 1.0, blue: 0.8 },
  'light_yellow': { red: 1.0, green: 1.0, blue: 0.8 },
  'white': { red: 1.0, green: 1.0, blue: 1.0 },
  'black': { red: 0.0, green: 0.0, blue: 0.0 }
};

function convertRangeToGridRange(spreadsheetId, rangeName, sheetId = 0) {
  try {
    const gridRange = { sheetId };
    
    // אם יש טווח עם :
    if (rangeName.includes(':')) {
      const [start, end] = rangeName.split(':');
      const [startCol, startRow] = parseCellReference(start);
      const [endCol, endRow] = parseCellReference(end);
      
      gridRange.startRowIndex = startRow - 1;
      gridRange.endRowIndex = endRow;
      gridRange.startColumnIndex = startCol - 1;
      gridRange.endColumnIndex = endCol;
    } else {
      // תא יחיד
      const [col, row] = parseCellReference(rangeName);
      gridRange.startRowIndex = row - 1;
      gridRange.endRowIndex = row;
      gridRange.startColumnIndex = col - 1;
      gridRange.endColumnIndex = col;
    }
    
    return gridRange;
  } catch (error) {
    return { sheetId };
  }
}

function parseCellReference(cellRef) {
  let colStr = '';
  let rowStr = '';
  
  for (const char of cellRef) {
    if (/[A-Za-z]/.test(char)) {
      colStr += char;
    } else if (/\d/.test(char)) {
      rowStr += char;
    }
  }
  
  // המרת אותיות לאינדקס עמודה
  let colIndex = 0;
  for (const char of colStr.toUpperCase()) {
    colIndex = colIndex * 26 + char.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  }
  
  const rowIndex = parseInt(rowStr) || 1;
  
  return [colIndex, rowIndex];
}

function getFormatFields(cellFormat) {
  const fields = [];
  
  if (cellFormat.textFormat) {
    if (cellFormat.textFormat.bold !== undefined) fields.push('textFormat.bold');
    if (cellFormat.textFormat.italic !== undefined) fields.push('textFormat.italic');
    if (cellFormat.textFormat.underline !== undefined) fields.push('textFormat.underline');
    if (cellFormat.textFormat.fontSize !== undefined) fields.push('textFormat.fontSize');
    if (cellFormat.textFormat.foregroundColor) fields.push('textFormat.foregroundColor');
  }
  
  if (cellFormat.backgroundColor) fields.push('backgroundColor');
  if (cellFormat.horizontalAlignment) fields.push('horizontalAlignment');
  
  return fields.length > 0 ? fields.join(',') : 'userEnteredFormat';
}

export async function formatCells(spreadsheetId, rangeName, formattingRules) {
  try {
    // קבלת sheet ID מהטווח
    let sheetName = 'Sheet1';
    let actualRange = rangeName;
    
    if (rangeName.includes('!')) {
      [sheetName, actualRange] = rangeName.split('!');
    }
    
    const sheetId = await getSheetIdByName(spreadsheetId, sheetName) || 0;
    
    // בנית אובייקט CellFormat
    const cellFormat = {};
    
    // עיצוב טקסט
    const textFormat = {};
    if (formattingRules.bold !== undefined) textFormat.bold = formattingRules.bold;
    if (formattingRules.italic !== undefined) textFormat.italic = formattingRules.italic;
    if (formattingRules.underline !== undefined) textFormat.underline = formattingRules.underline;
    if (formattingRules.fontSize) textFormat.fontSize = formattingRules.fontSize;
    if (formattingRules.foregroundColor) {
      const colorName = formattingRules.foregroundColor.toLowerCase();
      if (COLOR_MAP[colorName]) {
        textFormat.foregroundColor = COLOR_MAP[colorName];
      }
    }
    
    if (Object.keys(textFormat).length > 0) {
      cellFormat.textFormat = textFormat;
    }
    
    // צבע רקע
    if (formattingRules.backgroundColor) {
      const colorName = formattingRules.backgroundColor.toLowerCase();
      if (COLOR_MAP[colorName]) {
        cellFormat.backgroundColor = COLOR_MAP[colorName];
      }
    }
    
    // יישור אופקי
    if (formattingRules.horizontalAlignment) {
      const alignment = formattingRules.horizontalAlignment.toUpperCase();
      if (['LEFT', 'CENTER', 'RIGHT'].includes(alignment)) {
        cellFormat.horizontalAlignment = alignment;
      }
    }
    
    // בנית הבקשה
    const sheets = getSheetService();
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          repeatCell: {
            range: convertRangeToGridRange(spreadsheetId, actualRange, sheetId),
            cell: {
              userEnteredFormat: cellFormat
            },
            fields: getFormatFields(cellFormat)
          }
        }]
      }
    });
    
    return `העיצוב הוחל בהצלחה על הטווח '${rangeName}'.`;
  } catch (error) {
    return `אירעה שגיאה בעיצוב: ${error.message}`;
  }
}

// === תרשימים ===

const CHART_TYPES = {
  'COLUMN': 'COLUMN',
  'BAR': 'BAR',
  'PIE': 'PIE',
  'LINE': 'LINE',
  'SCATTER': 'SCATTER'
};

export async function createChart(spreadsheetId, sheetName, sourceRange, chartType, title) {
  try {
    const sheetId = await getSheetIdByName(spreadsheetId, sheetName);
    if (sheetId === null) return `שגיאה: לא נמצא גיליון בשם '${sheetName}'.`;
    
    const upperChartType = chartType.toUpperCase();
    if (!CHART_TYPES[upperChartType]) {
      return `שגיאה: סוג תרשים '${chartType}' לא נתמך. סוגים זמינים: ${Object.keys(CHART_TYPES).join(', ')}`;
    }
    
    // המרת הטווח לגריד
    const actualRange = sourceRange.includes('!') ? sourceRange.split('!')[1] : sourceRange;
    const sourceGridRange = convertRangeToGridRange(spreadsheetId, actualRange, sheetId);
    
    // בנית אובייקט התרשים
    let chartSpec;
    
    if (upperChartType === 'PIE') {
      chartSpec = {
        title,
        pieChart: {
          domain: sourceGridRange,
          series: sourceGridRange
        }
      };
    } else {
      chartSpec = {
        title,
        basicChart: {
          chartType: CHART_TYPES[upperChartType],
          domains: [sourceGridRange],
          series: [sourceGridRange],
          headerCount: 1
        }
      };
    }
    
    const sheets = getSheetService();
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          addChart: {
            chart: {
              spec: chartSpec,
              position: {
                overlayPosition: {
                  anchorCell: {
                    sheetId,
                    rowIndex: 0,
                    columnIndex: 0
                  }
                }
              }
            }
          }
        }]
      }
    });
    
    return `התרשים '${title}' נוצר בהצלחה בגיליון '${sheetName}'.`;
  } catch (error) {
    return `אירעה שגיאה ביצירת התרשים: ${error.message}`;
  }
}

// === תרגום הודעות לפקודות API ===

export async function translateToApiCall(userPrompt, spreadsheetId) {
  const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
  });

  const systemPrompt = `
    You are an expert AI assistant that translates natural language into a SINGLE JavaScript function call for Google Sheets.
    The user's original prompt is: "${userPrompt}"
    The spreadsheet_id is always: '${spreadsheetId}'

    **CHOOSE ONE FUNCTION FROM THIS LIST:**

    --- DATA & CONTENT TOOLS ---
    - getData(spreadsheet_id, range_name): To read and display data from a range.
    - updateData(spreadsheet_id, range_name, values): To write or overwrite data.
    - appendData(spreadsheet_id, range_name, values): To add new rows to the end.
    - clearData(spreadsheet_id, range_name): To delete content from cells.
    - findAndReplace(spreadsheet_id, sheet_name, find_text, replace_text): Use for "replace" or "change" requests.
    - deleteRowsByKeyword(spreadsheet_id, sheet_name, keyword): Use to delete entire rows containing a word.
    - generateAndAddData(spreadsheet_id, sheet_name, header_range, num_rows, user_prompt): Use to invent realistic data.

    --- SHEET STRUCTURE & ORGANIZATION TOOLS ---
    - createSheet(spreadsheet_id, name): To create a new sheet (tab).
    - deleteSheet(spreadsheet_id, name): To delete a sheet.
    - renameSheet(spreadsheet_id, old_name, new_name): To rename a sheet.
    - duplicateSheet(spreadsheet_id, source_sheet_name, new_sheet_name): To copy an entire sheet.
    - sortSheet(spreadsheet_id, sheet_name, column_index, order): Use to sort a whole sheet. column_index is 1-based (A=1, B=2...). order must be 'ASCENDING' or 'DESCENDING'.

    --- FORMATTING TOOLS ---
    - formatCells(spreadsheet_id, range_name, formatting_rules): Use to change cell appearance (colors, font, alignment). formatting_rules is an object.

    --- ADVANCED FEATURES ---
    - createChart(spreadsheet_id, sheet_name, source_range, chart_type, title): Use to create charts from data.

    Return ONLY a JSON object with function name and parameters:
    {"function": "getData", "params": {"spreadsheet_id": "...", "range_name": "..."}}
  `;

  try {
    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [{ role: "system", content: systemPrompt }],
      temperature: 0.0,
    });

    return JSON.parse(response.choices[0].message.content);
  } catch (error) {
    throw new Error(`Error communicating with OpenAI: ${error.message}`);
  }
}

// === ביצוע פעולות ===

export async function executeSheetOperation(apiCall, spreadsheetId) {
  const { function: funcName, params } = apiCall;
  
  switch (funcName) {
    case 'getData':
      return await getData(params.spreadsheet_id, params.range_name);
    case 'updateData':
      return await updateData(params.spreadsheet_id, params.range_name, params.values);
    case 'appendData':
      return await appendData(params.spreadsheet_id, params.range_name, params.values);
    case 'clearData':
      return await clearData(params.spreadsheet_id, params.range_name);
    case 'findAndReplace':
      return await findAndReplace(params.spreadsheet_id, params.sheet_name, params.find_text, params.replace_text);
    case 'deleteRowsByKeyword':
      return await deleteRowsByKeyword(params.spreadsheet_id, params.sheet_name, params.keyword);
    case 'generateAndAddData':
      return await generateAndAddData(params.spreadsheet_id, params.sheet_name, params.header_range, params.num_rows, params.user_prompt);
    case 'createSheet':
      return await createSheet(params.spreadsheet_id, params.name);
    case 'deleteSheet':
      return await deleteSheet(params.spreadsheet_id, params.name);
    case 'renameSheet':
      return await renameSheet(params.spreadsheet_id, params.old_name, params.new_name);
    case 'duplicateSheet':
      return await duplicateSheet(params.spreadsheet_id, params.source_sheet_name, params.new_sheet_name);
    case 'sortSheet':
      return await sortSheet(params.spreadsheet_id, params.sheet_name, params.column_index, params.order);
    case 'formatCells':
      return await formatCells(params.spreadsheet_id, params.range_name, params.formatting_rules);
    case 'createChart':
      return await createChart(params.spreadsheet_id, params.sheet_name, params.source_range, params.chart_type, params.title);
    default:
      throw new Error(`Unknown function: ${funcName}`);
  }
}