const { translateToApiCall, executeSheetOperation } = require('../../lib/sheets');

exports.handler = async (event, context) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers,
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  try {
    const { message, spreadsheet_id } = JSON.parse(event.body);

    if (!message || !spreadsheet_id) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ reply: 'שגיאה: הודעה או מזהה גיליון חסרים.' })
      };
    }

    const apiCall = await translateToApiCall(message, spreadsheet_id);
    const result = await executeSheetOperation(apiCall, spreadsheet_id);
    
    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ reply: result })
    };
  } catch (error) {
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ reply: `אופס, משהו השתבש. השגיאה: ${error.message}` })
    };
  }
}; 