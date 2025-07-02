import os
import json
import re
import traceback
import ast
from flask import Flask, render_template, request, jsonify
from openai import OpenAI
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

app = Flask(__name__)

# --- הגדרות ותצורה (ללא שינוי) ---
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'credentials.json'
OPENAI_API_KEY = "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"
SPREADSHEET_SERVICE = None

# --- פונקציות עזר ופונקציות בסיסיות ---

def get_sheet_service():
    """ מתחבר לשירות של גוגל ויוצר אובייקט שירות גלובלי לשימוש חוזר. """
    global SPREADSHEET_SERVICE
    if SPREADSHEET_SERVICE is None:
        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        SPREADSHEET_SERVICE = build('sheets', 'v4', credentials=creds)
    return SPREADSHEET_SERVICE

def _get_sheet_id_by_name(spreadsheet_id, sheet_name):
    """ פונקציית עזר פנימית חשובה: ממירה שם גיליון למזהה המספרי שלו (sheetId). """
    try:
        service = get_sheet_service()
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        for sheet in sheets:
            if sheet.get('properties', {}).get('title', '') == sheet_name:
                return sheet.get('properties', {}).get('sheetId', '')
        return None # לא נמצא גיליון עם שם כזה
    except HttpError:
        return None

def _get_raw_data(spreadsheet_id, range_name):
    # ... (ללא שינוי)
    try:
        service = get_sheet_service()
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        return result.get('values', [])
    except HttpError as err:
        raise err

# --- קטגוריה 1 + 2: כלים קיימים ---
def get_data(spreadsheet_id, range_name):
    try:
        values = _get_raw_data(spreadsheet_id, range_name)
        if not values: return "לא נמצאו נתונים."
        table_html = '<table border="1" style="border-collapse: collapse; width: 100%; text-align: right;">'
        for row in values:
            table_html += '<tr>' + ''.join(f'<td style="padding: 5px;">{cell}</td>' for cell in row) + '</tr>'
        table_html += '</table>'
        return table_html
    except HttpError as err: return f"אירעה שגיאת API: {err}"

def update_data(spreadsheet_id, range_name, values):
    try:
        body = {'values': values}
        result = get_sheet_service().spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_name, valueInputOption='USER_ENTERED', body=body).execute()
        return f"בוצע בהצלחה. עודכנו {result.get('updatedCells')} תאים."
    except HttpError as err: return f"אירעה שגיאת API: {err}"

def append_data(spreadsheet_id, range_name, values):
    try:
        body = {'values': values}
        result = get_sheet_service().spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_name, valueInputOption='USER_ENTERED', body=body, insertDataOption='INSERT_ROWS').execute()
        return "השורות נוספו בהצלחה לסוף הגיליון."
    except HttpError as err: return f"אירעה שגיאת API: {err}"

def clear_data(spreadsheet_id, range_name):
    try:
        get_sheet_service().spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=range_name).execute()
        return f"הטווח '{range_name}' נוקה בהצלחה."
    except HttpError as err: return f"אירעה שגיאת API: {err}"

def delete_rows_by_keyword(spreadsheet_id, sheet_name, keyword):
    try:
        full_range = f"{sheet_name}!A:Z"
        all_data = _get_raw_data(spreadsheet_id, full_range)
        if not all_data: return "הגיליון ריק, אין מה למחוק."
        data_to_keep = [row for row in all_data if not any(keyword.lower() in str(cell).lower() for cell in row)]
        rows_deleted_count = len(all_data) - len(data_to_keep)
        if rows_deleted_count == 0: return f"לא נמצאו שורות המכילות את המילה '{keyword}'."
        get_sheet_service().spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=full_range).execute()
        if data_to_keep: update_data(spreadsheet_id, f"{sheet_name}!A1", data_to_keep)
        return f"הושלם. נמחקו {rows_deleted_count} שורות שהכילו את המילה '{keyword}'."
    except HttpError as err: return f"אירעה שגיאת API: {err}"
    
def find_and_replace(spreadsheet_id, sheet_name, find_text, replace_text):
    try:
        full_range = f"{sheet_name}!A:Z"
        all_data = _get_raw_data(spreadsheet_id, full_range)
        if not all_data: return "הגיליון ריק, אין מה להחליף."
        new_data = [[str(cell).replace(find_text, replace_text) for cell in row] for row in all_data]
        replacements_count = sum(row.count(replace_text) for row in new_data) - sum(row.count(replace_text) for row in all_data if isinstance(row, list))
        if replacements_count == 0: return f"לא נמצאו מופעים של המילה '{find_text}'."
        update_data(spreadsheet_id, f"{sheet_name}!A1", new_data)
        return f"הושלם. בוצעו {replacements_count} החלפות."
    except HttpError as err: return f"אירעה שגיאת API במהלך החלפת הנתונים: {err}"
    
def generate_and_add_data(spreadsheet_id, sheet_name, header_range, num_rows, user_prompt):
    try:
        headers = _get_raw_data(spreadsheet_id, header_range)
        if not headers: return "שגיאה: לא הצלחתי לקרוא את הכותרות."
        headers_list = headers[0]
        client = OpenAI(api_key=OPENAI_API_KEY)
        data_generation_prompt = f"Based on user request '{user_prompt}' and headers {headers_list}, generate {num_rows} realistic data rows. Response MUST be ONLY a valid Python list of lists."
        response = client.chat.completions.create(model="gpt-4o", messages=[{"role": "system", "content": data_generation_prompt}], temperature=0.7)
        generated_data_str = response.choices[0].message.content
        values_to_add = ast.literal_eval(extract_code_from_markdown(generated_data_str))
        last_header_row = int(re.findall(r'\d+', header_range)[-1])
        append_range = f"{sheet_name}!A{last_header_row + 1}"
        return append_data(spreadsheet_id, append_range, values_to_add)
    except Exception as e: return f"אירעה שגיאה כללית במהלך יצירת הנתונים: {e}"

# --- ⭐️ קטגוריה 3: כלים חדשים לניהול גיליונות ⭐️ ---

def create_sheet(spreadsheet_id, name):
    """יוצר גיליון (טאב) חדש."""
    try:
        body = {'requests': [{'addSheet': {'properties': {'title': name}}}]}
        get_sheet_service().spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        return f"הגיליון '{name}' נוצר בהצלחה."
    except HttpError as err: return f"אירעה שגיאה ביצירת הגיליון: {err}"

def delete_sheet(spreadsheet_id, name):
    """מוחק גיליון (טאב) לפי שם."""
    sheet_id = _get_sheet_id_by_name(spreadsheet_id, name)
    if sheet_id is None: return f"שגיאה: לא נמצא גיליון בשם '{name}'."
    try:
        body = {'requests': [{'deleteSheet': {'sheetId': sheet_id}}]}
        get_sheet_service().spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        return f"הגיליון '{name}' נמחק בהצלחה."
    except HttpError as err: return f"אירעה שגיאה במחיקת הגיליון: {err}"

def rename_sheet(spreadsheet_id, old_name, new_name):
    """משנה שם של גיליון (טאב)."""
    sheet_id = _get_sheet_id_by_name(spreadsheet_id, old_name)
    if sheet_id is None: return f"שגיאה: לא נמצא גיליון בשם '{old_name}'."
    try:
        body = {'requests': [{'updateSheetProperties': {'properties': {'sheetId': sheet_id, 'title': new_name}, 'fields': 'title'}}]}
        get_sheet_service().spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        return f"שם הגיליון שונה מ-'{old_name}' ל-'{new_name}'."
    except HttpError as err: return f"אירעה שגיאה בשינוי השם: {err}"

def duplicate_sheet(spreadsheet_id, source_sheet_name, new_sheet_name):
    """משכפל גיליון קיים."""
    sheet_id = _get_sheet_id_by_name(spreadsheet_id, source_sheet_name)
    if sheet_id is None: return f"שגיאה: לא נמצא גיליון בשם '{source_sheet_name}' לשכפול."
    try:
        body = {'requests': [{'duplicateSheet': {'sourceSheetId': sheet_id, 'newSheetName': new_sheet_name}}]}
        get_sheet_service().spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        return f"הגיליון '{source_sheet_name}' שוכפל בהצלחה לגיליון חדש בשם '{new_sheet_name}'."
    except HttpError as err: return f"אירעה שגיאה בשכפול הגיליון: {err}"

# --- ⭐️ קטגוריה 4: כלי חדש למיון נתונים ⭐️ ---

def sort_sheet(spreadsheet_id, sheet_name, column_index, order):
    """ממיין גיליון שלם לפי עמודה."""
    sheet_id = _get_sheet_id_by_name(spreadsheet_id, sheet_name)
    if sheet_id is None: return f"שגיאה: לא נמצא גיליון בשם '{sheet_name}'."
    try:
        # API is 0-indexed, user provides 1-indexed.
        body = {'requests': [{'sortRange': {
            'range': {'sheetId': sheet_id}, 
            'sortSpecs': [{'dimensionIndex': column_index - 1, 'sortOrder': order}]
        }}]}
        get_sheet_service().spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        return f"הגיליון '{sheet_name}' מוין לפי עמודה {column_index} בסדר {order}."
    except HttpError as err: return f"אירעה שגיאה במיון: {err}"

# --- ⭐️ קטגוריה 5: עיצוב מתקדם ⭐️ ---

def format_cells(spreadsheet_id, range_name, formatting_rules):
    """מעצב תאים בגיליון לפי כללי עיצוב פשוטים."""
    
    # מילון mapping של צבעים לערכי RGB
    COLOR_MAP = {
        'red': {'red': 1.0, 'green': 0.0, 'blue': 0.0},
        'green': {'red': 0.0, 'green': 1.0, 'blue': 0.0},
        'blue': {'red': 0.0, 'green': 0.0, 'blue': 1.0},
        'yellow': {'red': 1.0, 'green': 1.0, 'blue': 0.0},
        'orange': {'red': 1.0, 'green': 0.65, 'blue': 0.0},
        'purple': {'red': 0.5, 'green': 0.0, 'blue': 0.5},
        'pink': {'red': 1.0, 'green': 0.75, 'blue': 0.8},
        'brown': {'red': 0.65, 'green': 0.16, 'blue': 0.16},
        'gray': {'red': 0.5, 'green': 0.5, 'blue': 0.5},
        'light_blue': {'red': 0.8, 'green': 0.9, 'blue': 1.0},
        'light_green': {'red': 0.8, 'green': 1.0, 'blue': 0.8},
        'light_yellow': {'red': 1.0, 'green': 1.0, 'blue': 0.8},
        'white': {'red': 1.0, 'green': 1.0, 'blue': 1.0},
        'black': {'red': 0.0, 'green': 0.0, 'blue': 0.0}
    }
    
    try:
        # בנית אובייקט CellFormat
        cell_format = {}
        
        # עיצוב טקסט
        text_format = {}
        if 'bold' in formatting_rules:
            text_format['bold'] = formatting_rules['bold']
        if 'italic' in formatting_rules:
            text_format['italic'] = formatting_rules['italic']
        if 'underline' in formatting_rules:
            text_format['underline'] = formatting_rules['underline']
        if 'fontSize' in formatting_rules:
            text_format['fontSize'] = formatting_rules['fontSize']
        if 'foregroundColor' in formatting_rules:
            color_name = formatting_rules['foregroundColor'].lower()
            if color_name in COLOR_MAP:
                text_format['foregroundColor'] = COLOR_MAP[color_name]
        
        if text_format:
            cell_format['textFormat'] = text_format
        
        # צבע רקע
        if 'backgroundColor' in formatting_rules:
            color_name = formatting_rules['backgroundColor'].lower()
            if color_name in COLOR_MAP:
                cell_format['backgroundColor'] = COLOR_MAP[color_name]
        
        # יישור אופקי
        if 'horizontalAlignment' in formatting_rules:
            alignment = formatting_rules['horizontalAlignment'].upper()
            if alignment in ['LEFT', 'CENTER', 'RIGHT']:
                cell_format['horizontalAlignment'] = alignment
        
        # בנית הבקשה לAPI
        requests = [{
            'repeatCell': {
                'range': _convert_range_to_grid_range(spreadsheet_id, range_name),
                'cell': {
                    'userEnteredFormat': cell_format
                },
                'fields': _get_format_fields(cell_format)
            }
        }]
        
        body = {'requests': requests}
        get_sheet_service().spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        
        return f"העיצוב הוחל בהצלחה על הטווח '{range_name}'."
        
    except HttpError as err: 
        return f"אירעה שגיאת API בעיצוב: {err}"
    except Exception as e:
        return f"אירעה שגיאה כללית בעיצוב: {e}"

def _convert_range_to_grid_range(spreadsheet_id, range_name):
    """ממיר range כמו 'Sheet1!A1:B2' לאובייקט GridRange."""
    try:
        # פיצול הטווח לשם גיליון וטווח
        if '!' in range_name:
            sheet_name, cell_range = range_name.split('!', 1)
        else:
            sheet_name = 'Sheet1'  # ברירת מחדל
            cell_range = range_name
        
        # קבלת sheet_id
        sheet_id = _get_sheet_id_by_name(spreadsheet_id, sheet_name)
        if sheet_id is None:
            sheet_id = 0  # ברירת מחדל לגיליון הראשון
        
        grid_range = {'sheetId': sheet_id}
        
        # המרת טווח תאים לאינדקסים
        if ':' in cell_range:
            start_cell, end_cell = cell_range.split(':')
            start_col, start_row = _parse_cell_reference(start_cell)
            end_col, end_row = _parse_cell_reference(end_cell)
            
            grid_range.update({
                'startRowIndex': start_row - 1,
                'endRowIndex': end_row,
                'startColumnIndex': start_col - 1,
                'endColumnIndex': end_col
            })
        else:
            # תא יחיד
            col, row = _parse_cell_reference(cell_range)
            grid_range.update({
                'startRowIndex': row - 1,
                'endRowIndex': row,
                'startColumnIndex': col - 1,
                'endColumnIndex': col
            })
        
        return grid_range
        
    except Exception:
        # אם יש שגיאה, נחזיר טווח כולל כברירת מחדל
        return {'sheetId': 0}

def _parse_cell_reference(cell_ref):
    """ממיר התייחסות תא כמו 'A1' לאינדקסים (עמודה, שורה)."""
    col_str = ''
    row_str = ''
    
    for char in cell_ref:
        if char.isalpha():
            col_str += char
        elif char.isdigit():
            row_str += char
    
    # המרת אותיות לאינדקס עמודה (A=1, B=2, ..., Z=26, AA=27)
    col_index = 0
    for char in col_str.upper():
        col_index = col_index * 26 + ord(char) - ord('A') + 1
    
    row_index = int(row_str) if row_str else 1
    
    return col_index, row_index

def _get_format_fields(cell_format):
    """יוצר מחרוזת fields עבור batchUpdate בהתבסס על העיצוב שהוגדר."""
    fields = []
    
    if 'textFormat' in cell_format:
        text_fields = []
        if 'bold' in cell_format['textFormat']:
            text_fields.append('textFormat.bold')
        if 'italic' in cell_format['textFormat']:
            text_fields.append('textFormat.italic')
        if 'underline' in cell_format['textFormat']:
            text_fields.append('textFormat.underline')
        if 'fontSize' in cell_format['textFormat']:
            text_fields.append('textFormat.fontSize')
        if 'foregroundColor' in cell_format['textFormat']:
            text_fields.append('textFormat.foregroundColor')
        fields.extend(text_fields)
    
    if 'backgroundColor' in cell_format:
        fields.append('backgroundColor')
    
    if 'horizontalAlignment' in cell_format:
        fields.append('horizontalAlignment')
    
    return ','.join(fields) if fields else 'userEnteredFormat'

# --- ⭐️ קטגוריה 6: תרשימים ⭐️ ---

def create_chart(spreadsheet_id, sheet_name, source_range, chart_type, title):
    """יוצר תרשים בגיליון מנתונים קיימים."""
    
    # מילון סוגי תרשימים נתמכים
    CHART_TYPES = {
        'COLUMN': 'COLUMN',
        'BAR': 'BAR', 
        'PIE': 'PIE',
        'LINE': 'LINE',
        'SCATTER': 'SCATTER'
    }
    
    try:
        # קבלת sheet_id
        sheet_id = _get_sheet_id_by_name(spreadsheet_id, sheet_name)
        if sheet_id is None:
            return f"שגיאה: לא נמצא גיליון בשם '{sheet_name}'."
        
        # בדיקת סוג תרשים תקין
        chart_type = chart_type.upper()
        if chart_type not in CHART_TYPES:
            return f"שגיאה: סוג תרשים '{chart_type}' לא נתמך. סוגים זמינים: {', '.join(CHART_TYPES.keys())}"
        
        # המרת הטווח לגריד
        source_grid_range = _convert_range_to_grid_range(spreadsheet_id, source_range)
        
        # בנית אובייקט התרשים
        chart_spec = {
            'title': title,
            'basicChart': {
                'chartType': CHART_TYPES[chart_type],
                'domains': [source_grid_range],
                'series': [source_grid_range],
                'headerCount': 1  # השורה הראשונה כותרות
            }
        }
        
        # אם זה תרשים עוגה, צריך להגדיר אחרת
        if chart_type == 'PIE':
            chart_spec = {
                'title': title,
                'pieChart': {
                    'domain': source_grid_range,
                    'series': source_grid_range
                }
            }
        
        # בקשת יצירת התרשים
        requests = [{
            'addChart': {
                'chart': {
                    'spec': chart_spec,
                    'position': {
                        'overlayPosition': {
                            'anchorCell': {
                                'sheetId': sheet_id,
                                'rowIndex': 0,
                                'columnIndex': 0
                            }
                        }
                    }
                }
            }
        }]
        
        body = {'requests': requests}
        get_sheet_service().spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        
        return f"התרשים '{title}' נוצר בהצלחה בגיליון '{sheet_name}'."
        
    except HttpError as err:
        return f"אירעה שגיאת API ביצירת התרשים: {err}"
    except Exception as e:
        return f"אירעה שגיאה כללית ביצירת התרשים: {e}"

# --- מוח ה-AI: הנחיית מערכת מעודכנת עם כל הכלים ---
def extract_code_from_markdown(text):
    match = re.search(r"```(?:python)?\s*\n?(.*?)\n?```", text, re.DOTALL)
    if match: return match.group(1).strip()
    return text.strip()

def translate_to_api_call(user_prompt, spreadsheet_id):
    client = OpenAI(api_key=OPENAI_API_KEY)
    
    # ⭐️ המוח המורחב של הבוט ⭐️
    system_prompt = f"""
    You are an expert AI assistant that translates natural language into a SINGLE Python function call for Google Sheets.
    The user's original prompt is: "{user_prompt}"
    The spreadsheet_id is always: '{spreadsheet_id}'

    **CHOOSE ONE FUNCTION FROM THIS LIST:**

    --- DATA & CONTENT TOOLS ---
    - `get_data(spreadsheet_id, range_name)`: To read and display data from a range.
    - `update_data(spreadsheet_id, range_name, values)`: To write or overwrite data.
    - `append_data(spreadsheet_id, range_name, values)`: To add new rows to the end.
    - `clear_data(spreadsheet_id, range_name)`: To delete content from cells.
    - `find_and_replace(spreadsheet_id, sheet_name, find_text, replace_text)`: **Use for "replace" or "change" requests.**
    - `delete_rows_by_keyword(spreadsheet_id, sheet_name, keyword)`: **Use to delete entire rows containing a word.**
    - `generate_and_add_data(spreadsheet_id, sheet_name, header_range, num_rows, user_prompt)`: **Use to invent realistic data.**

    --- SHEET STRUCTURE & ORGANIZATION TOOLS ---
    - `create_sheet(spreadsheet_id, name)`: To create a new sheet (tab).
    - `delete_sheet(spreadsheet_id, name)`: To delete a sheet.
    - `rename_sheet(spreadsheet_id, old_name, new_name)`: To rename a sheet.
    - `duplicate_sheet(spreadsheet_id, source_sheet_name, new_sheet_name)`: To copy an entire sheet.
    - `sort_sheet(spreadsheet_id, sheet_name, column_index, order)`: **Use to sort a whole sheet.** `column_index` is 1-based (A=1, B=2...). `order` must be 'ASCENDING' or 'DESCENDING'.

    --- FORMATTING TOOLS ---
    - `format_cells(spreadsheet_id, range_name, formatting_rules)`: **Use to change cell appearance** (colors, font, alignment). `formatting_rules` is a dictionary.

    --- ADVANCED FEATURES ---
    - `create_chart(spreadsheet_id, sheet_name, source_range, chart_type, title)`: **Use to create charts** from data.

    **CRITICAL RULES:**
    1.  Infer all parameters from the user's prompt.
    2.  For `sort_sheet`, you must infer the column number. 'עמודה א' is 1, 'עמודה ב' is 2, etc.
    3.  Return ONLY the Python code for the single function call. No other text.

    **CRITICAL RULES FOR FORMATTING:**
    - Infer the formatting_rules dictionary from the user's request. Possible keys are: backgroundColor, foregroundColor, bold, italic, underline, fontSize, horizontalAlignment.
    - You can combine multiple rules in one call.

    **CRITICAL RULES FOR CHARTING:**
    - Infer the chart_type from the user's request. Choose one of: 'COLUMN', 'BAR', 'PIE', 'LINE', 'SCATTER'.
    - Infer the source_range of the data to be charted.

    **EXAMPLES:**
    User: "תמחק את הגיליון 'זמני'" -> `delete_sheet(spreadsheet_id='{spreadsheet_id}', name='זמני')`
    User: "מיין את גיליון 'לקוחות' לפי עמודה 5 בסדר עולה" -> `sort_sheet(spreadsheet_id='{spreadsheet_id}', sheet_name='לקוחות', column_index=5, order='ASCENDING')`
    User: "צבע את הרקע של A1 עד B2 בצהוב" -> `format_cells(spreadsheet_id='{spreadsheet_id}', range_name='Sheet1!A1:B2', formatting_rules={{'backgroundColor': 'yellow'}})`
    User: "הדגש את הטקסט בשורה 1, מרכז אותו, וקבע גודל גופן 12" -> `format_cells(spreadsheet_id='{spreadsheet_id}', range_name='Sheet1!1:1', formatting_rules={{'bold': True, 'horizontalAlignment': 'CENTER', 'fontSize': 12}})`
    User: "צור גרף עמודות מהנתונים בעמודות A ו-C" -> `create_chart(spreadsheet_id='{spreadsheet_id}', sheet_name='Sheet1', source_range='Sheet1!A:C', chart_type='COLUMN', title='גרף עמודות')`
    User: "אני רוצה גרף פאי על בסיס הנתונים ב-A1 עד B10, עם כותרת 'התפלגות'" -> `create_chart(spreadsheet_id='{spreadsheet_id}', sheet_name='Sheet1', source_range='Sheet1!A1:B10', chart_type='PIE', title='התפלגות')`
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "system", "content": system_prompt}],
            temperature=0.0
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error communicating with OpenAI: {e}"

# --- נתיבי ה-Flask (ללא שינוי) ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/chat', methods=['POST'])
def chat():
    data = request.json
    user_message = data.get('message')
    spreadsheet_id = data.get('spreadsheet_id')

    if not user_message or not spreadsheet_id:
        return jsonify({'reply': 'שגיאה: הודעה או מזהה גיליון חסרים.'}), 400

    print(f"\n[לוג] התקבלה בקשה מהמשתמש: '{user_message}'")
    
    # מאתחל את שירות ה-API בפעם הראשונה
    get_sheet_service() 
    
    raw_api_call = translate_to_api_call(user_message, spreadsheet_id)
    api_call_str = extract_code_from_markdown(raw_api_call)
    print(f"[לוג] פקודה שנוצרה (לאחר ניקוי): {api_call_str}")

    bot_response = ""
    try:
        if not ('(' in api_call_str and ')' in api_call_str):
             raise SyntaxError("ה-AI החזיר טקסט במקום פקודה. נסה לנסח מחדש.")
             
        # כאן ה-eval קורא לפונקציות החדשות שהוספנו
        result = eval(api_call_str)
        bot_response = str(result)
        print(f"[לוג] תוצאת הפעולה: {bot_response}")

    except Exception as e:
        error_details = traceback.format_exc()
        print(f"[שגיאה] אירעה שגיאה בביצוע הפעולה: {e}")
        print(error_details)
        bot_response = f"אופס, משהו השתבש.<br><strong>השגיאה:</strong> {e}<br>ייתכן שהפקודה שה-AI יצר לא הייתה נכונה או שהבקשה לא הייתה ברורה. נסה לנסח מחדש."

    return jsonify({'reply': bot_response})

if __name__ == '__main__':
    try:
        with open(SERVICE_ACCOUNT_FILE, 'r') as f:
            creds_data = json.load(f)
            service_email = creds_data.get('client_email')
            print("-" * 50)
            print("שרת הצ'אט מתחיל...")
            print(f"ודא ששיתפת את הגיליון שלך עם כתובת המייל: \n{service_email}\nעם הרשאות 'Editor'.")
            print("פתח את הדפדפן ועבור לכתובת: http://127.0.0.1:5000")
            print("כדי לעצור את השרת, לחץ Ctrl+C בחלון זה.")
            print("-" * 50)
    except FileNotFoundError:
        print(f"שגיאה קריטית: הקובץ '{SERVICE_ACCOUNT_FILE}' לא נמצא. לא ניתן להפעיל את השרת.")
    
    app.run(debug=True, use_reloader=False)