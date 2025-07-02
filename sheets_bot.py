import os
import json
import re  # <-- הוספה חדשה
from openai import OpenAI
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- הגדרות ותצורה (ללא שינוי) ---

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'credentials.json'
OPENAI_API_KEY = "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"
if not OPENAI_API_KEY:
    raise ValueError("לא נמצא מפתח API של OpenAI. הגדר אותו במשתנה OPENAI_API_KEY.")

# --- פונקציות עזר ל-Google Sheets API (ללא שינוי) ---

def get_sheet_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    return service.spreadsheets()

def get_data(spreadsheet_id, range_name):
    try:
        sheet = get_sheet_service()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])
        if not values:
            return "לא נמצאו נתונים."
        return values
    except HttpError as err:
        return f"אירעה שגיאה: {err}"

def update_data(spreadsheet_id, range_name, values):
    try:
        sheet = get_sheet_service()
        body = {'values': values}
        result = sheet.values().update(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption='USER_ENTERED', body=body).execute()
        return f"עודכנו {result.get('updatedCells')} תאים."
    except HttpError as err:
        return f"אירעה שגיאה: {err}"

def append_data(spreadsheet_id, range_name, values):
    try:
        sheet = get_sheet_service()
        body = {'values': values}
        result = sheet.values().append(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption='USER_ENTERED', body=body,
            insertDataOption='INSERT_ROWS').execute()
        return "הנתונים נוספו בהצלחה."
    except HttpError as err:
        return f"אירעה שגיאה: {err}"

def clear_data(spreadsheet_id, range_name):
    try:
        sheet = get_sheet_service()
        result = sheet.values().clear(spreadsheetId=spreadsheet_id, range=range_name).execute()
        return f"הטווח '{range_name}' נוקה בהצלחה."
    except HttpError as err:
        return f"אירעה שגיאה: {err}"

# --- פונקציות OpenAI (ללא שינוי) ---

def translate_to_api_call(user_prompt, spreadsheet_id):
    client = OpenAI(api_key=OPENAI_API_KEY)
    system_prompt = f"""
    אתה עוזר חכם שמתרגם בקשות בשפה טבעית לקריאות לפונקציות פייתון עבור Google Sheets.
    עליך להחזיר אך ורק קוד פייתון של קריאה לפונקציה אחת מתוך הרשימה הבאה:
    - `get_data(spreadsheet_id, range_name)`
    - `update_data(spreadsheet_id, range_name, values)`
    - `append_data(spreadsheet_id, range_name, values)`
    - `clear_data(spreadsheet_id, range_name)`

    ה-spreadsheet_id הוא תמיד: '{spreadsheet_id}'.
    טווחים צריכים להיות בפורמט 'Sheet1!A1:B10' או 'גיליון1!A:C'.
    נתונים לכתיבה (values) צריכים להיות בפורמט של רשימה של רשימות, למשל: [['ערך1', 'ערך2']].

    דוגמאות:
    משתמש: "מה יש בתאים A1 עד B2 בגיליון הראשון?"
    תשובתך: get_data(spreadsheet_id='{spreadsheet_id}', range_name='Sheet1!A1:B2')
    
    משתמש: "תוסיף שורה עם 'משה', 'לוי', '42' בסוף הגיליון"
    תשובתך: append_data(spreadsheet_id='{spreadsheet_id}', range_name='Sheet1!A1', values=[['משה', 'לוי', '42']])

    משתמש: "תכתוב 'שלום עולם' בתא C5"
    תשובתך: update_data(spreadsheet_id='{spreadsheet_id}', range_name='Sheet1!C5', values=[['שלום עולם']])

    משתמש: "תנקה את כל גיליון 2"
    תשובתך: clear_data(spreadsheet_id='{spreadsheet_id}', range_name='Sheet2')
    
    החזר רק את שורת הקוד. בלי הסברים, בלי מלל נוסף, בלי Markdown.
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.0,
            max_tokens=200
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error communicating with OpenAI: {e}"

# --- פונקציית עזר חדשה לניקוי הקוד ---

def extract_code_from_markdown(text):
    """
    מחלצת קוד פייתון מתוך בלוק Markdown.
    מתמודדת עם ```python ... ``` וגם עם ``` ... ```.
    """
    pattern = r"```(?:python)?\s*\n?(.*?)\n?```"
    match = re.search(pattern, text, re.DOTALL)
    if match:
        return match.group(1).strip()
    else:
        return text.strip()

# --- הלוגיקה הראשית של האפליקציה (עם התיקון) ---

def main():
    print("--- ברוכים הבאים לצ'אטבוט של Google Sheets ---")

    try:
        with open(SERVICE_ACCOUNT_FILE, 'r') as f:
            creds_data = json.load(f)
            service_email = creds_data.get('client_email')
            print(f"\nודא ששיתפת את הגיליון שלך עם כתובת המייל: \n{service_email}\nעם הרשאות 'Editor'.\n")
    except FileNotFoundError:
        print(f"שגיאה: הקובץ '{SERVICE_ACCOUNT_FILE}' לא נמצא. ודא שהוא נמצא באותה תיקייה.")
        return
        
    spreadsheet_id = input("אנא הזן את מזהה ה-Google Sheet שלך (Spreadsheet ID): ").strip()
    if not spreadsheet_id:
        print("לא הוזן ID. יוצא מהתוכנית.")
        return
        
    print("\nהצ'אטבוט מוכן! הקלד את בקשתך בעברית.")
    print("כדי לצאת, הקלד 'צא' או 'exit'.")
    print("-" * 50)

    while True:
        user_input = input("אני מקשיב >> ")
        if user_input.lower() in ['צא', 'exit']:
            break
            
        print("חושב...")
        # 1. תרגם את הבקשה לקוד
        raw_api_call = translate_to_api_call(user_input, spreadsheet_id)
        
        if raw_api_call.startswith("Error"):
            print(f"שגיאה בתקשורת עם OpenAI: {raw_api_call}")
            continue
            
        # 2. <-- התיקון! נקה את התשובה לפני השימוש
        api_call_str = extract_code_from_markdown(raw_api_call)
        
        print(f"הפקודה שנוצרה (לאחר ניקוי): {api_call_str}")
        
        # 3. הרץ את הקוד הנקי
        try:
            result = eval(api_call_str)
            print("\n--- תוצאה ---")
            if isinstance(result, list):
                for row in result:
                    print('\t'.join(map(str, row)))
            else:
                print(result)
            print("---------------\n")

        except Exception as e:
            print(f"\n!!! אירעה שגיאה בביצוע הפעולה: {e}")
            print("ייתכן שהבקשה לא הייתה ברורה מספיק או שהייתה שגיאה בנתונים.\n")

    print("תודה שהשתמשת בצ'אטבוט. להתראות!")

if __name__ == '__main__':
    main()