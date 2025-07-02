import { useState, useEffect } from 'react';
import styles from '../styles/Chat.module.css';

export default function ChatInterface() {
  const [messages, setMessages] = useState([
    {
      type: 'bot',
      content: 'שלום! אני כאן כדי לעזור לך לנהל את ה-Google Sheet שלך. אנא הדבק את מזהה הגיליון והתחל לתת לי הוראות.'
    }
  ]);
  const [inputMessage, setInputMessage] = useState('');
  const [spreadsheetId, setSpreadsheetId] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  // טען מזהה גיליון שמור
  useEffect(() => {
    const saved = localStorage.getItem('googlesheets-spreadsheet-id');
    if (saved) setSpreadsheetId(saved);
  }, []);

  // שמור מזהה גיליון
  useEffect(() => {
    if (spreadsheetId) {
      localStorage.setItem('googlesheets-spreadsheet-id', spreadsheetId);
    } else {
      localStorage.removeItem('googlesheets-spreadsheet-id');
    }
  }, [spreadsheetId]);

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!inputMessage.trim() || !spreadsheetId.trim() || isLoading) return;

    const userMessage = { type: 'user', content: inputMessage };
    setMessages(prev => [...prev, userMessage]);
    setInputMessage('');
    setIsLoading(true);

    try {
      const response = await fetch('/.netlify/functions/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: inputMessage,
          spreadsheet_id: spreadsheetId
        })
      });

      const data = await response.json();
      const botMessage = { type: 'bot', content: data.reply };
      setMessages(prev => [...prev, botMessage]);
    } catch (error) {
      const errorMessage = { type: 'bot', content: 'אירעה שגיאת תקשורת עם השרת.' };
      setMessages(prev => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <header className={styles.header}>
        <h1>צ'אטבוט לניהול Google Sheets</h1>
        <div className={styles.spreadsheetInput}>
          <label>מזהה גיליון (Spreadsheet ID):</label>
          <input
            type="text"
            value={spreadsheetId}
            onChange={(e) => setSpreadsheetId(e.target.value)}
            placeholder="הדבק כאן את מזהה הגיליון..."
            className={styles.idInput}
          />
        </div>
      </header>

      <div className={styles.chatWindow}>
        {messages.map((msg, index) => (
          <div key={index} className={`${styles.message} ${styles[msg.type + 'Message']}`}>
            <div 
              className={styles.messageContent}
              dangerouslySetInnerHTML={{ __html: msg.content }}
            />
          </div>
        ))}
        {isLoading && (
          <div className={`${styles.message} ${styles.botMessage}`}>
            <div className={styles.messageContent}>חושב...</div>
          </div>
        )}
      </div>

      <form onSubmit={handleSubmit} className={styles.chatForm}>
        <input
          type="text"
          value={inputMessage}
          onChange={(e) => setInputMessage(e.target.value)}
          placeholder={spreadsheetId ? "כתוב את בקשתך כאן..." : "יש להזין מזהה גיליון תחילה"}
          disabled={!spreadsheetId || isLoading}
          className={styles.messageInput}
        />
        <button 
          type="submit" 
          disabled={!spreadsheetId || isLoading}
          className={styles.sendButton}
        >
          שלח
        </button>
      </form>
    </div>
  );
} 