document.addEventListener('DOMContentLoaded', () => {
    const chatWindow = document.getElementById('chat-window');
    const chatForm = document.getElementById('chat-form');
    const messageInput = document.getElementById('message-input');
    const spreadsheetIdInput = document.getElementById('spreadsheet-id');
    const sendButton = chatForm.querySelector('button');

    // Enable/disable chat based on spreadsheet ID
    spreadsheetIdInput.addEventListener('input', () => {
        const spreadsheetId = spreadsheetIdInput.value.trim();
        const hasId = spreadsheetId !== '';
        
        // Save to localStorage
        if (hasId) {
            localStorage.setItem('googlesheets-spreadsheet-id', spreadsheetId);
        } else {
            localStorage.removeItem('googlesheets-spreadsheet-id');
        }
        
        messageInput.disabled = !hasId;
        sendButton.disabled = !hasId;
        if (hasId) {
            messageInput.placeholder = "כתוב את בקשתך כאן...";
        } else {
            messageInput.placeholder = "יש להזין מזהה גיליון תחילה";
        }
    });

    // Load saved spreadsheet ID from localStorage (after event listener is attached)
    const savedSpreadsheetId = localStorage.getItem('googlesheets-spreadsheet-id');
    if (savedSpreadsheetId) {
        spreadsheetIdInput.value = savedSpreadsheetId;
        // Trigger input event to enable chat
        spreadsheetIdInput.dispatchEvent(new Event('input'));
    }

    chatForm.addEventListener('submit', async (event) => {
        event.preventDefault();

        const userMessage = messageInput.value.trim();
        const spreadsheetId = spreadsheetIdInput.value.trim();

        if (!userMessage || !spreadsheetId) {
            return;
        }

        // Display user's message
        appendMessage(userMessage, 'user');
        messageInput.value = '';

        // Display "thinking" message from bot
        const thinkingMessage = appendMessage('חושב...', 'bot thinking');

        try {
            // Send message to Flask server
            const response = await fetch('/chat', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    message: userMessage,
                    spreadsheet_id: spreadsheetId,
                }),
            });

            const data = await response.json();
            
            // Replace "thinking" message with the actual reply
            thinkingMessage.innerHTML = data.reply;
            thinkingMessage.parentElement.classList.remove('thinking');

        } catch (error) {
            console.error('Error:', error);
            thinkingMessage.innerHTML = 'אירעה שגיאת תקשורת עם השרת.';
            thinkingMessage.parentElement.classList.remove('thinking');
        }
    });

    function appendMessage(content, type) {
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${type}-message`;
        
        const contentDiv = document.createElement('div');
        contentDiv.className = 'message-content';
        contentDiv.innerHTML = content; // Use innerHTML to render HTML tags from bot

        messageDiv.appendChild(contentDiv);
        chatWindow.appendChild(messageDiv);
        chatWindow.scrollTop = chatWindow.scrollHeight; // Auto-scroll to bottom

        return contentDiv; // Return the content element for modification
    }
});