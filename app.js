let apiKey = '';

// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Excel add-in loaded");
        loadSavedApiKey();
    }
});

// Load saved API key from localStorage
function loadSavedApiKey() {
    const savedKey = localStorage.getItem('deepseekApiKey');
    if (savedKey) {
        apiKey = savedKey;
        document.getElementById('apiKey').value = '••••••••••••••••';
    }
}

// Save API key
function saveApiKey() {
    const keyInput = document.getElementById('apiKey');
    apiKey = keyInput.value;
    localStorage.setItem('deepseekApiKey', apiKey);
    showStatus('API Key saved successfully!', 'success');
}

// Call DeepSeek API
async function callDeepSeekAPI(messages, temperature = 0.7) {
    if (!apiKey) {
        showStatus('Please enter your API key first', 'error');
        return null;
    }

    try {
        const response = await fetch('https://api.deepseek.com/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: "deepseek-chat",
                messages: messages,
                temperature: temperature,
                max_tokens: 2000
            })
        });

        if (!response.ok) {
            throw new Error(`API Error: ${response.status}`);
        }

        const data = await response.json();
        return data.choices[0].message.content;
    } catch (error) {
        console.error('Error calling DeepSeek API:', error);
        showStatus(`Error: ${error.message}`, 'error');
        return null;
    }
}

// Send message to DeepSeek
async function sendMessage() {
    const userInput = document.getElementById('userInput').value;
    if (!userInput.trim()) return;

    const chatOutput = document.getElementById('chatOutput');
    chatOutput.innerHTML += `<div class="user-message">You: ${userInput}</div>`;
    
    showStatus('Thinking...', 'info');
    
    const response = await callDeepSeekAPI([
        { role: "system", content: "You are an Excel AI assistant. Help users with data analysis, formulas, and Excel tasks." },
        { role: "user", content: userInput }
    ]);
    
    if (response) {
        chatOutput.innerHTML += `<div class="ai-message">DeepSeek: ${response}</div>`;
        showStatus('Response received', 'success');
    }
    
    document.getElementById('userInput').value = '';
    chatOutput.scrollTop = chatOutput.scrollHeight;
}

// Analyze selected data
async function analyzeSelectedData() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('values');
            await context.sync();
            
            const data = range.values;
            const dataString = JSON.stringify(data);
            
            showStatus('Analyzing data...', 'info');
            
            const analysis = await callDeepSeekAPI([
                { role: "system", content: "You are a data analyst. Analyze the following Excel data and provide insights:" },
                { role: "user", content: `Analyze this Excel data: ${dataString}. Provide key insights, patterns, and suggestions.` }
            ]);
            
            if (analysis) {
                document.getElementById('chatOutput').innerHTML += 
                    `<div class="ai-message">Data Analysis:<br>${analysis}</div>`;
            }
        });
    } catch (error) {
        showStatus(`Error: ${error.message}`, 'error');
    }
}

// Generate formula based on description
async function generateFormula() {
    const userInput = document.getElementById('userInput').value || 
        prompt("Describe what you want the formula to do:");
    
    if (!userInput) return;
    
    const formula = await callDeepSeekAPI([
        { role: "system", content: "You are an Excel expert. Generate Excel formulas based on user requests. Always explain the formula and provide an example." },
        { role: "user", content: `Create an Excel formula for: ${userInput}` }
    ]);
    
    if (formula) {
        document.getElementById('chatOutput').innerHTML += 
            `<div class="ai-message">Formula Suggestion:<br>${formula}</div>`;
        
        // Copy to clipboard
        navigator.clipboard.writeText(formula);
        showStatus('Formula copied to clipboard!', 'success');
    }
}

// Summarize data
async function summarizeData() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load('values');
            await context.sync();
            
            const data = usedRange.values;
            const summary = await callDeepSeekAPI([
                { role: "system", content: "Summarize the key points from this Excel data in a concise way:" },
                { role: "user", content: `Summarize: ${JSON.stringify(data.slice(0, 10))}` }
            ]);
            
            if (summary) {
                document.getElementById('chatOutput').innerHTML += 
                    `<div class="ai-message">Data Summary:<br>${summary}</div>`;
            }
        });
    } catch (error) {
        showStatus(`Error: ${error.message}`, 'error');
    }
}

// Show status messages
function showStatus(message, type = 'info') {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    
    setTimeout(() => {
        statusDiv.textContent = '';
        statusDiv.className = 'status';
    }, 5000);
}