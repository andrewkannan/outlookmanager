import fs from 'fs';
import { GoogleGenAI } from '@google/genai';
const apiKey = '';
const genAI = new GoogleGenAI({ apiKey: apiKey });

const emails = JSON.parse(fs.readFileSync('C:\\Outlook\\public\\historical_emails.json', 'utf8'));

// Only pass non-system logs to the AI as requested
const analyzableEmails = emails.filter(e => !e.isSystemLog);

async function run() {
    console.log(`Analyzing ${analyzableEmails.length} emails for pattern detection...`);
    
    const subjects = analyzableEmails.map(e => e.subject);
    
    const prompt = `
    You are an expert Kanban board organizer for an IT Operations Manager.
    Analyze the following list of email subjects from a historical PST file:
    ${JSON.stringify(subjects, null, 2)}
    
    Based on the recurring patterns in these subjects, suggest 2 or 3 NEW Kanban columns that would be highly useful for categorizing these emails.
    Do NOT suggest "System Logs", "Pending Task", or "Inbox" as those already exist.
    Return ONLY a JSON array of objects with 'title' (e.g. "Vendor Quotes", "HR Onboarding"), 'color' (a hex color), and 'reason' (why you suggest this based on the emails).
    Do not include markdown blocks, just the raw JSON array.
    `;
    
    try {
        const response = await genAI.models.generateContent({
            model: 'gemini-2.0-flash',
            contents: prompt,
            config: {
                temperature: 0.1,
                responseMimeType: 'application/json'
            }
        });
        
        let text = response.text;
        const suggestions = JSON.parse(text);
        
        fs.writeFileSync('C:\\Outlook\\public\\pst_analysis_results.json', JSON.stringify(suggestions, null, 2));
        console.log("Analysis complete. Suggested columns saved.");
    } catch (e) {
        console.error(e);
    }
}

run();
