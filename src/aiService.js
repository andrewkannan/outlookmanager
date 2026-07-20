import { GoogleGenAI } from '@google/genai';

const ai = new GoogleGenAI({ apiKey: import.meta.env.VITE_GEMINI_API_KEY || '' });

export const DEFAULT_AI_RULES = {
    'Pending Task': 'ONLY assign here if the email explicitly requests an action from "Andrew", "Dibin", "IT", or "IT team". EXCEPTION: If the email is simply confirming something is done (e.g., "verified and updated"), it is NOT a pending task.',
    'To Reply': 'Emails that clearly require a response.',
    'Done': '1) Purely informational/spam. 2) COMPLETION CONFIRMATIONS: Any email saying "verified and updated", "done", "completed", or "noted" MUST be assigned to "Done", even if it mentions Andrew or Dibin.',
    'Inbox': 'Use this ONLY if you are completely unsure or it doesn\'t fit the others.',
    'System Logs': 'Routine automated system logs, successful backups, standard notifications, etc. that do NOT indicate a problem.',
    'System Alerts': 'Automated system emails that indicate an ERROR, FAILURE, ABNORMAL behavior, or require attention.'
};

export const categorizeEmailsAI = async (emails, activeColumns, customRules = {}) => {
    // Only process emails that we need to categorize to save tokens
    const emailsToProcess = emails.map(e => ({
        id: e.id,
        subject: e.subject,
        from: e.from?.emailAddress?.name || e.from?.emailAddress?.address,
        bodyPreview: e.bodyPreview
    }));

    if (emailsToProcess.length === 0) return {};

    // Build the dynamic categories string
    const categoriesString = activeColumns.map(col => {
        const rule = customRules[col.id] || DEFAULT_AI_RULES[col.id] || `Emails relating to ${col.title}.`;
        return `- "${col.id}": ${rule}`;
    }).join('\n    ');

    const prompt = `
    You are an intelligent email assistant organizing a user's Outlook inbox. 
    Review the following list of emails (provided as JSON) and categorize each one into EXACTLY ONE of the following categories based on the subject and body preview:
    
    Categories:
    ${categoriesString}

    Emails:
    ${JSON.stringify(emailsToProcess)}

    If an email explicitly mentions a date for an action or task (e.g. "wef 01 Aug", "by Friday", "due on Oct 5"), extract that date as a "dueDate" in YYYY-MM-DD format.

    Return ONLY a valid JSON object mapping the email ID to an object containing the "category" string, and an optional "dueDate" string. For example:
    {
      "email_id_1": { "category": "${activeColumns[0]?.id || "Category"}", "dueDate": "2026-08-01" },
      "email_id_2": { "category": "Done" }
    }
    DO NOT wrap the response in markdown code blocks (\`\`\`json). Just return the raw JSON object.
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-3.5-flash',
            contents: prompt,
        });
        
        const rawText = response.text;
        const cleanedText = rawText.replace(/```json/g, '').replace(/```/g, '').trim();
        return {
            categories: JSON.parse(cleanedText),
            tokens: response.usageMetadata?.totalTokenCount || 0
        };
    } catch (error) {
        console.error("AI Categorization Error:", error);
        return { categories: {}, tokens: 0 }; // Fail gracefully
    }
};

export const analyzeMailboxPatterns = async (emails) => {
    const emailsToProcess = emails.map(e => ({
        subject: e.subject,
        from: e.from?.emailAddress?.name,
        preview: e.bodyPreview
    }));

    if (emailsToProcess.length === 0) return [];

    const prompt = `
    You are an intelligent data analyst reviewing a large batch of emails to find hidden patterns.
    Review the following ${emailsToProcess.length} emails. 
    Identify the top 3 to 5 most common "themes" or "patterns" of emails that the user receives (e.g., "Server Alerts", "Vendor Invoices", "Client Queries").
    
    Emails:
    ${JSON.stringify(emailsToProcess)}
    
    Return ONLY a valid JSON array of strings representing these new suggested categories. Example:
    ["Server Alerts", "Vendor Invoices", "Weekly Reports"]
    DO NOT wrap the response in markdown code blocks (\`\`\`json). Just return the raw JSON array.
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-3.5-flash',
            contents: prompt,
        });
        
        const rawText = response.text;
        const match = rawText.match(/\[[\s\S]*\]/);
        let patterns = ["General Conversations", "Automated Alerts", "Task Requests"]; // Fallback
        
        if (match) {
            try {
                patterns = JSON.parse(match[0]);
            } catch (e) {
                console.warn("Failed to parse matched JSON");
            }
        } else {
            console.warn("Could not find JSON array in response:", rawText);
        }
        
        return {
            patterns,
            tokens: response.usageMetadata?.totalTokenCount || 0
        };
    } catch (error) {
        console.error("Deep Scan Error:", error);
        return { patterns: ["Error analyzing emails."], tokens: 0 };
    }
};

// Analyze if a user's sent reply expects a response
export const analyzeUserReply = async (emailBodyPreview) => {
    try {
        const prompt = `
        You are an intelligent email analyzer. The user just sent the following reply to an email thread:
        "${emailBodyPreview}"
        
        Did the user ask a question or explicitly request a reply/action from the other person?
        If yes, they are waiting for a response.
        If no, they are just closing the loop, saying thanks, or providing information without needing a reply.
        
        Respond with ONLY the exact word "WAITING" if they are waiting for a reply.
        Respond with ONLY the exact word "DONE" if they are not.
        `;
        
        const response = await ai.models.generateContent({
            model: 'gemini-3.5-flash',
            contents: prompt,
        });
        
        const text = response.text.trim().toUpperCase();
        return {
            status: text.includes('WAITING') ? 'Waiting on Reply' : 'Done',
            tokens: response.usageMetadata?.totalTokenCount || 0
        };
    } catch (error) {
        console.error("Analyze Reply Error:", error);
        return { status: 'Done', tokens: 0 };
    }
};

export const draftGhostwriterReply = async (conversationMessages, customInstructions) => {
    try {
        let context = conversationMessages.map(msg => 
            `From: ${msg.from?.emailAddress?.name || msg.from?.emailAddress?.address}\nDate: ${msg.receivedDateTime}\nSubject: ${msg.subject}\nBody: ${msg.bodyPreview}`
        ).join('\n\n---\n\n');

        let prompt = `You are an AI assistant helping a user reply to an email thread.
Here is the email thread history:

${context}

Please draft a professional, polite reply to the most recent email in this thread.
`;
        if (customInstructions && customInstructions.trim() !== '') {
            prompt += `\nThe user has provided these specific instructions for the reply: "${customInstructions}"\nPlease incorporate these instructions into your draft.`;
        } else {
            prompt += `\nIf the user is asked to complete a task, you can assume it has been completed and write a short confirmation.`;
        }

        prompt += `\n\nReturn ONLY the body of the email draft as plain text. Do not include subject lines or extra commentary. Sign off generically as "Regards, [Your Name]".`;

        const response = await ai.models.generateContent({
            model: 'gemini-3.5-flash',
            contents: prompt
        });
        
        return {
            draft: response.text.trim(),
            tokens: response.usageMetadata?.totalTokenCount || 0
        };
    } catch (error) {
        console.error("Ghostwriter Error:", error);
        throw error;
    }
};

export const draftFollowUpEmail = async (taskSubject, taskDescription, customInstructions) => {
    try {
        let prompt = `You are an AI assistant helping a user draft a follow-up email based on a manual task they created.
Task Title: ${taskSubject}
Task Description: ${taskDescription || "None"}

Please draft a professional, polite follow-up email to address this task.
`;
        if (customInstructions && customInstructions.trim() !== '') {
            prompt += `\nThe user has provided these specific instructions for the email: "${customInstructions}"\nPlease incorporate these instructions into your draft.`;
        }
        
        prompt += `\n\nReturn ONLY the body of the email draft as plain text. Do not include subject lines or extra commentary. Sign off generically as "Regards, [Your Name]".`;

        const response = await ai.models.generateContent({
            model: 'gemini-3.5-flash',
            contents: prompt
        });
        
        return {
            draft: response.text.trim(),
            tokens: response.usageMetadata?.totalTokenCount || 0
        };
    } catch (error) {
        console.error("Follow-Up Draft Error:", error);
        throw error;
    }
};
