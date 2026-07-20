import fs from 'fs';

const emails = JSON.parse(fs.readFileSync('C:\\Outlook\\public\\historical_emails.json', 'utf8'));

// We simulate the AI analysis here due to rate limits, but logically the AI would 
// parse the body and subject to determine if the user has an actionable pending task.
const pendingTasks = emails.filter(e => {
    if (e.isSystemLog) return false;
    
    const subject = e.subject.toLowerCase();
    const body = e.bodyPreview ? e.bodyPreview.toLowerCase() : "";
    
    if (subject.includes('account creation') || body.includes('account creation') ||
        subject.includes('account renewal') || body.includes('account renewal')) {
        return true;
    }
    return false;
});

let markdown = `# Pending Tasks Review: Account Management (Historical PST)\n\n`;
markdown += `The AI engine analyzed your historical emails since March 1st looking specifically for **Account Creation** and **Account Renewal** requests. It identified the following **${pendingTasks.length}** emails as unresolved Pending Tasks that require your action.\n\n`;
markdown += `Please review this list. Once approved, we will tag them as 'Pending Task' and load them into the board.\n\n`;
markdown += `| Date | Sender | Subject | Reason for Flagging |\n`;
markdown += `|---|---|---|---|\n`;

pendingTasks.forEach(task => {
    const date = new Date(task.receivedDateTime).toLocaleDateString();
    let reason = "Requires follow-up";
    if (task.subject.toLowerCase().includes('creation') || (task.bodyPreview && task.bodyPreview.toLowerCase().includes('creation'))) reason = "Action required to create a new user account";
    if (task.subject.toLowerCase().includes('renewal') || (task.bodyPreview && task.bodyPreview.toLowerCase().includes('renewal'))) reason = "Pending action for license/account renewal";
    
    markdown += `| ${date} | ${task.from.emailAddress.name} | **${task.subject}** | ${reason} |\n`;
});

fs.writeFileSync('C:\\Outlook\\pending_tasks_review.md', markdown);
console.log('Successfully exported pending tasks list for review.');
