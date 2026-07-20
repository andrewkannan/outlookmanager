import fs from 'fs';

const emails = JSON.parse(fs.readFileSync('C:\\Outlook\\public\\historical_emails.json', 'utf8'));

emails.forEach(e => {
    if (e.isSystemLog) return;
    
    const subject = e.subject.toLowerCase();
    const body = e.bodyPreview ? e.bodyPreview.toLowerCase() : "";
    
    // Identify our approved "Pending Tasks" (Account Creation / Renewal)
    if (subject.includes('account creation') || body.includes('account creation') ||
        subject.includes('account renewal') || body.includes('account renewal')) {
        e.categories = ['Pending Task'];
    } 
    // Fallback categories for other historical emails
    else if (subject.includes('vendor') || subject.includes('quote') || subject.includes('invoice') || subject.includes('hardware') || subject.includes('negotiation')) {
        e.categories = ['Vendor Quotes & Procurement'];
    } else if (subject.includes('onboarding') || subject.includes('hire') || subject.includes('provision')) {
        e.categories = ['HR Onboarding'];
    } else {
        e.categories = ['General IT Inquiries'];
    }
});

fs.writeFileSync('C:\\Outlook\\public\\historical_emails.json', JSON.stringify(emails, null, 2));
console.log('Historical emails categorized and Pending Tasks injected successfully.');
