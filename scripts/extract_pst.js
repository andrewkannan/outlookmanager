import { PSTFile } from 'pst-extractor';
import fs from 'fs';

const startDate = new Date('2026-03-01T00:00:00Z');
const pstPath = 'C:\\Outlook\\Exchange\\andrewkannan@asiamedic.com.sg.001.pst';
const outputPath = 'C:\\Outlook\\public\\historical_emails.json';
const systemLogKeywords = ['hl7', 'it-ops@asiamedic.com.sg'];

try {
    console.log(`Opening PST file: ${pstPath}`);
    const pst = new PSTFile(pstPath);
    const rootFolder = pst.getRootFolder();
    
    let extractedEmails = [];
    
    function processFolder(folder) {
        if (folder.hasSubfolders) {
            const childFolders = folder.getSubFolders();
            for (let childFolder of childFolders) {
                processFolder(childFolder);
            }
        }
        
        if (folder.contentCount > 0) {
            let email = folder.getNextChild();
            while (email != null) {
                const receivedDate = email.clientSubmitTime || email.creationTime;
                
                if (receivedDate && receivedDate >= startDate) {
                    const subject = email.subject || "(No Subject)";
                    const rawBody = email.body || email.bodyHTML || "";
                    // Strip HTML tags for preview and clean up whitespace
                    const bodyPreview = rawBody.replace(/<[^>]*>?/gm, ' ').replace(/\s+/g, ' ').trim();
                    const senderName = email.senderName || "";
                    const senderEmail = email.senderEmailAddress || "";
                    
                    const isSystemLog = systemLogKeywords.some(kw => 
                        subject.toLowerCase().includes(kw) || senderEmail.toLowerCase().includes(kw)
                    );

                    let convId = email.descriptorNodeId.toString();
                    if (email.conversationId) {
                        convId = Buffer.isBuffer(email.conversationId) ? email.conversationId.toString('hex') : String(email.conversationId);
                    }
                    
                    const lowerSub = subject.toLowerCase();
                    const lowerBody = bodyPreview.toLowerCase();
                    const isPendingTask = lowerSub.includes('account creation') || lowerBody.includes('account creation') ||
                                          lowerSub.includes('account renewal') || lowerBody.includes('account renewal');

                    if (isPendingTask && !isSystemLog) {
                        extractedEmails.push({
                            id: 'historical_' + email.descriptorNodeId.toString(),
                            conversationId: convId,
                            subject: subject,
                            bodyPreview: bodyPreview.substring(0, 500),
                            body: {
                                contentType: 'text',
                                content: bodyPreview
                            },
                            from: {
                                emailAddress: {
                                    name: senderName,
                                    address: senderEmail
                                }
                            },
                            receivedDateTime: receivedDate.toISOString(),
                            importance: email.importance === 2 ? 'high' : 'normal',
                            categories: ['Pending Task'],
                            isSystemLog: false,
                            isHistorical: true
                        });
                    }
                }
                email = folder.getNextChild();
            }
        }
    }

    processFolder(rootFolder);
    
    // Sort oldest to newest
    extractedEmails.sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime));

    console.log(`Extracted ${extractedEmails.length} emails since ${startDate.toISOString()}`);
    fs.writeFileSync(outputPath, JSON.stringify(extractedEmails, null, 2));
    console.log(`Saved to ${outputPath}`);
    
} catch (e) {
    console.error(e);
}
