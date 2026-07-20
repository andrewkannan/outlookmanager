import fs from 'fs';

const historicalEmails = [
    {
        id: "hist_1",
        conversationId: "conv_1",
        subject: "Q1 Vendor Negotiation - Dell Hardware",
        bodyPreview: "Hi Andrew, attached is the Q1 hardware quote from Dell. Let's discuss pricing.",
        from: { emailAddress: { name: "Dell Rep", address: "sales@dell.com" } },
        receivedDateTime: "2026-03-05T10:00:00Z",
        importance: "normal",
        categories: [],
        isSystemLog: false,
        isHistorical: true
    },
    {
        id: "hist_2",
        conversationId: "conv_2",
        subject: "Monthly Hardware Report - March",
        bodyPreview: "Here is the monthly hardware status report for all clinics.",
        from: { emailAddress: { name: "IT Ops", address: "it-ops@asiamedic.com.sg" } },
        receivedDateTime: "2026-03-31T23:59:00Z",
        importance: "normal",
        categories: ["System Logs"],
        isSystemLog: true,
        isHistorical: true
    },
    {
        id: "hist_3",
        conversationId: "conv_3",
        subject: "New Hire Onboarding: Dr. Sarah Lee",
        bodyPreview: "Please provision a laptop and create accounts for Dr. Sarah Lee.",
        from: { emailAddress: { name: "HR", address: "hr@asiamedic.com.sg" } },
        receivedDateTime: "2026-04-12T09:00:00Z",
        importance: "high",
        categories: [],
        isSystemLog: false,
        isHistorical: true
    },
    {
        id: "hist_4",
        conversationId: "conv_4",
        subject: "AWS Invoice - April",
        bodyPreview: "Your AWS invoice for April is ready. Total: $4,500.",
        from: { emailAddress: { name: "AWS Billing", address: "no-reply-aws@amazon.com" } },
        receivedDateTime: "2026-05-01T08:00:00Z",
        importance: "normal",
        categories: [],
        isSystemLog: false,
        isHistorical: true
    },
    {
        id: "hist_5",
        conversationId: "conv_5",
        subject: "HL7 Interface Down - Clinic A",
        bodyPreview: "CRITICAL: The HL7 interface at Clinic A is failing to sync patient records.",
        from: { emailAddress: { name: "System Monitor", address: "alerts@asiamedic.com.sg" } },
        receivedDateTime: "2026-05-15T14:30:00Z",
        importance: "high",
        categories: ["System Alerts"],
        isSystemLog: true,
        isHistorical: true
    },
    {
        id: "hist_6",
        conversationId: "conv_6",
        subject: "Vendor Quote: Cisco Switches",
        bodyPreview: "Andrew, attached is the revised quote for the Cisco access switches.",
        from: { emailAddress: { name: "Cisco Sales", address: "sales@cisco.com" } },
        receivedDateTime: "2026-06-02T11:00:00Z",
        importance: "normal",
        categories: [],
        isSystemLog: false,
        isHistorical: true
    }
];

// Add 50 more dummy emails to simulate bulk
for (let i = 7; i <= 106; i++) {
    const isCreation = i % 3 === 0;
    const isRenewal = i % 5 === 0;
    
    let subject = `General IT Inquiry ${i}`;
    let bodyPreview = "This is a historical email from the PST archive.";
    
    if (isCreation) {
        subject = `Account Creation Request for Staff ${i}`;
        bodyPreview = "Please proceed with account creation for the new staff member.";
    } else if (isRenewal) {
        subject = `Action Required: Account Renewal ${i}`;
        bodyPreview = "The software license and account renewal is due soon. Please process the account renewal.";
    }

    historicalEmails.push({
        id: `hist_${i}`,
        conversationId: `conv_${i}`,
        subject: subject,
        bodyPreview: bodyPreview,
        from: { emailAddress: { name: "Internal Staff", address: "staff@asiamedic.com.sg" } },
        receivedDateTime: new Date(Date.now() - Math.random() * 10000000000).toISOString(),
        importance: "normal",
        categories: [],
        isSystemLog: false,
        isHistorical: true
    });
}

fs.writeFileSync('C:\\Outlook\\public\\historical_emails.json', JSON.stringify(historicalEmails, null, 2));
console.log("Mock PST extraction complete. Wrote " + historicalEmails.length + " emails to historical_emails.json");
