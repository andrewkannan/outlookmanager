import React, { useEffect, useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from '../authConfig';
import { getEmails, updateEmailCategory, getDeepEmails, getConversation, sendEmailReply, sendNewEmail } from '../graphApi';
import { categorizeEmailsAI, analyzeMailboxPatterns, analyzeUserReply, DEFAULT_AI_RULES, draftGhostwriterReply, draftFollowUpEmail } from '../aiService';
import { DragDropContext, Droppable, Draggable } from '@hello-pangea/dnd';

const COLUMNS = {
    INBOX: { id: 'Inbox', title: 'Inbox', color: 'var(--status-inbox)' },
    TASK: { id: 'Pending Task', title: 'Pending Task', color: 'var(--status-task)' },
    DONE: { id: 'Done', title: 'Done', color: 'var(--status-done)' },
    WAITING: { id: 'Waiting on Reply', title: 'Waiting on Reply', color: '#f59e0b' },
    ALERTS: { id: 'System Alerts', title: 'System Alerts', color: '#dc2626' },
    LOGS: { id: 'System Logs', title: 'System Logs', color: '#64748b' }
};

const TASK_SUBCATEGORIES = {
    'Account Creation': '#3b82f6',
    'Renewals': '#f59e0b',
    'Vendor Quotes & Procurement': '#8b5cf6',
    'HR Onboarding': '#14b8a6',
    'General IT Inquiries': '#6366f1'
};

export const KanbanBoard = () => {
    const { instance, accounts } = useMsal();
    const [emails, setEmails] = useState([]);
    const [loading, setLoading] = useState(true);
    const [isCategorizing, setIsCategorizing] = useState(false);
    const [isScanning, setIsScanning] = useState(false);
    const [autoPilotActive, setAutoPilotActive] = useState(true);
    const isProcessingRef = React.useRef(false);
    const [scanResults, setScanResults] = useState(null);
    const [showLogs, setShowLogs] = useState(false);
    const [showSettings, setShowSettings] = useState(false);
    const [selectedConversation, setSelectedConversation] = useState(null);
    const [conversationMessages, setConversationMessages] = useState([]);
    
    // Teach AI States
    const [showTeachModal, setShowTeachModal] = useState(false);
    const [teachEmail, setTeachEmail] = useState(null);
    const [teachCondition, setTeachCondition] = useState("");
    const [teachTarget, setTeachTarget] = useState("");
    const [showSystemLogs, setShowSystemLogs] = useState(false);
    const [showDraftModal, setShowDraftModal] = useState(false);
    const [draftContext, setDraftContext] = useState(null); // 'reply' or 'followup'
    const [draftInstructions, setDraftInstructions] = useState("");
    const [draftContent, setDraftContent] = useState("");
    const [isDrafting, setIsDrafting] = useState(false);
    const [isSending, setIsSending] = useState(false);
    
    // Manual Task States
    const [manualTasks, setManualTasks] = useState(() => {
        const saved = localStorage.getItem('manualTasks');
        return saved ? JSON.parse(saved) : [];
    });
    const [showManualTaskModal, setShowManualTaskModal] = useState(false);
    const [newTaskSubject, setNewTaskSubject] = useState("");
    const [newTaskBody, setNewTaskBody] = useState("");
    const [newTaskTargetCol, setNewTaskTargetCol] = useState("");
    const [newTaskDueDate, setNewTaskDueDate] = useState("");
    const [pendingFilter, setPendingFilter] = useState('All');
    const [searchQuery, setSearchQuery] = useState("");

    const [aiLogs, setAiLogs] = useState(() => {
        const saved = localStorage.getItem('aiLogs');
        return saved ? JSON.parse(saved) : [];
    });
    const [activeColumns, setActiveColumns] = useState(() => {
        const saved = localStorage.getItem('activeColumns');
        let parsed = saved ? JSON.parse(saved) : Object.values(COLUMNS);
        
        // Remove old subcategories if they exist from a legacy save
        parsed = parsed.filter(c => c.id !== 'To Reply' && !Object.keys(TASK_SUBCATEGORIES).includes(c.id));
        
        if (!parsed.find(c => c.id === 'System Alerts')) {
            parsed.push(COLUMNS.ALERTS);
        }
        if (!parsed.find(c => c.id === 'Waiting on Reply')) {
            parsed.push(COLUMNS.WAITING);
        }

        // Hide System Log from board
        parsed = parsed.filter(c => c.id !== COLUMNS.LOGS.id);
        
        return parsed;
    });
    const [manualCollapsed, setManualCollapsed] = useState(() => {
        const saved = localStorage.getItem('manualCollapsed');
        return saved ? JSON.parse(saved) : {};
    });

    const [customAiRules, setCustomAiRules] = useState(() => {
        const saved = localStorage.getItem('customAiRules');
        return saved ? JSON.parse(saved) : {};
    });
    const [showRulesModal, setShowRulesModal] = useState(false);

    useEffect(() => {
        localStorage.setItem('manualCollapsed', JSON.stringify(manualCollapsed));
    }, [manualCollapsed]);

    useEffect(() => {
        localStorage.setItem('customAiRules', JSON.stringify(customAiRules));
    }, [customAiRules]);

    const [emailDueDates, setEmailDueDates] = useState(() => {
        const saved = localStorage.getItem('emailDueDates');
        return saved ? JSON.parse(saved) : {};
    });

    useEffect(() => {
        localStorage.setItem('emailDueDates', JSON.stringify(emailDueDates));
    }, [emailDueDates]);

    // Save columns to local storage whenever they change
    useEffect(() => {
        localStorage.setItem('activeColumns', JSON.stringify(activeColumns));
    }, [activeColumns]);

    useEffect(() => {
        localStorage.setItem('aiLogs', JSON.stringify(aiLogs));
    }, [aiLogs]);

    const logAction = (action, tokens, details) => {
        setAiLogs(prev => [{ timestamp: new Date().toISOString(), action, tokens, details }, ...prev].slice(0, 50));
    };

    const processConversations = (fetchedEmails, currentCols) => {
            const conversationsMap = new Map();
            fetchedEmails.forEach(email => {
                if (!conversationsMap.has(email.conversationId)) {
                    conversationsMap.set(email.conversationId, {
                        ...email,
                        latestCategories: email.categories || [],
                        allCategories: email.categories || []
                    });
                } else {
                    const existing = conversationsMap.get(email.conversationId);
                    if (email.categories) existing.allCategories.push(...email.categories);
                }
            });

            return Array.from(conversationsMap.values()).map(conv => {
                let category = COLUMNS.INBOX.id;
                let subCategory = conv.allCategories.find(c => Object.keys(TASK_SUBCATEGORIES).includes(c));
                
                const matchedCategory = conv.latestCategories.find(c => currentCols.some(col => col.id === c) || c === 'To Reply');
                
                if (matchedCategory) {
                    category = matchedCategory === 'To Reply' ? COLUMNS.TASK.id : matchedCategory;
                } else if (subCategory) {
                    category = COLUMNS.TASK.id;
                }
                
                return { ...conv, boardCategory: category, subCategory: subCategory };
            });
    };

    const fetchMails = async () => {
        try {
            // Merge Historical PST Emails first for instant UI loading
            let historicalEmails = [];
            try {
                const histRes = await fetch('/historical_emails.json');
                if (histRes.ok) historicalEmails = await histRes.json();
            } catch(e) { console.error("Could not load historical emails", e); }
            
            const currentCols = JSON.parse(localStorage.getItem('activeColumns')) || Object.values(COLUMNS);
            
            // Render historical tasks immediately so user doesn't get stuck waiting for Auth
            setEmails(processConversations(historicalEmails, currentCols));
            setLoading(false);

            // Fetch live Graph API emails in the background
            try {
                const response = await instance.acquireTokenSilent({
                    ...loginRequest,
                    account: accounts[0]
                });
                const fetchedEmails = await getEmails(response.accessToken);
                const allEmails = [...fetchedEmails, ...historicalEmails];
                setEmails(processConversations(allEmails, currentCols));
            } catch (authErr) {
                console.error("Auth error fetching live emails:", authErr);
                // If silent auth fails for any reason (timeout, interaction required), force a popup to heal the session
                try {
                    const response = await instance.acquireTokenPopup({
                        ...loginRequest,
                        account: accounts[0]
                    });
                    const fetchedEmails = await getEmails(response.accessToken);
                    const allEmails = [...fetchedEmails, ...historicalEmails];
                    setEmails(processConversations(allEmails, currentCols));
                } catch (popupErr) {
                    console.error("Popup auth also failed:", popupErr);
                }
            }
        } catch (error) {
            console.error(error);
            setLoading(false);
        }
    };

    useEffect(() => {
        if (accounts.length > 0) {
            fetchMails();
        }
    }, [accounts]);

    // Background Auto-Pilot Loop
    useEffect(() => {
        if (accounts.length === 0 || !autoPilotActive) return;

        const autoPilotLoop = async () => {
            if (isProcessingRef.current) return;
            isProcessingRef.current = true;

            try {
                // 1. Fetch fresh emails
                const response = await instance.acquireTokenSilent({
                    ...loginRequest,
                    account: accounts[0]
                });
                const fetchedEmails = await getEmails(response.accessToken);
                
                const currentCols = JSON.parse(localStorage.getItem('activeColumns')) || Object.values(COLUMNS);
                const processedEmails = processConversations(fetchedEmails, currentCols);
                
                let updatedEmails = [...processedEmails];
                let hasChanges = false;
                let autoRoutedCount = 0;

                // 2. Check if user replied to pending tasks
                const userEmail = accounts[0].username;
                for (let i = 0; i < updatedEmails.length; i++) {
                    const e = updatedEmails[i];
                    const isSentToSelf = e.toRecipients?.length === 1 && e.toRecipients[0].emailAddress.address === userEmail;
                    if (e.from?.emailAddress?.address === userEmail && !isSentToSelf && e.boardCategory !== 'Done' && e.boardCategory !== 'Waiting on Reply') {
                        // User replied. Analyze the reply to see if they are waiting or done.
                        const { status, tokens } = await analyzeUserReply(e.bodyPreview);
                        e.boardCategory = status;
                        hasChanges = true;
                        updateEmailCategory(response.accessToken, e.id, status).catch(console.error);
                        logAction("Auto-Resolve Reply", tokens, `Analyzed your sent reply. Routed to: ${status}`);
                    }
                }

                // 3. Silently Categorize Inbox emails
                const inboxEmails = updatedEmails.filter(e => e.boardCategory === COLUMNS.INBOX.id);
                if (inboxEmails.length > 0) {
                    const { categories: categoryMap, tokens } = await categorizeEmailsAI(inboxEmails, currentCols, customAiRules);
                    let categorizedCount = 0;
                    let newDueDates = {};
                    
                    updatedEmails = updatedEmails.map(email => {
                        const aiResult = categoryMap[email.id];
                        if (aiResult) {
                            const isObj = typeof aiResult === 'object';
                            const categoryString = isObj ? aiResult.category : aiResult;
                            if (isObj && aiResult.dueDate) {
                                newDueDates[email.id] = aiResult.dueDate;
                            }
                            
                            const targetColumnId = currentCols.find(c => c.id === categoryString)?.id;
                            if (targetColumnId && targetColumnId !== COLUMNS.INBOX.id) {
                                hasChanges = true;
                                categorizedCount++;
                                updateEmailCategory(response.accessToken, email.id, targetColumnId).catch(console.error);
                                return { ...email, boardCategory: targetColumnId };
                            }
                        }
                        return email;
                    });

                    if (Object.keys(newDueDates).length > 0) {
                        setEmailDueDates(prev => ({ ...prev, ...newDueDates }));
                    }

                    if (hasChanges) {
                        setEmails(updatedEmails);
                        logAction("Auto-Pilot Categorization", tokens, `Routed ${categorizedCount} emails from Inbox.`);
                    } else if (tokens > 0) {
                        logAction("Auto-Pilot Scan", tokens, "No emails matched new rules.");
                    }
                }
            } catch (error) {
                console.error("AutoPilot Error:", error);
            }
            
            isProcessingRef.current = false;
        };

        const intervalId = setInterval(autoPilotLoop, 15000); // Check every 15 seconds for snappier experience
        return () => clearInterval(intervalId);
    }, [accounts, instance, autoPilotActive]);

    const handleAiCategorize = async () => {
        setIsCategorizing(true);
        // Find emails currently in the Inbox
        const inboxEmails = emails.filter(e => e.boardCategory === COLUMNS.INBOX.id);
        
        if (inboxEmails.length === 0) {
            setIsCategorizing(false);
            return; // Nothing to categorize
        }

        const { categories: categoryMap, tokens } = await categorizeEmailsAI(inboxEmails, activeColumns, customAiRules);
        
        // Update local state based on AI response
        const newEmails = [...emails];
        let hasChanges = false;
        let count = 0;

        let newDueDates = {};

        newEmails.forEach(email => {
            const aiResult = categoryMap[email.id];
            if (aiResult) {
                const isObj = typeof aiResult === 'object';
                const categoryString = isObj ? aiResult.category : aiResult;
                if (isObj && aiResult.dueDate) {
                    newDueDates[email.id] = aiResult.dueDate;
                }

                const targetColumnId = activeColumns.find(c => c.id === categoryString)?.id;
                if (targetColumnId && targetColumnId !== COLUMNS.INBOX.id && email.boardCategory === COLUMNS.INBOX.id) {
                    email.boardCategory = targetColumnId;
                    hasChanges = true;
                    count++;
                    
                    // Fire off background graph API updates
                    instance.acquireTokenSilent({
                        ...loginRequest,
                        account: accounts[0]
                    }).then(response => {
                        updateEmailCategory(response.accessToken, email.id, targetColumnId);
                    }).catch(console.error);
                }
            }
        });

        if (Object.keys(newDueDates).length > 0) {
            setEmailDueDates(prev => ({ ...prev, ...newDueDates }));
        }

        if (hasChanges) {
            setEmails(newEmails);
            logAction("Manual Categorization", tokens, `Manually categorized ${count} emails.`);
        } else {
            logAction("Manual Categorization", tokens, `Scanned ${inboxEmails.length} emails. No changes made.`);
        }
        setIsCategorizing(false);
    };

    const handleRescanPendingTasks = async () => {
        setIsCategorizing(true);
        const pendingEmails = emails.filter(e => e.boardCategory === COLUMNS.TASK.id);
        
        if (pendingEmails.length === 0) {
            setIsCategorizing(false);
            return;
        }

        const batchSize = 20;
        const newEmails = [...emails];
        let hasChanges = false;
        let totalMoved = 0;
        let totalTokens = 0;
        let newDueDates = {};
        let detailsLog = [];

        try {
            const tokenResponse = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });

            for (let i = 0; i < pendingEmails.length; i += batchSize) {
                const batch = pendingEmails.slice(i, i + batchSize);
                const { categories: categoryMap, tokens } = await categorizeEmailsAI(batch, activeColumns, customAiRules);
                totalTokens += tokens;

                newEmails.forEach(email => {
                    if (!batch.find(b => b.id === email.id)) return;

                    const aiResult = categoryMap[email.id];
                    if (aiResult) {
                        const isObj = typeof aiResult === 'object';
                        const categoryString = isObj ? aiResult.category : aiResult;
                        if (isObj && aiResult.dueDate) {
                            newDueDates[email.id] = aiResult.dueDate;
                        }

                        const targetColumnId = activeColumns.find(c => c.id === categoryString)?.id;
                        if (targetColumnId && targetColumnId !== COLUMNS.TASK.id && email.boardCategory === COLUMNS.TASK.id) {
                            email.boardCategory = targetColumnId;
                            hasChanges = true;
                            totalMoved++;
                            const subjectSnippet = email.subject ? (email.subject.length > 40 ? email.subject.substring(0, 40) + '...' : email.subject) : 'No Subject';
                            detailsLog.push(`- Moved "${subjectSnippet}" to ${targetColumnId}`);
                            
                            updateEmailCategory(tokenResponse.accessToken, email.id, targetColumnId).catch(console.error);
                        }
                    }
                });
            }

            if (Object.keys(newDueDates).length > 0) {
                setEmailDueDates(prev => ({ ...prev, ...newDueDates }));
            }

            if (hasChanges) {
                setEmails(newEmails);
                logAction("Pending Task Rescan", totalTokens, `Categorized ${totalMoved} emails:\n${detailsLog.join('\n')}`);
            } else {
                logAction("Pending Task Rescan", totalTokens, `Scanned ${pendingEmails.length} pending tasks. No changes made.`);
            }
        } catch (error) {
            console.error("Batch Rescan Error:", error);
            logAction("Pending Task Rescan", totalTokens, "Rescan failed or was interrupted.");
        }
        
        setIsCategorizing(false);
    };

    const handleCardClick = async (email) => {
        if (email.isManual) return; // Do not fetch conversation for manual tasks
        
        setSelectedConversation(email);
        setConversationMessages([]);
        try {
            const response = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });
            
            let msgs = [];
            if (email.conversationId) {
                msgs = await getConversation(response.accessToken, email.conversationId);
            }
            if (!msgs || msgs.length === 0) {
                msgs = [email];
            }
            setConversationMessages(msgs);
        } catch (error) {
            console.error("Failed to load conversation", error);
            setConversationMessages([email]);
        }
    };

    const handleSaveTeachRule = async () => {
        if (!teachCondition || !teachTarget || !teachEmail) return;

        // Update custom rules
        const currentRule = customAiRules[teachTarget] || '';
        const newRule = currentRule ? `${currentRule} ALSO: If ${teachCondition}, assign to this category.` : `If ${teachCondition}, assign to this category.`;
        setCustomAiRules(prev => ({
            ...prev,
            [teachTarget]: newRule
        }));

        // Move email locally
        const updatedEmails = emails.map(e => 
            e.id === teachEmail.id ? { ...e, boardCategory: teachTarget } : e
        );
        setEmails(updatedEmails);
        
        // Ensure manual moves persist
        const newManual = [...latestCategories];
        const existingIdx = newManual.findIndex(c => c.id === teachEmail.id);
        if (existingIdx > -1) {
            newManual[existingIdx].category = teachTarget;
        } else {
            newManual.push({ id: teachEmail.id, category: teachTarget, manual: true });
        }
        setLatestCategories(newManual);
        localStorage.setItem('latestCategories', JSON.stringify(newManual));

        try {
            const response = await instance.acquireTokenSilent(loginRequest);
            await updateEmailCategory(response.accessToken, teachEmail.id, teachTarget);
            logAction("AI Trained", 0, `Added new rule to ${teachTarget}`);
        } catch (error) {
            console.error(error);
        }

        // Close modal
        setShowTeachModal(false);
        setTeachCondition("");
        setTeachTarget("");
        setTeachEmail(null);
    };

    const handleSaveManualTask = () => {
        if (!newTaskSubject.trim() || !newTaskTargetCol) return;
        
        const newTask = {
            id: 'manual_' + Date.now(),
            isManual: true,
            subject: newTaskSubject,
            bodyPreview: newTaskBody,
            boardCategory: newTaskTargetCol,
            receivedDateTime: new Date().toISOString()
        };
        
        const updatedTasks = [...manualTasks, newTask];
        setManualTasks(updatedTasks);
        localStorage.setItem('manualTasks', JSON.stringify(updatedTasks));
        
        if (newTaskDueDate) {
            setEmailDueDates(prev => ({ ...prev, [newTask.id]: newTaskDueDate }));
        }

        setShowManualTaskModal(false);
        setNewTaskSubject("");
        setNewTaskBody("");
        setNewTaskDueDate("");
    };

    const handleDeepScan = async () => {
        setIsScanning(true);
        try {
            const response = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });
            const deepEmails = await getDeepEmails(response.accessToken);
            const { patterns, tokens } = await analyzeMailboxPatterns(deepEmails);
            setScanResults(patterns);
            logAction("Deep Analytics Scan", tokens, `Scanned ${deepEmails.length} emails for patterns.`);
        } catch (error) {
            console.error(error);
            alert("Failed to run deep scan.");
        }
        setIsScanning(false);
    };

    const handleGenerateDraft = async () => {
        setIsDrafting(true);
        try {
            if (draftContext?.type === 'reply') {
                const result = await draftGhostwriterReply(conversationMessages, draftInstructions);
                setDraftContent(result.draft);
                logAction("Ghostwriter", result.tokens, `Drafted reply for: ${selectedConversation?.subject}`);
            } else if (draftContext?.type === 'followup') {
                const result = await draftFollowUpEmail(draftContext.task.subject, draftContext.task.bodyPreview, draftInstructions);
                setDraftContent(result.draft);
                logAction("Ghostwriter", result.tokens, `Drafted follow-up for task: ${draftContext.task.subject}`);
            }
        } catch (error) {
            console.error("Draft generation failed:", error);
            alert("Failed to generate draft. Please try again.");
        }
        setIsDrafting(false);
    };

    const handleSendDraft = async () => {
        if (!draftContent) return;
        setIsSending(true);
        try {
            const response = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });

            // Convert plain text draft to basic HTML for sending
            const htmlContent = draftContent.replace(/\n/g, '<br/>');

            if (draftContext?.type === 'reply') {
                // Send reply to the most recent message in the thread
                const latestMessage = conversationMessages[conversationMessages.length - 1] || selectedConversation;
                await sendEmailReply(response.accessToken, latestMessage.id, htmlContent);
                alert("Reply sent successfully!");
                // Optionally move to Done
                if (selectedConversation) {
                    await updateEmailCategory(response.accessToken, selectedConversation.id, 'Done');
                    const newEmails = [...emails];
                    const idx = newEmails.findIndex(e => e.id === selectedConversation.id);
                    if (idx > -1) newEmails[idx].boardCategory = 'Done';
                    setEmails(newEmails);
                }
            } else if (draftContext?.type === 'followup') {
                // For manual task, we need a "To" email address. Since we don't have one stored, we'll ask the user.
                const toEmail = prompt("Please enter the email address to send this follow-up to:");
                if (toEmail) {
                    await sendNewEmail(response.accessToken, toEmail, `Follow up: ${draftContext.task.subject}`, htmlContent);
                    alert("Follow-up email sent successfully!");
                }
            }
            setShowDraftModal(false);
            setDraftContent("");
            setDraftInstructions("");
            setDraftContext(null);
        } catch (error) {
            console.error("Sending failed:", error);
            alert("Failed to send email. Make sure you have granted the Mail.Send permission.");
        }
        setIsSending(false);
    };

    const onDragEnd = async (result) => {
        if (!result.destination) return;
        
        const { source, destination, draggableId, type } = result;

        if (type === 'column') {
            if (source.index === destination.index) return;
            const newCols = Array.from(activeColumns);
            const [removed] = newCols.splice(source.index, 1);
            newCols.splice(destination.index, 0, removed);
            setActiveColumns(newCols);
            localStorage.setItem('activeColumns', JSON.stringify(newCols));
            return;
        }

        if (source.droppableId === destination.droppableId) return;

        // Handle Manual Task Dragging
        if (draggableId.startsWith('manual_')) {
            const newTasks = Array.from(manualTasks);
            const taskIndex = newTasks.findIndex(t => t.id === draggableId);
            newTasks[taskIndex].boardCategory = destination.droppableId;
            setManualTasks(newTasks);
            localStorage.setItem('manualTasks', JSON.stringify(newTasks));
            return;
        }

        // Update local state immediately for snappy UI
        const newEmails = Array.from(emails);
        const emailIndex = newEmails.findIndex(e => e.id === draggableId);
        newEmails[emailIndex].boardCategory = destination.droppableId;
        setEmails(newEmails);

        // Update in Microsoft Graph
        try {
            const response = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });
            await updateEmailCategory(response.accessToken, draggableId, destination.droppableId);
        } catch (error) {
            console.error("Failed to update category in Graph:", error);
            // Revert state if error (simple implementation)
            fetchMails(); 
        }
    };

    if (loading) {
        return <div className="flex-center" style={{ height: '100vh' }}>Loading your mailbox...</div>;
    }

    return (
        <div style={{ padding: '1rem 2rem', height: '100vh', display: 'flex', flexDirection: 'column' }}>
            <header style={{ marginBottom: '1rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                    <h1 className="text-gradient" style={{ margin: 0, fontSize: '1.5rem' }}>Mailbox Manager</h1>
                    <button 
                        onClick={() => setAutoPilotActive(!autoPilotActive)}
                        style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0 }}
                        title={autoPilotActive ? 'Auto-Pilot: ON' : 'Auto-Pilot: OFF'}
                    >
                        <div className={autoPilotActive ? "blinking-dot" : ""} style={{ 
                            width: '12px', height: '12px', borderRadius: '50%', 
                            backgroundColor: autoPilotActive ? '#10b981' : '#ef4444',
                            boxShadow: `0 0 8px ${autoPilotActive ? '#10b981' : '#ef4444'}`
                        }}></div>
                    </button>
                </div>
                <div style={{ display: 'flex', gap: '1rem', alignItems: 'center' }}>
                    <div style={{ position: 'relative' }}>
                        <span style={{ position: 'absolute', left: '0.8rem', top: '50%', transform: 'translateY(-50%)', opacity: 0.5 }}>🔍</span>
                        <input 
                            type="text" 
                            placeholder="Search board..." 
                            value={searchQuery}
                            onChange={(e) => setSearchQuery(e.target.value)}
                            style={{ padding: '0.4rem 0.8rem 0.4rem 2rem', borderRadius: '20px', border: '1px solid var(--border-light)', background: 'var(--bg-secondary)', color: 'white', fontSize: '0.85rem', width: '200px' }}
                        />
                    </div>
                    
                    <button 
                        className="btn" 
                        style={{ padding: '0.4rem 0.8rem', display: 'flex', alignItems: 'center', gap: '0.4rem', border: '1px solid rgba(16, 185, 129, 0.5)', background: 'rgba(16, 185, 129, 0.15)', color: '#10b981', borderRadius: '6px', fontSize: '0.85rem', fontWeight: 600 }}
                        onClick={() => { setNewTaskTargetCol(COLUMNS.TASK.id); setShowManualTaskModal(true); }}
                        title="Add Manual Task"
                    >
                        <span>+</span> Create Task
                    </button>
                    
                    <span style={{ fontSize: '0.9rem', color: 'var(--text-secondary)' }}>{accounts[0]?.name}</span>
                    
                    <div style={{ position: 'relative' }}>
                        <button 
                            className="btn btn-outline" 
                            style={{ padding: '0.5rem', border: 'none', background: 'transparent' }} 
                            onClick={() => setShowSettings(!showSettings)}
                        >
                            ⚙️
                        </button>
                        
                        {showSettings && (
                            <div className="glass-panel" style={{ 
                                position: 'absolute', right: 0, top: '100%', marginTop: '0.5rem', 
                                zIndex: 50, display: 'flex', flexDirection: 'column', gap: '0.5rem', 
                                padding: '1rem', minWidth: '220px', background: 'rgba(22, 25, 32, 0.95)'
                            }}>
                                <button className="btn btn-outline" style={{ justifyContent: 'flex-start' }} onClick={() => { setShowSettings(false); setShowLogs(true); }}>
                                    📋 AI Logs & Cost
                                </button>
                                <button className="btn btn-outline" style={{ justifyContent: 'flex-start' }} onClick={() => { setShowSettings(false); handleDeepScan(); }} disabled={isScanning}>
                                    🔍 {isScanning ? 'Scanning...' : 'Deep Scan Mailbox'}
                                </button>
                                <button className="btn btn-outline" style={{ justifyContent: 'flex-start' }} onClick={() => { setShowSettings(false); setShowSystemLogs(true); }}>
                                    🗄️ System Emails
                                </button>
                                <button className="btn btn-outline" style={{ justifyContent: 'flex-start' }} onClick={() => { setShowSettings(false); setShowRulesModal(true); }}>
                                    🧠 AI Workflow Rules
                                </button>
                                <button className="btn btn-primary" style={{ justifyContent: 'flex-start' }} onClick={() => { setShowSettings(false); handleAiCategorize(); }} disabled={isCategorizing}>
                                    ✨ Force Categorize Inbox
                                </button>
                                <button className="btn btn-primary" style={{ justifyContent: 'flex-start', background: 'var(--accent-purple)' }} onClick={() => { setShowSettings(false); handleRescanPendingTasks(); }} disabled={isCategorizing}>
                                    ✨ Rescan Pending Tasks
                                </button>
                                <div style={{ height: '1px', background: 'var(--border-light)', margin: '0.5rem 0' }}></div>
                                <button className="btn btn-outline" style={{ justifyContent: 'flex-start', color: '#ef4444', border: 'none' }} onClick={() => instance.logoutRedirect()}>
                                    Logout
                                </button>
                            </div>
                        )}
                    </div>
                </div>
            </header>

            {showLogs && (
                <div style={{ 
                    position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, 
                    background: 'rgba(0,0,0,0.7)', zIndex: 100, display: 'flex', 
                    alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(5px)'
                }}>
                    <div className="glass-panel" style={{ padding: '2rem', maxWidth: '600px', width: '100%', maxHeight: '80vh', display: 'flex', flexDirection: 'column' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
                            <h2>📋 AI Activity Logs</h2>
                            <div style={{ background: 'rgba(139, 92, 246, 0.2)', padding: '0.5rem 1rem', borderRadius: '8px', color: 'var(--accent-primary)', fontWeight: 'bold' }}>
                                Total Tokens: {aiLogs.reduce((acc, log) => acc + (log.tokens || 0), 0).toLocaleString()}
                            </div>
                        </div>
                        <p style={{ color: 'var(--text-secondary)', marginBottom: '1.5rem', fontSize: '0.9rem' }}>
                            This log tracks every AI action and the exact token cost for full transparency. You can use these insights to optimize how often Auto-Pilot runs.
                        </p>
                        <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '0.5rem', marginBottom: '1.5rem' }}>
                            {aiLogs.length === 0 ? <p>No AI activity yet.</p> : aiLogs.map((log, i) => (
                                <div key={i} style={{ background: 'var(--bg-secondary)', padding: '0.75rem', borderRadius: '0.5rem', fontSize: '0.85rem' }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.25rem' }}>
                                        <strong style={{ color: 'var(--accent-purple)' }}>{log.action}</strong>
                                        <span style={{ color: 'var(--text-muted)' }}>{new Date(log.timestamp).toLocaleTimeString()}</span>
                                    </div>
                                    <div style={{ color: 'var(--text-primary)', whiteSpace: 'pre-wrap', lineHeight: '1.4' }}>{log.details}</div>
                                    <div style={{ color: 'var(--status-done)', marginTop: '0.25rem', fontSize: '0.8rem' }}>Tokens used: {log.tokens?.toLocaleString()}</div>
                                </div>
                            ))}
                        </div>
                        <button className="btn btn-primary" onClick={() => setShowLogs(false)} style={{ width: '100%' }}>
                            Close Logs
                        </button>
                    </div>
                </div>
            )}

            {showSystemLogs && (
                <div style={{ 
                    position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, 
                    background: 'rgba(0,0,0,0.7)', zIndex: 100, display: 'flex', 
                    alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(5px)'
                }}>
                    <div className="glass-panel" style={{ padding: '2rem', maxWidth: '600px', width: '100%', maxHeight: '80vh', display: 'flex', flexDirection: 'column' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
                            <h2>🗄️ System Emails</h2>
                        </div>
                        <p style={{ color: 'var(--text-secondary)', marginBottom: '1.5rem', fontSize: '0.9rem' }}>
                            These emails were automatically classified as automated system messages.
                        </p>
                        <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '0.5rem', marginBottom: '1.5rem' }}>
                            {emails.filter(e => e.boardCategory === COLUMNS.LOGS.id).length === 0 ? <p>No system logs found.</p> : emails.filter(e => e.boardCategory === COLUMNS.LOGS.id).map(email => (
                                <div key={email.id} style={{ background: 'var(--bg-secondary)', padding: '0.75rem', borderRadius: '0.5rem', fontSize: '0.85rem' }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.25rem' }}>
                                        <strong style={{ color: 'var(--accent-purple)' }}>{email.from?.emailAddress?.name || email.from?.emailAddress?.address}</strong>
                                        <span style={{ color: 'var(--text-muted)' }}>{new Date(email.receivedDateTime).toLocaleDateString()}</span>
                                    </div>
                                    <div style={{ color: 'var(--text-primary)', fontWeight: 600, marginBottom: '0.25rem' }}>{email.subject}</div>
                                    <div style={{ color: 'var(--text-secondary)', display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden' }}>{email.bodyPreview}</div>
                                </div>
                            ))}
                        </div>
                        <button className="btn btn-primary" onClick={() => setShowSystemLogs(false)} style={{ width: '100%' }}>
                            Close
                        </button>
                    </div>
                </div>
            )}

            {scanResults && (
                <div style={{ 
                    position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, 
                    background: 'rgba(0,0,0,0.7)', zIndex: 100, display: 'flex', 
                    alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(5px)'
                }}>
                    <div className="glass-panel" style={{ padding: '2rem', maxWidth: '500px', width: '100%' }}>
                        <h2>🔍 AI Deep Scan Results</h2>
                        <p>Based on a deep analysis of your recent history, here are the most common email patterns you receive. You might want to create columns for these!</p>
                        <ul style={{ marginBottom: '1.5rem', paddingLeft: '0', listStyle: 'none' }}>
                            {scanResults.map((pattern, i) => (
                                <li key={i} style={{ marginBottom: '0.75rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--bg-secondary)', padding: '0.75rem', borderRadius: '0.5rem' }}>
                                    <span style={{ fontSize: '1rem', fontWeight: 500 }}>{pattern}</span>
                                    <button 
                                        className="btn btn-primary" 
                                        style={{ padding: '0.25rem 0.75rem', fontSize: '0.8rem' }}
                                        onClick={() => {
                                            if (!activeColumns.find(c => c.id === pattern)) {
                                                setActiveColumns([...activeColumns, { id: pattern, title: pattern, color: '#14b8a6' }]);
                                            }
                                        }}
                                    >
                                        ➕ Add
                                    </button>
                                </li>
                            ))}
                        </ul>
                        <button className="btn btn-outline" onClick={() => setScanResults(null)} style={{ width: '100%' }}>
                            Close
                        </button>
                    </div>
                </div>
            )}

            {selectedConversation && (
                <div style={{ 
                    position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, 
                    background: 'rgba(0,0,0,0.8)', zIndex: 200, display: 'flex', 
                    alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(8px)'
                }} onClick={() => setSelectedConversation(null)}>
                    <div className="glass-panel" style={{ padding: '2rem', maxWidth: '800px', width: '90%', height: '85vh', display: 'flex', flexDirection: 'column' }} onClick={e => e.stopPropagation()}>
                        <div style={{ borderBottom: '1px solid var(--border-light)', paddingBottom: '1rem', marginBottom: '1rem' }}>
                            <h2 style={{ margin: '0 0 0.5rem 0', fontSize: '1.25rem' }}>{selectedConversation.subject}</h2>
                            <div style={{ display: 'flex', gap: '0.5rem' }}>
                                <span style={{ background: 'var(--accent-purple)', padding: '0.2rem 0.6rem', borderRadius: '4px', fontSize: '0.75rem' }}>{selectedConversation.boardCategory}</span>
                            </div>
                        </div>
                        
                        <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                            {conversationMessages.length === 0 ? (
                                <div className="flex-center" style={{ height: '100%' }}>Loading thread...</div>
                            ) : (
                                conversationMessages.map(msg => {
                                    const isMe = msg.from?.emailAddress?.address === accounts[0].username;
                                    return (
                                        <div key={msg.id} style={{ 
                                            background: isMe ? 'rgba(16, 185, 129, 0.1)' : 'var(--bg-secondary)', 
                                            padding: '1rem', borderRadius: '8px',
                                            borderLeft: isMe ? '4px solid #10b981' : '4px solid var(--accent-purple)'
                                        }}>
                                            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.5rem', fontSize: '0.85rem' }}>
                                                <strong>{msg.from?.emailAddress?.name || msg.from?.emailAddress?.address}</strong>
                                                <span style={{ color: 'var(--text-muted)' }}>{new Date(msg.receivedDateTime).toLocaleString()}</span>
                                            </div>
                                            {msg.body?.contentType === 'html' ? (
                                                <iframe 
                                                    srcDoc={msg.body.content} 
                                                    title="email-body"
                                                    style={{ width: '100%', minHeight: '500px', border: 'none', background: 'white', borderRadius: '4px', marginTop: '0.5rem' }}
                                                    sandbox="allow-same-origin allow-popups"
                                                />
                                            ) : (
                                                <div style={{ whiteSpace: 'pre-wrap', fontSize: '0.9rem', color: 'var(--text-primary)', lineHeight: '1.5' }}>
                                                    {(msg.body?.content || msg.bodyPreview || "")
                                                        .replace(/P\s*\{[^\}]+\}/gi, '')
                                                        .replace(/&nbsp;/g, ' ')
                                                        .replace(/&amp;/g, '&')
                                                        .replace(/&quot;/g, '"')
                                                        .replace(/(Company Name:|Package:|Packages:|Payment Mode:|Recognition Mode:|Type of Service:|Valid Period:|Account Category:|Sales Person:|Hardcopy Contract with:|Thank you\.|Regards,)/gi, '\n$1')
                                                        .trim()
                                                    }
                                                </div>
                                            )}
                                        </div>
                                    )
                                })
                            )}
                        </div>
                        <div style={{ display: 'flex', gap: '1rem', marginTop: '1rem' }}>
                            <button className="btn btn-outline" onClick={() => setSelectedConversation(null)} style={{ flex: 1 }}>
                                Close Thread
                            </button>
                            <button className="btn btn-primary" onClick={() => {
                                setDraftContext({ type: 'reply' });
                                setShowDraftModal(true);
                            }} style={{ flex: 1, background: 'var(--accent-primary)' }}>
                                ✨ Draft Reply (AI)
                            </button>
                        </div>
                    </div>
                </div>
            )}

                            {showTeachModal && teachEmail && (
                <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.8)', zIndex: 300, display: 'flex', alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(8px)' }} onClick={() => setShowTeachModal(false)}>
                    <div className="glass-panel" style={{ padding: '2rem', maxWidth: '500px', width: '90%' }} onClick={e => e.stopPropagation()}>
                        <h2 style={{ margin: '0 0 1rem 0' }}>🧠 Teach the AI</h2>
                        <p style={{ fontSize: '0.9rem', color: 'var(--text-secondary)', marginBottom: '1.5rem' }}>
                            The AI left this email in the Inbox. Define a rule so it learns what to do next time!
                        </p>
                        
                        <div style={{ background: 'var(--bg-secondary)', padding: '1rem', borderRadius: '8px', marginBottom: '1.5rem', fontSize: '0.85rem' }}>
                            <strong>Subject:</strong> {teachEmail.subject}<br/>
                            <strong>From:</strong> {teachEmail.from?.emailAddress?.address}
                        </div>

                        <div style={{ marginBottom: '1rem' }}>
                            <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem', fontWeight: 600 }}>If the email is like this (Condition):</label>
                            <input 
                                type="text" 
                                value={teachCondition}
                                onChange={e => setTeachCondition(e.target.value)}
                                placeholder="e.g. It is an AWS invoice"
                                style={{ width: '100%', padding: '0.8rem', borderRadius: '6px', border: '1px solid var(--border-light)', background: 'var(--bg-secondary)', color: 'white', boxSizing: 'border-box' }}
                            />
                        </div>

                        <div style={{ marginBottom: '2rem' }}>
                            <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem', fontWeight: 600 }}>Route it to:</label>
                            <select 
                                value={teachTarget}
                                onChange={e => setTeachTarget(e.target.value)}
                                style={{ width: '100%', padding: '0.8rem', borderRadius: '6px', border: '1px solid var(--border-light)', background: 'var(--bg-secondary)', color: 'white', boxSizing: 'border-box' }}
                            >
                                <option value="" disabled>Select a column...</option>
                                {activeColumns.map(c => (
                                    <option key={c.id} value={c.id}>{c.title}</option>
                                ))}
                            </select>
                        </div>

                        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '1rem' }}>
                            <button className="btn btn-outline" onClick={() => setShowTeachModal(false)}>Cancel</button>
                            <button className="btn btn-primary" onClick={handleSaveTeachRule} disabled={!teachCondition || !teachTarget}>Save Rule & Move Email</button>
                        </div>
                    </div>
                </div>
            )}

            {showManualTaskModal && (
                <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.8)', zIndex: 300, display: 'flex', alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(8px)' }} onClick={() => setShowManualTaskModal(false)}>
                    <div className="glass-panel" style={{ padding: '2rem', maxWidth: '500px', width: '90%' }} onClick={e => e.stopPropagation()}>
                        <h2 style={{ margin: '0 0 1.5rem 0' }}>📝 Create Manual Task</h2>
                        
                        <div style={{ marginBottom: '1rem' }}>
                            <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem', fontWeight: 600 }}>Task Title:</label>
                            <input 
                                type="text" 
                                value={newTaskSubject}
                                onChange={e => setNewTaskSubject(e.target.value)}
                                placeholder="e.g. Call vendor to discuss pricing"
                                style={{ width: '100%', padding: '0.8rem', borderRadius: '6px', border: '1px solid var(--border-light)', background: 'var(--bg-secondary)', color: 'white', boxSizing: 'border-box' }}
                                autoFocus
                            />
                        </div>

                        <div style={{ marginBottom: '1rem' }}>
                            <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem', fontWeight: 600 }}>Follow-Up Date (Optional):</label>
                            <input 
                                type="date" 
                                value={newTaskDueDate}
                                onChange={e => setNewTaskDueDate(e.target.value)}
                                style={{ width: '100%', padding: '0.8rem', borderRadius: '6px', border: '1px solid var(--border-light)', background: 'var(--bg-secondary)', color: 'white', boxSizing: 'border-box' }}
                            />
                        </div>

                        <div style={{ marginBottom: '2rem' }}>
                            <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem', fontWeight: 600 }}>Description (Optional):</label>
                            <textarea 
                                value={newTaskBody}
                                onChange={e => setNewTaskBody(e.target.value)}
                                placeholder="Add any details..."
                                style={{ width: '100%', padding: '0.8rem', borderRadius: '6px', border: '1px solid var(--border-light)', background: 'var(--bg-secondary)', color: 'white', boxSizing: 'border-box', minHeight: '80px', fontFamily: 'inherit' }}
                            />
                        </div>

                        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '1rem' }}>
                            <button className="btn btn-outline" onClick={() => setShowManualTaskModal(false)}>Cancel</button>
                            <button className="btn btn-primary" onClick={handleSaveManualTask} disabled={!newTaskSubject.trim()}>Create Task</button>
                        </div>
                    </div>
                </div>
            )}

            {showDraftModal && (
                <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.8)', zIndex: 400, display: 'flex', alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(8px)' }} onClick={() => !isDrafting && !isSending && setShowDraftModal(false)}>
                    <div className="glass-panel" style={{ padding: '2rem', maxWidth: '600px', width: '90%', display: 'flex', flexDirection: 'column', gap: '1rem' }} onClick={e => e.stopPropagation()}>
                        <h2 style={{ margin: 0 }}>✨ AI Ghostwriter</h2>
                        
                        <div>
                            <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem', fontWeight: 600 }}>Custom Instructions (Optional):</label>
                            <input 
                                type="text" 
                                value={draftInstructions}
                                onChange={e => setDraftInstructions(e.target.value)}
                                placeholder="e.g. Tell them we can only do a 20% discount..."
                                style={{ width: '100%', padding: '0.8rem', borderRadius: '6px', border: '1px solid var(--border-light)', background: 'var(--bg-secondary)', color: 'white', boxSizing: 'border-box' }}
                            />
                        </div>

                        <button className="btn btn-primary" onClick={handleGenerateDraft} disabled={isDrafting} style={{ width: '100%' }}>
                            {isDrafting ? 'Generating...' : 'Generate Draft'}
                        </button>

                        {draftContent && (
                            <>
                                <div>
                                    <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem', fontWeight: 600 }}>Draft Preview (Edit if needed):</label>
                                    <textarea 
                                        value={draftContent}
                                        onChange={e => setDraftContent(e.target.value)}
                                        style={{ width: '100%', padding: '0.8rem', borderRadius: '6px', border: '1px solid var(--border-light)', background: 'var(--bg-secondary)', color: 'white', boxSizing: 'border-box', minHeight: '150px', fontFamily: 'inherit' }}
                                    />
                                </div>
                                <div style={{ display: 'flex', gap: '1rem' }}>
                                    <button className="btn btn-outline" onClick={() => {
                                        navigator.clipboard.writeText(draftContent);
                                        alert('Copied to clipboard!');
                                    }} style={{ flex: 1 }}>
                                        📋 Copy
                                    </button>
                                    <button className="btn btn-primary" onClick={handleSendDraft} disabled={isSending} style={{ flex: 2, background: '#10b981' }}>
                                        {isSending ? 'Sending...' : '📤 Send Email'}
                                    </button>
                                </div>
                            </>
                        )}
                        
                        <button className="btn btn-outline" onClick={() => setShowDraftModal(false)} disabled={isDrafting || isSending} style={{ marginTop: '0.5rem' }}>
                            Cancel
                        </button>
                    </div>
                </div>
            )}

            <DragDropContext onDragEnd={onDragEnd}>
                <Droppable droppableId="board" type="column" direction="horizontal">
                    {(providedBoard) => (
                        <div 
                            ref={providedBoard.innerRef}
                            {...providedBoard.droppableProps}
                            style={{ display: 'flex', gap: '1.5rem', flex: 1, overflowX: 'auto', paddingBottom: '1rem' }}
                        >
                            {activeColumns.map((column, colIndex) => {
                                const allItems = [...emails, ...manualTasks];
                                let columnEmailsAll = allItems.filter(e => e.boardCategory === column.id);
                                
                                if (searchQuery.trim() !== '') {
                                    const q = searchQuery.toLowerCase();
                                    columnEmailsAll = columnEmailsAll.filter(e => 
                                        (e.subject && e.subject.toLowerCase().includes(q)) || 
                                        (e.bodyPreview && e.bodyPreview.toLowerCase().includes(q)) ||
                                        (e.from?.emailAddress?.name && e.from.emailAddress.name.toLowerCase().includes(q))
                                    );
                                }
                                
                                let categoryCounts = {};
                                let totalPending = 0;
                                if (column.id === COLUMNS.TASK.id) {
                                    totalPending = columnEmailsAll.length;
                                    columnEmailsAll.forEach(e => {
                                        if (e.subCategory) {
                                            categoryCounts[e.subCategory] = (categoryCounts[e.subCategory] || 0) + 1;
                                        }
                                    });
                                }

                                // Apply filter if we are in the Pending Task column
                                if (column.id === COLUMNS.TASK.id) {
                                    if (pendingFilter !== 'All') {
                                        columnEmailsAll = columnEmailsAll.filter(e => e.subCategory === pendingFilter);
                                    }
                                    
                                    // Sort by due date (items with due dates first, sorted chronologically)
                                    columnEmailsAll.sort((a, b) => {
                                        const dateA = emailDueDates[a.id];
                                        const dateB = emailDueDates[b.id];
                                        if (dateA && dateB) {
                                            return new Date(dateA) - new Date(dateB);
                                        }
                                        if (dateA) return -1;
                                        if (dateB) return 1;
                                        return 0;
                                    });
                                }
                                
                                const columnEmails = columnEmailsAll.slice(0, 50); // Cap at 50 for performance
                                
                                const isAutoCollapsed = column.id === COLUMNS.INBOX.id && columnEmailsAll.length === 0;
                                const isManuallyCollapsed = manualCollapsed[column.id] === true;
                                const isManuallyExpanded = manualCollapsed[column.id] === false;
                                const isCollapsed = isManuallyCollapsed || (isAutoCollapsed && !isManuallyExpanded);

                                return (
                                    <Draggable key={column.id} draggableId={column.id} index={colIndex}>
                                        {(providedCol, snapshotCol) => (
                                            <div
                                                ref={providedCol.innerRef}
                                                {...providedCol.draggableProps}
                                                style={{ 
                                                    flex: isCollapsed ? '0 0 60px' : (column.id === COLUMNS.TASK.id ? '2 1 0' : '1 1 0'), 
                                                    minWidth: isCollapsed ? '60px' : (column.id === COLUMNS.TASK.id ? '440px' : '220px'), 
                                                    display: 'flex', 
                                                    flexDirection: 'column', 
                                                    height: '100%',
                                                    opacity: isCollapsed ? 0.7 : 1,
                                                    transition: 'all 0.3s ease',
                                                    ...providedCol.draggableProps.style
                                                }}
                                            >
                                                <div 
                                                    {...providedCol.dragHandleProps}
                                                    style={{ 
                                                        marginBottom: '1rem', display: 'flex', alignItems: 'center', gap: '0.5rem', 
                                                        cursor: snapshotCol.isDragging ? 'grabbing' : 'grab',
                                                        padding: '0.5rem', borderRadius: '4px', background: snapshotCol.isDragging ? 'var(--bg-secondary)' : 'transparent',
                                                        flexDirection: isCollapsed ? 'column' : 'row',
                                                        userSelect: 'none'
                                                    }}
                                                    onDoubleClick={(e) => {
                                                        e.stopPropagation();
                                                        setManualCollapsed(prev => ({ ...prev, [column.id]: !isCollapsed }));
                                                    }}
                                                    title="Double-click to expand/collapse"
                                                >
                                                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', flex: 1, flexDirection: isCollapsed ? 'column' : 'row' }}>
                                                        <div style={{ width: '12px', height: '12px', borderRadius: '50%', backgroundColor: column.color, flexShrink: 0 }}></div>
                                                        <h3 style={{ 
                                                            margin: 0, fontSize: '1rem', 
                                                            writingMode: isCollapsed ? 'vertical-rl' : 'horizontal-tb',
                                                            transform: isCollapsed ? 'rotate(180deg)' : 'none',
                                                            whiteSpace: 'nowrap'
                                                        }}>
                                                            {column.title} {!isCollapsed && <span style={{ color: 'var(--text-muted)', fontSize: '0.8rem' }}>({columnEmailsAll.length > 50 ? `50 of ${columnEmailsAll.length}` : columnEmailsAll.length})</span>}
                                                        </h3>
                                                    </div>
                                                </div>
                                                
                                                {/* Filter Bar for Pending Task */}
                                                {column.id === COLUMNS.TASK.id && (
                                                    <div style={{ display: 'flex', gap: '0.5rem', flexWrap: 'wrap', paddingBottom: '0.5rem', marginBottom: '0.5rem' }}>
                                                        <button 
                                                            onClick={() => setPendingFilter('All')}
                                                            style={{
                                                                padding: '0.2rem 0.6rem', fontSize: '0.75rem', borderRadius: '12px', cursor: 'pointer', whiteSpace: 'nowrap',
                                                                border: pendingFilter === 'All' ? '1px solid var(--accent-primary)' : '1px solid var(--border-light)',
                                                                background: pendingFilter === 'All' ? 'var(--accent-primary)' : 'transparent',
                                                                color: pendingFilter === 'All' ? 'white' : 'var(--text-primary)', transition: 'all 0.2s'
                                                            }}
                                                        >
                                                            All ({totalPending})
                                                        </button>
                                                        {Object.entries(TASK_SUBCATEGORIES).map(([sub, color]) => (
                                                            <button 
                                                                key={sub}
                                                                onClick={() => setPendingFilter(sub)}
                                                                style={{
                                                                    padding: '0.2rem 0.6rem', fontSize: '0.75rem', borderRadius: '12px', cursor: 'pointer', whiteSpace: 'nowrap',
                                                                    border: `1px solid ${color}`,
                                                                    background: pendingFilter === sub ? color : 'transparent',
                                                                    color: pendingFilter === sub ? 'white' : color,
                                                                    transition: 'all 0.2s'
                                                                }}
                                                            >
                                                                {sub} ({categoryCounts[sub] || 0})
                                                            </button>
                                                        ))}
                                                    </div>
                                                )}
                                                
                                                <Droppable droppableId={column.id} type="email">
                                                    {(provided) => (
                                                        <div 
                                                            {...provided.droppableProps}
                                                            ref={provided.innerRef}
                                                            style={{ 
                                                                flex: 1,
                                                                display: 'flex', 
                                                                flexDirection: 'column', 
                                                                padding: isCollapsed ? '0' : '1rem',
                                                                background: isCollapsed ? 'transparent' : 'rgba(22, 25, 32, 0.4)',
                                                                border: isCollapsed ? '1px dashed var(--border-light)' : '1px solid var(--border-light)',
                                                                borderRadius: 'var(--radius-lg)'
                                                            }}
                                                        >
                                        
                                        <div style={{ 
                                            flex: 1, overflowY: 'auto', overflowX: 'hidden', 
                                            display: (column.id === COLUMNS.TASK.id && !isCollapsed) ? 'grid' : 'flex', 
                                            gridTemplateColumns: (column.id === COLUMNS.TASK.id && !isCollapsed) ? '1fr 1fr' : undefined,
                                            alignContent: 'start',
                                            flexDirection: (column.id === COLUMNS.TASK.id && !isCollapsed) ? undefined : 'column',
                                            gap: '0.75rem', 
                                            paddingRight: '0.25rem' 
                                        }}>
                                            {columnEmails.map((email, index) => (
                                                <Draggable key={email.id} draggableId={email.id} index={index}>
                                                    {(provided, snapshot) => (
                                                        <div
                                                            ref={provided.innerRef}
                                                            {...provided.draggableProps}
                                                            {...provided.dragHandleProps}
                                                            className="glass-panel"
                                                            onClick={() => handleCardClick(email)}
                                                            style={{
                                                                padding: '0.75rem',
                                                                background: snapshot.isDragging ? 'var(--bg-secondary)' : 'var(--glass-bg)',
                                                                cursor: 'grab',
                                                                borderLeft: `4px solid ${column.color}`,
                                                                boxSizing: 'border-box',
                                                                ...provided.draggableProps.style
                                                            }}
                                                        >
                                                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '0.5rem', gap: '0.5rem' }}>
                                                                <div style={{ fontWeight: 500, fontSize: '0.8rem', color: 'var(--text-primary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1 }}>
                                                                    {email.isManual ? 'User Created' : (email.from?.emailAddress?.name || email.from?.emailAddress?.address)}
                                                                </div>
                                                                <div style={{ display: 'flex', gap: '0.25rem', alignItems: 'center', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                                                                    {emailDueDates[email.id] && (() => {
                                                                        const dueDate = new Date(emailDueDates[email.id]);
                                                                        const today = new Date();
                                                                        today.setUTCHours(0, 0, 0, 0);
                                                                        const isOverdue = dueDate < today;
                                                                        return (
                                                                            <span style={{ background: isOverdue ? '#ef4444' : '#ec4899', color: 'white', padding: '0.1rem 0.4rem', borderRadius: '4px', fontSize: '0.7rem', fontWeight: 600, flexShrink: 0 }}>
                                                                                {isOverdue ? '⚠️ Overdue:' : '📅 Due:'} {dueDate.toLocaleDateString(undefined, { month: 'short', day: 'numeric', timeZone: 'UTC' })}
                                                                            </span>
                                                                        );
                                                                    })()}
                                                                    {email.subCategory && (
                                                                        <span style={{ background: TASK_SUBCATEGORIES[email.subCategory] || '#475569', color: 'white', padding: '0.1rem 0.4rem', borderRadius: '4px', fontSize: '0.7rem', fontWeight: 600, flexShrink: 0 }}>
                                                                            {email.subCategory}
                                                                        </span>
                                                                    )}
                                                                    {email.isManual && (
                                                                        <span style={{ background: '#6366f1', color: 'white', padding: '0.1rem 0.4rem', borderRadius: '4px', fontSize: '0.7rem', fontWeight: 600, flexShrink: 0 }}>
                                                                            📝 Manual Task
                                                                        </span>
                                                                    )}
                                                                    {column.id === COLUMNS.INBOX.id && !email.isManual && (
                                                                        <button 
                                                                            title="Teach AI how to route this email"
                                                                            onClick={(e) => { e.stopPropagation(); setTeachEmail(email); setShowTeachModal(true); }}
                                                                            style={{ background: 'rgba(255,255,255,0.1)', border: 'none', borderRadius: '50%', width: '20px', height: '20px', display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', fontSize: '0.7rem' }}
                                                                        >
                                                                            🧠
                                                                        </button>
                                                                    )}
                                                                    {email.importance === 'high' && !email.isManual && (
                                                                        <span style={{ background: '#ef4444', color: 'white', padding: '0.1rem 0.4rem', borderRadius: '4px', fontSize: '0.7rem', fontWeight: 600, flexShrink: 0 }}>
                                                                            🔥 High
                                                                        </span>
                                                                    )}
                                                                </div>
                                                            </div>
                                                            <div style={{ fontSize: '0.875rem', fontWeight: 600, marginBottom: '0.5rem', color: 'var(--text-primary)', overflowWrap: 'break-word', wordBreak: 'break-word' }}>
                                                                {email.subject || '(No Subject)'}
                                                            </div>
                                                            <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden', marginBottom: '0.5rem', overflowWrap: 'break-word', wordBreak: 'break-all' }}>
                                                                {email.bodyPreview}
                                                            </div>
                                                            {email.isManual && (
                                                                <button 
                                                                    onClick={(e) => { e.stopPropagation(); setDraftContext({ type: 'followup', task: email }); setShowDraftModal(true); }}
                                                                    style={{ width: '100%', padding: '0.4rem', background: 'var(--accent-primary)', color: 'white', border: 'none', borderRadius: '4px', fontSize: '0.75rem', cursor: 'pointer', marginBottom: '0.5rem' }}
                                                                >
                                                                    ✨ Draft Follow-Up
                                                                </button>
                                                            )}
                                                            <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', textAlign: 'right' }}>
                                                                {new Date(email.receivedDateTime).toLocaleString()}
                                                            </div>
                                                        </div>
                                                    )}
                                                </Draggable>
                                            ))}
                                            {provided.placeholder}
                                        </div>
                                    </div>
                                )}
                            </Droppable>
                        </div>
                    )}
                </Draggable>
                                );
                            })}
                            {providedBoard.placeholder}
                        </div>
                    )}
                </Droppable>
            </DragDropContext>
        </div>
    );
};
