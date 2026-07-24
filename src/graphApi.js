import { Client } from "@microsoft/microsoft-graph-client";

// Get an instance of the Microsoft Graph client
export const getGraphClient = (accessToken) => {
    return Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
};

// Fetch recent emails
export const getEmails = async (accessToken) => {
    const client = getGraphClient(accessToken);
    try {
        const response = await client
            .api('/me/messages')
            .select('id,conversationId,subject,bodyPreview,from,toRecipients,receivedDateTime,categories,importance')
            .top(100)
            .orderby('receivedDateTime DESC')
            .get();
        return response.value;
    } catch (error) {
        console.error("Error fetching emails:", error);
        throw error;
    }
};

// Fetch full conversation by conversationId
export const getConversation = async (accessToken, conversationId) => {
    const client = getGraphClient(accessToken);
    try {
        const response = await client
            .api('/me/messages')
            .filter(`conversationId eq '${conversationId}'`)
            .select('id,subject,bodyPreview,body,from,receivedDateTime')
            .orderby('receivedDateTime ASC')
            .get();
        return response.value;
    } catch (error) {
        console.error("Error fetching conversation:", error);
        throw error;
    }
};

// Update an email's categories (used as tags for our columns)
export const updateEmailCategory = async (accessToken, messageId, categoryName) => {
    const client = getGraphClient(accessToken);
    try {
        await client
            .api(`/me/messages/${messageId}`)
            .patch({
                categories: [categoryName]
            });
    } catch (error) {
        console.error("Error updating email category:", error);
        throw error;
    }
};

// Fetch a larger batch of emails for deep analysis
export const getDeepEmails = async (accessToken) => {
    const client = getGraphClient(accessToken);
    try {
        const response = await client
            .api('/me/messages')
            .select('id,conversationId,subject,bodyPreview,from,toRecipients,receivedDateTime,categories,importance')
            .top(200)
            .orderby('receivedDateTime DESC')
            .get();
        return response.value;
    } catch (error) {
        console.error("Error fetching deep emails:", error);
        throw error;
    }
};

export const sendEmailReply = async (accessToken, messageId, htmlContent) => {
    const client = getGraphClient(accessToken);
    try {
        await client
            .api(`/me/messages/${messageId}/reply`)
            .post({
                message: {
                    body: {
                        contentType: "html",
                        content: htmlContent
                    }
                }
            });
    } catch (error) {
        console.error("Error sending reply:", error);
        throw error;
    }
};

export const sendNewEmail = async (accessToken, toEmail, subject, htmlContent) => {
    const client = getGraphClient(accessToken);
    try {
        await client
            .api('/me/sendMail')
            .post({
                message: {
                    subject: subject,
                    body: {
                        contentType: "html",
                        content: htmlContent
                    },
                    toRecipients: [
                        {
                            emailAddress: {
                                address: toEmail
                            }
                        }
                    ]
                }
            });
    } catch (error) {
        console.error("Error sending new email:", error);
        throw error;
    }
};
