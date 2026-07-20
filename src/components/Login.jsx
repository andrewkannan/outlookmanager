import React from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from '../authConfig';
import { Mail, ArrowRight } from 'lucide-react';

export const Login = () => {
    const { instance } = useMsal();

    const handleLogin = () => {
        instance.loginRedirect(loginRequest).catch((e) => {
            console.error(e);
        });
    };

    return (
        <div style={{ display: 'flex', minHeight: '100vh', alignItems: 'center', justifyContent: 'center', padding: '2rem' }}>
            <div className="glass-panel" style={{ maxWidth: '480px', width: '100%', padding: '3rem', textAlign: 'center' }}>
                <div style={{ 
                    width: '64px', height: '64px', borderRadius: '1rem', 
                    background: 'rgba(59, 130, 246, 0.1)', display: 'flex', 
                    alignItems: 'center', justifyContent: 'center', margin: '0 auto 1.5rem auto'
                }}>
                    <Mail size={32} color="#3b82f6" />
                </div>
                
                <h1 style={{ marginBottom: '1rem' }}>Mailbox <span className="text-gradient">Manager</span></h1>
                <p style={{ marginBottom: '2rem' }}>
                    Connect your Outlook account to automatically organize your emails into a Kanban board using AI.
                </p>
                
                <button onClick={handleLogin} className="btn btn-primary" style={{ width: '100%', padding: '0.75rem', fontSize: '1rem' }}>
                    Sign in with Microsoft
                    <ArrowRight size={18} />
                </button>
                
                <div style={{ marginTop: '2rem', fontSize: '0.875rem', color: 'var(--text-muted)' }}>
                    Requires Azure AD App Registration. <br/>
                    Ensure you have set the Client ID in the configuration.
                </div>
            </div>
        </div>
    );
};
