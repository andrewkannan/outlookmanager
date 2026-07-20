import React from 'react';
import ReactDOM from 'react-dom/client';
import { PublicClientApplication } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import { msalConfig } from './authConfig';
import App from './App.jsx';
import './index.css';

const msalInstance = new PublicClientApplication(msalConfig);

// Initialize MSAL before rendering
msalInstance.initialize().then(() => {
  // This is required for MSAL to process the redirect response from Azure AD
  msalInstance.handleRedirectPromise().then((response) => {
    // If response exists, it means we just came back from login
    console.log("MSAL redirect handled.");
  }).catch(err => {
    console.error("MSAL redirect error:", err);
  });

  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </React.StrictMode>,
  );
});
