import { LogLevel } from "@azure/msal-browser";

// This will be replaced with the actual Client ID the user provides
export const msalConfig = {
    auth: {
        clientId: "4ffc4be3-1436-4f0d-9a82-5590b8683ee3",
        authority: "https://login.microsoftonline.com/da892747-7b26-4692-b33c-607c95a2917f", // Tenant-specific endpoint
        redirectUri: import.meta.env.VITE_REDIRECT_URI || "http://localhost:5173/",
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {	
        loggerOptions: {	
            loggerCallback: (level, message, containsPii) => {	
                if (containsPii) {		
                    return;		
                }		
                switch (level) {
                    case LogLevel.Error:
                        console.error(message);
                        return;
                    case LogLevel.Info:
                        console.info(message);
                        return;
                    case LogLevel.Verbose:
                        console.debug(message);
                        return;
                    case LogLevel.Warning:
                        console.warn(message);
                        return;
                    default:
                        return;
                }	
            }	
        }	
    }
};

export const loginRequest = {
    scopes: ["User.Read", "Mail.ReadWrite", "Mail.Send"]
};

export const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages"
};
