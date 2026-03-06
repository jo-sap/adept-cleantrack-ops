
import React, { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import { MsalProvider } from "@azure/msal-react";
import { msalInstance } from "./lib/msal";

const rootElement = document.getElementById('root');
if (!rootElement) {
  throw new Error("Could not find root element to mount to");
}

const root = createRoot(rootElement);

// Initialize MSAL before mounting the provider
msalInstance.initialize().then(() => {
  root.render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>
  );
}).catch(err => {
  console.error("MSAL Initialization Failed", err);
  // Fallback mount so the application still functions without MSAL
  root.render(
    <StrictMode>
      <App />
    </StrictMode>
  );
});
