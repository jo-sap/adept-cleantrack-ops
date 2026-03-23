
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

const MSAL_INIT_TIMEOUT_MS = 10_000;

function renderApp(useMsalProvider: boolean) {
  root.render(
    <StrictMode>
      {useMsalProvider ? (
        <MsalProvider instance={msalInstance}>
          <App />
        </MsalProvider>
      ) : (
        <App />
      )}
    </StrictMode>
  );
}

// Immediate feedback so a hung MSAL init does not look like a broken tab
root.render(
  <StrictMode>
    <div
      style={{
        minHeight: '100vh',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontFamily: 'Inter, system-ui, sans-serif',
        color: '#6b7280',
        fontSize: 14,
      }}
    >
      Loading CleanTrack…
    </div>
  </StrictMode>
);

const initOk = msalInstance.initialize().then(
  () => true as const,
  () => false as const
);
const timedOut = new Promise<boolean>((resolve) => {
  setTimeout(() => resolve(false), MSAL_INIT_TIMEOUT_MS);
});

Promise.race([initOk, timedOut])
  .then((ok) => {
    if (ok) {
      renderApp(true);
    } else {
      console.warn(
        "[CleanTrack] MSAL initialize() did not complete successfully before timeout — continuing without MsalProvider. Sign-in may not work until you refresh."
      );
      renderApp(false);
    }
  })
  .catch((err) => {
    console.error("MSAL Initialization Failed", err);
    renderApp(false);
  });
