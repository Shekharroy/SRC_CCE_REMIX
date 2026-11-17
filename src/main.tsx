import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.tsx';
import { ErrorBoundary } from './components/ErrorBoundary.tsx';
import './index.css';

// Prevent third-party browser extension injection errors (e.g. chrome-extension://*/frame_ant.js)
// from surfacing as unhandled app crashes.
window.addEventListener('error', (event) => {
  if (
    typeof event.filename === 'string' &&
    event.filename.startsWith('chrome-extension://')
  ) {
    event.preventDefault();
  }
});

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>
);
