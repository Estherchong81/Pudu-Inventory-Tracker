import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

// Expose GROQ key from Vite env to window for runtime access
window.__GROQ_KEY__ = import.meta.env.VITE_GROQ_API_KEY || '';

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
)
