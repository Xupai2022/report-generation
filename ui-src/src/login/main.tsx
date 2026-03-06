import React from 'react';
import { createRoot } from 'react-dom/client';
import { LoginApp } from '../pages/LoginApp';
import '../shared/base.css';

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <LoginApp />
  </React.StrictMode>,
);
