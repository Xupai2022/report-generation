import React from 'react';
import { createRoot } from 'react-dom/client';
import { AdminApp } from '../pages/AdminApp';
import '../shared/base.css';

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <AdminApp />
  </React.StrictMode>,
);
