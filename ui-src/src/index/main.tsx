import React from 'react';
import { createRoot } from 'react-dom/client';
import { IndexApp } from '../pages/IndexApp';
import '../shared/base.css';

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <IndexApp />
  </React.StrictMode>,
);
