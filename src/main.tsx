import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import './index.css';

if (window.location.protocol === 'file:') {
  const warning = document.createElement('div');
  warning.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; background: #EF4444; color: white; padding: 20px; text-align: center; z-index: 9999; font-family: sans-serif;';
  warning.innerHTML = '<h1>⚠️ 无法直接打开 HTML 文件</h1><p>由于浏览器安全限制，React 应用必须通过本地服务器运行。请使用 <code>npx serve -s dist</code> 或 <code>npm run dev</code> 运行。</p>';
  document.body.appendChild(warning);
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
