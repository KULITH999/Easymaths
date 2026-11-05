// Paste your Google Apps Script Web App URL here.
// Deploy Code.gs as Web App (Execute as: Me, Access: Anyone) and copy the /exec URL.
window.EASY_MATHS_API_URL = 'https://script.google.com/macros/s/AKfycbwEW8h1-nqAh5jICCDaBxNo35LVUM0qqiQkFB--_N4JnquMtgMaCFCRF2aItUYzK4FPsg/exec';

// Universal API helper used by index.html and admin.html.
window.callApi = async function (action, ...args) {
  const url = window.EASY_MATHS_API_URL;
  if (!url || url.indexOf('PASTE_YOUR') === 0) {
    throw new Error('API URL not configured. Edit public/config.js first.');
  }
  const res = await fetch(url, {
    method: 'POST',
    // text/plain avoids CORS preflight (Apps Script does not handle OPTIONS)
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action: action, args: args })
  });
  if (!res.ok) throw new Error('Network error: ' + res.status);
  return res.json();
};