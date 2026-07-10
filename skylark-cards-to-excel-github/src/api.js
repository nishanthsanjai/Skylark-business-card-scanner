import { API_URL } from './config.js';

/**
 * All backend calls go through the Apps Script web app.
 * The body is sent as a plain string (no JSON content-type header) so the
 * browser skips the CORS preflight that Apps Script can't answer.
 */
async function call(action, payload = {}) {
  if (!API_URL || API_URL.startsWith('PASTE_')) {
    throw new Error('Backend not configured yet — set API_URL in src/config.js');
  }
  const res = await fetch(API_URL, {
    method: 'POST',
    body: JSON.stringify({ action, ...payload }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error);
  return data;
}

export const api = {
  signup: (username, password) => call('signup', { username, password }),
  login: (username, password) => call('login', { username, password }),
  listContacts: () => call('listContacts'),
  processCard: (imageB64, mediaType, owner) => call('processCard', { imageB64, mediaType, owner }),
  updateContact: (serial, patch) => call('updateContact', { serial, patch }),
  deleteContact: (serial) => call('deleteContact', { serial }),
};

/** Resize/compress a photo before upload: faster, cheaper, kinder to mobile data */
export function compressImage(file, maxDim = 1600, quality = 0.82) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
      const dataUrl = canvas.toDataURL('image/jpeg', quality);
      resolve({ b64: dataUrl.split(',')[1], mediaType: 'image/jpeg', preview: dataUrl });
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Could not read image')); };
    img.src = url;
  });
}
