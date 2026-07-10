/**
 * SKYLARK DRONES — Cards to Excel · Backend
 * ------------------------------------------------
 * Paste this into Apps Script bound to your Google Sheet
 * (open the Sheet → Extensions → Apps Script), then:
 *   1. Project Settings → Script properties → add:
 *        ANTHROPIC_API_KEY = sk-ant-...
 *   2. Deploy → New deployment → Web app
 *        Execute as: Me · Who has access: Anyone
 *   3. Copy the web app URL into the frontend's src/config.js
 *
 * The bound spreadsheet is the ONE shared sheet for the whole team.
 * Tabs "Contacts" and "Users" are created automatically.
 */

var MODEL = 'claude-sonnet-4-6';
var CONTACT_HEADERS = [
  'Serial Number', 'Date Added', 'Name of Person', 'Name of Company',
  'Company Sector', 'Phone Number', 'Email ID', 'Lead Owner'
];

/* ---------------- routing ---------------- */

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ error: 'Bad request' });
  }
  try {
    switch (req.action) {
      case 'signup':        return json_(signup_(req));
      case 'login':         return json_(login_(req));
      case 'listContacts':  return json_(listContacts_());
      case 'processCard':   return json_(processCard_(req));
      case 'updateContact': return json_(updateContact_(req));
      case 'deleteContact': return json_(deleteContact_(req));
      default:              return json_({ error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ error: String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  return json_({ ok: true, service: 'Skylark Cards to Excel backend' });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- sheets ---------------- */

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function contactsSheet_() { return sheet_('Contacts', CONTACT_HEADERS); }
function usersSheet_()    { return sheet_('Users', ['Username Key', 'Display Name', 'Password Hash', 'Created At']); }

/* ---------------- auth ---------------- */

function hash_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function signup_(req) {
  var name = String(req.username || '').trim();
  var pass = String(req.password || '');
  if (!name || !pass) return { error: 'Enter your full name and a password.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = usersSheet_();
    var key = name.toLowerCase();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === key) return { error: 'That name already has an account. Use Login instead.' };
    }
    sh.appendRow([key, name, hash_(pass), new Date().toISOString()]);
    return { ok: true, displayName: name };
  } finally {
    lock.releaseLock();
  }
}

function login_(req) {
  var name = String(req.username || '').trim().toLowerCase();
  var pass = String(req.password || '');
  var data = usersSheet_().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name && String(data[i][2]) === hash_(pass)) {
      return { ok: true, displayName: String(data[i][1]) };
    }
  }
  return { error: 'Username or password is incorrect.' };
}

/* ---------------- contacts ---------------- */

function rowToContact_(row) {
  return {
    serial: Number(row[0]),
    date: String(row[1]),
    name: String(row[2]),
    company: String(row[3]),
    sector: String(row[4]),
    phone: String(row[5]),
    email: String(row[6]),
    owner: String(row[7])
  };
}

function listContacts_() {
  var sh = contactsSheet_();
  var data = sh.getDataRange().getValues();
  var contacts = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== '' && data[i][0] !== null) contacts.push(rowToContact_(data[i]));
  }
  return { ok: true, contacts: contacts, sheetUrl: SpreadsheetApp.getActive().getUrl() };
}

function processCard_(req) {
  if (!req.imageB64) return { error: 'No image received' };
  var owner = String(req.owner || 'Unknown');
  var mediaType = String(req.mediaType || 'image/jpeg');

  var card = extractCard_(req.imageB64, mediaType);          // AI: read the card
  var sector = card.company ? findSector_(card.company) : ''; // AI: web search the sector

  var lock = LockService.getScriptLock();                     // serial numbers stay unique
  lock.waitLock(15000);
  try {
    var sh = contactsSheet_();
    var data = sh.getDataRange().getValues();
    var maxSerial = 0;
    var duplicate = false;
    var normPhone = card.phone.replace(/\D/g, '');
    for (var i = 1; i < data.length; i++) {
      var s = Number(data[i][0]);
      if (s > maxSerial) maxSerial = s;
      var exPhone = String(data[i][5]).replace(/\D/g, '');
      var exEmail = String(data[i][6]).toLowerCase();
      if ((card.email && exEmail && exEmail === card.email.toLowerCase()) ||
          (normPhone && exPhone && exPhone === normPhone)) duplicate = true;
    }
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy');
    var contact = {
      serial: maxSerial + 1, date: dateStr,
      name: card.name, company: card.company, sector: sector,
      phone: card.phone, email: card.email, owner: owner
    };
    sh.appendRow([contact.serial, contact.date, contact.name, contact.company,
                  contact.sector, contact.phone, contact.email, contact.owner]);
    return { ok: true, contact: contact, duplicate: duplicate };
  } finally {
    lock.releaseLock();
  }
}

function findRowBySerial_(sh, serial) {
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(serial)) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

function updateContact_(req) {
  var sh = contactsSheet_();
  var row = findRowBySerial_(sh, req.serial);
  if (row === -1) return { error: 'Contact not found' };
  var p = req.patch || {};
  var current = sh.getRange(row, 1, 1, 8).getValues()[0];
  var updated = [
    current[0], current[1],
    p.name !== undefined ? p.name : current[2],
    p.company !== undefined ? p.company : current[3],
    p.sector !== undefined ? p.sector : current[4],
    p.phone !== undefined ? p.phone : current[5],
    p.email !== undefined ? p.email : current[6],
    current[7]
  ];
  sh.getRange(row, 1, 1, 8).setValues([updated]);
  return { ok: true };
}

function deleteContact_(req) {
  var sh = contactsSheet_();
  var row = findRowBySerial_(sh, req.serial);
  if (row === -1) return { error: 'Contact not found' };
  sh.deleteRow(row);
  return { ok: true };
}

/* ---------------- Claude API ---------------- */

function anthropic_(payload) {
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('ANTHROPIC_API_KEY is not set in Script properties');
  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var body = JSON.parse(res.getContentText());
  if (code >= 400) {
    throw new Error('AI request failed: ' + (body.error && body.error.message ? body.error.message : code));
  }
  return body;
}

function extractCard_(b64, mediaType) {
  var data = anthropic_({
    model: MODEL,
    max_tokens: 1000,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mediaType, data: b64 } },
        { type: 'text', text: 'This is a photo of a business card. Extract the contact details and respond with ONLY a JSON object, no markdown fences, no other text: {"name": "person\'s full name", "company": "company name", "phone": "primary phone number", "email": "email address"}. Use an empty string "" for anything not present on the card. If the image is not a business card or is unreadable, respond with exactly {"error": "not a business card"}.' }
      ]
    }]
  });
  var text = (data.content || []).filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; }).join('\n');
  var clean = text.replace(/```json|```/g, '').trim();
  var s = clean.indexOf('{'), e = clean.lastIndexOf('}');
  if (s === -1 || e === -1) throw new Error('Could not read details from this image');
  var parsed = JSON.parse(clean.substring(s, e + 1));
  if (parsed.error) throw new Error("This image doesn't look like a business card");
  return {
    name: parsed.name || '', company: parsed.company || '',
    phone: parsed.phone || '', email: parsed.email || ''
  };
}

function findSector_(company) {
  try {
    var data = anthropic_({
      model: MODEL,
      max_tokens: 1000,
      messages: [{
        role: 'user',
        content: 'Search the web for the company "' + company + '" and identify which business sector / industry it operates in. Respond with ONLY the sector name in 2 to 5 words (e.g. "Renewable Energy", "Drone Surveying & GIS", "IT Services"). No other text, no punctuation at the end. If you cannot find the company, respond with exactly: Unknown'
      }],
      tools: [{ type: 'web_search_20250305', name: 'web_search' }]
    });
    var texts = (data.content || []).filter(function (b) { return b.type === 'text'; })
      .map(function (b) { return b.text.trim(); }).filter(Boolean);
    var last = texts.length ? texts[texts.length - 1] : '';
    var sector = last.split('\n').pop().trim();
    if (!sector || /unknown/i.test(sector) || sector.length > 60) return 'Unknown';
    return sector;
  } catch (err) {
    return 'Unknown';
  }
}
