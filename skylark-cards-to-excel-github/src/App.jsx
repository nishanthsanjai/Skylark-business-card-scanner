import React, { useState, useEffect, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { api, compressImage } from './api.js';
import { LOGO_SRC } from './logo.js';

/* ============================================================
   SKYLARK DRONES — Cards to Excel
   Business-card scanner → one shared Google Sheet for everyone
   ============================================================ */

const ORANGE = '#FF4F00';
const INK = '#1E2126';
const PAPER = '#FAF9F6';
const SESSION_KEY = 'skylark_cte_session';

/* ---------- Excel export (from the live shared data) ---------- */
function downloadExcel(contacts) {
  const rows = [...contacts]
    .sort((a, b) => a.serial - b.serial)
    .map((c) => ({
      'Serial Number': c.serial,
      'Date Added': c.date,
      'Name of Person': c.name,
      'Name of Company': c.company,
      'Company Sector': c.sector,
      'Phone Number': c.phone,
      'Email ID': c.email,
      'Lead Owner': c.owner,
    }));
  const ws = XLSX.utils.json_to_sheet(rows, {
    header: ['Serial Number', 'Date Added', 'Name of Person', 'Name of Company', 'Company Sector', 'Phone Number', 'Email ID', 'Lead Owner'],
  });
  ws['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 22 }, { wch: 24 }, { wch: 22 }, { wch: 16 }, { wch: 28 }, { wch: 18 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Contacts');
  XLSX.writeFile(wb, 'Skylark_Contacts.xlsx');
}

/* ---------- tiny UI atoms ---------- */
const styles = {
  shell: { minHeight: '100vh', background: PAPER, fontFamily: "'Open Sans', 'Segoe UI', system-ui, sans-serif", color: INK, display: 'flex', justifyContent: 'center' },
  frame: { width: '100%', maxWidth: 480, display: 'flex', flexDirection: 'column', minHeight: '100vh' },
  h: { fontFamily: "'Raleway', 'Segoe UI', system-ui, sans-serif", fontWeight: 800, letterSpacing: '-0.02em' },
};

function Logo({ size = 28 }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
      <img src={LOGO_SRC} alt="Skylark Drones logo" style={{ height: size * 1.15, width: 'auto', borderRadius: size * 0.22, display: 'block' }} />
      <div style={{ ...styles.h, fontSize: size * 0.55, lineHeight: 1 }}>
        SKYLARK <span style={{ color: ORANGE }}>DRONES</span>
      </div>
    </div>
  );
}

function Btn({ children, onClick, variant = 'solid', disabled, style: extra }) {
  const base = {
    border: 'none', borderRadius: 12, padding: '14px 20px', fontSize: 16, fontWeight: 700,
    fontFamily: "'Raleway', sans-serif", cursor: disabled ? 'not-allowed' : 'pointer',
    opacity: disabled ? 0.5 : 1, transition: 'transform 0.08s ease', width: '100%',
  };
  const variants = {
    solid: { background: ORANGE, color: '#fff' },
    dark: { background: INK, color: '#fff' },
    ghost: { background: 'transparent', color: INK, border: `2px solid ${INK}22` },
    danger: { background: 'transparent', color: '#C0392B', border: '2px solid #C0392B33' },
  };
  return (
    <button
      style={{ ...base, ...variants[variant], ...extra }}
      onClick={onClick}
      disabled={disabled}
      onMouseDown={(e) => !disabled && (e.currentTarget.style.transform = 'scale(0.98)')}
      onMouseUp={(e) => (e.currentTarget.style.transform = 'scale(1)')}
    >
      {children}
    </button>
  );
}

function Input({ label, type = 'text', value, onChange, autoFocus }) {
  return (
    <label style={{ display: 'block', marginBottom: 14 }}>
      <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 6, color: '#555', fontFamily: "'Raleway', sans-serif" }}>{label}</div>
      <input
        type={type}
        value={value}
        autoFocus={autoFocus}
        onChange={(e) => onChange(e.target.value)}
        style={{ width: '100%', boxSizing: 'border-box', padding: '13px 14px', fontSize: 16, borderRadius: 10, border: `2px solid ${INK}1F`, background: '#fff', outline: 'none' }}
        onFocus={(e) => (e.target.style.borderColor = ORANGE)}
        onBlur={(e) => (e.target.style.borderColor = `${INK}1F`)}
      />
    </label>
  );
}

/* ============================================================
   MAIN APP
   ============================================================ */
export default function App() {
  const [user, setUser] = useState(() => {
    try {
      const s = JSON.parse(localStorage.getItem(SESSION_KEY));
      return s && s.username ? { username: s.username } : null;
    } catch { return null; }
  });
  const [screen, setScreen] = useState(user ? 'home' : 'auth');
  const [contacts, setContacts] = useState([]);
  const [sheetUrl, setSheetUrl] = useState('');
  const [loadErr, setLoadErr] = useState('');

  /* auth state */
  const [authMode, setAuthMode] = useState(null);
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [authErr, setAuthErr] = useState('');
  const [authBusy, setAuthBusy] = useState(false);

  /* capture state */
  const [queue, setQueue] = useState([]); // [{b64, mediaType, preview}]
  const [askKeepClicking, setAskKeepClicking] = useState(false);
  const [jobs, setJobs] = useState([]);
  const [lastAdded, setLastAdded] = useState([]);

  const singleCamRef = useRef(null);
  const multiCamRef = useRef(null);
  const galleryRef = useRef(null);

  const refresh = useCallback(async () => {
    try {
      setLoadErr('');
      const r = await api.listContacts();
      setContacts(r.contacts || []);
      setSheetUrl(r.sheetUrl || '');
      return r.contacts || [];
    } catch (e) {
      setLoadErr(e.message);
      return [];
    }
  }, []);

  useEffect(() => { if (user) refresh(); }, [user, refresh]);

  /* ---------- auth ---------- */
  async function handleSignup() {
    setAuthErr('');
    if (!username.trim() || !password) return setAuthErr('Enter your full name and a password.');
    setAuthBusy(true);
    try {
      const r = await api.signup(username, password);
      localStorage.setItem(SESSION_KEY, JSON.stringify({ username: r.displayName }));
      setUser({ username: r.displayName });
      setPassword('');
      setScreen('home'); // straight to page 2 — no need to log in after sign-up
    } catch (e) {
      setAuthErr(e.message);
    } finally {
      setAuthBusy(false);
    }
  }

  async function handleLogin() {
    setAuthErr('');
    if (!username.trim() || !password) return setAuthErr('Enter your username and password.');
    setAuthBusy(true);
    try {
      const r = await api.login(username, password);
      localStorage.setItem(SESSION_KEY, JSON.stringify({ username: r.displayName }));
      setUser({ username: r.displayName });
      setPassword('');
      setScreen('home');
    } catch (e) {
      setAuthErr(e.message);
    } finally {
      setAuthBusy(false);
    }
  }

  function handleLogout() {
    localStorage.removeItem(SESSION_KEY);
    setUser(null);
    setUsername('');
    setPassword('');
    setAuthMode('login'); // login page opens as soon as someone logs out
    setScreen('auth');
  }

  /* ---------- capture flows ---------- */
  async function addFilesToQueue(fileList) {
    const files = Array.from(fileList || []);
    const items = [];
    for (const f of files) {
      try { items.push(await compressImage(f)); } catch { /* skip unreadable file */ }
    }
    setQueue((q) => [...q, ...items]);
    return items;
  }

  function onSingleShot(e) {
    const files = e.target.files;
    if (!files || !files.length) return;
    (async () => {
      const items = await addFilesToQueue(files);
      e.target.value = '';
      startProcessing(items); // single picture: uploads automatically
    })();
  }

  function onMultiShot(e) {
    const files = e.target.files;
    if (!files || !files.length) return;
    (async () => {
      await addFilesToQueue(files);
      e.target.value = '';
      setScreen('multiCapture');
      setAskKeepClicking(true); // "Keep clicking?"
    })();
  }

  function onGalleryPick(e) {
    const files = e.target.files;
    if (!files || !files.length) return;
    (async () => {
      await addFilesToQueue(files);
      e.target.value = '';
      setScreen('gallery');
    })();
  }

  function removeFromQueue(idx) {
    setQueue((q) => q.filter((_, i) => i !== idx));
  }

  /* ---------- processing pipeline ---------- */
  async function startProcessing(items) {
    const list = items && items.length ? items : queue;
    if (!list.length) return;
    setScreen('processing');
    setJobs(list.map((it, i) => ({ id: i, preview: it.preview, status: 'waiting', detail: 'Waiting…', result: null })));

    const added = [];
    for (let i = 0; i < list.length; i++) {
      setJobs((j) => j.map((x) => (x.id === i ? { ...x, status: 'reading', detail: 'Reading card & searching sector…' } : x)));
      try {
        const r = await api.processCard(list[i].b64, list[i].mediaType, user.username);
        added.push({ ...r.contact, duplicate: r.duplicate });
        setJobs((j) => j.map((x) => (x.id === i ? { ...x, status: 'done', detail: r.contact.sector || 'Done', result: r.contact } : x)));
      } catch (err) {
        setJobs((j) => j.map((x) => (x.id === i ? { ...x, status: 'error', detail: err.message || 'Failed to read this image' } : x)));
      }
    }

    setLastAdded(added);
    setQueue([]);
    await refresh();
    setScreen('final');
  }

  /* ============================================================ */
  return (
    <div style={styles.shell}>
      <style>{fontCss}</style>
      <div style={styles.frame}>
        <input ref={singleCamRef} type="file" accept="image/*" capture="environment" hidden onChange={onSingleShot} />
        <input ref={multiCamRef} type="file" accept="image/*" capture="environment" hidden onChange={onMultiShot} />
        <input ref={galleryRef} type="file" accept="image/*" multiple hidden onChange={onGalleryPick} />

        {screen === 'auth' && (
          <AuthScreen
            mode={authMode}
            setMode={(m) => { setAuthMode(m); setAuthErr(''); }}
            username={username} setUsername={setUsername}
            password={password} setPassword={setPassword}
            err={authErr} busy={authBusy}
            onSignup={handleSignup} onLogin={handleLogin}
          />
        )}

        {screen === 'home' && (
          <HomeScreen
            user={user}
            count={contacts.length}
            loadErr={loadErr}
            sheetUrl={sheetUrl}
            onLogout={handleLogout}
            onSingle={() => singleCamRef.current?.click()}
            onMulti={() => { setQueue([]); multiCamRef.current?.click(); }}
            onGallery={() => { setQueue([]); galleryRef.current?.click(); }}
            onViewSheet={async () => { await refresh(); setLastAdded([]); setScreen('final'); }}
            onDownload={async () => { const c = await refresh(); if (c.length) downloadExcel(c); }}
          />
        )}

        {screen === 'multiCapture' && (
          <MultiCaptureScreen
            queue={queue}
            asking={askKeepClicking}
            onKeepClicking={() => { setAskKeepClicking(false); multiCamRef.current?.click(); }}
            onDoneClicking={() => { setAskKeepClicking(false); startProcessing(); }}
            onRemove={removeFromQueue}
            onAddMore={() => multiCamRef.current?.click()}
            onScan={() => startProcessing()}
            onCancel={() => { setQueue([]); setAskKeepClicking(false); setScreen('home'); }}
          />
        )}

        {screen === 'gallery' && (
          <GalleryScreen
            queue={queue}
            onRemove={removeFromQueue}
            onAddMore={() => galleryRef.current?.click()}
            onScan={() => startProcessing()}
            onCancel={() => { setQueue([]); setScreen('home'); }}
          />
        )}

        {screen === 'processing' && <ProcessingScreen jobs={jobs} />}

        {screen === 'final' && (
          <FinalScreen
            contacts={contacts}
            lastAdded={lastAdded}
            sheetUrl={sheetUrl}
            onDownload={() => contacts.length && downloadExcel(contacts)}
            onBack={() => setScreen('home')}
            onEdit={async (serial, patch) => { await api.updateContact(serial, patch); await refresh(); }}
            onDelete={async (serial) => { await api.deleteContact(serial); await refresh(); }}
            onRefresh={refresh}
          />
        )}
      </div>
    </div>
  );
}

/* ============================================================
   SCREENS
   ============================================================ */

function AuthScreen({ mode, setMode, username, setUsername, password, setPassword, err, busy, onSignup, onLogin }) {
  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '48px 28px 32px' }}>
      <div style={{ display: 'flex', justifyContent: 'center' }}>
        <Logo size={32} />
      </div>

      <div style={{ marginTop: 72, textAlign: 'center' }}>
        <div style={{ ...styles.h, fontSize: 40, lineHeight: 1.05 }}>
          Cards to <span style={{ color: ORANGE }}>Excel</span>
        </div>
        <div style={{ marginTop: 10, color: '#666', fontSize: 15 }}>
          Snap a business card. It lands in the team sheet.
        </div>
      </div>

      {!mode && (
        <div style={{ marginTop: 56, textAlign: 'center', fontSize: 20, fontFamily: "'Raleway', sans-serif", fontWeight: 700 }}>
          <span role="button" tabIndex={0} onClick={() => setMode('signup')} onKeyDown={(e) => e.key === 'Enter' && setMode('signup')} style={{ color: ORANGE, textDecoration: 'underline', cursor: 'pointer' }}>
            Sign up
          </span>
          <span style={{ margin: '0 16px', color: '#bbb' }}>·</span>
          <span role="button" tabIndex={0} onClick={() => setMode('login')} onKeyDown={(e) => e.key === 'Enter' && setMode('login')} style={{ color: INK, textDecoration: 'underline', cursor: 'pointer' }}>
            Log in
          </span>
        </div>
      )}

      {mode && (
        <div style={{ marginTop: 40, background: '#fff', borderRadius: 16, padding: 22, boxShadow: '0 2px 16px rgba(0,0,0,0.06)' }}>
          <div style={{ ...styles.h, fontSize: 20, marginBottom: 16 }}>
            {mode === 'signup' ? 'Create your account' : 'Welcome back'}
          </div>
          <Input label={mode === 'signup' ? 'Full name (this becomes your Lead Owner name)' : 'Username (full name)'} value={username} onChange={setUsername} autoFocus />
          <Input label="Password" type="password" value={password} onChange={setPassword} />
          {err && <div style={{ color: '#C0392B', fontSize: 14, marginBottom: 12 }}>{err}</div>}
          <Btn onClick={mode === 'signup' ? onSignup : onLogin} disabled={busy}>
            {busy ? 'One moment…' : mode === 'signup' ? 'Sign up' : 'Log in'}
          </Btn>
          <div style={{ marginTop: 14, textAlign: 'center', fontSize: 14, color: '#666' }}>
            {mode === 'signup' ? (
              <>Already have an account?{' '}
                <span onClick={() => setMode('login')} style={{ color: ORANGE, cursor: 'pointer', fontWeight: 700 }}>Log in</span></>
            ) : (
              <>New here?{' '}
                <span onClick={() => setMode('signup')} style={{ color: ORANGE, cursor: 'pointer', fontWeight: 700 }}>Sign up</span></>
            )}
          </div>
        </div>
      )}

      <div style={{ marginTop: 'auto', textAlign: 'center', fontSize: 12, color: '#999', paddingTop: 24 }}>
        You stay signed in on this device until you log out.
      </div>
    </div>
  );
}

function HomeScreen({ user, count, loadErr, sheetUrl, onLogout, onSingle, onMulti, onGallery, onViewSheet, onDownload }) {
  const [pickCameraMode, setPickCameraMode] = useState(false);
  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '24px 24px 32px' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <Logo size={24} />
        <button onClick={onLogout} style={{ background: 'none', border: 'none', color: '#888', fontSize: 13, fontWeight: 700, cursor: 'pointer', fontFamily: "'Raleway', sans-serif" }}>
          Log out
        </button>
      </div>

      <div style={{ marginTop: 32 }}>
        <div style={{ fontSize: 14, color: '#888' }}>Hi {user.username.split(' ')[0]},</div>
        <div style={{ ...styles.h, fontSize: 28, marginTop: 4 }}>Add cards to the team sheet</div>
      </div>

      {loadErr && (
        <div style={{ marginTop: 16, background: '#FDEDEC', border: '2px solid #C0392B33', borderRadius: 12, padding: 12, fontSize: 13, color: '#922B21' }}>
          {loadErr}
        </div>
      )}

      <div style={{ marginTop: 24, background: INK, color: '#fff', borderRadius: 16, padding: 18 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
          <div style={{ fontFamily: "'Raleway', sans-serif", fontWeight: 800, fontSize: 16 }}>Team contact sheet</div>
          <div style={{ fontSize: 13, color: '#aaa' }}>{count} contact{count === 1 ? '' : 's'}</div>
        </div>
        <div style={{ display: 'flex', gap: 10, marginTop: 14 }}>
          <Btn variant="solid" onClick={onDownload} disabled={!count} style={{ padding: '11px 12px', fontSize: 14 }}>
            Download Excel
          </Btn>
          <Btn variant="ghost" onClick={onViewSheet} style={{ padding: '11px 12px', fontSize: 14, color: '#fff', border: '2px solid #ffffff33' }}>
            View sheet
          </Btn>
        </div>
        {sheetUrl && (
          <a href={sheetUrl} target="_blank" rel="noreferrer" style={{ display: 'block', marginTop: 12, fontSize: 13, color: ORANGE, fontWeight: 700, textDecoration: 'none' }}>
            Open the live Google Sheet ↗
          </a>
        )}
      </div>

      <div style={{ marginTop: 28, display: 'flex', flexDirection: 'column', gap: 14 }}>
        {!pickCameraMode ? (
          <>
            <BigAction icon="📷" title="Take pictures" sub="Use your camera to snap cards" onClick={() => setPickCameraMode(true)} />
            <BigAction icon="🖼️" title="Select from gallery" sub="Pick one or more photos you already have" onClick={onGallery} />
          </>
        ) : (
          <>
            <div style={{ fontSize: 14, fontWeight: 700, color: '#666', fontFamily: "'Raleway', sans-serif" }}>How many cards?</div>
            <BigAction icon="🃏" title="Take single picture" sub="Snap one card — it uploads automatically" onClick={() => { setPickCameraMode(false); onSingle(); }} />
            <BigAction icon="🗂️" title="Take multiple pictures" sub="Keep clicking cards one after another" onClick={() => { setPickCameraMode(false); onMulti(); }} />
            <button onClick={() => setPickCameraMode(false)} style={{ background: 'none', border: 'none', color: '#888', fontSize: 14, cursor: 'pointer', padding: 8 }}>
              ← Back
            </button>
          </>
        )}
      </div>

      <div style={{ marginTop: 'auto', fontSize: 12, color: '#999', textAlign: 'center', paddingTop: 24 }}>
        Your phone will ask permission to use the camera the first time.
      </div>
    </div>
  );
}

function BigAction({ icon, title, sub, onClick }) {
  return (
    <button
      onClick={onClick}
      style={{ display: 'flex', alignItems: 'center', gap: 16, background: '#fff', border: `2px solid ${INK}14`, borderRadius: 16, padding: '18px 18px', cursor: 'pointer', textAlign: 'left', width: '100%' }}
    >
      <div style={{ fontSize: 28 }}>{icon}</div>
      <div>
        <div style={{ fontFamily: "'Raleway', sans-serif", fontWeight: 800, fontSize: 17, color: INK }}>{title}</div>
        <div style={{ fontSize: 13, color: '#777', marginTop: 2 }}>{sub}</div>
      </div>
      <div style={{ marginLeft: 'auto', color: ORANGE, fontSize: 22, fontWeight: 700 }}>›</div>
    </button>
  );
}

function Thumbs({ queue, onRemove }) {
  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10, marginTop: 16 }}>
      {queue.map((it, i) => (
        <div key={i} style={{ position: 'relative', borderRadius: 12, overflow: 'hidden', aspectRatio: '1', border: `2px solid ${INK}14` }}>
          <img src={it.preview} alt={`Card ${i + 1}`} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
          <button
            onClick={() => onRemove(i)}
            aria-label="Remove picture"
            style={{ position: 'absolute', top: 6, right: 6, width: 26, height: 26, borderRadius: '50%', background: 'rgba(0,0,0,0.65)', color: '#fff', border: 'none', cursor: 'pointer', fontSize: 14, lineHeight: '26px', padding: 0 }}
          >
            ✕
          </button>
        </div>
      ))}
    </div>
  );
}

function MultiCaptureScreen({ queue, asking, onKeepClicking, onDoneClicking, onRemove, onAddMore, onScan, onCancel }) {
  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '24px 24px 32px' }}>
      <div style={{ ...styles.h, fontSize: 24 }}>Cards clicked so far</div>
      <div style={{ color: '#777', fontSize: 14, marginTop: 4 }}>{queue.length} picture{queue.length === 1 ? '' : 's'} — tap ✕ to remove any</div>

      <Thumbs queue={queue} onRemove={onRemove} />

      {!asking && (
        <div style={{ marginTop: 20, display: 'flex', flexDirection: 'column', gap: 10 }}>
          <Btn variant="ghost" onClick={onAddMore}>+ Click another card</Btn>
          <Btn onClick={onScan} disabled={!queue.length}>Scan &amp; upload {queue.length ? `(${queue.length})` : ''}</Btn>
          <Btn variant="danger" onClick={onCancel}>Cancel</Btn>
        </div>
      )}

      {asking && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24, zIndex: 10 }}>
          <div style={{ background: '#fff', borderRadius: 18, padding: 24, width: '100%', maxWidth: 340, textAlign: 'center' }}>
            <div style={{ ...styles.h, fontSize: 22 }}>Keep clicking?</div>
            <div style={{ color: '#777', fontSize: 14, marginTop: 6 }}>
              {queue.length} card{queue.length === 1 ? '' : 's'} captured so far
            </div>
            <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
              <Btn variant="ghost" onClick={onDoneClicking}>No, upload</Btn>
              <Btn onClick={onKeepClicking}>Yes</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function GalleryScreen({ queue, onRemove, onAddMore, onScan, onCancel }) {
  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '24px 24px 110px', position: 'relative' }}>
      <div style={{ ...styles.h, fontSize: 24 }}>Selected from gallery</div>
      <div style={{ color: '#777', fontSize: 14, marginTop: 4 }}>Tap ✕ on a picture to remove it from the selection</div>

      <Thumbs queue={queue} onRemove={onRemove} />

      <div style={{ marginTop: 18, display: 'flex', gap: 10 }}>
        <Btn variant="ghost" onClick={onAddMore}>+ Add more</Btn>
        <Btn variant="danger" onClick={onCancel}>Cancel</Btn>
      </div>

      {queue.length > 0 && (
        <button
          onClick={onScan}
          style={{ position: 'fixed', bottom: 24, right: 24, zIndex: 10, background: ORANGE, color: '#fff', border: 'none', borderRadius: 999, padding: '16px 24px', fontSize: 16, fontWeight: 800, fontFamily: "'Raleway', sans-serif", boxShadow: '0 6px 20px rgba(255,79,0,0.4)', cursor: 'pointer' }}
        >
          Start Scanning ({queue.length})
        </button>
      )}
    </div>
  );
}

function ProcessingScreen({ jobs }) {
  const done = jobs.filter((j) => j.status === 'done' || j.status === 'error').length;
  return (
    <div style={{ flex: 1, padding: '32px 24px' }}>
      <div style={{ ...styles.h, fontSize: 24 }}>Scanning cards…</div>
      <div style={{ color: '#777', fontSize: 14, marginTop: 4 }}>{done} of {jobs.length} finished — details upload automatically</div>

      <div style={{ marginTop: 8, height: 6, background: `${INK}12`, borderRadius: 3 }}>
        <div style={{ height: '100%', width: `${jobs.length ? (done / jobs.length) * 100 : 0}%`, background: ORANGE, borderRadius: 3, transition: 'width 0.4s ease' }} />
      </div>

      <div style={{ marginTop: 20, display: 'flex', flexDirection: 'column', gap: 12 }}>
        {jobs.map((j) => (
          <div key={j.id} style={{ display: 'flex', gap: 12, alignItems: 'center', background: '#fff', borderRadius: 14, padding: 12, border: `2px solid ${INK}0F` }}>
            <img src={j.preview} alt="" style={{ width: 52, height: 52, borderRadius: 10, objectFit: 'cover' }} />
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 700, fontSize: 15, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {j.result?.name || (j.status === 'error' ? "Couldn't read this card" : 'Reading…')}
              </div>
              <div style={{ fontSize: 13, color: j.status === 'error' ? '#C0392B' : '#777', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {j.status === 'done' ? `${j.result.company || '—'} · ${j.result.sector || '—'}` : j.detail}
              </div>
            </div>
            <div style={{ fontSize: 18 }}>
              {j.status === 'done' ? '✅' : j.status === 'error' ? '⚠️' : <Spinner />}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function Spinner() {
  return (
    <span style={{ display: 'inline-block', width: 18, height: 18, border: `3px solid ${INK}22`, borderTopColor: ORANGE, borderRadius: '50%', animation: 'cte-spin 0.8s linear infinite' }} />
  );
}

function FinalScreen({ contacts, lastAdded, sheetUrl, onDownload, onBack, onEdit, onDelete, onRefresh }) {
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [busy, setBusy] = useState(false);
  const sorted = [...contacts].sort((a, b) => b.serial - a.serial);
  const dupCount = lastAdded.filter((c) => c.duplicate).length;

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '24px 20px 32px' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <button onClick={onBack} style={{ background: 'none', border: 'none', color: ORANGE, fontWeight: 800, fontSize: 15, cursor: 'pointer', fontFamily: "'Raleway', sans-serif", padding: 0 }}>
          ← Scan more
        </button>
        <button onClick={onRefresh} style={{ background: 'none', border: 'none', color: '#888', fontSize: 13, cursor: 'pointer' }}>
          ⟳ Refresh
        </button>
      </div>

      {lastAdded.length > 0 && (
        <div style={{ marginTop: 16, background: '#E9F7EF', border: '2px solid #27AE6033', borderRadius: 14, padding: 14, fontSize: 14 }}>
          ✅ <b>{lastAdded.length}</b> contact{lastAdded.length === 1 ? '' : 's'} added to the team sheet.
          {dupCount > 0 && (
            <div style={{ marginTop: 6, color: '#B7791F' }}>
              ⚠️ {dupCount} may be {dupCount === 1 ? 'a duplicate' : 'duplicates'} (matching phone/email already in the sheet) — worth a quick check below.
            </div>
          )}
        </div>
      )}

      <div style={{ marginTop: 18, display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
        <div style={{ ...styles.h, fontSize: 22 }}>Team contact sheet</div>
        <div style={{ fontSize: 13, color: '#888' }}>{sorted.length} total</div>
      </div>
      <div style={{ fontSize: 13, color: '#888', marginTop: 2 }}>One sheet, shared by everyone. Tap a contact to edit it.</div>

      <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
        <Btn onClick={onDownload} disabled={!sorted.length}>⬇ Download Excel (.xlsx)</Btn>
        {sheetUrl && (
          <a href={sheetUrl} target="_blank" rel="noreferrer" style={{ textAlign: 'center', fontSize: 14, color: ORANGE, fontWeight: 700, textDecoration: 'none', padding: 6 }}>
            Open the live Google Sheet ↗
          </a>
        )}
      </div>

      <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 10 }}>
        {sorted.length === 0 && (
          <div style={{ textAlign: 'center', color: '#999', padding: '40px 0', fontSize: 14 }}>
            No contacts yet. Scan your first card to start the sheet.
          </div>
        )}
        {sorted.map((c) => (
          <div key={c.serial} onClick={() => setEditing({ ...c })} style={{ background: '#fff', borderRadius: 14, padding: '12px 14px', border: `2px solid ${INK}0F`, cursor: 'pointer' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8 }}>
              <div style={{ fontWeight: 800, fontFamily: "'Raleway', sans-serif", fontSize: 15 }}>
                #{c.serial} · {c.name || '—'}
              </div>
              <div style={{ fontSize: 12, color: '#999', whiteSpace: 'nowrap' }}>{c.date}</div>
            </div>
            <div style={{ fontSize: 13, color: '#555', marginTop: 3 }}>
              {c.company || '—'}{c.sector ? ` · ${c.sector}` : ''}
            </div>
            <div style={{ fontSize: 13, color: '#777', marginTop: 3 }}>
              {c.phone || '—'} · {c.email || '—'}
            </div>
            <div style={{ fontSize: 12, color: ORANGE, marginTop: 4, fontWeight: 700 }}>Lead owner: {c.owner}</div>
          </div>
        ))}
      </div>

      {editing && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20, zIndex: 20 }}>
          <div style={{ background: '#fff', borderRadius: 18, padding: 22, width: '100%', maxWidth: 400, maxHeight: '88vh', overflowY: 'auto' }}>
            <div style={{ ...styles.h, fontSize: 20, marginBottom: 14 }}>Edit contact #{editing.serial}</div>
            <Input label="Name of person" value={editing.name} onChange={(v) => setEditing({ ...editing, name: v })} />
            <Input label="Company" value={editing.company} onChange={(v) => setEditing({ ...editing, company: v })} />
            <Input label="Company sector" value={editing.sector} onChange={(v) => setEditing({ ...editing, sector: v })} />
            <Input label="Phone number" value={editing.phone} onChange={(v) => setEditing({ ...editing, phone: v })} />
            <Input label="Email ID" value={editing.email} onChange={(v) => setEditing({ ...editing, email: v })} />
            <div style={{ display: 'flex', gap: 10, marginTop: 6 }}>
              <Btn variant="ghost" onClick={() => setEditing(null)} disabled={busy}>Cancel</Btn>
              <Btn disabled={busy} onClick={async () => {
                setBusy(true);
                try {
                  const { serial, ...patch } = editing;
                  await onEdit(serial, patch);
                  setEditing(null);
                } finally { setBusy(false); }
              }}>{busy ? 'Saving…' : 'Save changes'}</Btn>
            </div>
            <div style={{ marginTop: 12, textAlign: 'center' }}>
              {confirmDelete === editing.serial ? (
                <Btn variant="danger" disabled={busy} onClick={async () => {
                  setBusy(true);
                  try {
                    await onDelete(editing.serial);
                    setConfirmDelete(null);
                    setEditing(null);
                  } finally { setBusy(false); }
                }}>
                  Tap again to permanently delete
                </Btn>
              ) : (
                <button onClick={() => setConfirmDelete(editing.serial)} style={{ background: 'none', border: 'none', color: '#C0392B', fontSize: 13, cursor: 'pointer' }}>
                  Delete this contact
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const fontCss = `
@import url('https://fonts.googleapis.com/css2?family=Raleway:wght@600;700;800&family=Open+Sans:wght@400;600;700&display=swap');
@keyframes cte-spin { to { transform: rotate(360deg); } }
* { -webkit-tap-highlight-color: transparent; }
`;
