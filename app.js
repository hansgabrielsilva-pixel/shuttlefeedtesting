// =====================================================
// SHUTTLEFEED front-end - data is stored in Google Sheets
// through a Google Apps Script web app (see Code.gs)
// =====================================================

// Only needed if you host this page OUTSIDE Apps Script. Leave '' when the page is served by Apps Script.
const CONFIG = { API_URL: 'PASTE_YOUR_EXEC_URL_HERE' };

// Pictures. Put a Google Drive file ID (file shared as "Anyone with the link") or a full image URL.
// Leave '' to hide a picture.
const IMAGES = { logo: '1_tJv4EeRAP_WyFMjXsHRqyzm4bBJaAEK', bg: '1MXtWPB7D8PBoEeXk__7uHozS54Nq1yTV', shuttles: '1uWWlKl0OhlUI--cdqUUH2IzNR837nafW', smash: '1ZH1nP35P8yAUaYIX5Wliwfk84XzK1Oe2', service: '1WkSogeesZHrtIobvtie-tZV5alD1EY1H' };
const TOKEN_KEY = 'shuttlefeedToken';

// --- Light / dark mode (saved on this device; the <head> script applies it before the page paints) ---
const THEME_KEY = 'shuttlefeedTheme';
function currentTheme() { return document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light'; }
function syncThemeUi() {
    const dark = currentTheme() === 'dark';
    setText('theme-label', dark ? 'Dark mode' : 'Light mode');
    const btn = document.getElementById('theme-btn');
    if (btn) { btn.setAttribute('aria-checked', dark ? 'true' : 'false'); btn.title = dark ? 'Switch to light mode' : 'Switch to dark mode'; }
    const tile = document.querySelector('.theme-tile');
    if (tile) tile.setAttribute('aria-checked', dark ? 'true' : 'false');
}
let themeAnimTimer = null;
function applyTheme(theme) {
    const root = document.documentElement;
    // brief cross-fade of colors while switching (skipped if the user prefers reduced motion)
    if (!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches)) {
        root.classList.add('theme-anim');
        clearTimeout(themeAnimTimer);
        themeAnimTimer = setTimeout(() => root.classList.remove('theme-anim'), 450);
    }
    if (theme === 'dark') root.setAttribute('data-theme', 'dark');
    else root.removeAttribute('data-theme');
    try { localStorage.setItem(THEME_KEY, theme); } catch (e) {}
    syncThemeUi();
}
function toggleTheme() { applyTheme(currentTheme() === 'dark' ? 'light' : 'dark'); }

// SIMULATION VARIABLES
let simulationTimer;
let shuttlesLaunched = 0;
let totalShuttles = 50;
let selectedRange = 'MEDIUM';           // SHORT | MEDIUM | LONG - how far the shuttle is thrown
const RANGE_LABELS = { SHORT: 'Short', MEDIUM: 'Medium', LONG: 'Long' };
let currentMode = 'Smash';

// --- Dashboard: last drill (for Quick Start) ---
const DEFAULT_DRILL = { mode: 'Smash', range: 'MEDIUM', interval: 3, count: 24, direction: 0, angle: 35 };

// --- Smash and Service drills are fixed: 3 seconds between shuttles, 24 shuttles. Remove a mode from `modes` to unlock it. ---
const LOCKED_DRILL = { modes: ['Smash', 'Service'], interval: 3, count: 24 };
// Old sessions were saved as 'Low Service' / 'High Service'; there is now a single 'Service' mode
function normMode(m) { return /service/i.test(m || '') ? 'Service' : (m || 'Smash'); }
function isLockedMode() { return LOCKED_DRILL.modes.includes(currentMode); }
function getInterval() { return isLockedMode() ? LOCKED_DRILL.interval : (parseFloat(document.getElementById('interval').value) || 3); }
// --- Launch settings lock: once START is pressed, range and direction can't be changed until the session is stopped or finished ---
let controlsLocked = false;
function setControlsLocked(on) {
    controlsLocked = !!on;
    const card = document.querySelector('.params-card');
    if (card) card.classList.toggle('locked', controlsLocked);
    const pill = document.getElementById('lock-pill');
    if (pill) pill.classList.toggle('hidden', !controlsLocked);
}
function blockedWhileLocked() {
    if (!controlsLocked) return false;
    toast('Settings are locked while launching. Press STOP to change them.');
    return true;
}
function getCount() { return isLockedMode() ? LOCKED_DRILL.count : (parseInt(document.getElementById('count').value) || 50); }
let lastDrill = { ...DEFAULT_DRILL };

// --- Logged-in instructor + data loaded from the server ---
let currentUser = null;
let currentFullName = '';             // shown in the dashboard greeting
let authToken = null;
let historyCache = [];
let statsCache = { sessions: 0, shuttles: 0 };
let activeSession = null;                 // Promise -> server id of the running session
let lastSession = null;                   // same promise, kept after the session ends so late score fixes can still be saved
let shotLog = [];                         // one entry per launched shuttle: null (not scored) | 'hit' | 'miss'
let resultsSaveTimer = null;
let pendingFinish = Promise.resolve();    // resolves when the last session has been saved

// --- Small on-screen message (replaces alert()) ---
function toast(msg, ms = 3200) {
    let t = document.getElementById('sf-toast');
    if (!t) {
        t = document.createElement('div');
        t.id = 'sf-toast';
        t.style.cssText = 'position:fixed;left:50%;bottom:30px;transform:translateX(-50%);background:var(--sec-green);color:#fff;padding:12px 22px;border-radius:25px;font-size:15px;z-index:2000;box-shadow:0 6px 20px rgba(0,0,0,.3);transition:opacity .3s;max-width:90%;text-align:center;pointer-events:none;opacity:0;';
        document.body.appendChild(t);
    }
    t.textContent = msg;
    t.style.opacity = '1';
    clearTimeout(t._hide);
    t._hide = setTimeout(() => { t.style.opacity = '0'; }, ms);
}

// --- API layer: talks to Code.gs (google.script.run inside Apps Script, fetch if hosted elsewhere) ---
function apiConfigured() { return !!CONFIG.API_URL && !CONFIG.API_URL.startsWith('PASTE'); }
function hasAppsScript() { return typeof google !== 'undefined' && google.script && google.script.run; }
function hasServer() { return hasAppsScript() || apiConfigured(); }

function callServer(req) {
    return new Promise((resolve, reject) => {
        if (hasAppsScript()) {
            google.script.run
                .withSuccessHandler(resolve)
                .withFailureHandler(err => reject(new Error('Cannot reach the server. ' + ((err && err.message) || ''))))
                .apiCall(req);
        } else if (apiConfigured()) {
            // text/plain keeps this a "simple" request, so the browser skips the CORS preflight Apps Script can't answer
            fetch(CONFIG.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(req) })
                .then(r => r.json()).then(resolve)
                .catch(() => reject(new Error('Cannot reach the server. Check the Wi-Fi connection.')));
        } else {
            reject(new Error('Not connected to the server.'));
        }
    });
}

async function api(action, payload = {}) {
    let timer;
    const timeout = new Promise((_, rej) => { timer = setTimeout(() => rej(new Error('The server took too long to respond.')), 25000); });
    let data;
    try {
        data = await Promise.race([callServer({ token: authToken, ...payload, action }), timeout]);
    } finally {
        clearTimeout(timer);
    }
    if (!data || typeof data !== 'object') throw new Error('Unexpected response from the server.');
    if (!data.ok) {
        if (data.code === 'AUTH' && authToken) handleAuthExpired();
        throw new Error(data.error || 'Request failed.');
    }
    return data;
}

// --- Pictures (Drive file IDs or URLs from IMAGES above) ---
function imgUrl(v) {
    if (!v) return '';
    return /^(https?:|data:)/.test(v) ? v : 'https://lh3.googleusercontent.com/d/' + v;
}
function applyImages() {
    document.querySelectorAll('img[data-img]').forEach(img => {
        const u = imgUrl(IMAGES[img.dataset.img]);
        if (u) { img.onerror = () => { img.style.display = 'none'; }; img.src = u; }
    });
    const bg = imgUrl(IMAGES.bg);
    if (bg) document.documentElement.style.setProperty('--bg-img', `url("${bg}")`);
}

function applyUserData(d) {
    d = d || {};
    lastDrill = { ...DEFAULT_DRILL, ...(d.lastDrill || {}) };
    lastClass = d.lastClass || null;
    historyCache = d.history || [];
    statsCache = d.stats || { sessions: 0, shuttles: 0 };
}

function resetLocalState() {
    currentUser = null;
    authToken = null;
    currentClass = null;
    lastClass = null;
    activeSession = null;
    lastSession = null;
    shotLog = [];
    clearTimeout(resultsSaveTimer);
    lastDrill = { ...DEFAULT_DRILL };
    historyCache = [];
    statsCache = { sessions: 0, shuttles: 0 };
    try { localStorage.removeItem(TOKEN_KEY); } catch (e) {}
    setUserName('');
    ['username', 'password', 'new-fullname', 'new-username', 'new-password'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
}

function handleAuthExpired() {
    resetLocalState();
    navigate('login-screen');
    toast('Session expired. Please log in again.');
}

function logout() {
    const token = authToken;
    navigate('login-screen');                 // also stops the machine and closes any running session
    const closing = pendingFinish;
    resetLocalState();
    if (token) closing.then(() => api('logout', { token })).catch(() => {});
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.innerText = text;
}

// Title used in the greeting when the instructor's name doesn't start with one (e.g. "Dr. Ana Reyes" keeps "Dr.")
const DEFAULT_TITLE = "Ma'am";

function setUserName(fullName, fallback) {
    currentFullName = (fullName || '').trim();
    setText('sidebar-name', currentFullName || fallback || 'Instructor');
}

// "Maria Santos Reyes" -> "Ma'am Reyes"   |   "Reyes, Maria" -> "Ma'am Reyes"   |   "Sir Jose Cruz" -> "Sir Cruz"
function greetingName() {
    let n = currentFullName;
    if (!n) return 'Instructor';
    let title = DEFAULT_TITLE;
    const m = n.match(/^(ma'?am|madam|sir|mr|mrs|ms|prof|dr|engr|atty)\.?\s+/i);
    if (m) {
        const t = m[1].toLowerCase();
        title = /^ma'?am$/.test(t) ? "Ma'am" : t.charAt(0).toUpperCase() + t.slice(1) + (/^(madam|sir)$/.test(t) ? '' : '.');
        n = n.slice(m[0].length).trim();
    }
    let last;
    if (n.includes(',')) last = n.split(',')[0].trim();
    else {
        const parts = n.split(/\s+/).filter(p => p && !/^(jr|sr|ii|iii|iv)\.?$/i.test(p));
        last = parts[parts.length - 1];
    }
    return last ? `${title} ${last}` : 'Instructor';
}

function setTextIfChanged(id, text) {
    const el = document.getElementById(id);
    if (el && el.textContent !== text) el.textContent = text;
}

// Kiosk clock + date + greeting (the greeting also flips at noon and 6 PM)
function updateClock() {
    const now = new Date();
    const hour = now.getHours();
    const part = hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
    setTextIfChanged('dash-greeting-text', `${part}, ${greetingName()}`);
    setTextIfChanged('clock-hm', `${hour % 12 || 12}:${String(now.getMinutes()).padStart(2, '0')}`);
    setTextIfChanged('clock-ap', hour < 12 ? 'AM' : 'PM');
    setTextIfChanged('dash-date', now.toLocaleDateString('en-PH', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }));
}
function tickClock() {
    const screen = document.getElementById('dashboard-screen');
    if (screen && screen.classList.contains('active')) updateClock();
}

// Number card: counts up only when the value actually changes
function setCount(id, value) {
    const el = document.getElementById(id);
    if (!el || el.dataset.count === String(value)) return;
    el.dataset.count = value;
    animateCounts(el.parentElement);
}

function refreshDashboard() {
    updateClock();
    setText('qs-detail', `${normMode(lastDrill.mode)} · ${RANGE_LABELS[lastDrill.range] || ''} · ${LOCKED_DRILL.modes.includes(normMode(lastDrill.mode)) ? LOCKED_DRILL.count : lastDrill.count} shuttles`);
    setText('class-chip-label', currentClass ? `${classLabel(currentClass)} · PHED 4` : 'Set your class');
    setCount('dash-sessions', statsCache.sessions || 0);
    setCount('dash-shuttles', statsCache.shuttles || 0);
    let hits = 0, misses = 0;
    (getHistory() || []).forEach(it => { hits += it.hits || 0; misses += it.misses || 0; });
    setText('dash-accuracy', (hits + misses) ? Math.round(hits / (hits + misses) * 100) + '%' : '—');
    renderHistory();
}

// --- Session logging (saved to the Sessions sheet) ---
function beginLoggedSession(intervalSec) {
    if (activeSession) finishLoggedSession('STOPPED');
    resetShotLog();                      // a fresh START begins a fresh score sheet
    lastDrill = {
        mode: currentMode, range: selectedRange, interval: intervalSec, count: totalShuttles,
        direction: parseInt(document.getElementById('h-dir').value) || 0,
        angle: parseInt(document.getElementById('v-ang').value) || 0
    };
    refreshDashboard();
    activeSession = api('startSession', {
        cls: classLabel(currentClass), mode: lastDrill.mode, range: lastDrill.range,
        direction: lastDrill.direction, angle: lastDrill.angle,
        interval: intervalSec, planned: totalShuttles
    }).then(r => r.id).catch(e => { toast('Session not saved: ' + e.message); return null; });
    lastSession = activeSession;
}

// Called on STOP / completion. Saves how many shuttles were ACTUALLY launched.
function finishLoggedSession(status) {
    if (!activeSession) return;
    const pending = activeSession;
    const token = authToken;
    const launched = shuttlesLaunched;
    const { hits, misses } = shotCounts();
    activeSession = null;
    pendingFinish = (async () => {
        const id = await pending;
        if (!id) return;
        try {
            const r = await api('endSession', { token, id, launched, status, hits, misses });
            if (token === authToken) { applyUserData(r.data); refreshDashboard(); }
        } catch (e) {
            toast('Could not save the session: ' + e.message);
        }
    })();
}

// Loads the last drill's settings and opens the control screen (instructor still presses START)
function quickStart() {
    document.getElementById('interval').value = lastDrill.interval;
    document.getElementById('count').value = lastDrill.count;
    document.getElementById('h-dir').value = lastDrill.direction ?? 0;
    document.getElementById('v-ang').value = lastDrill.angle ?? 30;
    setupSession(normMode(lastDrill.mode));
    setRange(lastDrill.range);
}

// --- Class details (program / year / section) ---
let currentClass = null;
let lastClass = null;
let classSetupFromDashboard = false;

function classLabel(c) { return c ? `${c.program} ${c.year}-${c.section}` : 'No class set'; }

function togglePassword() {
    const input = document.getElementById('password');
    const btn = document.getElementById('toggle-password');
    const show = input.type === 'password';
    input.type = show ? 'text' : 'password';
    btn.querySelector('.eye-on').style.display = show ? 'none' : '';
    btn.querySelector('.eye-off').style.display = show ? '' : 'none';
    btn.classList.toggle('showing', show);
    btn.setAttribute('aria-label', show ? 'Hide password' : 'Show password');
    btn.title = show ? 'Hide password' : 'Show password';
    input.focus();
}

function prepareClassForm() {
    document.getElementById('class-year').value = '2';   // year level is fixed to 2nd Year
    const hasLast = !!lastClass;
    document.getElementById('continue-class').classList.toggle('hidden', !hasLast);
    document.getElementById('or-divider').classList.toggle('hidden', !hasLast);
    document.getElementById('class-dots').classList.toggle('hidden', classSetupFromDashboard);
    document.getElementById('class-error').innerText = '';
    if (hasLast) document.getElementById('continue-class-label').innerText = classLabel(lastClass);
}

let authBusy = false;
function setAuthButtons(disabled) {
    document.querySelectorAll('#login-screen .btn-primary, #signup-screen .btn-primary').forEach(b => { b.disabled = disabled; });
}

// --- Login animation: a shuttlecock rally between two rackets ---
const SHUTTLE_SVG = '<svg viewBox="0 0 40 64" xmlns="http://www.w3.org/2000/svg"><path d="M14 44 L3 4 Q20 -3 37 4 L26 44 Z" fill="#fff" stroke="#C9A98D" stroke-width="1.5" stroke-linejoin="round"/><path d="M14 44 L9 3 M20 44 L20 0 M26 44 L31 3" stroke="#C9A98D" stroke-width="1.2" fill="none"/><path d="M6 15 Q20 10 34 15" stroke="#C9A98D" stroke-width="1.2" fill="none"/><ellipse cx="20" cy="52" rx="9" ry="9" fill="#7C0005"/><path d="M12 45 Q20 50 28 45" stroke="#4d0003" stroke-width="1.5" fill="none"/></svg>';
const RACKET_SVG = '<svg viewBox="0 0 40 80" xmlns="http://www.w3.org/2000/svg"><defs><clipPath id="rk-clip"><ellipse cx="20" cy="21" rx="14" ry="18"/></clipPath></defs><g clip-path="url(#rk-clip)" stroke="#C9A98D" stroke-width="1"><path d="M13 0 V44 M20 0 V44 M27 0 V44 M0 12 H40 M0 21 H40 M0 30 H40"/></g><ellipse cx="20" cy="21" rx="14" ry="18" fill="none" stroke="#7C0005" stroke-width="3.5"/><path d="M20 39 V50" stroke="#7C0005" stroke-width="3" stroke-linecap="round"/><rect x="17" y="50" width="6" height="28" rx="3" fill="#2C3812"/></svg>';
const RALLY_MS = 3700, RALLY_PASS_S = RALLY_MS / 4000;   // 4 passes across the screen in total
let rallyTimer = null, rallyRAF = null, introDone = null;

function playLoginAnimation(title, done) {
    const overlay = document.getElementById('rally-overlay');
    const reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (!overlay || reduce) { done(); return; }
    document.getElementById('rally-title').innerText = title;
    overlay.querySelectorAll('.rally-stage').forEach(n => n.remove());

    const H = window.innerHeight, W = window.innerWidth;
    const half = Math.min(W * 0.3, 340);                 // half the width of the rally
    const x0 = W / 2 - half, x1 = W / 2 + half;          // where each racket meets the shuttle
    const base = H * 0.7;                                // height of the hit
    const apex = Math.min(H * 0.34, 240);                // how high each pass arcs
    const sw = W < 600 ? 30 : 38, sh = sw * 1.6;         // shuttle size

    const stage = document.createElement('div');
    stage.className = 'rally-stage';
    const ground = document.createElement('div');
    ground.className = 'rally-ground';
    ground.style.top = Math.round(base + 72) + 'px';
    const mkRacket = (cx, tilt, kick) => {               // cx = where the racket head sits
        const r = document.createElement('div');
        r.className = 'rally-racket';
        r.style.setProperty('--tilt', tilt + 'deg');
        r.style.setProperty('--kick', kick + 'deg');
        r.style.left = Math.round(cx - 23 + (tilt > 0 ? -24 : 24)) + 'px';
        r.style.top = Math.round(base - 20) + 'px';
        r.innerHTML = RACKET_SVG;
        return r;
    };
    const racketL = mkRacket(x0 - 14, 25, 35), racketR = mkRacket(x1 + 14, -25, -35);
    const shuttle = document.createElement('div');
    shuttle.className = 'rally-shuttle';
    shuttle.style.width = sw + 'px';
    shuttle.innerHTML = SHUTTLE_SVG;
    stage.append(ground, racketL, racketR, shuttle);
    overlay.appendChild(stage);

    const t0 = performance.now();
    let lastPass = -1, lastTrail = 0;
    function frame(now) {
        const t = (now - t0) / 1000;
        const pass = Math.floor(t / RALLY_PASS_S);
        const u = t / RALLY_PASS_S - pass;               // 0..1 progress of this pass
        const rightward = pass % 2 === 0;
        const xa = rightward ? x0 : x1, xb = rightward ? x1 : x0;
        const x = xa + (xb - xa) * u;
        const y = base - 4 * apex * u * (1 - u);         // parabola: hit -> peak -> hit
        const vx = xb - xa, vy = -4 * apex * (1 - 2 * u);
        const ang = Math.atan2(-vx, vy) * 180 / Math.PI; // cork leads the way, like a real shuttle
        shuttle.style.transform = 'translate(' + (x - sw / 2).toFixed(1) + 'px,' + (y - sh / 2).toFixed(1) + 'px) rotate(' + ang.toFixed(1) + 'deg)';
        if (pass !== lastPass) {                         // a new hit: swing the racket that sent it
            lastPass = pass;
            const r = rightward ? racketL : racketR;
            r.classList.remove('swing'); void r.offsetWidth; r.classList.add('swing');
        }
        if (now - lastTrail > 45) {                      // little fading dots behind the shuttle
            lastTrail = now;
            const d = document.createElement('div');
            d.className = 'rally-trail';
            d.style.left = x.toFixed(1) + 'px'; d.style.top = y.toFixed(1) + 'px';
            d.addEventListener('animationend', () => d.remove());
            stage.insertBefore(d, shuttle);
        }
        rallyRAF = requestAnimationFrame(frame);
    }
    rallyRAF = requestAnimationFrame(frame);

    introDone = done;
    overlay.classList.add('show');
    rallyTimer = setTimeout(skipLoginAnimation, RALLY_MS);
}

function skipLoginAnimation() {                   // runs automatically at the end, or when the user taps the screen
    const overlay = document.getElementById('rally-overlay');
    if (!introDone) return;
    clearTimeout(rallyTimer);
    cancelAnimationFrame(rallyRAF);
    const done = introDone; introDone = null;
    done();                                       // go to the next screen under the overlay
    overlay.classList.remove('show');             // then fade the overlay away
}

async function startIntro(source) {    // after login / sign up
    if (authBusy) return;
    const isSignup = source === 'signup';
    const val = id => { const el = document.getElementById(id); return el ? el.value : ''; };
    const username = val(isSignup ? 'new-username' : 'username').trim().toLowerCase();
    const password = val(isSignup ? 'new-password' : 'password');
    const fullName = val('new-fullname').trim();
    if (!username) return toast('Please enter your Instructor ID.');
    if (!password) return toast('Please enter your password.');
    if (isSignup && !fullName) return toast('Please enter your full name.');

    authBusy = true;
    setAuthButtons(true);
    try {
        const res = await api(isSignup ? 'signup' : 'login', { token: null, username, password, fullName });
        authToken = res.token;
        try { localStorage.setItem(TOKEN_KEY, authToken); } catch (e) {}
        currentUser = res.user.username;
        applyUserData(res.data);
        setUserName(res.user.fullName, res.user.username);
        const hour = new Date().getHours();
        const part = hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
        document.getElementById('intro-greeting').innerText = `${part}, Instructor!`;
        classSetupFromDashboard = false;
        currentClass = null;
        ['password', 'new-password'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
        playLoginAnimation(isSignup ? 'Account created!' : 'Welcome back!', () => {
            navigate('onboarding-intro');
            if (isSignup) toast('Account created!');
        });
    } catch (e) {
        toast(e.message);
    } finally {
        authBusy = false;
        setAuthButtons(false);
    }
}

function startClassSetup() {            // intro -> class details            
    classSetupFromDashboard = false;
    prepareClassForm();
    navigate('onboarding-welcome');
}

function changeClass() {                // from the dashboard chip
    classSetupFromDashboard = true;
    prepareClassForm();
    navigate('onboarding-welcome');
}

function setCurrentClass(c) {
    currentClass = lastClass = c;
    api('saveClass', { program: c.program, year: c.year, section: c.section }).catch(e => toast('Class not saved: ' + e.message));
    navigate(classSetupFromDashboard ? 'dashboard-screen' : 'onboarding-goals');
}

function continueClass() { setCurrentClass(lastClass); }

function saveClassAndContinue() {
    const program = document.getElementById('class-program').value.trim().toUpperCase();
    const section = document.getElementById('class-section').value.trim().toUpperCase();
    const year = '2';   // fixed to 2nd Year
    if (!program || !section) {
        document.getElementById('class-error').innerText = 'Please enter the program and section.';
        return;
    }
    if (!/^[A-Z]$/.test(section)) {
        document.getElementById('class-error').innerText = 'Section must be a single letter (e.g. A, B, C).';
        return;
    }
    setCurrentClass({ program, year, section });
}

// --- Recent sessions (saved on every START) ---
function getHistory() {
    return historyCache;
}

function renderHistory() {
    const list = document.getElementById('history-list');
    if (!list) return;
    list.innerHTML = '';
    // One entry per class (most recent session first), showing only the class and the date last attended
    const seen = new Set();
    const items = getHistory().filter(it => {
        if (!it.cls || seen.has(it.cls)) return false;
        seen.add(it.cls);
        return true;
    }).slice(0, 5);
    if (!items.length) {
        const li = document.createElement('li');
        li.textContent = 'No classes yet';
        list.appendChild(li);
        return;
    }
    items.forEach(it => {
        const li = document.createElement('li');
        const top = document.createElement('b'); top.textContent = it.cls;
        const day = document.createElement('em'); day.textContent = it.ts ? new Date(it.ts).toLocaleDateString('en-PH', { month: 'short', day: 'numeric' }) : '';
        li.append(top, day);
        list.appendChild(li);
    });

}

// --- NEW: Multi-Option Toggler (for Training Focus Goals) ---
function toggleOption(element) {
    if(element.classList.contains('active-selection')) {
        element.classList.remove('active-selection');
    } else {
        element.classList.add('active-selection');
    }
}

// --- Universal Sidebar Manager ---
function closeAllSidebars() {
    document.getElementById('user-sidebar').classList.remove('open');
    document.getElementById('settings-sidebar').classList.remove('open');
    document.getElementById('sidebar-overlay').classList.remove('show');
}

function openSidebar(id) {
    closeAllSidebars();
    const sidebar = document.getElementById(id);
    sidebar.classList.add('open');
    document.getElementById('sidebar-overlay').classList.add('show');
    if (id === 'user-sidebar') {
        document.getElementById('sb-class-label').innerText = currentClass ? `${classLabel(currentClass)} · PHED 4` : 'No class set';
        renderHistory();
        const nums = document.querySelectorAll('#user-sidebar .sb-stat b');
        if (nums[0]) nums[0].dataset.count = statsCache.sessions;
        if (nums[1]) nums[1].dataset.count = statsCache.shuttles;
    }
    animateCounts(sidebar);
}

function toggleSidebar(id) {
    if (document.getElementById(id).classList.contains('open')) closeAllSidebars();
    else openSidebar(id);
}

function toggleUserSidebar() { toggleSidebar('user-sidebar'); }
function toggleSettingsSidebar() { toggleSidebar('settings-sidebar'); }

// Count-up animation for numbers marked with data-count
function animateCounts(root) {
    root.querySelectorAll('[data-count]').forEach(el => {
        const target = parseInt(el.dataset.count) || 0;
        const start = performance.now();
        const dur = 800;
        (function tick(now) {
            const p = Math.min((now - start) / dur, 1);
            el.textContent = Math.round(target * (1 - Math.pow(1 - p, 3)));
            if (p < 1) requestAnimationFrame(tick);
        })(start);
    });
}

// Tap the IP tile to copy it
function copyIP(tile) {
    const ip = document.getElementById('ip-text').innerText;
    try { navigator.clipboard.writeText(ip).catch(() => {}); } catch (e) {}
    const hint = tile.querySelector('em');
    tile.classList.add('copied');
    hint.innerText = 'Copied ✔';
    setTimeout(() => { tile.classList.remove('copied'); hint.innerText = 'Tap to copy'; }, 1500);
}

function popEl(id) {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.remove('bump');
    void el.offsetWidth;
    el.classList.add('bump');
}
function setRange(v) {
    if (blockedWhileLocked()) return;
    const s = String(v).toUpperCase();
    selectedRange = RANGE_LABELS[s] ? s : (s === 'LOW' ? 'SHORT' : s === 'HIGH' ? 'LONG' : 'MEDIUM');
    document.querySelectorAll('#range-opts button').forEach(b => b.classList.toggle('on', b.dataset.r === selectedRange));
    setText('range-name', RANGE_LABELS[selectedRange]);
    popEl('range-name');
    setText('current-speed-display', RANGE_LABELS[selectedRange]);
}
function updateSliders() {
    const h = document.getElementById('h-dir').value;
    const v = document.getElementById('v-ang').value;
    const hVal = document.getElementById('h-val');

    hVal.innerText = h + '°';
    const vEl = document.getElementById('v-val');
    vEl.innerText = v + '°';
    setText('current-angle-display', currentMode === 'Smash' ? SMASH_DIR_NAMES[smashDir] : FIXED_ANGLES[currentMode] ? (+h > 0 ? 'Right' : 'Left') : v + '°');   // Left/Right/Mixed in Smash, Low/High or degrees elsewhere

    // Live aim preview (needles rotate with the sliders)
    document.getElementById('needle-h').style.transform = `rotate(${h}deg)`;
    document.getElementById('needle-v').style.transform = `rotate(${-v}deg)`;

    // Interactive pop
    hVal.style.transform = "scale(1.15)";
    vEl.style.transform = "scale(1.15)";
    setTimeout(() => {
        hVal.style.transform = "scale(1)";
        vEl.style.transform = "scale(1)";
    }, 100);
}

// +/- stepper buttons for Interval and Shuttle Count
function stepValue(id, delta) {
    if (blockedWhileLocked()) return;
    const el = document.getElementById(id);
    const min = parseInt(el.min) || 1;
    el.value = Math.max(min, (parseInt(el.value) || min) + delta);
    if (id === 'count') updateCounterDisplay();
}

// UI Navigation
function navigate(screenId) {
    document.querySelectorAll('.screen').forEach(screen => {
        screen.classList.remove('active');
    });
    document.getElementById(screenId).classList.add('active');
    if (screenId === 'dashboard-screen') refreshDashboard();
    
    // Auto-hide sidebars if open during navigation
    closeAllSidebars();
    
    // If leaving the session screen, ensure the machine is stopped safely
    if(screenId !== 'session-screen') {
        sendCommand('STOP');
    }
}

// Setup Session based on chosen mode
// --- Hit / miss scoring (all modes). The instructor taps HIT or MISS; no sensors involved. ---
function resultsEnabled() { return true; }   // scoring is on for every mode (Smash, Service)

function shotCounts() {
    return {
        hits: shotLog.filter(r => r === 'hit').length,
        misses: shotLog.filter(r => r === 'miss').length
    };
}

function resetShotLog() { shotLog = []; renderResults(); }

// Called every time the machine launches a shuttle
function registerShot() { shotLog.push(null); renderResults(); fireAnim(); }

// HIT / MISS buttons score the oldest launched shot that has no call yet
function scoreShot(result) {
    const i = shotLog.indexOf(null);
    if (i === -1) {
        toast(shotLog.length ? 'Every launched shot is already scored.' : 'No shot has been launched yet.');
        return;
    }
    shotLog[i] = result;
    resultsChanged();
}

// Tap a numbered shot to fix a wrong call (hit <-> miss)
function flipShot(i) {
    if (shotLog[i] === 'hit') shotLog[i] = 'miss';
    else if (shotLog[i] === 'miss') shotLog[i] = 'hit';
    else return;
    resultsChanged();
}

// Clears the most recent call
function undoResult() {
    let i = shotLog.length - 1;
    while (i >= 0 && shotLog[i] === null) i--;
    if (i < 0) { toast('Nothing to undo.'); return; }
    shotLog[i] = null;
    resultsChanged();
}

function resultsChanged() {
    renderResults();
    scheduleResultsSave();
}

function renderResults() {
    const strip = document.getElementById('shot-strip');
    if (!strip) return;
    const { hits, misses } = shotCounts();
    const scored = hits + misses;
    setText('res-hits', hits);
    setText('res-misses', misses);
    setText('res-acc', scored ? `${Math.round(hits / scored * 100)}% accuracy · ${hits} of ${scored} scored` : 'No shots scored yet');

    // colour bar: share of the planned shots that were hits / misses (fixed size, only the fill changes)
    const bar = document.getElementById('res-bar');
    if (bar) {
        const total = Math.max(totalShuttles, shotLog.length, 1);
        bar.children[0].style.width = (hits / total * 100) + '%';
        bar.children[1].style.width = (misses / total * 100) + '%';
    }

    const keepScroll = strip.scrollLeft;
    strip.innerHTML = '';
    if (!shotLog.length) {
        const empty = document.createElement('span');
        empty.className = 'shot-empty';
        empty.textContent = 'Shots appear here. Tap a scored shot to switch it between hit and miss.';
        strip.appendChild(empty);
        return;
    }
    const next = shotLog.indexOf(null);
    let target = null;
    shotLog.forEach((r, i) => {
        const chip = document.createElement('button');
        chip.className = 'shot-chip' + (r ? ' ' + r : '') + (i === next ? ' next' : '');
        chip.textContent = (i + 1) + (r === 'hit' ? '✓' : r === 'miss' ? '✗' : '');
        chip.setAttribute('aria-label', `Shot ${i + 1}: ${r || 'not scored'}`);
        if (r) chip.onclick = () => flipShot(i);
        strip.appendChild(chip);
        if (i === next) target = chip;
    });
    // keep the strip where it was; only scroll when the shot being scored is out of view
    strip.scrollLeft = keepScroll;
    target = target || strip.lastElementChild;
    if (target) {
        const left = target.offsetLeft, right = left + target.offsetWidth;
        if (left < strip.scrollLeft + 12 || right > strip.scrollLeft + strip.clientWidth - 12) {
            strip.scrollLeft = left - strip.clientWidth / 2;
        }
    }
}

// Saves the latest scores to the Sessions sheet (waits 0.7 s so quick taps are sent together).
// Works during the session AND after it ends, so a wrong call can still be fixed.
function scheduleResultsSave() {
    if (!lastSession) return;
    const pending = lastSession;
    const token = authToken;
    const { hits, misses } = shotCounts();
    clearTimeout(resultsSaveTimer);
    resultsSaveTimer = setTimeout(async () => {
        const id = await pending;
        if (!id || token !== authToken) return;
        try {
            const r = await api('saveResults', { id, hits, misses });
            if (token === authToken) { historyCache = r.data.history; statsCache = r.data.stats; refreshDashboard(); }
        } catch (e) {
            toast('Results not saved: ' + e.message);
        }
    }, 700);
}

// --- Interactive aim pads: drag on the picture to aim ---
function setAim(axis, deg) {
    const el = document.getElementById(axis === 'h' ? 'h-dir' : 'v-ang');
    el.value = Math.max(+el.min, Math.min(+el.max, Math.round(deg)));
    updateSliders();
}
function nudgeAim(axis, d) {
    if (blockedWhileLocked()) return;
    const el = document.getElementById(axis === 'h' ? 'h-dir' : 'v-ang');
    setAim(axis, d === 0 ? (axis === 'h' ? 0 : 30) : (+el.value) + d);
}
function aimFromPointer(e, axis) {
    const svg = document.getElementById(axis === 'h' ? 'pad-h' : 'pad-v');
    const pt = svg.createSVGPoint();
    pt.x = e.clientX; pt.y = e.clientY;
    const p = pt.matrixTransform(svg.getScreenCTM().inverse());
    setAim(axis, axis === 'h' ? Math.atan2(p.x - 50, 55 - p.y) * 180 / Math.PI
                              : Math.atan2(55 - p.y, p.x - 10) * 180 / Math.PI);
}
function initAimPads() {
    [['pad-h', 'h'], ['pad-v', 'v']].forEach(([id, axis]) => {
        const pad = document.getElementById(id);
        let down = false;
        pad.addEventListener('pointerdown', e => { down = true; pad.setPointerCapture(e.pointerId); aimFromPointer(e, axis); });
        pad.addEventListener('pointermove', e => { if (down) aimFromPointer(e, axis); });
        ['pointerup', 'pointercancel'].forEach(t => pad.addEventListener(t, () => { down = false; }));
    });
}
// A little shuttle flies along the aim line whenever one is launched
function fireAnim() {
    document.querySelectorAll('.angle-opts button.on').forEach(b => { b.classList.remove('kick'); void b.offsetWidth; b.classList.add('kick'); });
    ['pad-h', 'pad-v'].forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.classList.remove('fire');
        void el.getBoundingClientRect();
        el.classList.add('fire');
    });
}

// --- Previous scores (last 20 sessions) ---
const HM = { filter: 'all', sort: 'new', open: null };
const HM_FILTERS = [['all', 'All'], ['Smash', 'Smash'], ['Service', 'Service']];

function hmEl(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
}
function hmRows() {
    return historyCache.slice(0, 20).map(it => {
        const hits = it.hits || 0, misses = it.misses || 0, scored = hits + misses;
        const count = it.count || 0;
        return {
            it, id: String(it.id || it.ts), hits, misses, scored, count,
            total: Math.max(count, scored, 1),
            acc: scored ? Math.round(hits / scored * 100) : null,
            when: it.ts ? new Date(it.ts) : null
        };
    });
}
function hmModeIs(r, key) { return key === 'Service' ? /service/i.test(r.it.mode || '') : r.it.mode === key; }
function hmTier(acc) { return acc === null ? 't-none' : acc >= 70 ? 't-great' : acc >= 50 ? 't-ok' : 't-low'; }
function hmTierLabel(acc) { return acc >= 70 ? 'Great aim' : acc >= 50 ? 'Getting there' : 'Keep practicing'; }
function hmDayLabel(d) {
    if (!d) return 'Earlier';
    const start = x => new Date(x.getFullYear(), x.getMonth(), x.getDate()).getTime();
    const diff = Math.round((start(new Date()) - start(d)) / 86400000);
    if (diff === 0) return 'Today';
    if (diff === 1) return 'Yesterday';
    return d.toLocaleDateString('en-PH', { weekday: 'short', month: 'short', day: 'numeric' });
}
function hmRing(acc, size) {
    const s = hmEl('span', 'hm-ring ' + hmTier(acc));
    s.style.setProperty('--p', acc === null ? 0 : acc);
    s.style.width = s.style.height = size + 'px';
    s.innerHTML = '<svg viewBox="0 0 36 36"><circle class="rb" cx="18" cy="18" r="15.9155"/><circle class="rf" cx="18" cy="18" r="15.9155"/></svg>';
    s.append(hmEl('i', '', acc === null ? '–' : acc + '%'));
    return s;
}
function hmCount(el, to, suffix) {
    suffix = suffix || '';
    if (!to || window.matchMedia('(prefers-reduced-motion: reduce)').matches) { el.textContent = to + suffix; return; }
    const t0 = performance.now(), dur = 700;
    (function tick(now) {
        const p = Math.min((now - t0) / dur, 1), e = 1 - Math.pow(1 - p, 3);
        el.textContent = Math.round(to * e) + suffix;
        if (p < 1) requestAnimationFrame(tick);
    })(t0);
}

function openHistory() {
    HM.filter = 'all'; HM.sort = 'new'; HM.open = null;
    hmRender();
    document.getElementById('history-modal').classList.remove('hidden');
    const body = document.querySelector('#history-modal .hm-body');
    if (body) body.scrollTop = 0;
}
function closeHistory() { document.getElementById('history-modal').classList.add('hidden'); }
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeHistory(); });

function hmRender() {
    const rows = hmRows();
    setText('hm-sub', rows.length ? `Your last ${rows.length} session${rows.length > 1 ? 's' : ''}` : 'No finished sessions yet');
    hmRenderStats(rows);
    hmRefresh();
}
// filter / sort changes only redraw the chips and the list (the total stays put)
function hmRefresh() {
    const rows = hmRows();
    hmRenderChips(rows);
    hmRenderList(rows);
    document.querySelectorAll('#hm-sort button').forEach(b => b.classList.toggle('on', b.dataset.s === HM.sort));
}

function hmRenderStats(rows) {
    const box = document.getElementById('hm-stats');
    box.innerHTML = '';
    const total = Math.max((statsCache && statsCache.sessions) || 0, rows.length);
    const card = hmEl('div', 'hm-total');
    const txt = hmEl('div');
    const v = hmEl('b', 'hm-big', '0');
    txt.append(hmEl('small', '', 'Total training sessions'), v);
    card.append(hmEl('div', 'hm-total-ic', '🏸'), txt);
    if (total > rows.length) card.append(hmEl('em', '', `Showing the latest ${rows.length}`));
    box.append(card);
    hmCount(v, total);
}

function hmRenderChips(rows) {
    const box = document.getElementById('hm-chips');
    box.innerHTML = '';
    HM_FILTERS.forEach(([key, label]) => {
        const n = key === 'all' ? rows.length : rows.filter(r => hmModeIs(r, key)).length;
        if (key !== 'all' && n === 0) return;
        const b = hmEl('button', 'hm-chip' + (HM.filter === key ? ' on' : ''));
        b.append(document.createTextNode(label + ' '), hmEl('span', '', n));
        b.onclick = () => { HM.filter = key; hmRefresh(); };
        box.append(b);
    });
}
function hmSetSort(s) { HM.sort = s; hmRefresh(); }

function hmRenderList(rows) {
    const list = document.getElementById('history-full');
    list.innerHTML = '';
    const scoredRows = rows.filter(r => r.acc !== null);
    const bestId = scoredRows.length > 1 ? scoredRows.reduce((b, r) => (r.acc > b.acc ? r : b), scoredRows[0]).id : null;

    let view = rows.filter(r => HM.filter === 'all' ? true : hmModeIs(r, HM.filter));
    if (HM.sort === 'best') view = view.slice().sort((a, b) => (b.acc === null ? -1 : b.acc) - (a.acc === null ? -1 : a.acc));

    if (!rows.length) { list.append(hmEmpty('🏸', 'No finished sessions yet', 'Run a drill and score the shots to see your results here.')); return; }
    if (!view.length) {
        const e = hmEmpty('🔍', 'Nothing matches this filter', '');
        const b = hmEl('button', 'hm-chip on', 'Show all'); b.onclick = () => { HM.filter = 'all'; hmRefresh(); };
        e.append(b); list.append(e); return;
    }
    let lastDay = null, n = 0;
    view.forEach((r, idx) => {
        if (HM.sort === 'new') {
            const day = hmDayLabel(r.when);
            if (day !== lastDay) { list.append(hmEl('div', 'hm-day', day)); lastDay = day; }
        }
        list.append(hmCard(r, r.id === bestId, Math.min(n++, 10), HM.sort === 'best' ? idx + 1 : 0));
    });
}
function hmEmpty(icon, title, sub) {
    const e = hmEl('div', 'hm-empty');
    e.append(hmEl('div', 'hm-empty-ic', icon), hmEl('b', '', title));
    if (sub) e.append(hmEl('small', '', sub));
    return e;
}

function hmCard(r, isBest, i, rank) {
    const it = r.it;
    const card = hmEl('div', 'hm-item ' + hmTier(r.acc) + (HM.open === r.id ? ' open' : ''));
    card.dataset.id = r.id;
    card.style.setProperty('--i', i);

    // collapsed row: what / when / how far - tap it to open the session
    const row = hmEl('div', 'hm-row');
    row.tabIndex = 0; row.setAttribute('role', 'button'); row.setAttribute('aria-expanded', HM.open === r.id ? 'true' : 'false');
    const isSmash = it.mode === 'Smash';
    const badge = hmEl('div', 'hm-badge', rank ? '#' + rank : '');
    if (!rank) {
        const picUrl = imgUrl(isSmash ? IMAGES.smash : IMAGES.service);
        if (picUrl) {
            const pic = document.createElement('img');
            pic.alt = isSmash ? 'Smash' : 'Service';
            pic.onerror = () => { badge.textContent = isSmash ? '⚡' : '🏸'; pic.remove(); };   // fallback if the image fails
            pic.src = picUrl;
            badge.append(pic);
        } else {
            badge.textContent = isSmash ? '⚡' : '🏸';   // no image set in IMAGES
        }
    }
    const info = hmEl('div', 'hm-info');
    const title = hmEl('b', '', `${it.mode} · ${it.cls}`);
    if (isBest) title.append(hmEl('span', 'hm-best', '🏆 Best'));
    const dt = r.when;
    const timeTxt = dt ? (HM.sort === 'best' ? dt.toLocaleString('en-PH', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }) : dt.toLocaleTimeString('en-PH', { hour: 'numeric', minute: '2-digit' })) : '';
    info.append(title, hmEl('small', '', `${timeTxt} · ${RANGE_LABELS[it.range] || ''} range · ${r.count} shuttles`));
    row.append(badge, info, hmEl('span', 'hm-chev', '›'));
    row.onclick = () => hmToggle(r.id);
    row.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); hmToggle(r.id); } };

    // dropdown: accuracy of THIS session + the other stats
    const more = hmEl('div', 'hm-more'), inner = hmEl('div', 'hm-more-in');
    if (r.acc === null) {
        inner.append(hmEl('div', 'hm-noscore', 'No shots were scored in this session.'));
    } else {
        const top = hmEl('div', 'hm-acc-top');
        const side = hmEl('div', 'hm-acc-side');
        side.append(hmEl('small', '', 'Accuracy'), hmEl('b', '', hmTierLabel(r.acc)), hmEl('span', '', `${r.hits} of ${r.scored} scored shots hit`));
        top.append(hmRing(r.acc, 92), side);
        inner.append(top);
        const grid = hmEl('div', 'hm-grid');
        [['Hits', r.hits, 'h'], ['Misses', r.misses, 'm'], ['Not scored', Math.max(r.count - r.scored, 0), 'n']].forEach(([lab, val, k]) => {
            const c = hmEl('div', 'hm-cell ' + k); c.append(hmEl('b', '', val), hmEl('small', '', lab)); grid.append(c);
        });
        const bar = hmEl('div', 'hm-seg-bar');
        [['h', r.hits], ['m', r.misses]].forEach(([k, v]) => { const s = hmEl('i', 'hm-seg ' + k); s.style.setProperty('--w', (v / r.total * 100) + '%'); bar.append(s); });
        inner.append(grid, bar);
    }
    const tags = hmEl('div', 'hm-tags');
    tags.append(hmEl('span', '', `🎯 ${RANGE_LABELS[it.range] || '-'} range`), hmEl('span', '', `🏸 ${r.count} launched`), hmEl('span', '', `🎓 ${it.cls}`));
    if (it.planned && r.count < it.planned) tags.append(hmEl('span', 'warn', `Stopped early · ${r.count} of ${it.planned}`));
    const again = hmEl('button', 'hm-again', 'Repeat this drill ▶');
    again.onclick = ev => { ev.stopPropagation(); repeatDrill(it); };
    inner.append(tags, again);
    more.append(inner);
    card.append(row, more);
    return card;
}

function hmToggle(id) {
    HM.open = HM.open === id ? null : id;
    document.querySelectorAll('#history-full .hm-item').forEach(c => {
        const on = c.dataset.id === HM.open;
        c.classList.toggle('open', on);
        c.querySelector('.hm-row').setAttribute('aria-expanded', on ? 'true' : 'false');
    });
}
function repeatDrill(it) {
    const st = document.getElementById('machine-state-display').innerText.toUpperCase();
    if (st.includes('RUN') || st.includes('SHOT') || st.includes('PAUSE')) { toast('Stop the current session first.'); return; }
    closeHistory();
    document.getElementById('count').value = it.planned || it.count || 50;
    setupSession(normMode(it.mode));
    setRange(it.range);
    toast(`${normMode(it.mode)} · ${RANGE_LABELS[it.range] || ''} range loaded. Press START when ready.`);
}

// --- Keyboard shortcuts on the control screen ---
function onSessionScreen(e) {
    return document.getElementById('session-screen').classList.contains('active') &&
           document.getElementById('history-modal').classList.contains('hidden') &&
           !(e.ctrlKey || e.metaKey || e.altKey) && !/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName);
}
document.addEventListener('keydown', e => {
    if (!onSessionScreen(e)) return;
    const k = e.key.toLowerCase();
    if (k === 'h') scoreShot('hit');
    else if (k === 'm') scoreShot('miss');
    else if (k === 'u') undoResult();
    else if (k === 's') sendCommand('STOP');
    else if (k === ' ') {
        const st = document.getElementById('machine-state-display').innerText.toUpperCase();
        sendCommand(st.includes('PAUSE') ? 'RESUME' : (st.includes('RUN') || st.includes('SHOT')) ? 'PAUSE' : 'START');
    } else return;
    e.preventDefault();
});
document.addEventListener('keyup', e => { if (e.key === ' ' && onSessionScreen(e)) e.preventDefault(); });

// --- Smash mode: the angle is fixed (assumed degrees - tune on the real machine); the instructor picks the DIRECTION ---
const SMASH_ANGLE = 35;
const SMASH_DIR_NAMES = { LEFT: 'Left', RIGHT: 'Right', MIXED: 'Mixed' };
let smashDir = 'MIXED';                 // LEFT | RIGHT | MIXED (MIXED = shuttles alternate between left and right in a random order)
let mixedQueue = [];
function smashDirFromAim(h) { return h < 0 ? 'LEFT' : h > 0 ? 'RIGHT' : 'MIXED'; }
function setSmashDirection(d) {
    if (blockedWhileLocked()) return;
    smashDir = SMASH_DIR_NAMES[d] ? d : 'MIXED';
    setAim('h', smashDir === 'MIXED' ? 0 : POSITION_DIRS[smashDir]);
    document.querySelectorAll('#smash-dir-opts button').forEach(b => b.classList.toggle('on', b.dataset.d === smashDir));
    setText('smash-dir-name', SMASH_DIR_NAMES[smashDir]);
    popEl('smash-dir-name');
}
// MIXED: half the shuttles go left, half go right, shuffled
function buildMixedQueue(n) {
    const q = [];
    for (let i = 0; i < n; i++) q.push(i % 2 ? 'RIGHT' : 'LEFT');
    for (let i = q.length - 1; i > 0; i--) { const k = Math.floor(Math.random() * (i + 1)); [q[i], q[k]] = [q[k], q[i]]; }
    return q;
}
// --- Service mode: the instructor picks the DIRECTION (Left or Right); the angle is fixed and not shown ---
const POSITION_DIRS = { LEFT: -30, RIGHT: 30 };     // assumed aim for each position (degrees) - tune on the machine
// Angle is fixed in the two service modes (degrees are assumed - tune on the machine)
const FIXED_ANGLES = { 'Service': { angle: 55, name: 'High' } };
function setPosition(p) {
    if (blockedWhileLocked()) return;
    setAim('h', POSITION_DIRS[p]);
    document.querySelectorAll('#position-opts button').forEach(b => b.classList.toggle('on', b.dataset.p === p));
    setText('position-name', p === 'LEFT' ? 'Left' : 'Right');
    popEl('position-name');
}

function applyModeLayout() {
    const smash = currentMode === 'Smash';
    const fixed = FIXED_ANGLES[currentMode];                 // set for Service
    document.getElementById('aim-pads').classList.toggle('hidden', smash || !!fixed);
    document.getElementById('fixed-tiles').classList.toggle('hidden', !fixed);
    document.getElementById('smash-dir-tile').classList.toggle('hidden', !smash);
    const locked = isLockedMode();
    document.getElementById('free-inputs').classList.toggle('hidden', locked);      // interval + shuttles steppers
    document.getElementById('locked-inputs').classList.toggle('hidden', !locked);   // fixed 3 sec / 24 shuttles
    setText('current-angle-label', (smash || fixed) ? 'Direction' : 'Angle');
    if (fixed) {
        setAim('v', fixed.angle);
        setPosition(+document.getElementById('h-dir').value > 0 ? 'RIGHT' : 'LEFT');
    }
    if (smash) {
        setAim('v', SMASH_ANGLE);                         // angle is fixed in Smash
        setSmashDirection(smashDirFromAim(+document.getElementById('h-dir').value));
    }
}

function setupSession(modeName) {
    setControlsLocked(false);                 // a fresh control screen is always editable
    currentMode = modeName;
    document.getElementById('session-title').innerText = modeName.toUpperCase() + " MODE CONTROL";
    setText('session-class', currentClass ? `${classLabel(currentClass)} · PHED 4` : 'No class set');
    shuttlesLaunched = 0;
    lastSession = null;
    document.getElementById('results-card').classList.toggle('hidden', !resultsEnabled());
    document.getElementById('right-column').classList.toggle('has-results', resultsEnabled());
    resetShotLog();
    applyModeLayout();
    updateCounterDisplay();
    resetButtons();
    updateSliders();
    navigate('session-screen');
}

// Handle UI button states
function resetButtons() {
    setControlsLocked(false);                 // session stopped / finished: settings can be changed again
    document.getElementById('btn-start').classList.remove('hidden');
    document.getElementById('btn-pause').classList.add('hidden');
    document.getElementById('btn-resume').classList.add('hidden');
    document.getElementById('btn-stop').classList.add('hidden');
    document.getElementById('machine-state-display').innerText = "IDLE";
}

function updateCounterDisplay() {
    totalShuttles = getCount();
    document.getElementById('launch-counter').innerText = `${shuttlesLaunched} / ${totalShuttles}`;

    // Update progress ring (circumference = 326.7)
    const ratio = Math.min(shuttlesLaunched / totalShuttles, 1);
    document.getElementById('progress-ring').style.strokeDashoffset = 326.7 * (1 - ratio);
}

// Keeps the server's login check ready, so even the first START press is fast
let lastMachineWarm = 0;
setInterval(function () {
    if (!hasServer() || !authToken) return;
    if (Date.now() - lastMachineWarm < 240000) return;       // refresh every 4 minutes
    lastMachineWarm = Date.now();
    api('machineCommand', { command: 'PING' }).catch(function () { lastMachineWarm = 0; });
}, 5000);

// ---- USB CABLE CONTROL (Web Serial: Chrome on Android / desktop, page must be on https) ----
let usbWriter = null, usbTimer = null;

function usbSetButton(connected) {
    const b = document.getElementById('btn-usb');
    if (b) b.innerText = connected ? 'USB connected' : 'Connect USB';
}

async function connectUsb() {
    if (!('serial' in navigator)) { toast('This browser has no USB (Web Serial) support.'); return; }
    try {
        const port = await navigator.serial.requestPort();      // must run from a button tap
        await port.open({ baudRate: 115200 });
        const enc = new TextEncoderStream();
        enc.readable.pipeTo(port.writable);
        usbWriter = enc.writable.getWriter();
        clearInterval(usbTimer);
        usbTimer = setInterval(() => usbSend('PING'), 1000);     // keeps the Arduino failsafe fed
        usbSetButton(true);
        toast('USB machine connected');
    } catch (e) { toast('USB not connected: ' + e.message); }
}

function usbSend(line) {
    if (!usbWriter) return;
    usbWriter.write(line + '\n').catch(() => {
        usbWriter = null;
        clearInterval(usbTimer);
        usbSetButton(false);
        toast('USB cable disconnected');
    });
}

// Sends the pressed button to the machine (Arduino) through the Apps Script backend
function sendToMachine(command) {
    if (usbWriter) { usbSend(command + ',' + selectedRange); return; }   // USB cable connected: send over the cable
    if (!hasServer() || !authToken) return;
    const t0 = performance.now();
    api('machineCommand', {
        command: command,
        range: selectedRange,
        direction: (currentMode === 'Smash') ? smashDir : 'MIXED',
        interval: getInterval(),
        count: getCount()
    }).then(() => {
        const ms = Math.round(performance.now() - t0);
        console.log('[MACHINE] ' + command + ' saved on the server in ' + ms + ' ms');
        toast('Machine ' + command + ' sent (' + ms + ' ms)');   // temporary: remove after testing
    }).catch(e => toast('Machine not reached: ' + e.message));
}

// Logic to simulate hardware response
function sendCommand(command) {
    const btnStart = document.getElementById('btn-start');
    const btnPause = document.getElementById('btn-pause');
    const btnResume = document.getElementById('btn-resume');
    const btnStop = document.getElementById('btn-stop');
    const statusText = document.getElementById('machine-state-display');
    const stateBox = statusText.parentElement; // Used to trigger CSS flash animation

    let intervalSec = getInterval();
    totalShuttles = getCount();

    console.log(`[SYSTEM MOCKUP] Command: ${command} | Range: ${selectedRange} | Interval: ${intervalSec}s` + (currentMode === 'Smash' ? ` | Direction: ${smashDir}` : ''));
    sendToMachine(command);                   // real machine: tell the Arduino

    if (command === 'START' || command === 'RESUME') {
        if (command === 'START') {
            try { beginLoggedSession(intervalSec); } catch (e) { console.warn('Could not start session log:', e); }
            if (currentMode === 'Smash') mixedQueue = buildMixedQueue(totalShuttles);
        }
        setControlsLocked(true);                  // locked from START until STOP (stays locked while paused)
        btnStart.classList.add('hidden');
        btnResume.classList.add('hidden');
        btnPause.classList.remove('hidden');
        btnStop.classList.remove('hidden');
        statusText.innerText = "LAUNCHING";

        clearInterval(simulationTimer);
        simulationTimer = setInterval(() => {
            shuttlesLaunched++;
            if (currentMode === 'Smash' && smashDir === 'MIXED') {          // this shuttle goes left or right
                const side = mixedQueue.shift() || (Math.random() < 0.5 ? 'LEFT' : 'RIGHT');
                setAim('h', POSITION_DIRS[side]);
                setText('smash-dir-name', 'Mixed · ' + SMASH_DIR_NAMES[side]);
                setText('current-angle-display', 'Mixed · ' + SMASH_DIR_NAMES[side]);
            }
            if (resultsEnabled()) registerShot();
            updateCounterDisplay();
            
            // Visual Animation Trigger
            statusText.innerText = "SHOT FIRED!";
            stateBox.classList.remove('flash-animation');
            void stateBox.offsetWidth; // Force CSS reflow to restart animation
            stateBox.classList.add('flash-animation');

            setTimeout(() => { 
                if(!btnPause.classList.contains('hidden')) statusText.innerText = "LAUNCHING"; 
            }, 500);

            if (shuttlesLaunched >= totalShuttles) {
                sendCommand('STOP');
                setTimeout(() => toast('Training session complete!'), 200);
            }
        }, intervalSec * 1000); 

    } else if (command === 'PAUSE') {
        btnPause.classList.add('hidden');
        btnResume.classList.remove('hidden');
        statusText.innerText = "PAUSED";
        stateBox.classList.remove('flash-animation');
        clearInterval(simulationTimer);
    } else if (command === 'STOP' || command === 'EMERGENCY_STOP') {
        finishLoggedSession(command === 'EMERGENCY_STOP' ? 'EMERGENCY_STOP' : (shuttlesLaunched >= totalShuttles ? 'COMPLETED' : 'STOPPED'));
        resetButtons();
        statusText.innerText = command === 'EMERGENCY_STOP' ? "EMERGENCY STOP" : "STOPPED";
        stateBox.classList.remove('flash-animation');
        clearInterval(simulationTimer);
        shuttlesLaunched = 0;
        updateCounterDisplay();
        if (currentMode === 'Smash') setSmashDirection(smashDir);          // back to "Mixed" / the chosen side
    }
}

try { refreshDashboard(); } catch (e) { console.warn(e); }
setInterval(tickClock, 1000);

// Stay logged in after a page refresh (token is valid for 12 hours)
(async function resumeLogin() {
    let t = null;
    try { t = localStorage.getItem(TOKEN_KEY); } catch (e) {}
    if (!t || !hasServer()) return;
    authToken = t;
    try {
        const r = await api('me');
        currentUser = r.user.username;
        applyUserData(r.data);
        currentClass = lastClass;
        setUserName(r.user.fullName, r.user.username);
        navigate('dashboard-screen');
    } catch (e) {
        if (authToken === t) { authToken = null; try { localStorage.removeItem(TOKEN_KEY); } catch (_) {} }
    }
})();

initAimPads();
setRange(selectedRange);
applyImages();
syncThemeUi();
