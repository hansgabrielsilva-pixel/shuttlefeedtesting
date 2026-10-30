// =====================================================
// SHUTTLEFEED front-end - data is stored in Google Sheets
// through a Google Apps Script web app (see Code.gs)
// =====================================================

// Only needed if you host this page OUTSIDE Apps Script. Leave '' when the page is served by Apps Script.
const CONFIG = { API_URL: 'https://script.google.com/macros/s/AKfycbz9OgCyPSPYmw5INRl_t2_taCs6hR49Kw4pnowhS3UI41v2ryjOy1odtuSs3IXZs3k/exec' };

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




// ===== Settings > Help & support / Terms & policies / User's manual =====
// Edit the text here. Each doc has groups; each group has items { ic, q, a } where "a" is HTML.
const SUPPORT_CONTACT = { name: 'SHUTTLEFEED Support Team', email: '', note: 'CvSU Carmona - PHED 4 Instructor Console' };   // put the team's email in "email" to show a contact card

const INFO_DOCS = {
    help: {
        title: 'Help & support',
        sub: 'Quick answers and troubleshooting',
        search: true,
        lead: 'Stuck? Tap a question below. If your problem is not listed, contact the support team.',
        groups: [
            { name: 'Getting started', items: [
                { ic: '🔑', q: 'I cannot log in', a: '<p>Check that your <b>Instructor ID</b> and <b>password</b> are typed correctly (the eye icon shows the password). After several wrong attempts the account is locked for <b>10 minutes</b>, so wait and try again.</p><p>No account yet? Tap <b>Sign up</b> on the login screen.</p>' },
                { ic: '⏱️', q: 'It says my session expired', a: '<p>A login lasts <b>12 hours</b>. Log in again; your class, scores and history are saved and will still be there.</p>' },
                { ic: '🎓', q: 'How do I change my class?', a: '<p>Tap the <b>class card</b> on the Home screen, or open your profile (top left) and tap <b>Change</b> beside Current class.</p>' }
            ]},
            { name: 'Feeder and connection', items: [
                { ic: '📶', q: 'The feeder does not respond', a: '<ol><li>Make sure the feeder is powered on and the Arduino UNO R4 WiFi is running.</li><li>Check the <b>Wi-Fi status</b> in Settings shows Connected.</li><li>Confirm the phone or laptop is on the <b>same Wi-Fi network</b> as the feeder.</li><li>If it still fails, restart the feeder and reload the page.</li></ol>' },
                { ic: '🔌', q: '"Connect machine (USB)" does not work', a: '<ul><li>USB connection works in <b>Chrome or Edge</b> on a laptop. Other browsers cannot use it.</li><li>Use a USB <b>data</b> cable (some cables only charge).</li><li>Tap the button, choose the Arduino in the list, and wait for it to say connected.</li><li>If the browser blocks USB access, open the page in its own tab instead of inside an embedded frame.</li></ul>' }
            ]},
            { name: 'During a drill', items: [
                { ic: '🔒', q: 'Why are my settings locked?', a: '<p>Launch settings are locked while a drill is running so they cannot change mid-session. Press <b>STOP</b> to unlock them.</p>' },
                { ic: '🏸', q: 'The shuttles are not launching or are uneven', a: '<ul><li>Check there are shuttles in the feeder and nothing is jammed.</li><li>Try a different <b>range</b>; each range uses a different motor speed.</li><li>Motor vibration or noise can shift the aim, so make sure the feeder is on a stable surface.</li></ul>' },
                { ic: '📊', q: 'My scores are missing from History', a: '<p>Scores are saved when a session ends. Press <b>STOP</b> to finish the drill, then open <b>History</b> in the session screen. The last 20 sessions are shown.</p>' }
            ]}
        ],
        showContact: true,
        feedback: true
    },

    terms: {
        title: 'Terms & policies',
        sub: 'Please read before using SHUTTLEFEED',
        search: true,
        lead: 'SHUTTLEFEED is a capstone design project of CvSU Carmona, built to help PHED 4 instructors run consistent badminton feeding drills.',
        groups: [
            { name: 'Terms of use', items: [
                { ic: '📜', q: 'Acceptance of terms', a: '<p>By creating an account or logging in you agree to these terms. If you do not agree, please do not use the system.</p>' },
                { ic: '👤', q: 'Who may use it', a: '<p>SHUTTLEFEED is intended for <b>authorized instructors</b> and project personnel of CvSU Carmona. You are responsible for keeping your password private and for all activity under your account.</p>' },
                { ic: '✅', q: 'Acceptable use', a: '<ul><li>Use the system only for badminton training and teaching.</li><li>Do not try to break into other accounts, overload the server or tamper with the hardware.</li><li>Do not share your account with students or other people.</li></ul>' },
                { ic: '🛠️', q: 'Availability', a: '<p>The system is a prototype provided as-is. It may be updated, paused or changed without notice, and it depends on Wi-Fi/USB, the feeder hardware and Google services being available.</p>' }
            ]},
            { name: 'Privacy policy', items: [
                { ic: '🗂️', q: 'What we store', a: '<ul><li>Your <b>name</b>, <b>Instructor ID</b> and last login time.</li><li>Your <b>class</b> entries (program, year, section).</li><li>Your <b>session records</b>: date, mode, range, direction, shuttle count and hit/miss results.</li></ul>' },
                { ic: '🔐', q: 'How it is protected', a: '<p>Passwords are never stored as plain text; they are <b>salted and hashed</b>. Logins use a temporary token that expires after 12 hours, and repeated failed logins are temporarily blocked.</p><p>Data is kept in the project\'s Google Sheet and is accessible only to the project administrators.</p>' },
                { ic: '🎯', q: 'How it is used', a: '<p>Your data is used only to run the app: showing your history and statistics, and for academic evaluation of the project. It is <b>not sold or shared</b> for advertising.</p>' },
                { ic: '🧹', q: 'Your choices', a: '<p>You may ask the administrators to correct or delete your account and records. Your theme preference (light/dark) is saved only on your own device.</p>' }
            ]},
            { name: 'Safety policy', items: [
                { ic: '⚠️', q: 'Hardware and court safety', a: '<div class="info-note warn"><b>The feeder launches shuttlecocks at speed and uses spinning motors.</b></div><ul><li>Keep hands, hair and loose clothing away from the launcher while it is powered.</li><li>Press <b>STOP</b> (or the <kbd>S</kbd> key) before reloading shuttles or adjusting the feeder.</li><li>Keep people out of the line of fire and wear eye protection if needed.</li><li>Place the feeder on a flat, stable surface and unplug it when not in use.</li><li>Only trained personnel should handle wiring, batteries and the Arduino.</li></ul>' },
                { ic: '⚖️', q: 'Limitation of liability', a: '<p>The developers and CvSU Carmona are not liable for injury, damage or data loss arising from misuse of the system or hardware. Always supervise the drill.</p>' }
            ]}
        ],
        foot: 'Last updated October 2026'
    },

    manual: {
        title: 'User\'s manual',
        sub: 'How to use the SHUTTLEFEED software',
        search: true,
        lead: 'A step-by-step guide to the SHUTTLEFEED instructor console, from logging in to reviewing your scores.',
        groups: [
            { name: 'Basics', items: [
                { ic: '1️⃣', q: 'Create an account and log in', a: '<ol><li>On the login screen tap <b>Sign up</b>.</li><li>Enter your <b>Full Name</b>, <b>Instructor ID</b> and a <b>Password</b>, then tap <b>Create Account</b>.</li><li>Next time, just enter your ID and password and tap <b>Log in</b>.</li></ol>' },
                { ic: '2️⃣', q: 'Set up your class', a: '<ol><li>Enter the <b>Program</b> (e.g. BSCpE), <b>Year</b> and <b>Section</b>.</li><li>Tap <b>Next</b>, then <b>Finish Setup</b> and <b>Go to Dashboard</b>.</li><li>Returning instructors can tap <b>Continue with</b> to reuse the last class.</li></ol>' },
                { ic: '3️⃣', q: 'Know the Home screen', a: '<ul><li>Shows the greeting, clock and your <b>class</b> card (tap it to change).</li><li><b>START A DRILL</b> opens the drill list; <b>Quick Start</b> repeats your last drill.</li><li><b>Top left</b> opens your profile; the <b>gear</b> opens Settings, where you can switch between light and dark mode.</li></ul>' }
            ]},
            { name: 'Running a drill', items: [
                { ic: '⚡', q: 'Choose a mode', a: '<p>From <b>START A DRILL</b> pick one:</p><ul><li><b>Smash Mode</b>: high-speed smashes (40 to 70 m/s). You choose range, direction, interval and shuttle count.</li><li><b>Service Mode</b>: consistent serves. You choose range and direction (Left or Right); interval and shuttle count are fixed.</li></ul>' },
                { ic: '🎛️', q: 'Adjust the launch settings', a: '<ul><li><b>Range</b>: Short, Medium or Long (how far it throws).</li><li><b>Direction</b>: where shuttles are thrown. Smash offers Left, Right or Mixed (Mixed alternates left and right in random order); Service offers Left or Right.</li><li><b>Interval</b> and <b>Shuttles</b> (Smash): tap <b>+</b> / <b>−</b> to change them. In Service they are fixed at 3 seconds and 24 shuttles.</li></ul>' },
                { ic: '▶️', q: 'Start, pause, resume and stop', a: '<ol><li>(Laptop only) tap <b>Connect machine (USB)</b> if you are using a USB cable.</li><li>Tap <b>START ▶</b> to begin the drill.</li><li><b>PAUSE ⏸</b> holds the drill, <b>RESUME ▶</b> continues it.</li><li><b>STOP ■</b> ends it. Settings are locked while the drill is running.</li></ol>' },
                { ic: '⌨️', q: 'Keyboard shortcuts', a: '<p>On the session screen of a laptop:</p><ul><li><kbd>Space</kbd> start / pause / resume</li><li><kbd>S</kbd> stop</li><li><kbd>H</kbd> mark a hit &nbsp; <kbd>M</kbd> mark a miss &nbsp; <kbd>U</kbd> undo</li></ul>' }
            ]},
            { name: 'Scores and history', items: [
                { ic: '🎯', q: 'Record shot results', a: '<p>After each shuttle tap <b>HIT ✓</b> or <b>MISS ✗</b>. The bar and shot strip show your running accuracy. Made a mistake? Tap <b>↶ Undo</b> to clear the last score.</p>' },
                { ic: '📈', q: 'View previous scores', a: '<p>Tap <b>History</b> on the session screen to see your last 20 sessions. Filter them, sort by <b>Newest</b> or <b>Best</b>, and tap a row to see its details.</p>' },
                { ic: '👤', q: 'Your profile', a: '<p>The profile panel (top left) shows your current class, sessions run, shuttles launched and classes attended. Use <b>Log Out</b> when you are finished.</p>' }
            ]},
            { name: 'Settings', items: [
                { ic: '⚙️', q: 'System settings', a: '<ul><li><b>Wi-Fi status</b>: whether the feeder is connected.</li><li><b>Hardware / Firmware</b>: the board and version in use.</li><li><b>Appearance</b>: switch between light and dark mode. It is remembered on your device.</li></ul>' }
            ]}
        ]
    }
};

// ===== Kiosk help center (full screen). Used by the Settings links and by the manual panel on the welcome screen =====
const INFO_META = {
    help:   { tab: 'Help & support',   short: 'Help',   ic: '❓' },
    terms:  { tab: 'Terms & policies', short: 'Terms',  ic: '🛡️' },
    manual: { tab: 'User\'s manual',   short: 'Manual', ic: '📖' }
};
const INFO_IDLE_MS = 120000;           // kiosk behaviour: close after 2 minutes without a touch
let infoReturnFocus = null, infoIdleTimer = null, infoCurrent = null;

// Two-pane viewer: topic list on the left, big readable page on the right, Previous / Next at the bottom.
// In a narrow space (phone, or the welcome-screen panel) it becomes list -> page with a Back button.
function buildKiosk(root, d, opt) {
    opt = opt || {};
    root.innerHTML = '';
    root.classList.add('kb');
    const topics = [];
    d.groups.forEach(g => g.items.forEach(it => topics.push({ g: g.name, it: it })));

    const wrap = document.createElement('div');
    wrap.className = 'kb-wrap';
    wrap.dataset.view = 'list';

    // ---- left: topic list ----
    const list = document.createElement('nav');
    list.className = 'kb-list';
    list.setAttribute('aria-label', d.title + ' topics');
    const head = document.createElement('div');
    head.className = 'kb-doc';
    head.innerHTML = '<h4>' + d.title + '</h4>' + (d.lead ? '<p>' + d.lead + '</p>' : '');
    list.appendChild(head);
    const btns = [];
    let idx = 0;
    d.groups.forEach(g => {
        const lab = document.createElement('div');
        lab.className = 'kb-glabel'; lab.textContent = g.name;
        list.appendChild(lab);
        g.items.forEach(it => {
            const i = idx++;
            const b = document.createElement('button');
            b.type = 'button'; b.className = 'kb-topic';
            b.innerHTML = '<span class="kb-tic">' + it.ic + '</span><span class="kb-tt">' + it.q + '</span><span class="kb-chev" aria-hidden="true">›</span>';
            b.addEventListener('click', () => { show(i, i >= cur ? 1 : -1); wrap.dataset.view = 'detail'; });
            list.appendChild(b); btns.push(b);
        });
    });
    if (d.showContact && opt.contact !== false) {
        const c = SUPPORT_CONTACT;
        const mail = c.email ? '<a href="mailto:' + c.email + '">' + c.email + '</a>' : 'Ask your course instructor or the project team.';
        const box = document.createElement('div');
        box.className = 'info-contact';
        box.innerHTML = '<span class="info-ic">💬</span><div><b>Still need help?</b><small>' + c.name + ' · ' + c.note + '</small><small>' + mail + '</small></div>';
        list.appendChild(box);
    }
    if (d.foot) { const f = document.createElement('p'); f.className = 'info-foot'; f.textContent = d.foot; list.appendChild(f); }

    // ---- right: reading page ----
    const detail = document.createElement('section');
    detail.className = 'kb-detail';
    detail.innerHTML =
        '<button type="button" class="kb-back">‹ All topics</button>' +
        '<div class="kb-scroll"><article class="kb-card"></article></div>' +
        '<div class="kb-pager">' +
            '<button type="button" class="kb-prev">‹ Previous</button>' +
            '<div class="kb-prog"><span class="kb-count"></span><i><b></b></i></div>' +
            '<button type="button" class="kb-next">Next ›</button>' +
        '</div>';
    const card = detail.querySelector('.kb-card'), scroller = detail.querySelector('.kb-scroll');
    const prev = detail.querySelector('.kb-prev'), next = detail.querySelector('.kb-next');
    const count = detail.querySelector('.kb-count'), bar = detail.querySelector('.kb-prog b');
    let cur = 0;

    function show(i, dir) {
        if (i < 0 || i >= topics.length) return;
        cur = i;
        const t = topics[i];
        const fb = d.feedback ? '<div class="info-fb"><span>Was this helpful?</span><button type="button" data-v="up">👍 Yes</button><button type="button" data-v="down">👎 No</button></div>' : '';
        card.innerHTML = '<div class="kb-card-head"><span class="kb-cic">' + t.it.ic + '</span><div><small>' + t.g + '</small><h4>' + t.it.q + '</h4></div></div><div class="kb-body">' + t.it.a + '</div>' + fb;
        card.classList.remove('go-next', 'go-prev'); void card.offsetWidth;
        card.classList.add(dir < 0 ? 'go-prev' : 'go-next');
        scroller.scrollTop = 0;
        btns.forEach((b, k) => { b.classList.toggle('on', k === i); b.setAttribute('aria-current', k === i ? 'true' : 'false'); });
        count.textContent = 'Topic ' + (i + 1) + ' of ' + topics.length;
        bar.style.width = ((i + 1) / topics.length * 100) + '%';
        prev.disabled = i === 0; next.disabled = i === topics.length - 1;
        try { btns[i].scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } catch (e) {}
    }
    prev.addEventListener('click', () => show(cur - 1, -1));
    next.addEventListener('click', () => show(cur + 1, 1));
    detail.querySelector('.kb-back').addEventListener('click', () => { wrap.dataset.view = 'list'; });
    root._kb = { prev: () => show(cur - 1, -1), next: () => show(cur + 1, 1) };

    wrap.appendChild(list); wrap.appendChild(detail);
    root.appendChild(wrap);
    show(0, 1);
}

function renderInfoTabs(active) {
    const tabs = document.getElementById('info-tabs');
    tabs.innerHTML = '';
    ['help', 'terms', 'manual'].forEach(k => {
        const m = INFO_META[k];
        const b = document.createElement('button');
        b.type = 'button'; b.className = 'kiosk-tab' + (k === active ? ' on' : '');
        b.setAttribute('role', 'tab'); b.setAttribute('aria-selected', k === active ? 'true' : 'false');
        b.innerHTML = '<span class="kt-ic">' + m.ic + '</span><span class="kt-long">' + m.tab + '</span><span class="kt-short">' + m.short + '</span>';
        b.addEventListener('click', () => { if (k !== infoCurrent) switchInfo(k); });
        tabs.appendChild(b);
    });
}
function switchInfo(key) {
    const d = INFO_DOCS[key];
    if (!d) return;
    infoCurrent = key;
    renderInfoTabs(key);
    buildKiosk(document.getElementById('info-body'), d);
    infoIdleReset();
}
function openInfo(key) {
    if (!INFO_DOCS[key]) return;
    closeAllSidebars();
    infoReturnFocus = document.activeElement;
    switchInfo(key);
    document.getElementById('info-modal').classList.remove('hidden');
    infoIdleReset();
}
function closeInfo() {
    clearTimeout(infoIdleTimer);
    const m = document.getElementById('info-modal');
    if (m.classList.contains('hidden')) return;
    m.classList.add('hidden');
    try { if (infoReturnFocus && infoReturnFocus.focus) infoReturnFocus.focus(); } catch (e) {}
}
function infoIdleReset() {
    clearTimeout(infoIdleTimer);
    const m = document.getElementById('info-modal');
    if (m && !m.classList.contains('hidden')) infoIdleTimer = setTimeout(closeInfo, INFO_IDLE_MS);
}
['pointerdown', 'keydown', 'wheel', 'touchmove'].forEach(ev => document.getElementById('info-modal').addEventListener(ev, infoIdleReset, true));
document.getElementById('info-modal').addEventListener('scroll', infoIdleReset, true);

// "Was this helpful?" answers
document.addEventListener('click', e => {
    const b = e.target.closest && e.target.closest('.info-fb button');
    if (!b) return;
    const box = b.parentNode;
    box.classList.add('done');
    box.innerHTML = b.dataset.v === 'up' ? '✔ Glad that helped!' : 'Sorry about that. Please contact the support team.';
});
// keys: Esc closes, arrows turn the page
document.addEventListener('keydown', e => {
    const m = document.getElementById('info-modal');
    if (e.key === 'Escape') { closeInfo(); return; }
    if (m.classList.contains('hidden')) return;
    const kb = document.querySelector('#info-body.kb');
    if (!kb || !kb._kb) return;
    if (e.key === 'ArrowRight') kb._kb.next();
    if (e.key === 'ArrowLeft') kb._kb.prev();
});

// the User's manual on the right side of the welcome screen (wide screens only; hidden by CSS on narrow ones)
(function () {
    const el = document.getElementById('intro-manual-body');
    if (el && typeof INFO_DOCS !== 'undefined') buildKiosk(el, INFO_DOCS.manual, { contact: false });
})();


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
           document.getElementById('confirm-modal').classList.contains('hidden') &&
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
        { const act = st.includes('PAUSE') ? 'RESUME' : (st.includes('RUN') || st.includes('SHOT')) ? 'PAUSE' : 'START'; if (act === 'START') requestStart(); else sendCommand(act); }
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
    sendDirection(smashDir === 'MIXED' ? 'CENTER' : smashDir);          // servo preview
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
    sendDirection(p);                                                   // servo preview
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

// Logic to simulate hardware response
function sendCommand(command) {
    if ((command === 'START' || command === 'RESUME') && !usbIsConnected()) {
        toast('Connect the machine first (press the USB button).');
        return;
    }
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
        aimServoForStart();                       // servo: go to the chosen side before shuttles start
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
                // servo: swing to the NEXT shuttle's side after this shot has left
                if (mixedQueue.length) setTimeout(() => sendDirection(mixedQueue[0]), SERVO_MOVE_DELAY_MS);
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
        const summaryData = shuttlesLaunched > 0 ? Object.assign({
            launched: shuttlesLaunched, planned: totalShuttles,
            status: command === 'EMERGENCY_STOP' ? 'EMERGENCY_STOP' : (shuttlesLaunched >= totalShuttles ? 'COMPLETED' : 'STOPPED')
        }, shotCounts()) : null;
        finishLoggedSession(command === 'EMERGENCY_STOP' ? 'EMERGENCY_STOP' : (shuttlesLaunched >= totalShuttles ? 'COMPLETED' : 'STOPPED'));
        resetButtons();
        statusText.innerText = command === 'EMERGENCY_STOP' ? "EMERGENCY STOP" : "STOPPED";
        stateBox.classList.remove('flash-animation');
        clearInterval(simulationTimer);
        shuttlesLaunched = 0;
        updateCounterDisplay();
        if (currentMode === 'Smash') setSmashDirection(smashDir);          // back to "Mixed" / the chosen side
        if (summaryData) showDrillSummary(summaryData);
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

// ============================================================
//  USB SERIAL (Web Serial)  -  paste this ABOVE the line "initAimPads();"
//  at the very bottom of your JavaScript file.
//  Then DELETE the old function  sendToMachine(command) { ... }  (see the guide).
// ============================================================
let usbPort = null, usbWriter = null, usbReader = null, usbLoop = null;
let usbKeepAlive = null, usbClosing = false, usbReady = false;
const usbEncoder = new TextEncoder();

function usbIsConnected() { return !!(usbPort && usbWriter); }

function setUsbUi(text, on) {
    const b = document.getElementById('usb-btn');
    if (!b) return;
    b.innerText = text;
    b.classList.toggle('on', !!on);
}

async function usbSend(line) {
    if (!usbIsConnected()) throw new Error('USB not connected');
    await usbWriter.write(usbEncoder.encode(line + '\n'));
}

function handleUsbLine(line) {
    if (!line) return;
    console.log('[ARDUINO] ' + line);
    if (line === 'READY' || line === 'PONG') {
        if (!usbReady) { usbReady = true; setUsbUi('🔌 Machine ready (tap to disconnect)', true); }
    }
    if (line === 'FAILSAFE') toast('Machine stopped itself (no signal from the page).');
    if (line === 'EMPTY') setFeederEmpty(true);
    else if (line === 'LOADED') setFeederEmpty(false);
}

async function readUsbLoop() {
    const decoder = new TextDecoder();
    let buf = '';
    try {
        while (usbPort && usbPort.readable && !usbClosing) {
            usbReader = usbPort.readable.getReader();
            try {
                while (true) {
                    const { value, done } = await usbReader.read();
                    if (done) break;
                    buf += decoder.decode(value, { stream: true });
                    let i;
                    while ((i = buf.indexOf('\n')) >= 0) {
                        handleUsbLine(buf.slice(0, i).trim());
                        buf = buf.slice(i + 1);
                    }
                }
            } finally { usbReader.releaseLock(); }
        }
    } catch (e) { console.warn('USB read stopped:', e); }
}

async function connectUsb() {
    if (usbPort) { await disconnectUsb(); return; }
    if (!('serial' in navigator)) {
        toast('This browser cannot use USB. Use Chrome or Edge on a laptop.');
        return;
    }
    try {
        usbPort = await navigator.serial.requestPort();      // the browser shows a picker (choose the Arduino / COM port)
        await usbPort.open({ baudRate: 115200 });
        usbClosing = false; usbReady = false;
        usbWriter = usbPort.writable.getWriter();
        usbLoop = readUsbLoop();
        setUsbUi('🔌 USB connected - waiting for machine...', true);
        usbKeepAlive = setInterval(() => { usbSend('PING').catch(() => {}); }, 1000);
        usbSend('PING').catch(() => {});
        toast('USB connected');
    } catch (e) {
        usbPort = null; usbWriter = null;
        if (e && e.name !== 'NotFoundError') toast('Could not open USB: ' + e.message);   // NotFoundError = you closed the picker
    }
}

async function disconnectUsb() {
    usbClosing = true;
    clearInterval(usbKeepAlive);
    try { await usbSend('STOP'); } catch (e) {}
    try { if (usbReader) await usbReader.cancel(); } catch (e) {}
    try { if (usbLoop) await usbLoop; } catch (e) {}
    try { if (usbWriter) usbWriter.releaseLock(); } catch (e) {}
    try { if (usbPort) await usbPort.close(); } catch (e) {}
    usbPort = usbWriter = usbReader = usbLoop = null;
    usbReady = false;
    setUsbUi('🔌 Connect machine (USB)', false);
    setFeederEmpty(false, true);
}

// Cable pulled out
if ('serial' in navigator) {
    navigator.serial.addEventListener('disconnect', () => {
        usbClosing = true;
        clearInterval(usbKeepAlive);
        usbPort = usbWriter = usbReader = usbLoop = null;
        usbReady = false;
        setUsbUi('🔌 Connect machine (USB)', false);
        toast('USB cable disconnected. The machine stops by itself.');
        setFeederEmpty(false, true);
    });
}

// Sends the pressed button straight to the Arduino through the cable
function sendToMachine(command) {
    if (!usbIsConnected()) { toast('Machine not connected. Press the USB button first.'); return; }
    let line = command;
    if (command === 'START' || command === 'RESUME') line += ',' + selectedRange;
    const t0 = performance.now();
    usbSend(line).then(() => {
        console.log('[MACHINE] ' + line + ' written to USB in ' + Math.round(performance.now() - t0) + ' ms');
    }).catch(e => toast('Machine not reached: ' + e.message));
}

// ============================================================
//  DIRECTION SERVO (MG996R on PCA9685 channel 0)
// ============================================================
const SERVO_MOVE_DELAY_MS = 600;   // wait after a shot before the servo moves to the next side

// Tells the Arduino where to aim: 'LEFT' | 'RIGHT' | 'CENTER'
function sendDirection(side) {
    if (!usbIsConnected()) return;
    usbSend('DIR,' + side).catch(() => {});
}

// Aim for START / RESUME. Smash: Left / Right / Mixed (first shuttle's side). Service: Left / Right.
function aimServoForStart() {
    if (currentMode === 'Smash') {
        sendDirection(smashDir === 'MIXED' ? (mixedQueue[0] || 'LEFT') : smashDir);
    } else {
        sendDirection(+document.getElementById('h-dir').value > 0 ? 'RIGHT' : 'LEFT');
    }
}

initAimPads();
setRange(selectedRange);
applyImages();
syncThemeUi();

// ===== Launch confirmation: START opens a summary of the chosen controls; the drill only begins after the instructor confirms =====
const CONFIRM_REQUIRE_CLEAR = true;       // true = instructor must tick "court is clear" before Confirm works
const CONFIRM_COUNTDOWN = 3;              // seconds counted down after Confirm (0 = start right away)
let cfTimer = null, cfRunning = false;

function cfDirKey() {
    if (currentMode === 'Smash') return smashDir;                                   // LEFT | RIGHT | MIXED
    return (+document.getElementById('h-dir').value > 0) ? 'RIGHT' : 'LEFT';      // Service: Left or Right
}
function cfFmtTime(sec) {
    sec = Math.max(0, Math.round(sec));
    const m = Math.floor(sec / 60), s = sec % 60;
    return m ? m + ' min' + (s ? ' ' + s + ' sec' : '') : s + ' sec';
}
function cfIcon(sel) { const s = document.querySelector(sel); return s ? s.outerHTML : ''; }

function requestStart() {
    if (!usbIsConnected()) { toast('Connect the machine first (press the USB button).'); return; }
    openLaunchConfirm();
}
function openLaunchConfirm() {
    const dir = cfDirKey();
    const sec = getInterval(), n = getCount();
    const rangeIcon = cfIcon('#range-opts button[data-r="' + selectedRange + '"] svg');
    const dirIcon = currentMode === 'Smash' ? cfIcon('#smash-dir-opts button[data-d="' + dir + '"] svg') : cfIcon('#position-opts button[data-p="' + dir + '"] svg');

    setText('cf-mode', currentMode.toUpperCase() + ' MODE');
    setText('cf-class', document.getElementById('session-class').innerText || '');
    document.getElementById('cf-tiles').innerHTML =
        '<button type="button" class="cf-tile" data-edit="range"><span class="cf-k">Range</span><span class="cf-ic">' + rangeIcon + '</span><b>' + RANGE_LABELS[selectedRange] + '</b><em>Tap to change</em></button>' +
        '<button type="button" class="cf-tile" data-edit="dir"><span class="cf-k">Direction</span><span class="cf-ic">' + dirIcon + '</span><b>' + SMASH_DIR_NAMES[dir] + '</b><em>Tap to change</em></button>' +
        '<div class="cf-tile static"><span class="cf-k">Interval</span><span class="cf-ic big">⏱️</span><b>' + sec + ' sec</b><em>between shuttles</em></div>' +
        '<div class="cf-tile static"><span class="cf-k">Shuttles</span><span class="cf-ic big">🏸</span><b>' + n + '</b><em>to launch</em></div>';
    setText('cf-time', '≈ ' + cfFmtTime(sec * n));

    const chk = document.getElementById('cf-check');
    chk.classList.toggle('hidden', !CONFIRM_REQUIRE_CLEAR);
    chk.setAttribute('aria-checked', 'false'); chk.classList.remove('on');
    cfSyncGo();

    cfReset();
    document.getElementById('confirm-modal').classList.remove('hidden');
    setTimeout(() => { try { document.getElementById(CONFIRM_REQUIRE_CLEAR ? 'cf-check' : 'cf-go').focus(); } catch (e) {} }, 60);
}
function cfSyncGo() {
    const ok = !CONFIRM_REQUIRE_CLEAR || document.getElementById('cf-check').classList.contains('on');
    const go = document.getElementById('cf-go');
    go.disabled = !ok;
    go.classList.toggle('ready', ok);
}
function cfToggleClear() {
    const c = document.getElementById('cf-check');
    const on = !c.classList.contains('on');
    c.classList.toggle('on', on); c.setAttribute('aria-checked', on ? 'true' : 'false');
    cfSyncGo();
}
function cfReset() {
    clearInterval(cfTimer); cfTimer = null; cfRunning = false;
    document.getElementById('cf-card').classList.remove('counting');
}
function closeLaunchConfirm() {
    cfReset();
    document.getElementById('confirm-modal').classList.add('hidden');
}
function cfEdit(which) {
    closeLaunchConfirm();
    const target = which === 'range' ? document.getElementById('range-opts')
        : document.getElementById(currentMode === 'Smash' ? 'smash-dir-opts' : 'position-opts');
    const tile = target && target.closest('.speed-tile, .aim-pad');
    if (!tile) return;
    try { tile.scrollIntoView({ block: 'center', behavior: 'smooth' }); } catch (e) {}
    tile.classList.remove('cf-pulse'); void tile.offsetWidth; tile.classList.add('cf-pulse');
    setTimeout(() => tile.classList.remove('cf-pulse'), 2200);
}
function cfConfirm() {
    if (cfRunning) return;
    if (CONFIRM_REQUIRE_CLEAR && !document.getElementById('cf-check').classList.contains('on')) return;
    if (!usbIsConnected()) { closeLaunchConfirm(); toast('Connect the machine first (press the USB button).'); return; }
    if (CONFIRM_COUNTDOWN <= 0) { closeLaunchConfirm(); sendCommand('START'); return; }
    cfRunning = true;
    const card = document.getElementById('cf-card'), num = document.getElementById('cf-num');
    card.classList.add('counting');
    let left = CONFIRM_COUNTDOWN;
    const tick = () => {
        num.textContent = left > 0 ? left : 'GO!';
        num.classList.remove('pop'); void num.offsetWidth; num.classList.add('pop');
        if (left < 0) {
            clearInterval(cfTimer); cfTimer = null;
            closeLaunchConfirm(); sendCommand('START');
            return;
        }
        left--;
    };
    tick();
    cfTimer = setInterval(tick, 1000);
}
document.getElementById('cf-tiles').addEventListener('click', e => {
    const t = e.target.closest('[data-edit]');
    if (t) cfEdit(t.dataset.edit);
});
document.addEventListener('keydown', e => {
    const m = document.getElementById('confirm-modal');
    if (m.classList.contains('hidden')) return;
    if (e.key === 'Escape') closeLaunchConfirm();
    else if (e.key === 'Enter' && !cfRunning && !document.getElementById('cf-go').disabled && document.activeElement.id !== 'cf-check') { e.preventDefault(); cfConfirm(); }
});


// ============================================================
//  KIOSK FEATURES: drill summary, idle screen, empty-feeder alert
// ============================================================

// ---------- Drill summary (shown when a drill ends; returns Home after 15 s) ----------
const SUMMARY_MS = 15000;
let summaryTimer = null, summaryTick = null;

function activeScreenId() { const a = document.querySelector('.screen.active'); return a ? a.id : ''; }
function isLoggedInScreen() { const id = activeScreenId(); return !!id && id !== 'login-screen' && id !== 'signup-screen' && !id.startsWith('onboarding'); }

function showDrillSummary(d) {
    if (!isLoggedInScreen()) return;                           // e.g. logging out mid-drill
    const titles = { COMPLETED: ['✓ Drill complete', 'Nice work!', ''], STOPPED: ['■ Drill stopped', 'Drill stopped', 'stopped'], EMERGENCY_STOP: ['⚠ Emergency stop', 'Emergency stop', 'emergency'] };
    const t = titles[d.status] || titles.STOPPED;
    const badge = document.getElementById('sum-badge');
    badge.textContent = t[0]; badge.className = 'sum-badge ' + t[2];
    setText('sum-title', t[1]);
    setText('sum-sub', [currentMode, RANGE_LABELS[selectedRange] || '', currentClass ? classLabel(currentClass) : ''].filter(Boolean).join(' · '));
    setText('sum-launched', String(d.launched));
    setText('sum-of', 'of ' + d.planned + ' shuttles launched');
    const scored = d.hits + d.misses;
    document.getElementById('sum-stats').classList.toggle('hidden', scored === 0);
    setText('sum-hits', String(d.hits));
    setText('sum-misses', String(d.misses));
    setText('sum-acc', scored ? Math.round(d.hits / scored * 100) + '%' : '—');

    document.getElementById('summary-modal').classList.remove('hidden');
    const bar = document.getElementById('sum-bar');
    bar.classList.remove('run'); void bar.offsetWidth; bar.classList.add('run');
    const end = Date.now() + SUMMARY_MS;
    clearInterval(summaryTick); clearTimeout(summaryTimer);
    const tick = () => setText('sum-count', 'Returning to Home in ' + Math.max(0, Math.ceil((end - Date.now()) / 1000)) + 's');
    tick(); summaryTick = setInterval(tick, 250);
    summaryTimer = setTimeout(() => closeSummary(true), SUMMARY_MS);
}
function closeSummary(goHome) {
    clearInterval(summaryTick); clearTimeout(summaryTimer);
    const m = document.getElementById('summary-modal');
    if (m.classList.contains('hidden')) return;
    m.classList.add('hidden');
    if (goHome && isLoggedInScreen()) navigate('dashboard-screen');
}
function summaryRunAgain() {
    closeSummary(false);
    if (activeScreenId() !== 'session-screen') quickStart();    // reload the same drill settings
    requestStart();                                             // the usual "Ready to launch?" check still appears
}

// ---------- Idle / attract screen (2 minutes without a touch) ----------
const IDLE_MS = 120000;
let idleTimer = null, idleClockTimer = null;

function drillIsRunning() { const stopBtn = document.getElementById('btn-stop'); return !!(activeSession || (stopBtn && !stopBtn.classList.contains('hidden'))); }   // running or paused
function idleVisible() { return !document.getElementById('idle-screen').classList.contains('hidden'); }

function updateIdleClock() {
    const now = new Date(), h = now.getHours();
    setTextIfChanged('idle-hm', `${h % 12 || 12}:${String(now.getMinutes()).padStart(2, '0')}`);
    setTextIfChanged('idle-ap', h < 12 ? 'AM' : 'PM');
    setTextIfChanged('idle-date', now.toLocaleDateString('en-PH', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }));
}
function goIdle() {
    if (idleVisible()) return;
    if (drillIsRunning()) { resetIdleTimer(); return; }         // never interrupt a running or paused drill
    closeSummary(false);
    closeLaunchConfirm();
    closeInfo();
    if (typeof closeHistory === 'function') { try { closeHistory(); } catch (e) {} }
    dismissEmptyAlert();
    closeAllSidebars();
    if (isLoggedInScreen() && activeScreenId() !== 'dashboard-screen') navigate('dashboard-screen');
    updateIdleClock();
    clearInterval(idleClockTimer); idleClockTimer = setInterval(updateIdleClock, 1000);
    document.getElementById('idle-screen').classList.remove('hidden');
}
function wakeIdle() {
    clearInterval(idleClockTimer);
    document.getElementById('idle-screen').classList.add('hidden');
    resetIdleTimer();
}
function resetIdleTimer() {
    clearTimeout(idleTimer);
    if (!idleVisible()) idleTimer = setTimeout(goIdle, IDLE_MS);
}
['pointerdown', 'keydown', 'wheel', 'touchstart'].forEach(ev => document.addEventListener(ev, () => { if (!idleVisible()) resetIdleTimer(); }, true));
document.addEventListener('keydown', e => { if (idleVisible()) { e.preventDefault(); wakeIdle(); } }, true);
resetIdleTimer();

// ---------- No shuttlecocks left (IR sensor) ----------
let feederEmpty = false;
let audioCtx = null;

function getAudio() {
    try {
        audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
        if (audioCtx.state === 'suspended') audioCtx.resume();
    } catch (e) { return null; }
    return audioCtx;
}
// the browser only allows sound after a touch/click, so unlock it on the first one
document.addEventListener('pointerdown', () => { getAudio(); }, { once: true, capture: true });

function beepFor(ms, freq) {
    const ctx = getAudio();
    if (!ctx) return;
    const t = ctx.currentTime, dur = ms / 1000;
    const osc = ctx.createOscillator(), gain = ctx.createGain();
    osc.type = 'square'; osc.frequency.value = freq || 880;
    gain.gain.setValueAtTime(0.0001, t);
    gain.gain.exponentialRampToValueAtTime(0.25, t + 0.03);
    gain.gain.setValueAtTime(0.25, t + dur - 0.05);
    gain.gain.exponentialRampToValueAtTime(0.0001, t + dur);
    osc.connect(gain); gain.connect(ctx.destination);
    osc.start(t); osc.stop(t + dur + 0.02);
}

function setFeederEmpty(empty, silent) {
    if (empty === feederEmpty) return;                         // the Arduino repeats its status; only act on a change
    feederEmpty = empty;
    const m = document.getElementById('empty-modal');
    if (empty) {
        m.classList.remove('hidden');
        beepFor(1500, 880);                                    // 1.5-second beep
    } else {
        m.classList.add('hidden');
        if (!silent) toast('Shuttlecocks detected. The feeder is loaded.');
    }
}
function dismissEmptyAlert() { document.getElementById('empty-modal').classList.add('hidden'); }
