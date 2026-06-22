// api.js

// ─── 1. Backend URLs ────────────────────────────────────────────────────────

const LOCAL_API_URL = "http://127.0.0.1:8000";

const ONLINE_API_URL =
    "https://conversion-studio-five.vercel.app";


// ─── 2. Detect local or hosted frontend ─────────────────────────────────────

const _hostname = window.location.hostname;

const _isLocal =
    _hostname === "localhost" ||
    _hostname === "127.0.0.1";


// ─── 3. Select backend ──────────────────────────────────────────────────────

const API_BASE = (
    _isLocal
        ? LOCAL_API_URL
        : ONLINE_API_URL
)
    .trim()
    .replace(/\/+$/, "");


// ─── 4. Build API URL ───────────────────────────────────────────────────────

function apiURL(path) {
    const normalizedPath = String(path || "").startsWith("/")
        ? String(path)
        : `/${String(path || "")}`;

    return `${API_BASE}${normalizedPath}`;
}


// ─── 5. Export values to index.html ─────────────────────────────────────────

window.API_BASE = API_BASE;
window.apiURL = apiURL;


/*
 * Live Connect is available only when the application is running locally.
 *
 * Local Windows:
 * true
 *
 * Hosted Vercel:
 * false
 */
window.LIVE_CONNECT_AVAILABLE = _isLocal;


// ─── 6. Debug information ───────────────────────────────────────────────────

console.info(
    "%c[Conversion Studio] API",
    "color:#c9a84c;font-weight:700",
    `→ ${API_BASE} ${_isLocal ? "(local Windows mode)" : "(hosted conversion mode)"}`
);