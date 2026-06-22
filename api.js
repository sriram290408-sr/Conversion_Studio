// api.js

// ─── 1. Connection constants ────────────────────────────────────────────────

/**
 * Local FastAPI backend.
 */
const LOCAL_API_URL = "http://127.0.0.1:8000";

/*
* Online FastAPI backend.
*/
const ONLINE_API_URL =
    "https://conversion-studio-five.vercel.app/api";


// ─── 2. Environment detection ───────────────────────────────────────────────

const _hostname = window.location.hostname;

const _isLocal =
    _hostname === "127.0.0.1" ||
    _hostname === "localhost";


// ─── 3. Optional query-string configuration ─────────────────────────────────

/**
 * Online backend can also be configured through:
 *
 * https://conversion-studio-rho.vercel.app
 * ?api=https://your-backend.vercel.app
 */
const _queryParameter = new URLSearchParams(
    window.location.search
).get("api");

if (_queryParameter && !_isLocal) {
    const _normalizedQueryUrl = _queryParameter
        .trim()
        .replace(/\/+$/, "");

    localStorage.setItem(
        "conversion-studio-api-base",
        _normalizedQueryUrl
    );
}


// ─── 4. Stored online backend ────────────────────────────────────────────────

const _storedOnlineApiUrl = localStorage.getItem(
    "conversion-studio-api-base"
);


// ─── 5. API base resolution ─────────────────────────────────────────────────

/**
 * Local execution always uses LOCAL_API_URL.
 *
 * Online execution uses:
 * 1. Backend URL stored using ?api=
 * 2. ONLINE_API_URL constant
 */
const API_BASE = (
    _isLocal
        ? LOCAL_API_URL
        : (_storedOnlineApiUrl || ONLINE_API_URL)
)
    .trim()
    .replace(/\/+$/, "");


// ─── 6. Validation ───────────────────────────────────────────────────────────

if (
    !_isLocal &&
    (
        !API_BASE ||
        API_BASE.includes("YOUR-ACTUAL-BACKEND-DOMAIN")
    )
) {
    console.error(
        "[Conversion Studio] Backend URL is not configured. " +
        "Open api.js and replace ONLINE_API_URL with the actual backend URL."
    );
}

if (
    !_isLocal &&
    API_BASE === window.location.origin
) {
    console.error(
        "[Conversion Studio] The frontend URL is being used as the backend URL. " +
        "Set ONLINE_API_URL to the deployed FastAPI backend domain."
    );
}


// ─── 7. API URL helper ───────────────────────────────────────────────────────

/**
 * Creates the complete URL for a backend endpoint.
 *
 * Examples:
 *
 * apiURL("/upload")
 * apiURL("/health")
 * apiURL(`/download/${sessionId}`)
 *
 * @param {string} path Backend endpoint path.
 * @returns {string} Complete backend URL.
 */
function apiURL(path) {
    if (!API_BASE) {
        throw new Error(
            "Conversion Studio backend URL is not configured."
        );
    }

    const normalizedPath = String(path || "").startsWith("/")
        ? String(path)
        : `/${String(path || "")}`;

    return `${API_BASE}${normalizedPath}`;
}


// ─── 8. Export for index.html ────────────────────────────────────────────────

window.API_BASE = API_BASE;
window.apiURL = apiURL;


// ─── 9. Debug information ────────────────────────────────────────────────────

console.info(
    "%c[Conversion Studio] API",
    "color:#c9a84c;font-weight:700",
    `→ ${API_BASE} ${_isLocal ? "(local)" : "(online)"}`
);