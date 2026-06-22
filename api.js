// Conversion Studio API routing
// Standard conversion runs in Vercel. Excel COM live-connect runs on this Windows PC.

const CLOUD_API_URL = "https://conversion-studio-five.vercel.app";
const LOCAL_AGENT_API_URL = "http://127.0.0.1:8000";

const currentHost = window.location.hostname;
const isLocalPage = currentHost === "localhost" || currentHost === "127.0.0.1";

function normalizeBase(value) {
  return String(value || "").trim().replace(/\/+$/, "");
}

function normalizePath(path) {
  const value = String(path || "");
  return value.startsWith("/") ? value : `/${value}`;
}

const STANDARD_API_BASE = normalizeBase(isLocalPage ? LOCAL_AGENT_API_URL : CLOUD_API_URL);
const LIVE_API_BASE = normalizeBase(LOCAL_AGENT_API_URL);

function apiURL(path) {
  return `${STANDARD_API_BASE}${normalizePath(path)}`;
}

function liveApiURL(path) {
  return `${LIVE_API_BASE}${normalizePath(path)}`;
}

window.CLOUD_API_URL = normalizeBase(CLOUD_API_URL);
window.LOCAL_AGENT_API_URL = normalizeBase(LOCAL_AGENT_API_URL);
window.API_BASE = STANDARD_API_BASE;
window.LIVE_API_BASE = LIVE_API_BASE;
window.apiURL = apiURL;
window.liveApiURL = liveApiURL;
window.IS_LOCAL_WINDOWS_MODE = isLocalPage;

console.info("[Conversion Studio] Standard API:", STANDARD_API_BASE);
console.info("[Conversion Studio] Windows agent:", LIVE_API_BASE);
