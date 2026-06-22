const LOCAL_API_URL = "http://127.0.0.1:8000";
const HOSTED_API_URL = "https://conversion-studio-five.vercel.app";

const IS_LOCAL_WINDOWS_MODE =
    window.location.hostname === "localhost" ||
    window.location.hostname === "127.0.0.1";

const API_BASE = IS_LOCAL_WINDOWS_MODE
    ? LOCAL_API_URL
    : HOSTED_API_URL;

function apiURL(path) {
    const normalizedPath = path.startsWith("/")
        ? path
        : `/${path}`;

    return `${API_BASE}${normalizedPath}`;
}

// Export to global window namespace
window.LOCAL_API_URL = LOCAL_API_URL;
window.HOSTED_API_URL = HOSTED_API_URL;
window.IS_LOCAL_WINDOWS_MODE = IS_LOCAL_WINDOWS_MODE;
window.API_BASE = API_BASE;
window.apiURL = apiURL;