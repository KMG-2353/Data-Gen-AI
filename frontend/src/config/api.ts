// API Configuration
// Production: uses VITE_API_URL (set via Render env)
// Development: derives host from window.location so it works from any device
//   on the same network (Mac, iPhone, etc.)

function getApiBaseUrl(): string {
  const envUrl = import.meta.env.VITE_API_URL as string | undefined;

  // If an explicit URL is configured AND it's not localhost, use it as-is
  // (handles both production deployments and explicit overrides).
  if (envUrl && !envUrl.includes('localhost')) {
    return envUrl;
  }

  // In development (or when VITE_API_URL points to localhost), derive the
  // backend host from the current page's hostname so that iPhone / other
  // devices on the same network automatically hit the right backend.
  if (typeof window !== 'undefined') {
    return `http://${window.location.hostname}:8000`;
  }

  return 'http://localhost:8000';
}

const API_BASE_URL = getApiBaseUrl();

export const apiConfig = {
  baseURL: API_BASE_URL,
  endpoints: {
    upload: `${API_BASE_URL}/api/upload`,
    generate: `${API_BASE_URL}/api/generate`,
    download: (sessionId: string) => `${API_BASE_URL}/api/download/${sessionId}`,
  },
};

export default apiConfig;
