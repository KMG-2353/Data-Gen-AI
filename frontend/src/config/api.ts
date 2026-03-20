// API Configuration
// Production: https://data-gen-ai.onrender.com
// Development: http://localhost:8000
const API_BASE_URL = import.meta.env.VITE_API_URL || 'https://data-gen-ai.onrender.com';

export const apiConfig = {
  baseURL: API_BASE_URL,
  endpoints: {
    upload: `${API_BASE_URL}/api/upload`,
    analyze: `${API_BASE_URL}/api/analyze`,
    generate: `${API_BASE_URL}/api/generate`,
    download: (sessionId: string) => `${API_BASE_URL}/api/download/${sessionId}`,
    getRules: (sessionId: string) => `${API_BASE_URL}/api/rules/${sessionId}`,
    updateRule: `${API_BASE_URL}/api/rules/update`,
    repromptRule: `${API_BASE_URL}/api/rules/reprompt`,
    health: `${API_BASE_URL}/api/health`,
  },
};

// Warm up the backend server on app load
export const warmUpServer = async () => {
  try {
    await fetch(apiConfig.endpoints.health);
  } catch {
    // Silently ignore errors - this is just a warm-up call
  }
};

export default apiConfig;
