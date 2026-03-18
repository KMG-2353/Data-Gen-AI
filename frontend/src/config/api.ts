// API Configuration
const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:8000';

export const apiConfig = {
  baseURL: API_BASE_URL,
  endpoints: {
    upload: `${API_BASE_URL}/api/upload`,
    generate: `${API_BASE_URL}/api/generate`,
    download: (sessionId: string) => `${API_BASE_URL}/api/download/${sessionId}`,
  },
};

export default apiConfig;
