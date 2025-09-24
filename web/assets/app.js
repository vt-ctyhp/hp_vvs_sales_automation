(function () {
  const TOKEN_KEY = 'vvsapp_token';
  const EMAIL_KEY = 'vvsapp_email';
  const ROLE_KEY = 'vvsapp_role';

  async function login(email, password) {
    const response = await fetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password }),
    });
    if (!response.ok) {
      const data = await safeJson(response);
      throw new Error(data.error || 'Login failed');
    }
    const data = await response.json();
    localStorage.setItem(TOKEN_KEY, data.token);
    if (data.email) {
      localStorage.setItem(EMAIL_KEY, data.email);
    }
    if (data.role) {
      localStorage.setItem(ROLE_KEY, data.role);
    }
    return data;
  }

  function logout() {
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem(EMAIL_KEY);
    localStorage.removeItem(ROLE_KEY);
  }

  function getToken() {
    return localStorage.getItem(TOKEN_KEY);
  }

  function requireAuth() {
    const token = getToken();
    if (!token) {
      window.location.href = '/';
      throw new Error('Authentication required');
    }
    return token;
  }

  async function apiFetch(path, options = {}) {
    const token = getToken();
    const opts = Object.assign({ headers: {} }, options);
    const headers = new Headers(opts.headers);
    if (token) {
      headers.set('Authorization', `Bearer ${token}`);
    }
    if (!headers.has('Content-Type') && !(opts.body instanceof FormData)) {
      headers.set('Content-Type', 'application/json');
    }
    opts.headers = headers;

    const response = await fetch(path, opts);
    if (response.status === 401) {
      logout();
      window.location.href = '/';
      throw new Error('Unauthorized');
    }
    const data = await safeJson(response);
    if (!response.ok) {
      throw new Error((data && data.error) || response.statusText);
    }
    return data;
  }

  async function safeJson(response) {
    try {
      return await response.json();
    } catch (err) {
      return {};
    }
  }

  function formatDateTimeLocal(value) {
    if (!value) return '';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return '';
    }
    const pad = (n) => String(n).padStart(2, '0');
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
  }

  function parseDateTimeLocal(value) {
    if (!value) return null;
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return null;
    }
    return date.toISOString();
  }

  window.vvsapp = {
    login,
    logout,
    getToken,
    requireAuth,
    apiFetch,
    formatDateTimeLocal,
    parseDateTimeLocal,
  };
})();
