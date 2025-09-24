const healthOverall = document.querySelector('#health-overall');
const healthDb = document.querySelector('#health-db');
const healthTime = document.querySelector('#health-time');
const healthError = document.querySelector('#health-error');
const refreshHealthButton = document.querySelector('#refresh-health');
const loginForm = document.querySelector('#login-form');
const loginMessage = document.querySelector('#login-message');
const sessionOutput = document.querySelector('#session-output');
const copyrightYear = document.querySelector('#copyright-year');

const API_BASE = '/api';

function setHealthLoading() {
    healthOverall.textContent = 'Loading…';
    healthDb.textContent = 'Loading…';
    healthTime.textContent = 'Loading…';
    healthError.hidden = true;
}

async function fetchHealthStatus() {
    setHealthLoading();
    try {
        const response = await fetch(`${API_BASE}/health`, { cache: 'no-store' });
        if (!response.ok) {
            throw new Error(`Request failed with status ${response.status}`);
        }
        const payload = await response.json();
        healthOverall.textContent = payload.status ?? 'unknown';
        healthDb.textContent = payload.db ?? 'unknown';
        healthTime.textContent = payload.nowIso ?? 'unknown';
    } catch (error) {
        console.error('Health request failed', error);
        healthError.textContent = 'Unable to load system health. Please try again.';
        healthError.hidden = false;
        healthOverall.textContent = 'error';
        healthDb.textContent = 'error';
        healthTime.textContent = '--';
    }
}

function setLoginMessage(message, type = 'info') {
    if (!message) {
        loginMessage.hidden = true;
        return;
    }
    loginMessage.textContent = message;
    loginMessage.hidden = false;
    loginMessage.dataset.type = type;
    loginMessage.classList.toggle('error-message', type === 'error');
}

function renderSessionDetails(details) {
    if (!details) {
        sessionOutput.innerHTML = '<p>No active session. Log in to see your access token and role.</p>';
        return;
    }

    const formatted = JSON.stringify(details, null, 2);
    sessionOutput.innerHTML = `
        <p><strong>Authenticated as:</strong> ${details.email}</p>
        <p><strong>Role:</strong> ${details.role}</p>
        <details open>
            <summary>Access token</summary>
            <pre><code>${formatted}</code></pre>
        </details>
    `;
}

async function handleLoginSubmit(event) {
    event.preventDefault();
    setLoginMessage('Signing in…');

    const formData = new FormData(loginForm);
    const email = formData.get('email');
    const password = formData.get('password');

    if (!email || !password) {
        setLoginMessage('Please provide both email and password.', 'error');
        return;
    }

    try {
        const response = await fetch(`${API_BASE}/auth/login`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ email, password }),
        });

        if (!response.ok) {
            throw new Error('Invalid credentials');
        }

        const payload = await response.json();
        const sessionDetails = {
            email: payload.email,
            role: payload.role,
            token: payload.token,
        };
        localStorage.setItem('vvsapp-session', JSON.stringify(sessionDetails));
        setLoginMessage('Successfully signed in.');
        renderSessionDetails(sessionDetails);
        loginForm.reset();
    } catch (error) {
        console.error('Login failed', error);
        setLoginMessage(error.message || 'Login failed. Please check your credentials.', 'error');
    }
}

function restoreSession() {
    try {
        const saved = localStorage.getItem('vvsapp-session');
        if (!saved) {
            renderSessionDetails(null);
            return;
        }
        const details = JSON.parse(saved);
        renderSessionDetails(details);
    } catch (error) {
        console.warn('Failed to restore session from storage', error);
        renderSessionDetails(null);
    }
}

function init() {
    copyrightYear.textContent = new Date().getFullYear();
    restoreSession();
    fetchHealthStatus();
}

refreshHealthButton?.addEventListener('click', fetchHealthStatus);
loginForm?.addEventListener('submit', handleLoginSubmit);

document.addEventListener('DOMContentLoaded', init);
