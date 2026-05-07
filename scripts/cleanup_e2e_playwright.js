const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');
const { chromium } = require('playwright');

function requiredEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing required env var ${name}`);
  return value;
}

function optionalEnv(name, fallback) {
  return process.env[name] || fallback;
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function fileExists(target) {
  try {
    fs.accessSync(target);
    return true;
  } catch (_) {
    return false;
  }
}

async function saveShot(page, outDir, name) {
  const file = path.join(outDir, `${name}.png`);
  await page.screenshot({ path: file, fullPage: true });
  return file;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function optionalBoolEnv(name, fallback) {
  const value = process.env[name];
  if (value == null || value === '') return fallback;
  return !/^(0|false|no)$/i.test(String(value));
}

function chromeBinaryPath() {
  const candidates = [
    process.env.CHROME_BINARY,
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary'
  ].filter(Boolean);
  for (const candidate of candidates) {
    if (fileExists(candidate)) return candidate;
  }
  throw new Error('Could not find a Chrome binary. Set CHROME_BINARY if Chrome is installed in a nonstandard location.');
}

async function runCommand(cmd, args) {
  await new Promise((resolve, reject) => {
    const child = spawn(cmd, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let stderr = '';
    child.stderr.on('data', (chunk) => { stderr += String(chunk || ''); });
    child.on('error', reject);
    child.on('close', (code) => {
      if (code === 0) return resolve();
      reject(new Error(`${cmd} exited with code ${code}${stderr ? `: ${stderr.trim()}` : ''}`));
    });
  });
}

async function copyChromeProfile(sourceUserDataDir, profileName) {
  if (!sourceUserDataDir || !fileExists(sourceUserDataDir)) {
    throw new Error(`Chrome source user-data dir does not exist: ${sourceUserDataDir}`);
  }
  const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'cleanup-e2e-chrome-'));
  const sourceProfileDir = path.join(sourceUserDataDir, profileName);
  const sourceLocalState = path.join(sourceUserDataDir, 'Local State');
  if (!fileExists(sourceProfileDir)) {
    throw new Error(`Chrome profile directory not found: ${sourceProfileDir}`);
  }
  if (!fileExists(sourceLocalState)) {
    throw new Error(`Chrome Local State not found: ${sourceLocalState}`);
  }
  await runCommand('rsync', [
    '-a',
    '--delete',
    '--exclude=*/Cache/*',
    '--exclude=*/Code Cache/*',
    '--exclude=*/GPUCache/*',
    '--exclude=*/ShaderCache/*',
    '--exclude=*/Service Worker/*',
    '--exclude=*/Singleton*',
    '--exclude=*/lockfile',
    sourceLocalState,
    sourceProfileDir,
    `${tempRoot}/`
  ]);
  return tempRoot;
}

async function waitForDevtools(port, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`http://127.0.0.1:${port}/json/version`);
      if (res.ok) return await res.json();
    } catch (_) {}
    await sleep(500);
  }
  throw new Error(`Timed out waiting for Chrome DevTools on port ${port}.`);
}

async function launchChromeForCdp(userDataDir, profileName, headless) {
  const chromePath = chromeBinaryPath();
  const port = 9222 + Math.floor(Math.random() * 2000);
  const args = [
    `--user-data-dir=${userDataDir}`,
    `--profile-directory=${profileName}`,
    `--remote-debugging-port=${port}`,
    '--no-first-run',
    '--disable-background-networking',
    '--disable-default-apps',
    '--disable-sync',
    '--no-default-browser-check',
    'about:blank'
  ];
  if (headless) args.splice(args.length - 1, 0, '--headless=new');
  const proc = spawn(chromePath, args, { stdio: ['ignore', 'pipe', 'pipe'] });
  let stderr = '';
  proc.stderr.on('data', (chunk) => { stderr += String(chunk || ''); });
  proc.on('error', (err) => {
    stderr += `\n${err && err.message ? err.message : String(err)}`;
  });
  const version = await waitForDevtools(port, 30000).catch(async (err) => {
    try { proc.kill('SIGTERM'); } catch (_) {}
    throw new Error(`${err.message}\nChrome stderr:\n${stderr.trim()}`);
  });
  return {
    browserWs: version.webSocketDebuggerUrl,
    proc,
    port,
    stderrRef: () => stderr
  };
}

async function safeClose(resource) {
  if (!resource) return;
  try {
    await resource.close();
  } catch (_) {}
}

async function stopChromeLaunch(launchInfo) {
  if (!launchInfo || !launchInfo.proc) return;
  const proc = launchInfo.proc;
  if (proc.exitCode != null) return;
  await new Promise((resolve) => {
    const done = () => resolve();
    proc.once('close', done);
    try {
      proc.kill('SIGTERM');
    } catch (_) {
      resolve();
      return;
    }
    setTimeout(() => {
      if (proc.exitCode == null) {
        try { proc.kill('SIGKILL'); } catch (_) {}
      }
    }, 3000);
  });
}

async function waitForTaskList(page) {
  const loading = page.locator('#taskList .loading');
  try {
    await loading.last().waitFor({ state: 'hidden', timeout: 30000 });
  } catch (_) {}
}

async function gotoApp(page, webAppUrl, outDir, label) {
  await page.goto(webAppUrl, { waitUntil: 'load', timeout: 120000 });
  await page.waitForLoadState('domcontentloaded');
  try {
    await page.waitForLoadState('networkidle', { timeout: 15000 });
  } catch (_) {}

  const googleGate = await page.locator('text=Choose an account').count();
  if (googleGate) {
    await saveShot(page, outDir, `${label}-google-gate`);
    throw new Error('Google account chooser appeared; copied Chrome profile did not grant domain access to the web app.');
  }
}

async function showLoginIfNeeded(page) {
  const loginScreen = page.locator('#loginScreen');
  if (await loginScreen.isVisible().catch(() => false)) return;
  const logout = page.locator('#logoutButton');
  if (await logout.isVisible().catch(() => false)) {
    await logout.click();
    await loginScreen.waitFor({ state: 'visible', timeout: 30000 });
  }
}

async function login(page, outDir, email, password, label) {
  await showLoginIfNeeded(page);
  await page.locator('#loginEmail').fill(email);
  await page.locator('#loginPassword').fill(password);
  await saveShot(page, outDir, `${label}-login-ready`);
  await page.locator('#loginButton').click();
  await page.locator('#appShell').waitFor({ state: 'visible', timeout: 120000 });
  await page.locator('#tab-cleanup').waitFor({ state: 'visible', timeout: 120000 });
  await saveShot(page, outDir, `${label}-dashboard`);
}

async function openCleanupTask(page, outDir, customerName, label) {
  await page.locator('#tab-cleanup').click();
  await waitForTaskList(page);

  let card = page.locator('.task-card').filter({ hasText: customerName }).first();
  if (!(await card.count())) {
    const refresh = page.locator('#refreshQueueButton');
    if (await refresh.isVisible().catch(() => false)) {
      await refresh.click();
      await waitForTaskList(page);
      card = page.locator('.task-card').filter({ hasText: customerName }).first();
    }
  }
  await card.waitFor({ state: 'visible', timeout: 120000 });
  await saveShot(page, outDir, `${label}-task-card`);
  await card.locator('button.js-open-task').click();
  await page.locator('#detailPanel').waitFor({ state: 'visible', timeout: 120000 });
  await page.locator('#detailPanel').locator('text=' + customerName).first().waitFor({ state: 'visible', timeout: 120000 }).catch(() => {});
  await saveShot(page, outDir, `${label}-task-open`);
}

async function submitProposal(page, outDir, customerName) {
  await openCleanupTask(page, outDir, customerName, 'lyn');
  await page.locator('#cleanupSalesStage').selectOption({ label: 'Lost Lead' });
  await page.locator('#cleanupConvStatus').selectOption({ label: 'Lost Lead' });
  await page.locator('#cleanupCustomOrder').selectOption('');
  await page.locator('#cleanupCenterStone').selectOption({ label: 'No Center Stone' });
  await page.locator('#cleanupLostLeadReason').selectOption({ label: 'Unable to contact' });
  await page.locator('#cleanupLostLeadNotes').fill('Playwright E2E stale-cleanup test: customer marked lost after no response.');
  await page.locator('#cleanupNextSteps').fill('Playwright E2E test closed this stale customer as lost lead.');
  await page.locator('#cleanupNotes').fill('Submitted from Lyn account during cleanup E2E run.');
  await page.locator('#cleanupOwnersVerified').check();
  await page.locator('#cleanupContactVerified').check();
  await page.locator('#cleanupOpsStatusReviewed').check();
  await page.locator('#cleanupNextStepsCurrent').check();
  await saveShot(page, outDir, 'lyn-proposal-filled');
  const submit = page.locator('.js-complete-task');
  await submit.waitFor({ state: 'visible', timeout: 30000 });
  await submit.click();
  await waitForTaskList(page);
  await page.locator('.task-card').filter({ hasText: customerName }).first().waitFor({ state: 'detached', timeout: 120000 }).catch(() => {});
  await saveShot(page, outDir, 'lyn-after-submit');
}

async function approveProposal(page, outDir, customerName) {
  await openCleanupTask(page, outDir, customerName, 'mark');
  await page.locator('#cleanupDecision').selectOption({ label: 'Approve' });
  await saveShot(page, outDir, 'mark-approve-ready');
  const submit = page.locator('.js-complete-task');
  await submit.waitFor({ state: 'visible', timeout: 30000 });
  await submit.click();
  await waitForTaskList(page);
  await page.locator('.task-card').filter({ hasText: customerName }).first().waitFor({ state: 'detached', timeout: 120000 }).catch(() => {});
  await saveShot(page, outDir, 'mark-after-approve');
}

async function main() {
  const webAppUrl = requiredEnv('WEB_APP_URL');
  const customerName = requiredEnv('CUSTOMER_NAME');
  const lynEmail = requiredEnv('LYN_EMAIL');
  const lynPassword = requiredEnv('LYN_PASSWORD');
  const markEmail = requiredEnv('MARK_EMAIL');
  const markPassword = requiredEnv('MARK_PASSWORD');
  const sourceUserDataDir = requiredEnv('CHROME_USER_DATA_DIR');
  const profileName = optionalEnv('CHROME_PROFILE_NAME', 'Profile 1');
  const headless = optionalBoolEnv('PLAYWRIGHT_HEADLESS', true);
  const outDir = optionalEnv('OUTPUT_DIR', path.join(process.cwd(), 'artifacts', 'cleanup_e2e'));
  ensureDir(outDir);
  const userDataDir = await copyChromeProfile(sourceUserDataDir, profileName);
  const chromeLaunch = await launchChromeForCdp(userDataDir, profileName, headless);
  const browser = await chromium.connectOverCDP(`http://127.0.0.1:${chromeLaunch.port}`);
  const context = browser.contexts()[0];
  if (!context) throw new Error('Chrome launched for CDP, but Playwright did not find a browser context.');

  try {
    const page = context.pages()[0] || await context.newPage();
    try {
      await page.setViewportSize({ width: 1440, height: 1100 });
    } catch (_) {}
    await gotoApp(page, webAppUrl, outDir, 'initial');
    await login(page, outDir, lynEmail, lynPassword, 'lyn');
    await submitProposal(page, outDir, customerName);
    await page.locator('#logoutButton').click();
    await page.locator('#loginScreen').waitFor({ state: 'visible', timeout: 30000 });
    await saveShot(page, outDir, 'after-lyn-logout');
    await login(page, outDir, markEmail, markPassword, 'mark');
    await approveProposal(page, outDir, customerName);

    const result = {
      ok: true,
      webAppUrl,
      customerName,
      screenshotsDir: outDir
    };
    fs.writeFileSync(path.join(outDir, 'result.json'), JSON.stringify(result, null, 2));
    console.log(JSON.stringify(result, null, 2));
  } finally {
    await safeClose(browser);
    await stopChromeLaunch(chromeLaunch);
  }
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : String(err));
  process.exit(1);
});
