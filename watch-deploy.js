/**
 * Happy Drawing (SIMS) - Auto Push & Auto Deploy Watcher
 * Memantau Code.gs, index.html, dan appsscript.json.
 * Setiap kali file disimpan, otomatis menjalankan:
 * 1. clasp push
 * 2. clasp deploy --deploymentId AKfycbyEsKdyMBJzJl2E15bOnkm5A7MYLcBB4xJePka1mz9icDHspmxw6gWu2e8jesl_8QLa
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const DEPLOYMENT_ID = 'AKfycbyEsKdyMBJzJl2E15bOnkm5A7MYLcBB4xJePka1mz9icDHspmxw6gWu2e8jesl_8QLa';
const EXEC_URL = `https://script.google.com/macros/s/${DEPLOYMENT_ID}/exec`;
const WATCHED_FILES = new Set(['Code.gs', 'index.html', 'appsscript.json']);

function getTimestamp() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

console.log('\x1b[36m%s\x1b[0m', '=============================================================');
console.log('\x1b[33m%s\x1b[0m', '  HAPPY DRAWING (SIMS) - AUTO PUSH & AUTO DEPLOY WATCHER');
console.log('\x1b[36m%s\x1b[0m', '=============================================================');
console.log(`\x1b[90mTarget Deployment:\x1b[0m ${DEPLOYMENT_ID}`);
console.log(`\x1b[90mTarget Web App URL:\x1b[0m \x1b[32m${EXEC_URL}\x1b[0m`);
console.log(`\x1b[90mFile yang dipantau:\x1b[0m ${Array.from(WATCHED_FILES).join(', ')}`);
console.log('\x1b[35m%s\x1b[0m', 'Watcher aktif! Tekan Ctrl+C untuk menghentikan.\n');

let isDeploying = false;
let pendingTrigger = false;
let debounceTimer = null;

function triggerDeploy(changedFile) {
  if (isDeploying) {
    pendingTrigger = true;
    return;
  }

  isDeploying = true;
  const time = getTimestamp();

  console.log(`\x1b[33m[${time}] Perubahan terdeteksi pada: ${changedFile}\x1b[0m`);

  try {
    console.log(`[${getTimestamp()}] 1. Menjalankan clasp push...`);
    execSync('clasp push --force', { stdio: 'inherit' });

    console.log(`[${getTimestamp()}] 2. Menjalankan clasp deploy ke ${DEPLOYMENT_ID}...`);
    const desc = `"Auto Deploy ${getTimestamp()}"`;
    execSync(`clasp deploy --deploymentId ${DEPLOYMENT_ID} --description ${desc}`, { stdio: 'inherit' });

    console.log(`\x1b[32m[${getTimestamp()}] SUKSES DEPLOY KE PROD:\x1b[0m \x1b[4;32m${EXEC_URL}\x1b[0m\n`);
  } catch (err) {
    console.error(`\x1b[31m[${getTimestamp()}] Gagal deploy:\x1b[0m`, err.message || err);
  } finally {
    isDeploying = false;
    if (pendingTrigger) {
      pendingTrigger = false;
      triggerDeploy('perubahan tertunda');
    }
  }
}

// Watch directory for file changes
fs.watch('.', (eventType, filename) => {
  if (!filename) return;
  const baseName = path.basename(filename);

  if (WATCHED_FILES.has(baseName)) {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      triggerDeploy(baseName);
    }, 1200); // 1.2 detik debounce untuk meredam multiple event save
  }
});
