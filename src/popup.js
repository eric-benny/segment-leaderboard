/**
 * Strava Segment Leaderboard — Popup Script
 *
 * On open: detects the segment ID from the active tab (if on a Strava segment page).
 * On "Generate": sends a GENERATE message to the content script, which runs the
 * full leaderboard fetch + HTML build in the page context (with Strava session cookies).
 *
 * While generating, the content script sends PROGRESS messages back so this popup
 * can show a live progress bar.
 */

const segmentInput  = document.getElementById('segmentId');
const dateRangeSelect = document.getElementById('dateRange');
const generateBtn   = document.getElementById('generate');
const statusEl      = document.getElementById('status');
const hintEl        = document.getElementById('hint');

// Progress elements
const progressWidget = document.getElementById('progress-widget');
const progBar        = document.getElementById('prog-bar');
const progPct        = document.getElementById('prog-pct');
const progLabel      = document.getElementById('prog-label');

// Timeout constants for message display
const ERROR_DISPLAY_DURATION_MS = 4000;
const SUCCESS_DISPLAY_DURATION_MS = 2500;

let activeTabId     = null;
let isOnSegmentPage = false;

// ── On popup open: detect active tab ─────────────────────────────────────────
chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
  if (!tabs.length) return;
  const tab = tabs[0];
  activeTabId = tab.id;

  const segMatch = tab.url?.match(/strava\.com\/segments\/(\d+)(?:[/?#]|$)/);
  if (segMatch) {
    isOnSegmentPage = true;
    segmentInput.value = segMatch[1];
    hintEl.textContent = '✓ Segment detected from current tab.';
    hintEl.style.color = '#2E7D32';
  } else if (tab.url?.includes('strava.com')) {
    hintEl.textContent = 'Navigate to a Strava segment page, or enter an ID above.';
  }
});

// ── Listen for progress updates forwarded from the content script ────────────
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === 'PROGRESS') {
    setProgress(msg.pct, msg.label || '');
  }
});

// ── Generate button ───────────────────────────────────────────────────────────
generateBtn.addEventListener('click', async () => {
  const rawId = segmentInput.value.trim();
  if (!rawId || isNaN(parseInt(rawId))) {
    showStatus('Please enter a valid segment ID.', 'err');
    return;
  }
  const segmentId = parseInt(rawId);
  const dateRange = dateRangeSelect.value;

  if (isOnSegmentPage && activeTabId) {
    startProgress();
    showStatus('');

    chrome.tabs.sendMessage(activeTabId, { type: 'GENERATE', segmentId, dateRange }, (response) => {
      if (chrome.runtime.lastError) {
        endProgress();
        showStatus('Could not connect to the Strava page. Try refreshing it.', 'err');
        setTimeout(() => showStatus(''), ERROR_DISPLAY_DURATION_MS);
        return;
      }
      if (response?.ok) {
        finishProgress();
        setTimeout(() => { endProgress(); showStatus(''); }, SUCCESS_DISPLAY_DURATION_MS);
      } else {
        endProgress();
        showStatus(response?.error || 'Something went wrong.', 'err');
        setTimeout(() => showStatus(''), ERROR_DISPLAY_DURATION_MS);
      }
    });

  } else {
    // Open the segment page so the content script can access Strava session
    chrome.storage.session.set({ pendingGenerate: { segmentId, dateRange } });
    chrome.tabs.create({ url: `https://www.strava.com/segments/${segmentId}` });
    showStatus('Opened the segment page — use the extension popup to generate.', 'ok');
  }
});

// ── Progress helpers ──────────────────────────────────────────────────────────
function startProgress() {
  generateBtn.disabled = true;
  generateBtn.textContent = 'Generating…';
  progressWidget.style.display = 'flex';
  setProgress(0, 'Starting…');
}

function setProgress(pct, label) {
  progBar.style.width = `${pct}%`;
  progBar.classList.remove('success');
  progPct.textContent = `${pct}%`;
  progPct.style.color = '#FC4C02';
  progLabel.style.color = '';
  if (label) progLabel.textContent = label;
}

function finishProgress() {
  progBar.style.width = '100%';
  progBar.classList.add('success');
  progPct.textContent = '100%';
  progPct.style.color = '#2E7D32';
  progLabel.textContent = '✓ Downloaded!';
  progLabel.style.color = '#2E7D32';
}

function endProgress() {
  progressWidget.style.display = 'none';
  progBar.style.width = '0%';
  progBar.classList.remove('success');
  progLabel.style.color = '';
  progPct.style.color = '';
  generateBtn.disabled = false;
  generateBtn.textContent = 'Generate Leaderboard';
}

// ── Status helpers ────────────────────────────────────────────────────────────
function showStatus(text, type = '') {
  statusEl.textContent = text;
  statusEl.className = type;
}
