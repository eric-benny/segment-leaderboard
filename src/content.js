/**
 * Strava Segment Leaderboard — Content Script
 *
 * Runs automatically on https://www.strava.com/segments/* pages.
 * Responds to messages from the popup to fetch leaderboard data using the existing
 * Strava session (no separate auth needed) and downloads a self-contained HTML file.
 */
(function () {
  'use strict';

  // ── Configuration ────────────────────────────────────────────────────────────
  const CONFIG = {
    PAGE_DELAY_MS: 1500,           // Delay between page requests
    GENDER_DELAY_MS: 2000,         // Delay between gender fetches
    ATHLETES_PER_PAGE: 25,         // Strava's pagination size
    TOP_ATHLETES: 10,              // Number to display in HTML
    URL_REVOKE_DELAY_MS: 5000,     // Cleanup delay for blob URLs
    INITIAL_PAGE_ESTIMATE: 5,      // Better starting estimate for progress
    DEBUG_LOGGING: true            // Enable/disable debug logs
  };

  // ── Logging utility ──────────────────────────────────────────────────────────
  const logger = {
    debug: (...args) => {
      if (CONFIG.DEBUG_LOGGING) {
        console.log('[Strava Leaderboard]', ...args);
      }
    },
    error: (...args) => {
      console.error('[Strava Leaderboard]', ...args);
    },
    warn: (...args) => {
      console.warn('[Strava Leaderboard]', ...args);
    }
  };

  // ── Detect segment page ──────────────────────────────────────────────────────
  const pathMatch = window.location.pathname.match(/\/segments\/(\d+)(?:[/?#]|$)/);
  if (!pathMatch) return;

  const PAGE_SEGMENT_ID = parseInt(pathMatch[1]);
  logger.debug('Detected segment page:', PAGE_SEGMENT_ID);


  // ── Progress tracker ─────────────────────────────────────────────────────────
  /**
   * Tracks progress across 3 gender fetches + 1 build step.
   * Page totals start at 1 (optimistic estimate) and are updated after the first
   * fetch per gender reveals how many pages actually exist.
   * The bar never moves backwards — pct is monotonically non-decreasing.
   */
  class ProgressTracker {
    constructor(onUpdate) {
      this.onUpdate = onUpdate;
      this.phases = {
        overall: { done: 0, total: CONFIG.INITIAL_PAGE_ESTIMATE },
        men:     { done: 0, total: CONFIG.INITIAL_PAGE_ESTIMATE },
        women:   { done: 0, total: CONFIG.INITIAL_PAGE_ESTIMATE },
      };
      this.buildDone = false;
      this._floorPct = 0; // never go below this
      this._label = 'Starting…';
    }

    /** Update the label text and immediately fire the callback so it renders. */
    setLabel(label) {
      this._label = label;
      this._notify();
    }

    /** Call once after the first page fetch reveals the real total for a phase. */
    setTotal(phase, total) {
      this.phases[phase].total = Math.max(total, 1);
      this._notify();
    }

    pageComplete(phase) {
      this.phases[phase].done = Math.min(
        this.phases[phase].done + 1,
        this.phases[phase].total
      );
      this._notify();
    }

    buildComplete() {
      this.buildDone = true;
      this._label = 'Building file…';
      this._notify();
    }

    get pct() {
      const totalWork = Object.values(this.phases).reduce((s, p) => s + p.total, 0) + 1;
      const doneWork  = Object.values(this.phases).reduce((s, p) => s + p.done,  0) + (this.buildDone ? 1 : 0);
      const raw = Math.round((doneWork / totalWork) * 100);
      this._floorPct = Math.max(this._floorPct, raw);
      return this._floorPct;
    }

    _notify() { this.onUpdate(this.pct, this._label); }
  }


  // ── Messages from popup ──────────────────────────────────────────────────────
  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.type === 'GET_SEGMENT_ID') {
      sendResponse({ segmentId: PAGE_SEGMENT_ID, segmentName: getSegmentName() });
      return;
    }
    if (msg.type === 'GENERATE') {
      generateLeaderboard(msg.segmentId || PAGE_SEGMENT_ID, msg.dateRange || 'this_month')
        .then(() => sendResponse({ ok: true }))
        .catch(err => sendResponse({ ok: false, error: err.message }));
      return true;
    }
  });

  // ── Helpers ──────────────────────────────────────────────────────────────────
  function getSegmentName() {
    const el =
      document.querySelector('span[data-full-name]') ||
      document.querySelector('#js-full-name') ||
      document.querySelector('h1');
    return el ? (el.getAttribute('data-full-name') || el.textContent).trim() : `Segment ${PAGE_SEGMENT_ID}`;
  }

  // ── Core leaderboard generation ──────────────────────────────────────────────
  async function generateLeaderboard(segmentId, dateRange) {

    // Phase → ProgressTracker key
    const PHASE = { all: 'overall', M: 'men', F: 'women' };
    const LABELS = { all: 'overall', M: "men's", F: "women's" };

    // Guaranteed paint before the next network call — prevents the browser from
    // batching label updates together and skipping intermediate states.
    const nextFrame = () => new Promise(r => requestAnimationFrame(r));

    // Pause between requests to avoid Strava / CloudFront rate limiting (429s).
    const sleep = ms => new Promise(r => setTimeout(r, ms));

    const tracker = new ProgressTracker((pct, label) => {
      // Notify the popup if it's open
      try { chrome.runtime.sendMessage({ type: 'PROGRESS', pct, label }); } catch (_) {}
    });

    try {
      const segmentName = getSegmentName();
      const now = new Date();
      const monthYear = now.toLocaleString('en-US', { month: 'long', year: 'numeric' });
      const safeName  = segmentName.replace(/[^a-z0-9]/gi, '_').replace(/_+/g, '_');
      const monthTag  = now.toLocaleString('en-US', { month: 'long' }) + '_' + now.getFullYear();

      // Parse one page of the partial leaderboard HTML Strava returns
      function parseHTML(html, pageNum) {
        const doc = (new DOMParser()).parseFromString(html, 'text/html');
        const table = doc.querySelector('table.table-leaderboard');
        if (!table) return { rows: [], totalPages: pageNum };
        let totalPages = pageNum;
        doc.querySelectorAll('.pagination a').forEach(a => {
          const n = parseInt(a.textContent.trim());
          if (!isNaN(n) && n > totalPages) totalPages = n;
        });
        const rows = Array.from(table.querySelectorAll('tbody tr'))
          .filter(row => row.children.length >= 6) // Skip rows without enough columns (empty state messages)
          .map((row, i) => {
            const c = row.children;
            const rt = c[0].textContent.trim();
            return {
              Rank: (rt && !isNaN(parseInt(rt))) ? parseInt(rt) : (pageNum - 1) * CONFIG.ATHLETES_PER_PAGE + i + 1,
              Name: (c[1].querySelector('a') || c[1]).textContent.trim(),
              Date: c[2].textContent.trim(),
              Pace: c[3].textContent.trim().replace(/\s+/g, ' '),
              HR:   c[4].textContent.trim().replace(/\s+/g, ' '),
              Time: c[5].textContent.trim().replace(/\s+/g, ' '),
            };
          });
        return { rows, totalPages };
      }

      // Fetch all pages for one gender, updating the progress bar as we go
      async function fetchGender(gender) {
        const phase = PHASE[gender];
        const label = LABELS[gender];
        const url = p =>
          `/segments/${segmentId}/leaderboard?date_range=${dateRange}&gender=${gender}&page=${p}&per_page=25&partial=true`;
        const opts = { headers: { 'X-Requested-With': 'XMLHttpRequest' } };

        // Page 1 — reveals total page count.
        // setLabel fires the callback → DOM updates → nextFrame waits for the
        // browser to actually paint before the network call starts.
        tracker.setLabel(`Fetching ${label}…`);
        await nextFrame();

        let res1;
        try {
          res1 = await fetch(url(1), opts);
        } catch (err) {
          logger.error('Network error on page 1:', err);
          throw new Error('Network connection failed — check your internet connection');
        }

        if (res1.status === 401) {
          throw new Error('Please log into Strava in this browser tab');
        } else if (res1.status === 403) {
          throw new Error('This segment is private or unavailable to your account');
        } else if (res1.status === 404) {
          throw new Error('Segment not found — check the segment ID');
        } else if (res1.status === 429) {
          throw new Error('Rate limited by Strava — wait a minute and try again');
        } else if (!res1.ok) {
          throw new Error(`Strava returned HTTP ${res1.status} — please try again`);
        }

        const { rows: r1, totalPages } = parseHTML(await res1.text(), 1);
        logger.debug(`Gender ${gender}: Found ${totalPages} pages`);
        tracker.setTotal(phase, totalPages);
        tracker.pageComplete(phase);

        const all = [...r1];

        // Remaining pages — sleep between each to stay under CloudFront rate limit
        for (let p = 2; p <= totalPages; p++) {
          tracker.setLabel(`Fetching ${label} · page ${p} of ${totalPages}`);
          await nextFrame();
          await sleep(CONFIG.PAGE_DELAY_MS);

          let res;
          try {
            res = await fetch(url(p), opts);
          } catch (err) {
            logger.error(`Network error on page ${p}:`, err);
            throw new Error('Network connection failed — check your internet connection');
          }

          if (res.status === 401) {
            throw new Error('Please log into Strava in this browser tab');
          } else if (res.status === 403) {
            throw new Error('This segment is private or unavailable to your account');
          } else if (res.status === 404) {
            throw new Error('Segment not found — check the segment ID');
          } else if (res.status === 429) {
            throw new Error('Rate limited by Strava — wait a minute and try again');
          } else if (!res.ok) {
            throw new Error(`Strava returned HTTP ${res.status} — please try again`);
          }

          const { rows } = parseHTML(await res.text(), p);
          all.push(...rows);
          tracker.pageComplete(phase);
        }

        return all;
      }

      const overall = await fetchGender('all');
      await sleep(CONFIG.GENDER_DELAY_MS);
      const men     = await fetchGender('M');
      await sleep(CONFIG.GENDER_DELAY_MS);
      const women   = await fetchGender('F');

      if (overall.length === 0) {
        throw new Error('No data returned — are you logged into Strava?');
      }

      // Log warnings for empty gender-specific boards
      if (men.length === 0) {
        logger.warn('No male athletes found for this segment and time period');
      }
      if (women.length === 0) {
        logger.warn('No female athletes found for this segment and time period');
      }

      tracker.buildComplete(); // sets label to 'Building file…' and fires callback
      await nextFrame();

      const embeddedData = JSON.stringify({ overall, men, women, segmentName, monthYear });
      const htmlFile = buildHTML(segmentName, monthYear, now, embeddedData, overall, men, women);

      tracker.setLabel('Downloading…');

      // Trigger download
      const blob  = new Blob([htmlFile], { type: 'text/html' });
      const dlUrl = URL.createObjectURL(blob);
      const fname = `${safeName}_${monthTag}_leaderboard.html`;
      const a = Object.assign(document.createElement('a'), { href: dlUrl, download: fname });
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(dlUrl), CONFIG.URL_REVOKE_DELAY_MS);

    } catch (err) {
      logger.error('Leaderboard generation failed:', err);
      throw err; // propagate so the popup message handler gets { ok: false }
    }
  }

  // ── HTML escaping for XSS prevention ─────────────────────────────────────────
  /**
   * Escapes HTML special characters to prevent XSS attacks
   * @param {string} unsafe - Potentially unsafe user input
   * @returns {string} HTML-safe string
   */
  function escapeHtml(unsafe) {
    if (typeof unsafe !== 'string') return '';
    return unsafe
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // ── HTML builder — Run the Whites / White Mountain Throwdown branded ─────────
  function buildHTML(segmentName, monthYear, now, embeddedData, overall, men, women) {
    const medals = ['🥇', '🥈', '🥉'];

    function top10Rows(entries) {
      return entries.slice(0, CONFIG.TOP_ATHLETES).map((e, i) =>
        `<tr class="${i === 0 ? 'gold' : i < 3 ? 'top3' : i % 2 === 0 ? 'even' : ''}">` +
        `<td class="rank">${medals[i] || e.Rank}</td>` +
        `<td class="name" title="${escapeHtml(e.Name)}">${escapeHtml(e.Name)}</td>` +
        `<td class="data">${escapeHtml(e.Time)}</td>` +
        `<td class="pace">${escapeHtml(e.Pace)}</td></tr>`
      ).join('');
    }

    function tableBlock(title, icon, entries) {
      const athleteCount = entries.length > 0 ? `${entries.length} athletes` : 'No athletes';
      const tableBody = entries.length > 0
        ? top10Rows(entries)
        : '<tr><td colspan="4" style="text-align:center;color:#999;padding:20px;">No athletes in this category</td></tr>';

      return `<div class="col">` +
        `<div class="col-head"><span class="icon">${icon}</span><span class="title">${title}</span><span class="count">${athleteCount}</span></div>` +
        `<table><thead><tr><th>#</th><th>Athlete</th><th>Time</th><th>Pace</th></tr></thead>` +
        `<tbody>${tableBody}</tbody></table></div>`;
    }

    function individualCard(title, icon, entries, segmentName, monthYear) {
      const athleteCount = entries.length > 0 ? `${entries.length} athletes` : 'No athletes';
      const tableBody = entries.length > 0
        ? top10Rows(entries)
        : '<tr><td colspan="4" style="text-align:center;color:#999;padding:20px;">No athletes in this category</td></tr>';

      return `<div class="individual-card">` +
        `<div class="header">` +
        `${mountainSvg}` +
        `${rtwLogoImg}` +
        `<div class="header-text">` +
        `<div class="event-tag">White Mountain Throwdown</div>` +
        `<h1>${escapeHtml(segmentName)}</h1>` +
        `</div>` +
        `</div>` +
        `<div class="col">` +
        `<div class="col-head"><span class="icon">${icon}</span><span class="title">${title}</span><span class="count">${athleteCount}</span></div>` +
        `<table><thead><tr><th>#</th><th>Athlete</th><th>Time</th><th>Pace</th></tr></thead>` +
        `<tbody>${tableBody}</tbody></table></div>` +
        `<div class="footer">` +
        `<div class="rtw-footer"><span class="rtw-wordmark">Run the Whites</span></div>` +
        `<img class="strava-logo" src="data:image/png;base64,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" alt="Powered by Strava">` +
        `</div>` +
        `</div>`;
    }

    const generatedDate = now.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

    const mountainSvg = `<svg class="mtn-bg" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 120" preserveAspectRatio="xMidYMid slice">
      <path d="M0,120 L0,82 L55,36 L100,60 L158,16 L218,54 L256,33 L308,70 L350,46 L392,68 L438,40 L488,18 L538,54 L576,32 L622,66 L660,44 L700,62 L744,33 L792,58 L832,38 L876,60 L916,42 L960,52 L960,120 Z" fill="white" fill-opacity="0.06"/>
      <path d="M0,120 L28,98 L68,108 L118,86 L162,100 L208,88 L252,104 L298,90 L342,106 L388,92 L432,108 L478,93 L522,106 L566,91 L612,106 L658,93 L702,108 L748,92 L792,106 L838,94 L882,108 L924,96 L960,104 L960,120 Z" fill="white" fill-opacity="0.04"/>
    </svg>`;
    const rtwLogoImg = `<img class="rtw-logo" src="data:image/png;base64,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" alt="Run the Whites">`;

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${escapeHtml(segmentName)} — ${escapeHtml(monthYear)} · White Mountain Throwdown</title>
  <style>
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;background:#E2E8EF;display:flex;flex-direction:column;align-items:center;padding:32px 16px;min-height:100vh}
    .container{width:100%;max-width:960px;display:flex;flex-direction:column;gap:24px}
    .card{background:#fff;border-radius:12px;overflow:hidden;width:100%;box-shadow:0 10px 40px rgba(11,30,63,.22)}

    /* ── Header ──────────────────────────────────── */
    .header{background:#1A3A5C;color:#fff;padding:26px 28px 22px;position:relative;overflow:hidden;display:flex;align-items:center;gap:18px}
    .mtn-bg{position:absolute;bottom:0;left:0;width:100%;height:100%;pointer-events:none}

    /* RTW logo image */
    .rtw-logo{position:relative;z-index:1;height:64px;width:auto;flex-shrink:0;display:block}

    .header-text{position:relative;z-index:1;flex:1;min-width:0}
    .event-tag{font-size:9.5px;font-weight:800;letter-spacing:2.8px;color:#1B8CCE;text-transform:uppercase;margin-bottom:7px}
    .header h1{font-size:21px;font-weight:800;letter-spacing:-.3px;line-height:1.2;color:#fff;white-space:nowrap;text-wrap:auto}
    .header-sub{margin-top:5px;font-size:11.5px;color:rgba(255,255,255,.55);letter-spacing:.2px}

    /* ── Export Button Container ────────────────── */
    .export-container{padding:16px 28px;background:#F8FBFD;border-bottom:2px solid #E2EBF3;display:flex;justify-content:flex-end}
    .export-btn{background:rgba(27,140,206,.2);border:1.5px solid rgba(27,140,206,.5);color:#1A3A5C;padding:8px 14px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;white-space:nowrap;letter-spacing:.3px;transition:all .15s}
    .export-btn:hover{background:rgba(27,140,206,.35);color:#1A3A5C}
    .export-btn:disabled{opacity:.55;cursor:wait}

    /* ── Grid ────────────────────────────────────── */
    .grid{display:grid;grid-template-columns:repeat(3,1fr);border-top:3px solid #1B8CCE}
    .col{border-right:1px solid #E2EBF3}
    .col:last-child{border-right:none}
    .col-head{display:flex;align-items:center;gap:7px;padding:10px 14px 9px;border-bottom:2px solid #1B8CCE;background:#F0F7FC}
    .icon{font-size:15px}
    .title{font-size:11px;font-weight:800;color:#1A3A5C;flex:1;text-transform:uppercase;letter-spacing:.7px}
    .count{font-size:10px;color:#7A9AB5;font-weight:600}

    table{width:100%;border-collapse:collapse;font-size:12px}
    thead tr{background:#F7FBFF}
    th{padding:5px 10px;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;color:#8AAAC0;text-align:left;border-bottom:1px solid #E2EBF3}
    td{padding:7px 10px;border-bottom:1px solid #EDF2F8;color:#1A2332}
    tr.even td{background:#F8FBFD}
    tr.top3 td{font-weight:700;color:#1A3A5C;background:#EBF5FD}
    tr.gold td{font-weight:800;color:#1A3A5C;background:#DFF0FA}
    tr:last-child td{border-bottom:none}
    .rank{width:34px;text-align:center;color:#8AAAC0;font-size:13px}
    tr.top3 .rank,tr.gold .rank{color:#1B8CCE}
    .name{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:0;width:100%}
    .data{font-variant-numeric:tabular-nums;white-space:nowrap;font-weight:500}
    .pace{color:#8AAAC0;font-size:11px;white-space:nowrap}

    /* ── Footer ──────────────────────────────────── */
    .footer{padding:10px 24px;background:#1A3A5C;display:flex;align-items:center;justify-content:space-between}
    .rtw-footer{display:flex;align-items:center;gap:8px}
    .rtw-wordmark{font-size:10px;font-weight:900;letter-spacing:2.2px;color:#1B8CCE;text-transform:uppercase}
    .footer-date{font-size:10px;color:rgba(255,255,255,.35);letter-spacing:.2px}
    .strava-logo{height:14px;width:auto;display:block}

    /* ── Individual Shareable Cards ─────────────── */
    .individual-cards{display:flex;flex-direction:column;align-items:center}
    .individual-card{background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 10px 40px rgba(11,30,63,.22);width:350px;margin-bottom:20px}
    .individual-header{background:#1A3A5C;color:#fff;padding:20px 24px;position:relative;overflow:hidden}
    .individual-title{font-size:18px;font-weight:800;letter-spacing:-.2px;color:#fff;display:flex;align-items:center;gap:10px;margin-bottom:6px}
    .individual-subtitle{font-size:11px;color:rgba(255,255,255,.55);letter-spacing:.2px}
    .individual-card table{width:100%}
    .individual-card .col-head{border-radius:0}
    .individual-card .col{border:none}
  </style>
</head>
<body>
<div class="container">
<div class="card">
  <div class="header">
    ${mountainSvg}
    ${rtwLogoImg}
    <div class="header-text">
      <div class="event-tag">White Mountain Throwdown</div>
      <h1>${escapeHtml(segmentName)}</h1>
      <div class="header-sub">${escapeHtml(monthYear)} Monthly Challenge &mdash; Top 10 Leaderboard</div>
    </div>
  </div>
  <div class="grid">
    ${tableBlock('Overall', '🏃', overall)}
    ${tableBlock('Men', '👨', men)}
    ${tableBlock('Women', '👩', women)}
  </div>
  <div class="footer">
    <div class="rtw-footer">
      <span class="rtw-wordmark">Run the Whites</span>
    </div>
    <span class="footer-date">Generated ${generatedDate}</span>
    <img class="strava-logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAW0AAAAlCAYAAACJfbiXAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAA/CSURBVHgB7V1ddts2Fv5I2zlnnqqsYJjZQJ0VhFlBlXTeK3cBE7cLqJUsYJosYKbyvDd2VhBmBXE2MEFWUPdpzrEtYe4lAQuiiT+SkiOb3zk8okgQBAniw4eLCyDBgAH3CHKMjH+vgEwC53u00d/z5LT8HXCL+B/lzW6Kyd5bTHGHMH9e/hztvMVL9IAEAwbcURBBjy5TjHeAJ0TQOR3KXOETiTOZQFDYT1ig2DtFAQ8un+F1muAb3AIorcXu7zgu0/F35OkCP6ADFhJ/EiOUlRe9A5EsIHb59xQCG8DFM8zoXX63s8AjVyVKJPgbIkDv5WXoM8y/L39e0AvYRyAoH95RPpw2nbuoCPsgBf5Fz/W3Pt7lLgYMuGMoFVuCw3mCH6iwjGTgdVT4uKDuk5IZI8XR1XMiLKBwFfqESEZ6KoN1QSmukrTlomw5TNABSVKLm17enH70e1gscBxSkbUB5xkRdlnpXKU4pJ+pLSxXLpTWFwjEIuU6CD/6wlEacDHHo70UrxEOuTO3K+i0epW/0JZQOn4JSYcPKQYMuEMgpfSCCt1HVahH6IaSCOcpPhNxnWjTyj1E+R6SFO/pPXy++r6bom/CTrIkaWK5F9xKsoXdlWXYYHMWpz0k7ygNCW1HiACldWar0FllL6qKNNPpuByXLb5OGEh7wJ0BEcpvUpYqqStZN2HM5E1N819xv5ER+8z6fA+mylYYKbXdCDadEAG+QTi0ynWlAXN5Ix0+SGqFvbLGR2aeFCv3TUhMONMRgoG0B9wJKBKZYM0gsji8eob3uOco30OkbdkGU2VreNX2oqyce1PbfavsB2m5cQWQ1S7Ku6rtgbQHbD2oGTphEsGGQDbVnzCAMemquBtUtsbG1PY6VPYFqWxKX1N8ndX2QNoDth70EUcppC6ggvjywSnOMKAEV5ZdlGOTytbYlNremMpeXtxJbQ+kPWCrQZ1ipGs200HILoF3zYe4DyRJu0rTobI1RqSOrXH3obbXqLJd76ST2h5Ie8BWY7EoSTsWwtiCr0klnmHATZBybONZs5f4beKs5F1xd1XbbVQ2xfHKprKD4+ugtgfSHrDVoALybURwwQM3dt+ubAkR/+MFcEClsbBeSc34TQ0y2UbwIKao8ExYSRhprUttt1TZn6mimDWdiIyvtdoeBtcM2GpEDmw5ayJeZaPmbaaHUidY2iSlxBsyixw3RUiVwFMEYJ6WHidZSFi637tdGdSx2nroPVdWTccvxtgnKbdPao47GINdJ8l0lCECMSYVpY6tA5xYbdP7DfbLV/G9ukzwBYhW2f9xqeyo+JTajh2wNJD2gG1HjE/2eP4cR2SPPLYVvL9Ux6dE3jPuJCPV9EQN5mhEqPq+eo5gpCn+SH6/HVWvKzAi7zNKx8fQ6yhs8FB+9vYB4kwD88qU0lhBstqmOF+qiiYErLaP5mSX3kv7VdmR8Wm1XURcM5hHBmw9RExgUkrTcpDMM3zkOSx4BGWTbZHJ+8EJJmxCuY+TSZXkLcPJhIfRh4Zt5e3jsQE/eFvatgUCwWqbCPafCIekNLzsbMuug57rahxnWhqU9oDthqRClMR3gul5Rrj0JqlSwkRSdPxD6GRRdx5JVAffnyHhlMrO0ALKpFLYzlO/BKvtmAE/MWQp9ORcdbRU2RoJqsrjNPSCQWkP2GqUJNsXSPWQUfJIz7HBSvwezzfCCJ7pDoEqt5NPvV9tz4C1mJVKlW072VplL5HFzOcykPaArUasy1cErieLuo/kzbZ/RCji3R3/0P4uKlvD14FJZpoD9A+vyo70QKkjoWohmPQH88iArUaLTqhoKPIe832U7XTr0aRYZYqMFOM3RHxjGdlRiCt88gXpZeSqx+OCj189o3NJZPrt8Kps9DMi99Hlc0xDBm9p0s4dYVjFxA7bzWlj/1luXn1R1xdwK6JMba775er3zBIX329kiSNDWC0vcLOJNYK7qWhLT1/3Z+SR4dsid5xz5U2G5fMVcEPnU0hYL5hI6YMfJesdzj7iioEU6Dd9rUBym0jSm8qY2UdKRC+N4hrSrdGHyr6+n8e2Tc/wMumPtK0q+5It4pJMamknlX0Neo//oBbda1/HN5N2RpuvaSNom9Hm+1gzVB0BuSUOdoK3KZVDoPS15AQ/bDg/wbKT4UClpw5+jpG6T93PdQoEvdymuPfhf0dMZvx+bB0KE4SRCscxbTjuur9A9V5jBhnY0PZbyIxrf4I9nzmcdiWbAf10+LFCIeLGmom79D6h3v5Pu6fhHUd3HbxIhDdMn/myObXtVNlJcr3AQV946FsAghFq085URO89Ybgw5up/OVgBy0LJ57kJa8s8reBGaK6Rc2P/ScP5EZbqzWtfWwOY2E+wwdnmDGSoSPIzelIznntNcTMfCyzzms+NLNdP1a8A+lWsJXEv8Gghm5VRb0jv/Zza1+AJtDapsjW8tm3Zy7flVNl0jxz9KfoSSm2PXGHqpH1QXbeyPQauE57Drla1yhWonOAfq/h4/xGWCnSKZiVuNrubhiabRN3kqmOaL77AjgI3n9HcZnDjaS08twr4OYU67yIsqHCu+0/hRj2PzPzJgH7mOEZFqE3fgs6nKW4+J6ftXB1vqrwyLL8fvl6gZ2j/aoO8BfpH1scKJNsOJuwQG+xaZmH0eJKUKlx2asW5VXZK9UaKf6N/PHRNScsIUdpcSCdYqqhJQxg+lqn9A9xs8grcJLam+2hbzqPaudyIX5NCXguzX4trU+D0zIDrXmtO2xibg84f/YHlQD82Nsu9zEmT6s8psDTRsKkrq51/b4Rbqxo2B8dQx9pTJpiOhXgVyUbz+GvD+YJMYCGEzYOXsKbWX5q4Wzwd1bZVZdeXEesbPrUd4/Kne4ezhnN6kc13sNsomdw0seVo7tgT6reutHUBKbAk5HqheWKEuQ0Uxn6GzWMKd8XaF4Tn/GsVhj86s1BNsHwva5ktj+cN4dXR6x88qy4mmN0TPOU5N5jEqXl92FGJZ7iPoIpvZ4HHIV40nB9SRpkLBSIGmfAAKZd/cwe1LekbsS7Am1bjsaJs2ZEmO6fabuOnXe/ZND0rTjzXmgo4bzj/yXJOE/IMVcVgHtPIanHcJs5xO9DvJsd61klkTIz9ouG8WTmPscxL3bqaYU0toQcJTnhB33mKj77CvPM73mgl3sYOTmooeK6NOwWJN6HzrTgXAmiMGseUF3GrAkm32qdK4xViIe0jYtuobPpWfturVjsK5gWX2o4h7e/U75fa8VBbMoMTLdR+k91aF+YMS9LJjHt8wNLmvI/VF6fDCLiRoXrpTVuO9jBNPi5SGjnu37XJXRj7GbqB82dS21hdafXMZhABezoKtX+ktgxV/ndpslrBy16poemAWnw2dOVwbUrBeuzfXyuE2uIERloONPIKglJlx7X4BE/GxHkRWYE6RxNSJR6rtqWL6Fuo7HLBBFLOf0ROIWtV2/XBNRluEhcfM2tMl5oW8EPATihF7b5nWPVGEUY4Ps4k9xqrafapuAz2zroCfvPKD2huCehjp544Ro77C6CTK9l57T5dMIa9EhGAt3nMapu9WXIsK1QX0bcGj96Tto7PirynpFwKavK+I5UkmpYLU51aGQIROtfG1wpuYfAvk+teNZtf6PcymiclBzx1BWKVLSNVtlbwc1nOrhjeLyNLk9i7Jv/mS8lR48e9BP8NjO3Up7LTOJVd+rDLqiS9iZlCNrH4bddJW6siG1glrbMDSRj7OSoC1pn3wTj3QZ1n9c/kkRnnCrjhGiASYlqZOM4VQFDzrrAcF/h6cIbmAUo5lj7ZXHCF5Xo+zt+L9qbh/1P0DFZZUnrjrVRfWhY4PU2q4MmmqGCMSKHb3EytkMndmFCKlW30iNIEOV1zaLNrK5U9RTiEOeVpmaZnOI4g7pHNv5niAsVVqveA+CTZ63+2nUwR7Zd9vSwZES8un+MP2n2ThHvTPFQLNvxcS4cXTHIFqs6jqeW8RkgNosP8aYlLqP1M/ebq11Sghfrdx6pN3aeydZinlu0w8PoCqyM8hRGH8FwvHPc/QDdk6A9sHz+obZxGVmgCYe6FZqHuY+DPCpgcWGWhHTImH2VSyRCJvcXdGVxTkm9khx278dnmY3Et1tsEIrHTup18LqPjsC4CTHHJeYAnSeIY2dnSlr0SH1VMvHE5CDZJyYbl1kL8tNkPmQur7SMVxn4ON0yCFZYwWlGzyWFshC2MMLx/bsSn7eNfsH6wktYkq4kow9ehkjNjP6QCawOB5XPve8KeW/Y7QzXr9diAjSIJGLa9bSAzApf9mDwazRvWeAxYrPcG0sXNCr2FbXtkswGz2qaKxBefdbHeMo0tbdnmAVbbLWzbNxYjbuM9UodW4ozvPGHHxr6tEjA7I01XvzqOjTj3HeHWCVZ5+kPvfwBBPHRh4XfYK0lasHHC1NhL42ah6xMhw7a3DSVJxnYSKzOJeaiFyrZWgH2q7T1S2y4PjnWrbI2WantlMeI+SJthuprlljD8MjWxFbAr0zMjvK4EmgheHzMN++tSlzbwi9e15gTdvE+64si4f1uTQSh8raW1gzrSDiQ2T54hw7a3FV3NJC1VtjUP+1TbHpVrVdncgUimsIddVXZgOmxIqFVzPfqyL9JmkhBqn3uW6xmXASsLm7o660zi1WT8wRLu3HGtDbrTybbFqsc2ajvrcP96+nNUbpBTdV6gn87ipvfEZM0dVjp/C9wi9Fwj2FDlETpse5uxkJF+0oaZpE+VfR1GeqeVqMdpVds2letKB5kmeIsa1el7rjZq2xy23xdpM3QnHL+wGSp3Lybqj2pfqzOuWV3kanZGMgo0P1zdCyTUJLCv0mPbxoiDqbZz+NV25rn/oef6X2vh+R1rEhVA2OrgAXjRkLaPRvoEcPtmAlZj7L62WJ0moW8IHkV51wmbwe6QsoWZhP3k+1TZGi1GNcaqbasq5gUOLkgUyLhpIZy2cUc6fNCLAPdK2gKrHZYZVn10mdxsHih1mMp65gj3ztjfRCekDVO451VZNwSqgvYY61edYoP3CgYvNaXnGelxoqhzJjAetr13j9aM5MqJFG6UqVFGzm6ZRHTmxs4hEqO2Xel4kMaP6gx9ri5qO1F/M/V7jn46sDJU9mj2PGHfRPZ/LsIvX5lm1ZemLDCcGacLtngyz3kzfuE53+X+oeHbInOca3OvrMO1ncAfeLqDb3kKTSKhzBgxacM5heMOuQ/EQqd9EjXPiUJKNGjoO/uA2yYrssbPzxoxGf/OW7d76cUY+zvp9ZxCvYNVdkzfAK/qQhz619DwiwWObfl3WfnpT4n8fqEK+W+2dOhwMfeNea6LKv7JTvNU042YE5cmGDDgHoEV2CVVJIlRge6qSvaudjAOWIUanTi6SnC0dxJtw791/B/hqRR5x4BJ5wAAAABJRU5ErkJggg==" alt="Powered by Strava">
  </div>
</div>
<button class="export-btn" onclick="exportExcel(this)">&#11015; Export Excel</button>
<div class="individual-cards">
  ${individualCard('Overall', '🏃', overall, segmentName, monthYear)}
  ${individualCard('Men', '👨', men, segmentName, monthYear)}
  ${individualCard('Women', '👩', women, segmentName, monthYear)}
</div>
</div>
<script id="lb-data" type="application/json">${embeddedData}<\/script>
<script>
async function exportExcel(btn){
  btn.disabled=true;btn.textContent='⏳ Building…';
  try{
    if(!window.XLSX){await new Promise((res,rej)=>{const s=document.createElement('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';s.onload=res;s.onerror=rej;document.head.appendChild(s);});}
    const d=JSON.parse(document.getElementById('lb-data').textContent);
    function makeSheet(rows){const ws=XLSX.utils.json_to_sheet(rows,{header:['Rank','Name','Date','Pace','HR','Time']});ws['!cols']=[{wch:7},{wch:28},{wch:14},{wch:12},{wch:12},{wch:10}];return ws;}
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,makeSheet(d.overall),'Overall');
    XLSX.utils.book_append_sheet(wb,makeSheet(d.men),'Men');
    XLSX.utils.book_append_sheet(wb,makeSheet(d.women),'Women');
    const bytes=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const safe=d.segmentName.replace(/[^a-z0-9]/gi,'_').replace(/_+/g,'_');
    const month=new Date().toLocaleString('en-US',{month:'long'})+'_'+new Date().getFullYear();
    const url=URL.createObjectURL(blob);
    const a=Object.assign(document.createElement('a'),{href:url,download:safe+'_'+month+'_leaderboard.xlsx'});
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url),5000);
    btn.textContent='✓ Downloaded!';setTimeout(()=>{btn.textContent='⬇ Export Excel';btn.disabled=false;},3000);
  }catch(e){btn.textContent='✗ Error — retry';btn.disabled=false;console.error(e);}
}
<\/script>
</body></html>`;
  }

})();
