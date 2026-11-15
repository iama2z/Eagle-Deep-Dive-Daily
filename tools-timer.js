/* Desktop Tools: persistent timer popup. Pure logic is exported for tests (Node);
   the DOM part runs only in a browser. */
(function (root) {
  'use strict';

  const DOUBLE_WINDOW_MS = 1000;
  const MAX_SECONDS = 99 * 60 * 60;

  // Digit keys map to minutes: 1-9 -> 1-9, 0 -> 10.
  function minutesForKey(key) {
    if (!/^[0-9]$/.test(key)) return null;
    return key === '0' ? 10 : Number(key);
  }

  // Same key again within the window doubles the current duration; otherwise resets to the key's value.
  function nextDuration(prev, key, now) {
    const minutes = minutesForKey(key);
    if (minutes === null) return null;
    const seconds = minutes * 60;
    const quick = prev && prev.key === key && now - prev.at <= DOUBLE_WINDOW_MS;
    return { key, at: now, seconds: Math.min(quick ? prev.seconds * 2 : seconds, MAX_SECONDS) };
  }

  function formatTime(totalSeconds) {
    const s = Math.max(0, Math.ceil(totalSeconds));
    const h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), sec = s % 60;
    const pad = n => String(n).padStart(2, '0');
    return (h ? h + ':' + pad(m) : String(m)) + ':' + pad(sec);
  }

  const api = { DOUBLE_WINDOW_MS, minutesForKey, nextDuration, formatTime };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof document === 'undefined') return;
  root.EddTimer = api;

  const $ = id => document.getElementById(id);
  let el, display, label, startBtn, state = null, tick = null, endsAt = 0, remaining = 0,
    running = false, lastPress = null, lastBeepSecond = null, audioCtx = null, confetti = null;

  function build() {
    const style = document.createElement('style');
    style.textContent = `
#eddTimer{position:fixed;right:1.25rem;bottom:5.5rem;z-index:9000;width:17rem;padding:1rem;border-radius:1.25rem;background:#fff;color:#1a1a1a;border:1px solid #e7e0d0;box-shadow:0 20px 50px rgba(0,0,0,.25);font-family:inherit;text-align:center}
#eddTimer[hidden]{display:none}
#eddTimer .edd-time{font-size:3.25rem;font-weight:800;font-variant-numeric:tabular-nums;line-height:1.1;margin:.25rem 0}
#eddTimer .edd-label{font-size:.65rem;font-weight:800;letter-spacing:.12em;text-transform:uppercase;opacity:.55}
#eddTimer .edd-row{display:flex;gap:.4rem;justify-content:center;margin-top:.6rem}
#eddTimer button{border:1px solid #e7e0d0;background:#faf6ec;border-radius:.75rem;padding:.35rem .7rem;font-size:.75rem;font-weight:700;cursor:pointer}
#eddTimer button:hover{background:#f1ead8}
#eddTimer .edd-hint{font-size:.65rem;opacity:.55;margin-top:.6rem}
#eddTimer.edd-urgent .edd-time{color:#dc2626}
#eddTimer.edd-done .edd-time{color:#059669}
#eddTimer.edd-full{inset:0;right:0;bottom:0;width:auto;border-radius:0;display:flex;flex-direction:column;align-items:center;justify-content:center;background:#fffdf7}
#eddTimer.edd-full .edd-time{font-size:min(28vw,22rem)}
#eddTimer.edd-full .edd-label{font-size:1rem}
#eddTimer.edd-full button{font-size:1rem;padding:.6rem 1.2rem}
#eddConfetti{position:fixed;inset:0;width:100%;height:100%;pointer-events:none;z-index:9500}`;
    document.head.appendChild(style);
    el = document.createElement('div');
    el.id = 'eddTimer';
    el.hidden = true;
    el.setAttribute('role', 'timer');
    el.innerHTML = '<div class="edd-label">Timer</div>' +
      '<div class="edd-time" aria-live="off">0:00</div>' +
      '<div class="edd-row"><button type="button" data-a="toggle">Start</button>' +
      '<button type="button" data-a="reset">Reset</button>' +
      '<button type="button" data-a="full" aria-label="Full screen (F)">Full screen</button>' +
      '<button type="button" data-a="close" aria-label="Close timer">Exit</button></div>' +
      '<div class="edd-hint">Press 1-9 for minutes, 0 for 10. Tap again within 1s to double. F = full screen.</div>';
    document.body.appendChild(el);
    display = el.querySelector('.edd-time');
    label = el.querySelector('.edd-label');
    startBtn = el.querySelector('[data-a="toggle"]');
    el.addEventListener('click', e => {
      const a = e.target.closest('button')?.dataset.a;
      if (a === 'toggle') running ? pause() : start();
      else if (a === 'reset') setSeconds(state ? state.seconds : 0, false);
      else if (a === 'full') toggleFull();
      else if (a === 'close') close();
    });
    document.addEventListener('keydown', onKey);
  }

  function ctx() {
    if (!audioCtx) { const C = root.AudioContext || root.webkitAudioContext; if (C) audioCtx = new C(); }
    audioCtx?.state === 'suspended' && audioCtx.resume();
    return audioCtx;
  }
  function tone(freq, start, dur, vol, type) {
    const c = ctx(); if (!c) return;
    const o = c.createOscillator(), g = c.createGain(), t = c.currentTime + start;
    o.type = type || 'sine'; o.frequency.value = freq;
    g.gain.setValueAtTime(vol, t); g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
    o.connect(g); g.connect(c.destination); o.start(t); o.stop(t + dur);
  }
  const playTick = last => tone(last ? 1320 : 880, 0, 0.15, 0.2, 'square');
  // "Celebration" sound: rising arpeggio plus a sparkle.
  function playCelebration() {
    [523, 659, 784, 1047, 1319].forEach((f, i) => tone(f, i * 0.12, 0.5, 0.25, 'triangle'));
    [1047, 1319, 1568].forEach(f => tone(f, 0.75, 0.9, 0.18, 'sine'));
  }

  function startConfetti() {
    if (root.matchMedia?.('(prefers-reduced-motion: reduce)').matches) return;
    stopConfetti();
    const cv = confetti = document.createElement('canvas');
    cv.id = 'eddConfetti';
    cv.width = innerWidth; cv.height = innerHeight;
    document.body.appendChild(cv);
    const g = cv.getContext('2d'), colors = ['#dc2626', '#f59e0b', '#10b981', '#3b82f6', '#8b5cf6', '#ec4899'];
    const bits = Array.from({ length: 220 }, () => ({
      x: Math.random() * cv.width, y: -20 - Math.random() * cv.height * 0.6,
      w: 6 + Math.random() * 8, h: 8 + Math.random() * 10, vx: -1.5 + Math.random() * 3,
      vy: 2 + Math.random() * 4, r: Math.random() * 6, vr: -0.2 + Math.random() * 0.4,
      c: colors[(Math.random() * colors.length) | 0]
    }));
    const t0 = performance.now();
    (function frame(now) {
      if (confetti !== cv) return;
      g.clearRect(0, 0, cv.width, cv.height);
      let alive = false;
      bits.forEach(b => {
        b.x += b.vx; b.y += b.vy; b.r += b.vr;
        if (b.y < cv.height + 20) alive = true;
        g.save(); g.translate(b.x, b.y); g.rotate(b.r); g.fillStyle = b.c;
        g.fillRect(-b.w / 2, -b.h / 2, b.w, b.h); g.restore();
      });
      if (alive && now - t0 < 8000) requestAnimationFrame(frame); else stopConfetti();
    })(t0);
  }
  function stopConfetti() { confetti?.remove(); confetti = null; }

  function render() {
    display.textContent = formatTime(remaining);
    startBtn.textContent = running ? 'Pause' : 'Start';
    el.classList.toggle('edd-urgent', running && remaining > 0 && remaining <= 10);
    document.title = running ? formatTime(remaining) + ' – Timer' : baseTitle;
  }
  let baseTitle = '';

  function setSeconds(seconds, autostart) {
    stopConfetti();
    el.classList.remove('edd-done');
    label.textContent = 'Timer';
    remaining = seconds;
    lastBeepSecond = null;
    state = state ? Object.assign(state, { seconds }) : { seconds };
    autostart ? start() : (pause(), render());
  }
  function start() {
    if (remaining <= 0) { if (!state || !state.seconds) return; remaining = state.seconds; }
    ctx();
    stopConfetti();
    el.classList.remove('edd-done');
    label.textContent = 'Timer';
    running = true;
    endsAt = Date.now() + remaining * 1000;
    clearInterval(tick);
    tick = setInterval(update, 200);
    update();
  }
  function pause() {
    running = false; clearInterval(tick); tick = null; render();
  }
  function update() {
    remaining = Math.max(0, (endsAt - Date.now()) / 1000);
    const whole = Math.ceil(remaining);
    if (remaining > 0 && whole <= 10 && whole !== lastBeepSecond) { lastBeepSecond = whole; playTick(whole <= 3); }
    if (remaining <= 0) { finish(); return; }
    render();
  }
  function finish() {
    running = false; clearInterval(tick); tick = null; remaining = 0;
    el.classList.add('edd-done'); label.textContent = "Time's up!";
    render(); playCelebration(); startConfetti();
  }

  function isFull() { return el.classList.contains('edd-full'); }
  function toggleFull() {
    const full = !isFull();
    el.classList.toggle('edd-full', full);
    el.querySelector('[data-a="full"]').textContent = full ? 'Exit full screen' : 'Full screen';
    if (full) el.requestFullscreen?.().catch(() => {});
    else if (document.fullscreenElement) document.exitFullscreen?.().catch(() => {});
  }

  function close() {
    pause(); stopConfetti();
    if (isFull()) toggleFull();
    el.hidden = true; lastPress = null; state = null;
    document.title = baseTitle;
  }

  function open() {
    if (!el) { baseTitle = document.title; build(); }
    el.hidden = false;
    if (!state) { state = { seconds: 0 }; remaining = 0; render(); }
  }

  function onKey(e) {
    if (el.hidden || e.ctrlKey || e.metaKey || e.altKey) return;
    const t = e.target;
    if (t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName))) return;
    if (e.key === 'f' || e.key === 'F') { e.preventDefault(); toggleFull(); return; }
    if (e.key === 'Escape' && isFull()) { toggleFull(); return; }
    const next = nextDuration(lastPress, e.key, Date.now());
    if (!next) return;
    e.preventDefault();
    lastPress = next;
    setSeconds(next.seconds, true);
  }

  // Browser Esc exits the Fullscreen API on its own; keep the layout class in sync.
  document.addEventListener('fullscreenchange', () => {
    if (el && !document.fullscreenElement && isFull()) {
      el.classList.remove('edd-full');
      el.querySelector('[data-a="full"]').textContent = 'Full screen';
    }
  });

  root.openTimer = open;
  document.addEventListener('click', e => {
    if (e.target.closest('[data-open-timer]')) open();
  });
})(typeof window !== 'undefined' ? window : globalThis);
