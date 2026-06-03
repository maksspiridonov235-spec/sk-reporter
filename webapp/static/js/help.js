// ── Макросы ───────────────────────────────────────────────────────────────

(function initMacroDate() {
  const el = document.getElementById('macroReportDate');
  if (el) el.value = new Date().toISOString().slice(0, 10);
})();

(function initResultsPanelResize() {
  const handle = document.getElementById('resultsResizeHandle');
  const panel = document.getElementById('resultsPanel');
  if (!handle || !panel) return;

  const MIN = 280;
  const maxWidth = () => Math.min(560, Math.floor(window.innerWidth * 0.55));

  try {
    const saved = parseInt(localStorage.getItem('resultsPanelWidth'), 10);
    if (saved >= MIN && saved <= maxWidth()) panel.style.width = saved + 'px';
  } catch (_) {}

  handle.addEventListener('mousedown', (e) => {
    e.preventDefault();
    const startX = e.clientX;
    const startW = panel.offsetWidth;
    handle.classList.add('is-dragging');
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';

    function onMove(ev) {
      const w = Math.min(maxWidth(), Math.max(MIN, startW + (startX - ev.clientX)));
      panel.style.width = w + 'px';
    }
    function onUp() {
      handle.classList.remove('is-dragging');
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      try { localStorage.setItem('resultsPanelWidth', String(panel.offsetWidth)); } catch (_) {}
    }
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
})();

(function initPrepareBlockHelp() {
  const btn = document.getElementById('helpPrepareBtn');
  const closeBtn = document.getElementById('helpPrepareClose');
  const pop = document.getElementById('helpPreparePopover');
  const backdrop = document.getElementById('helpPrepareBackdrop');
  if (!btn || !pop) return;

  function setOpen(open) {
    pop.classList.toggle('is-open', open);
    if (backdrop) {
      backdrop.classList.toggle('is-open', open);
      backdrop.setAttribute('aria-hidden', open ? 'false' : 'true');
    }
    btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    refreshActivityEmptyHint();
  }

  function toggle() {
    setOpen(!pop.classList.contains('is-open'));
  }

  window.closePrepareHelp = () => setOpen(false);

  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    toggle();
  });
  if (closeBtn) {
    closeBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      setOpen(false);
    });
  }
  if (backdrop) {
    backdrop.addEventListener('click', () => setOpen(false));
  }

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && pop.classList.contains('is-open')) setOpen(false);
  });
})();
