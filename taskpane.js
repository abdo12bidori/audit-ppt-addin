/* ================================================================
   Audit Capture — PowerPoint Add-in v3.7
   ================================================================
   This is the boot file. All placement logic lives in placement.js.

   v3.7 changes:
     - Exposes window.__auditQuerySlide() so the content script can
       request the current slide state (pre-flight).
     - Updates the pre-flight hint on the page.
     - Manual "Rafraîchir" button.
     - Reset button also clears the helper tracker.
   ================================================================ */

const state = {
  imagesPlaced: 0,
  masterWarningShown: false,
};

/* ================================================================
   UI helpers
================================================================ */
function setStatus(text, cls = '') {
  const el = document.getElementById('status');
  if (el) {
    el.textContent = text;
    el.className = 'status ' + cls;
  }
}

function setPreflight(text, cls = '') {
  const el = document.getElementById('preflight');
  if (el) {
    el.textContent = text || '—';
    el.className = 'preflight ' + cls;
    el.style.display = text ? 'block' : 'none';
  }
}

function log(msg, cls = '') {
  console.log('[AuditCapture:addin]', msg);
  const el = document.getElementById('log');
  if (el) {
    const line = document.createElement('div');
    if (cls) line.className = cls;
    line.textContent = new Date().toLocaleTimeString() + ' — ' + msg;
    el.prepend(line);
  }
}

function updateStats(currentSlide, slideImages) {
  const cs = document.getElementById('currentSlide');
  const si = document.getElementById('slideImages');
  const ti = document.getElementById('totalImages');
  if (cs) cs.textContent = currentSlide;
  if (si) si.textContent = slideImages;
  if (ti) ti.textContent = state.imagesPlaced;
}

/* ================================================================
   Boot
================================================================ */
Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus('Cet add-in ne fonctionne que dans PowerPoint.', 'err');
    log('Wrong host', 'err');
    return;
  }

  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.2')) {
    log('⚠️ PowerPointApi 1.2 non supporté — slides.add indisponible', 'err');
  }

  setStatus('Prêt ✅ En attente de l\'extension…', 'ok');
  log('Add-in v3.7 démarré');

  listenForImages();
  exposeQueryApi();

  setTimeout(refreshStats, 1500);
  setTimeout(refreshStats, 3000);
  setTimeout(refreshPreflight, 2000);
});

/* ================================================================
   Pre-flight query API
================================================================ */
function exposeQueryApi() {
  /* The content script calls this via window.__auditQuerySlide() */
  window.__auditQuerySlide = async function () {
    try {
      if (window.AuditQuery && typeof window.AuditQuery.getSlideState === 'function') {
        return await window.AuditQuery.getSlideState();
      }
      return { ok: false, error: 'AuditQuery not loaded', summary: 'Erreur' };
    } catch (e) {
      return { ok: false, error: e.message, summary: 'Erreur' };
    }
  };

  /* And the content script can request a refresh of the visible hint */
  window.__auditRefreshPreflight = async function () {
    await refreshPreflight();
  };
}

async function refreshPreflight() {
  try {
    if (!window.AuditQuery) return;
    const result = await window.AuditQuery.getSlideState();
    if (result.ok) {
      setPreflight(result.summary, result.willCreateNewSlide ? 'warn' : '');
    } else {
      setPreflight('⚠️ ' + (result.error || 'Erreur'), 'warn');
    }
  } catch (e) {
    setPreflight('⚠️ ' + e.message, 'warn');
  }
}

/* ================================================================
   Listen for images — dedupe + serialize
================================================================ */
let __lastDataUrl = null;
let __lastDataUrlTs = 0;
let __queue = Promise.resolve();

function listenForImages() {
  const enqueue = (dataUrl, mode, source) => {
    const now = Date.now();
    if (dataUrl === __lastDataUrl && now - __lastDataUrlTs < 800) {
      log(`Image ignorée (doublon ${source})`);
      return;
    }
    __lastDataUrl = dataUrl;
    __lastDataUrlTs = now;

    __queue = __queue
      .then(() => placeImage(dataUrl, mode || 'normal'))
      .catch((e) => log('❌ placeImage rejected : ' + e.message, 'err'));
  };

  window.addEventListener('message', (e) => {
    if (!e.data || e.data.type !== 'AUDIT_ADD_IMAGE') return;
    const mode = e.data.mode || 'normal';
    log(`Image reçue (postMessage, mode=${mode})`);
    enqueue(e.data.dataUrl, mode, 'postMessage');
  });

  window.__auditPlaceImage = (dataUrl, mode) => {
    log(`Image reçue (direct, mode=${mode || 'normal'})`);
    enqueue(dataUrl, mode || 'normal', 'direct');
  };
}

/* ================================================================
   placeImage wrapper with 12s watchdog
================================================================ */
async function placeImage(dataUrl, mode) {
  try {
    setStatus('⏳ Placement…');

    const watchdog = new Promise((_, reject) =>
      setTimeout(
        () => reject(new Error('Timeout — Office.js ne répond pas (12 s)')),
        12000
      )
    );

    await Promise.race([runPlacement(dataUrl, mode), watchdog]);
    refreshStats();
    refreshPreflight();
  } catch (err) {
    console.error(err);
    log('❌ Erreur : ' + err.message, 'err');
    setStatus('Erreur : ' + err.message, 'err');
    refreshPreflight();
  }
}

/* ================================================================
   Refresh stats
================================================================ */
async function refreshStats() {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length === 0) return;

      const last = slides.items[slides.items.length - 1];
      last.shapes.load('items');
      await context.sync();

      const imgs = last.shapes.items.filter(
        (s) =>
          s.type === PowerPoint.ShapeType.image &&
          (s.name || '').startsWith('audit-img-')
      ).length;

      updateStats(slides.items.length, imgs);
    });
  } catch (e) {
    console.warn('[AuditCapture:addin] refreshStats failed:', e.message || e);
  }
}

/* ================================================================
   Buttons
================================================================ */
const resetBtn = document.getElementById('reset');
if (resetBtn) {
  resetBtn.addEventListener('click', () => {
    state.imagesPlaced = 0;
    placementState.imagesPlaced = 0;
    placementState.masterWarningShown = false;

    try {
      if (window.AuditHelpers && typeof window.AuditHelpers.resetHelperState === 'function') {
        window.AuditHelpers.resetHelperState();
      }
    } catch (e) {}

    updateStats('-', 0);
    setPreflight('');
    log('Session réinitialisée');
    refreshPreflight();
  });
}

const refreshBtn = document.getElementById('refresh');
if (refreshBtn) {
  refreshBtn.addEventListener('click', () => {
    log('Rafraîchissement de l\'état de la slide…');
    refreshStats();
    refreshPreflight();
  });
}