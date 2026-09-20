/* ================================================================
   Audit Capture — PowerPoint Add-in v3.3
   ================================================================
   This is the boot file. All placement logic lives in placement.js.
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
  log('Add-in v3.3 démarré');

  listenForImages();
  setTimeout(refreshStats, 500);
});

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
   placeImage wrapper with 15s watchdog
================================================================ */
async function placeImage(dataUrl, mode) {
  try {
    setStatus('⏳ Placement…');

    const watchdog = new Promise((_, reject) =>
      setTimeout(
        () => reject(new Error('Timeout — Office.js ne répond pas (15 s)')),
        15000
      )
    );

    await Promise.race([runPlacement(dataUrl, mode), watchdog]);
    refreshStats();
  } catch (err) {
    console.error(err);
    log('❌ Erreur : ' + err.message, 'err');
    setStatus('Erreur : ' + err.message, 'err');
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
  } catch (e) {}
}

/* ================================================================
   Reset button
================================================================ */
const resetBtn = document.getElementById('reset');
if (resetBtn) {
  resetBtn.addEventListener('click', () => {
    state.imagesPlaced = 0;
    placementState.imagesPlaced = 0;
    placementState.masterWarningShown = false;
    updateStats('-', 0);
    log('Session réinitialisée');
  });
}