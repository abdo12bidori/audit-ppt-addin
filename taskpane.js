/* ================================================================
   Audit Capture — PowerPoint Add-in v4.0 (template-driven)
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

  const templates = window.AuditTemplates && window.AuditTemplates.getList
    ? window.AuditTemplates.getList()
    : [];
  log(`Add-in v4.0 démarré — ${templates.length} template(s) chargé(s) : ${templates.map((t) => t.label).join(', ')}`);

  setStatus('Prêt ✅ En attente de l\'extension…', 'ok');

  listenForImages();
  exposeQueryApi();

  setTimeout(refreshStats, 1500);
});

/* ================================================================
   Pre-flight query API
================================================================ */
function exposeQueryApi() {
  window.__auditQuerySlide = async function () {
    return {
      ok: true,
      slideNumber: '?',
      summary: 'Templates chargés : ' + (window.AuditTemplatesList || []).join(', '),
    };
  };
}

/* ================================================================
   Listen for images — dedupe + serialize
================================================================ */
let __lastDataUrl = null;
let __lastDataUrlTs = 0;
let __queue = Promise.resolve();

function listenForImages() {
  const enqueue = (dataUrl, mode, templateKey, source) => {
    const now = Date.now();
    const dedupeKey = (templateKey || '') + '|' + dataUrl;
    if (dedupeKey === __lastDataUrl && now - __lastDataUrlTs < 800) {
      log(`Image ignorée (doublon ${source})`);
      return;
    }
    __lastDataUrl = dedupeKey;
    __lastDataUrlTs = now;

    __queue = __queue
      .then(() => placeImage(dataUrl, mode || 'normal', templateKey || 'street'))
      .catch((e) => log('❌ placeImage rejected : ' + e.message, 'err'));
  };

  window.addEventListener('message', (e) => {
    if (!e.data || e.data.type !== 'AUDIT_ADD_IMAGE') return;
    const mode = e.data.mode || 'normal';
    const templateKey = e.data.template || 'street';
    log(`Image reçue (postMessage, mode=${mode}, template=${templateKey})`);
    enqueue(e.data.dataUrl, mode, templateKey, 'postMessage');
  });

  window.__auditPlaceImage = (dataUrl, mode, templateKey) => {
    log(`Image reçue (direct, mode=${mode || 'normal'}, template=${templateKey || 'street'})`);
    enqueue(dataUrl, mode || 'normal', templateKey || 'street', 'direct');
  };
}

/* ================================================================
   placeImage wrapper with 20s watchdog
================================================================ */
async function placeImage(dataUrl, mode, templateKey) {
  try {
    setStatus(`⏳ Placement ${templateKey}…`);

    const watchdog = new Promise((_, reject) =>
      setTimeout(
        () => reject(new Error('Timeout — Office.js ne répond pas (20 s)')),
        20000
      )
    );

    await Promise.race([runPlacement(dataUrl, mode, templateKey), watchdog]);
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
  } catch (e) {
    console.warn('[AuditCapture:addin] refreshStats failed:', e.message || e);
  }
}

/* ================================================================
   Reset button
================================================================ */
const resetBtn = document.getElementById('reset');
if (resetBtn) {
  resetBtn.addEventListener('click', () => {
    state.imagesPlaced = 0;
    placementState.imagesPlaced = 0;
    updateStats('-', 0);
    log('Session réinitialisée');
  });
}