/* ================================================================
   Audit Capture — PowerPoint Add-in v5.0
   ================================================================ */

const state = { imagesPlaced: 0 };

function setStatus(text, cls = '') {
  const el = document.getElementById('status');
  if (el) { el.textContent = text; el.className = 'status ' + cls; }
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

function updateStats(slide, images) {
  const cs = document.getElementById('currentSlide');
  const si = document.getElementById('slideImages');
  const ti = document.getElementById('totalImages');
  if (cs) cs.textContent = slide;
  if (si) si.textContent = images;
  if (ti) ti.textContent = state.imagesPlaced;
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus('Cet add-in ne fonctionne que dans PowerPoint.', 'err');
    return;
  }

  const templates = window.AuditTemplates.getList();
  log(`Add-in v5.0 démarré — ${templates.length} catégorie(s) : ${templates.map(t => t.label).join(', ')}`);
  setStatus('Prêt ✅ En attente de l\'extension…', 'ok');

  /* Listen for images */
  window.addEventListener('message', (e) => {
    if (!e.data || e.data.type !== 'AUDIT_ADD_IMAGE') return;
    const mode = e.data.mode || 'normal';
    const templateKey = e.data.template || 'street';
    log(`Image reçue (template=${templateKey}, mode=${mode})`);
    enqueueImage(e.data.dataUrl, templateKey);
  });

  /* Direct call */
  window.__auditPlaceImage = (dataUrl, templateKey) => {
    log(`Image reçue (direct, template=${templateKey || 'street'})`);
    enqueueImage(dataUrl, templateKey || 'street');
  };

  /* Query API */
  window.__auditQuerySlide = async () => ({ ok: true, summary: 'Prêt' });

  setTimeout(refreshStats, 1500);
});

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
        (s) => s.type === PowerPoint.ShapeType.image && (s.name || '').startsWith('audit-img-')
      ).length;
      updateStats(slides.items.length, imgs);
    });
  } catch (e) {}
}

const resetBtn = document.getElementById('reset');
if (resetBtn) {
  resetBtn.addEventListener('click', () => {
    state.imagesPlaced = 0;
    if (placementState) placementState.imagesPlaced = 0;
    updateStats('-', 0);
    log('Session réinitialisée');
  });
}

/* Notify helper (may be missing) */
if (typeof notify !== 'function') {
  window.notify = function (msg) {
    try {
      chrome.notifications && chrome.notifications.create({
        type: 'basic',
        iconUrl: 'icon128.png',
        title: 'Audit Capture',
        message: msg,
      });
    } catch (e) {}
  };
}