/* ================================================================
   Audit Capture — PowerPoint Add-in v5.3
   ----------------------------------------------------------------
   v5.3: adds "📐 Mesurer les formes" button — dumps every shape's
         position and size so the slot coordinates can be read
         directly from the taskpane (no console needed).
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
  log(`Add-in v5.3 démarré — ${templates.length} catégorie(s) : ${templates.map(t => t.label).join(', ')}`);
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

/* ================================================================
   RESET button
================================================================ */
const resetBtn = document.getElementById('reset');
if (resetBtn) {
  resetBtn.addEventListener('click', () => {
    state.imagesPlaced = 0;
    if (typeof placementState !== 'undefined' && placementState) {
      placementState.imagesPlaced = 0;
    }
    updateStats('-', 0);
    log('Session réinitialisée');
  });
}

/* ================================================================
   MEASURE button — dump every shape on every slide
================================================================ */
const measureBtn = document.getElementById('measure');
const measureOut = document.getElementById('measure-output');

if (measureBtn && measureOut) {
  measureBtn.addEventListener('click', async () => {
    measureOut.style.display = 'block';
    measureOut.textContent = '⏳ Mesure en cours…';

    try {
      const result = await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const out = [];

        for (let si = 0; si < slides.items.length; si++) {
          const slide = slides.items[si];
          slide.shapes.load('items');
          await context.sync();

          for (let i = 0; i < slide.shapes.items.length; i++) {
            const s = slide.shapes.items[i];
            let text = '';
            try {
              if (s.type === PowerPoint.ShapeType.textBox ||
                  s.type === PowerPoint.ShapeType.geometricShape) {
                s.textFrame.load('textRange/text');
              }
            } catch (e) {}
            try { await context.sync(); } catch (e) {}
            try {
              text = (s.textFrame && s.textFrame.textRange && s.textFrame.textRange.text || '').substring(0, 30);
            } catch (e) {}

            out.push({
              slide: si + 1,
              idx: i,
              name: s.name || '(unnamed)',
              type: s.type,
              left: Math.round(s.left ?? 0),
              top: Math.round(s.top ?? 0),
              w: Math.round(s.width ?? 0),
              h: Math.round(s.height ?? 0),
              text: text,
            });
          }
        }
        return out;
      });

      /* Format as text */
      let txt = `Total: ${result.length} formes\n\n`;
      txt += 'slide | idx | name                     | type | left | top  | w    | h\n';
      txt += '------+-----+--------------------------+------+------+------+------+------\n';

      for (const r of result) {
        txt += `${String(r.slide).padEnd(5)} | ${String(r.idx).padEnd(3)} | ${r.name.substring(0, 24).padEnd(24)} | ${String(r.type).padEnd(4)} | ${String(r.left).padEnd(4)} | ${String(r.top).padEnd(4)} | ${String(r.w).padEnd(4)} | ${r.h}\n`;
        if (r.text) {
          txt += `      |     | text: "${r.text}"\n`;
        }
      }

      txt += '\n\n--- JSON ---\n';
      txt += JSON.stringify(result, null, 2);

      measureOut.textContent = txt;
      log(`📐 ${result.length} formes mesurées`);
    } catch (e) {
      measureOut.textContent = '❌ ' + (e.message || e);
      log('❌ Mesure échouée : ' + (e.message || e), 'err');
    }
  });
}

/* ================================================================
   Legacy measure function (still callable from console if needed)
================================================================ */
window.__measureSlots = async function () {
  return await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    const slide = slides.items[0];
    slide.shapes.load('items');
    await context.sync();

    const shapes = slide.shapes.items;
    const result = [];

    for (let i = 0; i < shapes.length; i++) {
      const s = shapes[i];
      let text = '';
      try {
        if (s.type === PowerPoint.ShapeType.textBox ||
            s.type === PowerPoint.ShapeType.geometricShape) {
          s.textFrame.load('textRange/text');
        }
      } catch (e) {}
      try { await context.sync(); } catch (e) {}
      try { text = (s.textFrame?.textRange?.text || '').substring(0, 40); } catch (e) {}

      result.push({
        index: i,
        name: s.name || '(unnamed)',
        type: s.type,
        left: Math.round(s.left ?? 0),
        top: Math.round(s.top ?? 0),
        width: Math.round(s.width ?? 0),
        height: Math.round(s.height ?? 0),
        text: text,
      });
    }
    return result;
  });
};

/* ================================================================
   Notify helper (kept for compatibility)
================================================================ */
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