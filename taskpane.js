/* ================================================================
   Audit Capture — PowerPoint Add-in v2.2
   - Uses setSelectedDataAsync (works everywhere)
   - Positions the image with Office.js
   - v2.1: replaced non-existent slide.duplicate() with slides.add()
   - v2.2: queue + dedupe + 15s watchdog to fix hangs when the
           extension sends the same image twice via postMessage
================================================================ */

const LAYOUT = {
  perSlide: 2,
  margin: 0.5,
  gap: 0.3,
  topOffset: 1.5,
  aspect: 4 / 3,
};

const state = {
  imagesPlaced: 0,
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
  log('Add-in v2.2 démarré');

  listenForImages();
  setTimeout(refreshStats, 500);
});

/* ================================================================
   Listen for images — with queue + dedupe
   The extension's postmessage.js posts to BOTH window and every
   iframe, so the Add-in often receives the SAME image twice.
   We dedupe by dataUrl hash + serialize all placements.
================================================================ */
let __lastDataUrl = null;
let __lastDataUrlTs = 0;
let __queue = Promise.resolve();

function listenForImages() {
  const enqueue = (dataUrl, source) => {
    const now = Date.now();
    if (dataUrl === __lastDataUrl && now - __lastDataUrlTs < 800) {
      log(`Image ignorée (doublon ${source})`);
      return;
    }
    __lastDataUrl = dataUrl;
    __lastDataUrlTs = now;

    __queue = __queue
      .then(() => placeImage(dataUrl))
      .catch((e) => log('❌ placeImage rejected : ' + e.message, 'err'));
  };

  window.addEventListener('message', (e) => {
    if (!e.data || e.data.type !== 'AUDIT_ADD_IMAGE') return;
    log('Image reçue (postMessage)');
    enqueue(e.data.dataUrl, 'postMessage');
  });

  window.__auditPlaceImage = (dataUrl) => {
    log('Image reçue (direct)');
    enqueue(dataUrl, 'direct');
  };

  /* WebSocket — DISABLED because there's no local relay running.
     It was spamming the console with ERR_CONNECTION_REFUSED.
     Uncomment if you set up ws://127.0.0.1:8765 later. */
  // connectWebSocket();
}

/* Kept for optional future use — do not call unless you run a local relay */
function connectWebSocket() {
  try {
    const ws = new WebSocket('ws://127.0.0.1:8765/ppt');
    ws.onopen = () => log('WebSocket connecté');
    ws.onmessage = (e) => {
      try {
        const msg = JSON.parse(e.data);
        if (msg.type === 'AUDIT_ADD_IMAGE') {
          log('Image reçue (WebSocket)');
          placeImage(msg.dataUrl);
        }
      } catch (err) {}
    };
    ws.onclose = () => {
      setTimeout(connectWebSocket, 3000);
    };
    ws.onerror = () => {};
  } catch (e) {}
}

/* ================================================================
   Wrapper with 15s watchdog
================================================================ */
async function placeImage(dataUrl) {
  try {
    setStatus('⏳ Placement…');

    const watchdog = new Promise((_, reject) =>
      setTimeout(
        () => reject(new Error('Timeout — Office.js ne répond pas (15 s)')),
        15000
      )
    );

    await Promise.race([runPlacement(dataUrl), watchdog]);

    refreshStats();
  } catch (err) {
    console.error(err);
    log('❌ Erreur : ' + err.message, 'err');
    setStatus('Erreur : ' + err.message, 'err');
  }
}

/* ================================================================
   Core — place an image
   1. Look at the LAST slide
   2. If it already has perSlide images → add a new slide (same layout)
   3. Insert image via setSelectedDataAsync
   4. Position the inserted image in the correct grid slot
================================================================ */
async function runPlacement(dataUrl) {
  await PowerPoint.run(async (context) => {
    /* ----------------------------------------------------------
       1. Load slides
    ---------------------------------------------------------- */
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slides.items.length === 0) {
      throw new Error('No slides in presentation');
    }

    let targetSlide = slides.items[slides.items.length - 1];
    targetSlide.shapes.load('items');
    await context.sync();

    const imagesOnTarget = targetSlide.shapes.items.filter(
      (s) => s.type === PowerPoint.ShapeType.image
    );

    log(`Slide ${slides.items.length} : ${imagesOnTarget.length} image(s)`);

    /* ----------------------------------------------------------
       2. If slide is full → add a NEW slide with the same layout
          ⚠️ slide.duplicate() does NOT exist in the JS API.
             Use slides.add({ layoutId }) instead.
    ---------------------------------------------------------- */
    if (imagesOnTarget.length >= LAYOUT.perSlide) {
      log('Slide pleine → ajout d\'une nouvelle slide');

      targetSlide.layout.load('id,name');
      await context.sync();

      const layoutId = targetSlide.layout.id;
      slides.add({ layoutId });
      await context.sync();

      slides.load('items');
      await context.sync();

      targetSlide = slides.items[slides.items.length - 1];
      targetSlide.shapes.load('items');
      await context.sync();

      log(`Nouvelle slide créée : ${slides.items.length}`);
    }

    /* ----------------------------------------------------------
       3. Get slide dimensions
    ---------------------------------------------------------- */
    const pageSetup = context.presentation.pageSetup;
    pageSetup.load(['slideWidth', 'slideHeight']);
    await context.sync();

    const slideW = pageSetup.slideWidth;
    const slideH = pageSetup.slideHeight;

    /* ----------------------------------------------------------
       4. Re-count images to pick the right slot
    ---------------------------------------------------------- */
    targetSlide.shapes.load('items');
    await context.sync();

    const currentImages = targetSlide.shapes.items.filter(
      (s) => s.type === PowerPoint.ShapeType.image
    );
    const slotIndex = currentImages.length;

    /* ----------------------------------------------------------
       5. Compute position (grid of `perSlide` columns)
    ---------------------------------------------------------- */
    const margin = LAYOUT.margin * 72;
    const gap = LAYOUT.gap * 72;
    const topOffset = LAYOUT.topOffset * 72;

    const cols = LAYOUT.perSlide;
    const cellW = (slideW - 2 * margin - gap * (cols - 1)) / cols;
    const cellH = cellW / LAYOUT.aspect;

    const col = slotIndex % cols;
    const left = margin + col * (cellW + gap);

    log(`Position : left=${left.toFixed(0)} top=${topOffset.toFixed(0)} w=${cellW.toFixed(0)} h=${cellH.toFixed(0)}`);

    /* ----------------------------------------------------------
       6. Insert the image via setSelectedDataAsync
    ---------------------------------------------------------- */
    const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');

    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        base64,
        { coercionType: Office.CoercionType.Image },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error('setSelectedDataAsync failed: ' + (result.error?.message || 'unknown')));
          }
        }
      );
    });

    log('Image insérée via setSelectedDataAsync');

    /* ----------------------------------------------------------
       7. Position the newly-inserted image
    ---------------------------------------------------------- */
    await new Promise((r) => setTimeout(r, 200));

    await PowerPoint.run(async (context2) => {
      const slides2 = context2.presentation.slides;
      slides2.load('items');
      await context2.sync();

      const lastSlide = slides2.items[slides2.items.length - 1];
      lastSlide.shapes.load('items');
      await context2.sync();

      const imgs = lastSlide.shapes.items.filter(
        (s) => s.type === PowerPoint.ShapeType.image
      );

      if (imgs.length === 0) {
        log('⚠️ No image found after insertion', 'err');
        return;
      }

      const img = imgs[imgs.length - 1];

      img.left = left;
      img.top = topOffset;
      img.width = cellW;
      img.height = cellH;

      await context2.sync();
      log(`✅ Image positionnée — slide ${slides2.items.length}, slot ${slotIndex + 1}`, 'ok');
    });

    state.imagesPlaced++;
    setStatus(`✅ Image placée (slot ${slotIndex + 1})`, 'ok');
  });
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
        (s) => s.type === PowerPoint.ShapeType.image
      ).length;

      updateStats(slides.items.length, imgs);
    });
  } catch (e) {}
}

/* ================================================================
   Reset button
================================================================ */
document.getElementById('reset').addEventListener('click', () => {
  state.imagesPlaced = 0;
  updateStats('-', 0);
  log('Session réinitialisée');
});
