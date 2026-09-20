/* ================================================================
   Audit Capture — PowerPoint Add-in
   Places screenshot images: 2 per slide, side by side
   Auto-duplicates the slide when the current one is full
================================================================ */

const LAYOUT = {
  perSlide: 2,           /* 2 images per slide */
  margin: 0.5,           /* inches */
  gap: 0.3,              /* inches between images */
  topOffset: 1.5,        /* inches from top (below title) */
  aspect: 4 / 3,         /* image aspect ratio */
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
  setStatus('Prêt ✅ En attente de l\'extension…', 'ok');
  log('Add-in démarré');

  listenForImages();
  setTimeout(refreshStats, 500);
});

/* ================================================================
   Listen for images
   Supports:
   1. window.postMessage (from the extension)
   2. Direct function call: window.__auditPlaceImage(dataUrl)
   3. WebSocket relay (if the extension uses one)
================================================================ */
function listenForImages() {
  /* Method 1: postMessage */
  window.addEventListener('message', (e) => {
    if (!e.data || e.data.type !== 'AUDIT_ADD_IMAGE') return;
    log('Image reçue (postMessage)');
    placeImage(e.data.dataUrl);
  });

  /* Method 2: direct function call */
  window.__auditPlaceImage = (dataUrl) => {
    log('Image reçue (direct)');
    placeImage(dataUrl);
  };

  /* Method 3: WebSocket (optional) */
  connectWebSocket();
}

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
      } catch (err) {
        log('JSON parse error', 'err');
      }
    };
    ws.onclose = () => {
      log('WebSocket fermé, retry 3s');
      setTimeout(connectWebSocket, 3000);
    };
  } catch (e) {
    /* WebSocket not required */
  }
}

/* ================================================================
   Core — place an image on the appropriate slide
================================================================ */
async function placeImage(dataUrl) {
  try {
    setStatus('⏳ Placement…');

    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      slides.load('items');
      await context.sync();

      if (slides.items.length === 0) {
        throw new Error('No slides in presentation');
      }

      /* Find target slide */
      let targetSlide = slides.items[slides.items.length - 1];
      targetSlide.shapes.load('items');
      await context.sync();

      const imagesOnTarget = targetSlide.shapes.items.filter(
        (s) => s.type === PowerPoint.ShapeType.image
      );

      log(`Slide ${slides.items.length} : ${imagesOnTarget.length} image(s)`);

      /* If slide is full → duplicate it */
      if (imagesOnTarget.length >= LAYOUT.perSlide) {
        log('Slide pleine → duplication');
        targetSlide.duplicate();
        await context.sync();

        slides.load('items');
        await context.sync();
        targetSlide = slides.items[slides.items.length - 1];

        /* Remove all images from the duplicate */
        targetSlide.shapes.load('items');
        await context.sync();
        const dupImages = targetSlide.shapes.items.filter(
          (s) => s.type === PowerPoint.ShapeType.image
        );
        for (const img of dupImages) {
          img.delete();
        }
        await context.sync();
        log(`Nouvelle slide créée : ${slides.items.length}`);
      }

      /* Re-count images on target slide */
      targetSlide.shapes.load('items');
      await context.sync();
      const currentImages = targetSlide.shapes.items.filter(
        (s) => s.type === PowerPoint.ShapeType.image
      );
      const slotIndex = currentImages.length;

      /* Calculate position */
      const pageSetup = presentation.pageSetup;
      pageSetup.load(['slideWidth', 'slideHeight']);
      await context.sync();

      const slideW = pageSetup.slideWidth;
      const slideH = pageSetup.slideHeight;

      const margin = LAYOUT.margin * 72;
      const gap = LAYOUT.gap * 72;
      const topOffset = LAYOUT.topOffset * 72;

      const cols = LAYOUT.perSlide;
      const cellW = (slideW - 2 * margin - gap * (cols - 1)) / cols;
      const cellH = cellW / LAYOUT.aspect;

      const col = slotIndex % cols;
      const left = margin + col * (cellW + gap);

      /* Insert image */
      const shape = targetSlide.shapes.addImage(dataUrl);
      shape.left = left;
      shape.top = topOffset;
      shape.width = cellW;
      shape.height = cellH;

      await context.sync();

      state.imagesPlaced++;
      log(`✅ Image placée — slide ${slides.items.length}, slot ${slotIndex + 1}`, 'ok');
      setStatus(`✅ Image placée (slot ${slotIndex + 1})`, 'ok');
    });

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
        (s) => s.type === PowerPoint.ShapeType.image
      ).length;

      updateStats(slides.items.length, imgs);
    });
  } catch (e) {
    /* ignore */
  }
}

/* ================================================================
   Reset button
================================================================ */
document.getElementById('reset').addEventListener('click', () => {
  state.imagesPlaced = 0;
  updateStats('-', 0);
  log('Session réinitialisée');
});