/* ================================================================
   Audit Capture — PowerPoint Add-in v3.1
   ================================================================
   What this does:
   - Receives an image via postMessage from the Chrome extension
   - Detects the free zone on the last slide (below header, above footer)
   - Places the image so that it ADAPTS to the slot:
       • Keeps aspect ratio (no distortion)
       • No quality loss (original bytes preserved, Office.js just
         scales the displayed shape — the image inside the file
         stays full resolution)
       • Centered in its slot → looks natural
   - Images share the slide side by side:
       1 = 100%   2 = 50% each   3 = 33% each   →  4th = new slide
   - mode:'end' → closes the current slide. The next image will
     start a fresh slide automatically, even if the row isn't full.
   - Uses only fractions of the slide size (no hardcoded pixels)
   ================================================================ */

/* ----------------------------------------------------------------
   Config — all fractions, no absolute measurements
---------------------------------------------------------------- */
const CFG = {
  MAX_PER_ROW: 3,
  MARGIN_FRAC: 0.04,
  GAP_FRAC: 0.015,
  BODY_TOP_PAD_FRAC: 0.02,
  BODY_BOT_PAD_FRAC: 0.02,
  HEADER_TOP_FRAC: 0.30,
  FOOTER_BOT_FRAC: 0.85,
  FALLBACK_HEADER_FRAC: 0.15,
  FALLBACK_FOOTER_FRAC: 0.90,
  NAMESPACE: 'audit-img-',
  CLOSED_FLAG: 'audit-closed',   // shape-name marker for a closed slide
};

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
  log('Add-in v3.1 démarré');

  listenForImages();
  setTimeout(refreshStats, 500);
});

/* ================================================================
   Listen for images — dedupe + serialize + mode
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

  /* WebSocket disabled — uncomment if you run a local relay */
  // connectWebSocket();
}

function connectWebSocket() {
  try {
    const ws = new WebSocket('ws://127.0.0.1:8765/ppt');
    ws.onopen = () => log('WebSocket connecté');
    ws.onmessage = (e) => {
      try {
        const msg = JSON.parse(e.data);
        if (msg.type === 'AUDIT_ADD_IMAGE') {
          placeImage(msg.dataUrl, msg.mode || 'normal');
        }
      } catch (err) {}
    };
    ws.onclose = () => setTimeout(connectWebSocket, 3000);
    ws.onerror = () => {};
  } catch (e) {}
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
   Decode image dimensions (off-DOM)
================================================================ */
function decodeImageDims(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => reject(new Error('Cannot decode image'));
    img.src = dataUrl;
  });
}

/* ================================================================
   scanSlide — inspect the slide (dynamic, no hardcoded pixels)
================================================================ */
function scanSlide(slide, slideW, slideH) {
  const shapes = slide.shapes.items || [];
  const header = [];
  const footer = [];
  const auditImages = [];
  let closed = false;

  for (const s of shapes) {
    if (!s || s.type === undefined) continue;
    const name = s.name || '';

    if (name.startsWith(CFG.CLOSED_FLAG)) {
      closed = true;
      continue;
    }

    if (name.startsWith(CFG.NAMESPACE)) {
      auditImages.push(s);
      continue;
    }

    const top = s.top ?? 0;
    const bottom = top + (s.height ?? 0);

    if (top < slideH * CFG.HEADER_TOP_FRAC) header.push({ bottom });
    else if (bottom > slideH * CFG.FOOTER_BOT_FRAC) footer.push({ top });
  }

  const headerBottom = header.length
    ? Math.max(...header.map((h) => h.bottom))
    : slideH * CFG.FALLBACK_HEADER_FRAC;

  const footerTop = footer.length
    ? Math.min(...footer.map((f) => f.top))
    : slideH * CFG.FALLBACK_FOOTER_FRAC;

  return { headerBottom, footerTop, auditImages, closed };
}

/* ================================================================
   computeGrid — dynamic slot layout for N images
   All sizes are fractions of the slide → no hardcoded measurements
================================================================ */
function computeGrid(scan, N, slideW, slideH) {
  const margin = slideW * CFG.MARGIN_FRAC;
  const gap = slideW * CFG.GAP_FRAC;
  const usableW = slideW - 2 * margin;

  const bodyTop = scan.headerBottom + slideH * CFG.BODY_TOP_PAD_FRAC;
  const bodyBottom = scan.footerTop - slideH * CFG.BODY_BOT_PAD_FRAC;
  const bodyH = Math.max(bodyBottom - bodyTop, slideH * 0.15);

  const cellW = (usableW - gap * (N - 1)) / N;
  const cellH = bodyH;

  const slots = [];
  for (let i = 0; i < N; i++) {
    slots.push({
      x: margin + i * (cellW + gap),
      y: bodyTop,
      w: cellW,
      h: cellH,
    });
  }
  return slots;
}

/* ================================================================
   fitContain — adapt image to slot with NO distortion
   - Keeps aspect ratio
   - Uses min scale → image fully visible inside slot
   - Centers the result (looks natural)
   - Office.js only changes the DISPLAY size; original bytes stay
     intact inside the PPTX → no quality loss
================================================================ */
function fitContain(slot, imgW, imgH) {
  const scale = Math.min(slot.w / imgW, slot.h / imgH);
  const w = imgW * scale;
  const h = imgH * scale;
  return {
    x: slot.x + (slot.w - w) / 2,
    y: slot.y + (slot.h - h) / 2,
    w,
    h,
  };
}

/* ================================================================
   Core — runPlacement
================================================================ */
async function runPlacement(dataUrl, mode) {
  /* 1. Decode image dims */
  let dims;
  try {
    dims = await decodeImageDims(dataUrl);
  } catch (e) {
    log('⚠️ Cannot decode image dims — fallback 4:3', 'err');
    dims = { w: 800, h: 600 };
  }
  log(`Image : ${dims.w}×${dims.h}, mode=${mode}`);

  await PowerPoint.run(async (context) => {
    /* 2. Slide dimensions */
    const pageSetup = context.presentation.pageSetup;
    pageSetup.load(['slideWidth', 'slideHeight']);
    await context.sync();
    const slideW = pageSetup.slideWidth;
    const slideH = pageSetup.slideHeight;

    /* 3. Get last slide */
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    if (slides.items.length === 0) throw new Error('No slides in presentation');

    let target = slides.items[slides.items.length - 1];
    target.shapes.load('items');
    await context.sync();

    /* 4. Master warning (once) */
    if (!state.masterWarningShown && slides.items.length === 1) {
      if (target.shapes.items.length <= 2) {
        log('⚠️ Le masque ne contient pas d\'en-tête. Les nouvelles diapositives seront vides.', 'err');
        state.masterWarningShown = true;
      }
    }

    /* 5. Scan */
    let scan = scanSlide(target, slideW, slideH);
    log(`Slide ${slides.items.length} — ${scan.auditImages.length} image(s), header→${scan.headerBottom.toFixed(0)}, footer→${scan.footerTop.toFixed(0)}, closed=${scan.closed}`);

    /* 6. Decide if we need a new slide */
    const isFull = scan.auditImages.length >= CFG.MAX_PER_ROW;
    const isClosed = scan.closed === true;
    const needsNewSlide = isFull || isClosed;

    if (needsNewSlide) {
      const reason = isClosed ? 'slide clôturée (END)' : `slide pleine (${scan.auditImages.length}/${CFG.MAX_PER_ROW})`;
      log(`${reason} → nouvelle slide`);

      target.layout.load('id');
      await context.sync();
      const layoutId = target.layout.id;

      slides.add({ layoutId });
      await context.sync();

      slides.load('items');
      await context.sync();
      target = slides.items[slides.items.length - 1];
      target.shapes.load('items');
      await context.sync();

      scan = scanSlide(target, slideW, slideH);
      log(`Nouvelle slide : ${slides.items.length} (${scan.auditImages.length} image(s) déjà)`);
    }

    /* 7. Compute grid for existing + new */
    const totalAfter = scan.auditImages.length + 1;
    const slots = computeGrid(scan, totalAfter, slideW, slideH);
    log(`Grille : ${totalAfter} slot(s), largeur ${slots[0].w.toFixed(0)}`);

    /* 8. Reposition existing images to slots 0..N-2, contain-fit */
    for (let i = 0; i < scan.auditImages.length; i++) {
      const img = scan.auditImages[i];
      const slot = slots[i];
      /* Recover the ratio from the shape's current width/height.
         Since we always placed them contain-fit, this stays correct. */
      const curW = img.width || 1;
      const curH = img.height || 1;
      const ratio = curH / curW;
      const fakeW = slot.w;
      const fakeH = fakeW * ratio;
      const fitted = fitContain(slot, fakeW, fakeH);
      img.left = fitted.x;
      img.top = fitted.y;
      img.width = fitted.w;
      img.height = fitted.h;
    }
    await context.sync();

    /* 9. Insert the new image */
    const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        base64,
        { coercionType: Office.CoercionType.Image },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error('setSelectedDataAsync failed: ' + (result.error?.message || 'unknown')));
        }
      );
    });
    log('Image insérée');

    /* 10. Find, tag and position the new image */
    await new Promise((r) => setTimeout(r, 250));

    await PowerPoint.run(async (context2) => {
      const slides2 = context2.presentation.slides;
      slides2.load('items');
      await context2.sync();

      const lastSlide = slides2.items[slides2.items.length - 1];
      lastSlide.shapes.load('items');
      await context2.sync();

      /* Find the newest = image with NO audit- prefix */
      const candidates = lastSlide.shapes.items.filter(
        (s) =>
          s.type === PowerPoint.ShapeType.image &&
          !(s.name || '').startsWith(CFG.NAMESPACE) &&
          !(s.name || '').startsWith(CFG.CLOSED_FLAG)
      );

      if (candidates.length === 0) {
        log('⚠️ Nouvelle image introuvable', 'err');
        return;
      }

      const newImg = candidates[candidates.length - 1];

      /* Tag it */
      const index = scan.auditImages.length + 1;
      try { newImg.name = `${CFG.NAMESPACE}${index}`; } catch (e) {}

      /* Position: contain-fit in its slot */
      const slot = slots[totalAfter - 1];
      const fitted = fitContain(slot, dims.w, dims.h);

      newImg.left = fitted.x;
      newImg.top = fitted.y;
      newImg.width = fitted.w;
      newImg.height = fitted.h;

      await context2.sync();
      log(`✅ Image ${index} placée — slide ${slides2.items.length}, slot ${totalAfter}/${CFG.MAX_PER_ROW}`, 'ok');

      /* 11. If END → mark the slide as closed */
      if (mode === 'end') {
        try {
          /* Add an invisible marker shape with the closed flag in its name */
          const marker = lastSlide.shapes.addTextBox(' ');
          marker.name = `${CFG.CLOSED_FLAG}-${Date.now()}`;
          marker.left = -100;
          marker.top = -100;
          marker.width = 1;
          marker.height = 1;
          await context2.sync();
          log('🔒 Slide clôturée (END) — la prochaine image ira sur une nouvelle slide');
        } catch (e) {
          log('⚠️ Impossible de marquer la slide comme clôturée : ' + e.message, 'err');
        }
      }
    });

    state.imagesPlaced++;
    setStatus(`✅ Image placée (${totalAfter}/${CFG.MAX_PER_ROW})${mode === 'end' ? ' — slide clôturée' : ''}`, 'ok');
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
        (s) =>
          s.type === PowerPoint.ShapeType.image &&
          (s.name || '').startsWith(CFG.NAMESPACE)
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
    state.masterWarningShown = false;
    updateStats('-', 0);
    log('Session réinitialisée');
  });
}
