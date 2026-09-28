/* ================================================================
   Placement engine v4.1 — dynamic 2-cell grid
   ================================================================
   - Slide body zone: { x:40, y:152, w:880, h:334 }
   - Max 2 images per slide
   - When 2 images → 2 columns side by side
   - Contain-fit inside each cell (no distortion, no overflow)
   - Wipe audit-img-* on the target slide before placing
   - Queue serializes processing
   - Retry every 3s when all slides full
================================================================ */

const BODY = { x: 40, y: 152, w: 880, h: 334 };
const GAP = 16;                 // gap between 2 cells (pt)
const MAX_PER_SLIDE = 2;
const NAMESPACE = 'audit-img-';
const RETRY_MS = 3000;

const placementState = { imagesPlaced: 0 };

/* ================================================================
   Decode image dimensions
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
   Contain-fit
================================================================ */
function containFit(cell, imgW, imgH) {
  let scale = Math.min(cell.w / imgW, cell.h / imgH);
  if (scale > 1) scale = 1;      // never enlarge
  const w = imgW * scale;
  const h = imgH * scale;
  return {
    x: cell.x + (cell.w - w) / 2,
    y: cell.y + (cell.h - h) / 2,
    w,
    h,
  };
}

/* ================================================================
   Grid layout — N cells for N images (N=1 or 2)
================================================================ */
function computeCells(N) {
  if (N === 1) {
    return [{ x: BODY.x, y: BODY.y, w: BODY.w, h: BODY.h }];
  }
  // N === 2 → 2 columns side by side
  const cellW = (BODY.w - GAP) / 2;
  return [
    { x: BODY.x,                     y: BODY.y, w: cellW, h: BODY.h },
    { x: BODY.x + cellW + GAP,       y: BODY.y, w: cellW, h: BODY.h },
  ];
}

/* ================================================================
   Insert via paste (clipboard + CDP Ctrl+V)
================================================================ */
async function insertViaPaste(base64) {
  const dataUrl = 'data:image/png;base64,' + base64;

  const writeClipboard = async (label) => {
    try {
      window.parent.postMessage({ type: 'AUDIT_WRITE_CLIPBOARD', dataUrl, ts: Date.now() }, '*');
      log(`→ [${label}] Écriture presse-papier`);
      await new Promise((r) => setTimeout(r, 500));
    } catch (e) { log(`⚠️ ${label}: ${e.message}`, 'err'); }
  };

  const pasteRequest = async (label) => {
    try {
      window.parent.postMessage({ type: 'AUDIT_PASTE_REQUEST', ts: Date.now() }, '*');
      log(`→ [${label}] Collage`);
    } catch (e) { log(`⚠️ ${label}: ${e.message}`, 'err'); }
  };

  await writeClipboard('1/4');
  await pasteRequest('2/4');
  await new Promise((r) => setTimeout(r, 900));
  await writeClipboard('3/4');
  await pasteRequest('4/4');
  await new Promise((r) => setTimeout(r, 900));
  return true;
}

/* ================================================================
   Scan slides → find first slide with room
   Returns { slideNumber, existingImages: [shape], needsWipe }
================================================================ */
async function findTargetSlide() {
  return await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      slide.shapes.load('items');
      await context.sync();

      const images = slide.shapes.items.filter(
        (s) => s.type === PowerPoint.ShapeType.image
      );

      // Only count images that overlap our body zone as "audit images"
      const auditImages = images.filter((s) => {
        const l = s.left ?? 0;
        const t = s.top ?? 0;
        const r = l + (s.width ?? 0);
        const b = t + (s.height ?? 0);
        return !(r < BODY.x || l > BODY.x + BODY.w ||
                 b < BODY.y || t > BODY.y + BODY.h);
      });

      if (auditImages.length < MAX_PER_SLIDE) {
        return {
          slideNumber: i + 1,
          existingCount: auditImages.length,
          existingImages: auditImages.map((s) => ({
            id: s.id,
            name: s.name || '',
            left: s.left ?? 0,
            top: s.top ?? 0,
            width: s.width ?? 0,
            height: s.height ?? 0,
          })),
        };
      }
    }
    return null;
  });
}

/* ================================================================
   Wipe audit-img-* images on a slide
   (leaves user's manual images alone)
================================================================ */
async function wipeAuditImages(slideNumber) {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < slideNumber) return { ok: true, deleted: 0 };

      const slide = slides.items[slideNumber - 1];
      slide.shapes.load('items');
      await context.sync();

      let deleted = 0;
      for (const s of slide.shapes.items) {
        if (!(s.name || '').startsWith(NAMESPACE)) continue;
        try { s.delete(); deleted++; } catch (e) {}
      }
      if (deleted > 0) await context.sync();
      return { ok: true, deleted };
    });
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/* ================================================================
   Reposition existing audit images to the new grid
================================================================ */
async function reflowExisting(slideNumber, cells, existingIds) {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < slideNumber) return { ok: false };

      const slide = slides.items[slideNumber - 1];
      slide.shapes.load('items');
      await context.sync();

      const audit = slide.shapes.items.filter((s) =>
        (s.name || '').startsWith(NAMESPACE)
      );

      for (let i = 0; i < audit.length && i < cells.length; i++) {
        const img = audit[i];
        const cell = cells[i];
        /* Recover aspect from current width/height */
        const curW = img.width || 1;
        const curH = img.height || 1;
        const aspect = curH / curW;
        const cellAspect = cell.h / cell.w;
        let w, h;
        if (aspect > cellAspect) {
          /* taller than cell → fit height */
          h = cell.h;
          w = h / aspect;
        } else {
          w = cell.w;
          h = w * aspect;
        }
        img.left = cell.x + (cell.w - w) / 2;
        img.top = cell.y + (cell.h - h) / 2;
        img.width = w;
        img.height = h;
      }
      await context.sync();
      return { ok: true };
    });
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/* ================================================================
   Focus a slide
================================================================ */
async function focusSlide(slideNumber) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < slideNumber) return;
      const target = slides.items[slideNumber - 1];
      context.presentation.setSelectedSlides([target.id]);
      await context.sync();
    });
  } catch (e) {}
}

/* ================================================================
   Position the newest untagged image on the target slide
================================================================ */
async function positionNewImage(slideNumber, fitted, templateKey) {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < slideNumber) return { ok: false };

      const slide = slides.items[slideNumber - 1];
      slide.shapes.load('items');
      await context.sync();

      const candidates = slide.shapes.items.filter(
        (s) =>
          s.type === PowerPoint.ShapeType.image &&
          !(s.name || '').startsWith(NAMESPACE)
      );

      if (candidates.length === 0) return { ok: false, reason: 'no new image' };

      const newImg = candidates[candidates.length - 1];

      for (let i = 0; i < candidates.length - 1; i++) {
        try { candidates[i].delete(); } catch (e) {}
      }

      try { newImg.name = NAMESPACE + templateKey; } catch (e) {}

      newImg.left = fitted.x;
      newImg.top = fitted.y;
      newImg.width = fitted.w;
      newImg.height = fitted.h;

      await context.sync();
      return { ok: true };
    });
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/* ================================================================
   Process one image
================================================================ */
async function processOneImage(dataUrl, templateKey) {
  const tpl = window.AuditTemplates.get(templateKey);
  if (!tpl) throw new Error('Template inconnu : ' + templateKey);

  const dims = await decodeImageDims(dataUrl);
  log(`Image : ${dims.w}×${dims.h} (${tpl.label})`);

  const target = await findTargetSlide();
  if (!target) {
    log('⚠️ Toutes les slides sont pleines', 'err');
    setStatus('⚠️ Slides pleines — dupliquez-en une', 'err');
    return { retry: true };
  }

  log(`Slide ${target.slideNumber} — ${target.existingCount} image(s) déjà`);

  /* Wipe old audit images on this slide (they'll be re-laid out) */
  if (target.existingCount > 0) {
    const w = await wipeAuditImages(target.slideNumber);
    if (w.ok && w.deleted > 0) log(`🗑 ${w.deleted} image(s) audit réinitialisée(s)`);
  }

  /* Compute grid for existingCount + 1 images */
  const N = (target.existingCount || 0) + 1;
  const cells = computeCells(N);
  log(`Grille : ${N} cellule(s)`);

  /* Reflow existing images to the first N-1 cells */
  if (target.existingCount > 0) {
    const reflow = await reflowExisting(target.slideNumber, cells, []);
    if (reflow.ok) log('↻ Images existantes repositionnées');
  }

  /* Compute fit for the new image in the last cell */
  const newCell = cells[N - 1];
  const fitted = containFit(newCell, dims.w, dims.h);
  log(`Cellule cible : x=${newCell.x.toFixed(0)} y=${newCell.y.toFixed(0)} w=${newCell.w.toFixed(0)} h=${newCell.h.toFixed(0)}`);
  log(`Fit : ${fitted.w.toFixed(0)}×${fitted.h.toFixed(0)}`);

  await focusSlide(target.slideNumber);

  const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
  await insertViaPaste(base64);

  await new Promise((r) => setTimeout(r, 1200));
  const pos = await positionNewImage(target.slideNumber, fitted, templateKey);

  if (pos.ok) {
    log(`✅ Image placée (slide ${target.slideNumber})`, 'ok');
    placementState.imagesPlaced++;
    setStatus(`✅ Image placée (slide ${target.slideNumber})`, 'ok');
  } else {
    log(`⚠️ Positionnement échoué`, 'err');
  }
  return { retry: false };
}

/* ================================================================
   Queue
================================================================ */
let __queue = [];
let __processing = false;
let __lastDataUrl = null;
let __lastDataUrlTs = 0;

async function processQueue() {
  if (__processing) return;
  __processing = true;

  while (__queue.length > 0) {
    const item = __queue.shift();
    try {
      const r = await processOneImage(item.dataUrl, item.templateKey);
      if (r && r.retry) {
        __queue.unshift(item);
        log(`⏳ Retry dans ${RETRY_MS}ms`);
        await new Promise((res) => setTimeout(res, RETRY_MS));
      }
    } catch (e) {
      log('❌ ' + e.message, 'err');
    }
  }

  __processing = false;
}

function enqueueImage(dataUrl, templateKey) {
  const now = Date.now();
  if (dataUrl === __lastDataUrl && now - __lastDataUrlTs < 800) {
    log('Image ignorée (doublon)');
    return;
  }
  __lastDataUrl = dataUrl;
  __lastDataUrlTs = now;

  __queue.push({ dataUrl, templateKey: templateKey || 'street' });
  log(`📥 File : ${__queue.length} image(s)`);

  if (!__processing) processQueue();
}