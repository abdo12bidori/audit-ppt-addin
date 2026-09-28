/* ================================================================
   Placement engine v4.0
   ================================================================
   - Placeholders are shapes named IMG_SLOT_* on the slide
     (drawn by the PPT author in the template)
   - Add-in finds the first empty placeholder across all slides
   - Deletes any existing image inside that placeholder
   - Contain-fit inside the placeholder (no overflow)
   - Queue serializes processing
   - Retry every 3s when all placeholders are full
================================================================ */

const PLACEHOLDER_PREFIX = 'IMG_SLOT';
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
   Contain-fit — inside a slot rect
================================================================ */
function containFit(slot, imgW, imgH) {
  let scale = Math.min(slot.w / imgW, slot.h / imgH);
  if (scale > 1) scale = 1;   /* never enlarge */
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
   Scan slides for the first empty placeholder
   ----------------------------------------------------------------
   A placeholder is a shape whose name starts with 'IMG_SLOT'.
   "Empty" means no audit-img-* image overlaps its rectangle.
   Returns { slideNumber, rect, placeholderName } or null.
================================================================ */
async function findFirstEmptyPlaceholder() {
  return await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      slide.shapes.load('items');
      await context.sync();

      const shapes = slide.shapes.items;
      const placeholders = shapes.filter(
        (s) => (s.name || '').startsWith(PLACEHOLDER_PREFIX)
      );
      const images = shapes.filter((s) => s.type === PowerPoint.ShapeType.image);

      for (const ph of placeholders) {
        const rect = {
          x: ph.left ?? 0,
          y: ph.top ?? 0,
          w: ph.width ?? 0,
          h: ph.height ?? 0,
        };
        if (rect.w < 50 || rect.h < 50) continue;

        const overlap = imagesOverlapsRect(images, rect);
        if (!overlap) {
          return {
            slideNumber: i + 1,
            rect,
            placeholderName: ph.name,
          };
        }
      }
    }
    return null;
  });
}

function imagesOverlapsRect(images, rect) {
  const PAD = 5;
  const rx1 = rect.x - PAD;
  const ry1 = rect.y - PAD;
  const rx2 = rect.x + rect.w + PAD;
  const ry2 = rect.y + rect.h + PAD;

  for (const img of images) {
    const il = img.left ?? 0;
    const it = img.top ?? 0;
    const ir = il + (img.width ?? 0);
    const ib = it + (img.height ?? 0);
    if (!(ir < rx1 || il > rx2 || ib < ry1 || it > ry2)) return true;
  }
  return false;
}

/* ================================================================
   Wipe any existing image inside the target placeholder
================================================================ */
async function wipeImagesInRect(slideNumber, rect) {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < slideNumber) return { ok: true, deleted: 0 };

      const slide = slides.items[slideNumber - 1];
      slide.shapes.load('items');
      await context.sync();

      const PAD = 5;
      const rx1 = rect.x - PAD;
      const ry1 = rect.y - PAD;
      const rx2 = rect.x + rect.w + PAD;
      const ry2 = rect.y + rect.h + PAD;

      let deleted = 0;
      for (const s of slide.shapes.items) {
        if (!s || s.type !== PowerPoint.ShapeType.image) continue;
        const il = s.left ?? 0;
        const it = s.top ?? 0;
        const ir = il + (s.width ?? 0);
        const ib = it + (s.height ?? 0);
        if (!(ir < rx1 || il > rx2 || ib < ry1 || it > ry2)) {
          try { s.delete(); deleted++; } catch (e) {}
        }
      }
      if (deleted > 0) await context.sync();
      return { ok: true, deleted };
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

  const slot = await findFirstEmptyPlaceholder();
  if (!slot) {
    log('⚠️ Toutes les zones image sont pleines', 'err');
    setStatus('⚠️ Zones pleines — dupliquez une slide', 'err');
    return { retry: true };
  }

  log(`Zone libre : slide ${slot.slideNumber}, ${slot.placeholderName}`);

  const w = await wipeImagesInRect(slot.slideNumber, slot.rect);
  if (w.ok && w.deleted > 0) log(`🗑 ${w.deleted} image(s) supprimée(s)`);

  const fitted = containFit(slot.rect, dims.w, dims.h);
  log(`Fit : ${fitted.w.toFixed(0)}×${fitted.h.toFixed(0)} @ (${fitted.x.toFixed(0)},${fitted.y.toFixed(0)})`);

  await focusSlide(slot.slideNumber);

  const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
  await insertViaPaste(base64);

  await new Promise((r) => setTimeout(r, 1200));
  const pos = await positionNewImage(slot.slideNumber, fitted, templateKey);

  if (pos.ok) {
    log(`✅ Image placée (slide ${slot.slideNumber})`, 'ok');
    placementState.imagesPlaced++;
    setStatus(`✅ Image placée (slide ${slot.slideNumber})`, 'ok');
  } else {
    log(`⚠️ Positionnement échoué`, 'err');
  }
  return { retry: false };
}

/* ================================================================
   Queue — serializes image processing
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