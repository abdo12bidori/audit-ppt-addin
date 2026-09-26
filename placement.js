/* ================================================================
   Placement engine v3.0
   ================================================================
   - Queue: one image at a time
   - Auto-detect variant from aspect ratio (for logging)
   - Scan slides for a free slot (left or right)
   - Wipe audit-img-* copies on duplicated slides
   - Contain-fit inside slot (no distortion, no enlargement)
   - Retry every 3s when all slides full
   ================================================================ */

const LAYOUT = {
  SLIDE_W: 960,
  SLIDE_H: 540,
  SLOT_LEFT:  { x: 18,  y: 120, w: 376, h: 315 },
  SLOT_RIGHT: { x: 401, y: 174, w: 338, h: 261 },
  NAMESPACE: 'audit-img-',
  MAX_PER_SLIDE: 2,
  RETRY_MS: 3000,
};

const placementState = {
  imagesPlaced: 0,
  masterWarningShown: false,
};

/* ================================================================
   Image decode
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
function containFit(slot, imgW, imgH) {
  let scale = Math.min(slot.w / imgW, slot.h / imgH);
  if (scale > 1) scale = 1;              /* never enlarge */
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
   insertViaPaste — clipboard write + CDP Ctrl+V
================================================================ */
async function insertViaPaste(base64) {
  const dataUrl = 'data:image/png;base64,' + base64;

  async function writeClipboard(label) {
    try {
      window.parent.postMessage({ type: 'AUDIT_WRITE_CLIPBOARD', dataUrl, ts: Date.now() }, '*');
      log(`→ [${label}] Demande d'écriture presse-papier`);
      await new Promise((r) => setTimeout(r, 500));
      return true;
    } catch (e) {
      log(`⚠️ [${label}] ${e.message}`, 'err');
      return false;
    }
  }

  async function pasteRequest(label) {
    try {
      window.parent.postMessage({ type: 'AUDIT_PASTE_REQUEST', ts: Date.now() }, '*');
      log(`→ [${label}] Demande de collage`);
      return true;
    } catch (e) {
      log(`⚠️ [${label}] ${e.message}`, 'err');
      return false;
    }
  }

  await writeClipboard('1/4 write');
  await pasteRequest('2/4 paste');
  await new Promise((r) => setTimeout(r, 900));
  await writeClipboard('3/4 rewrite');
  await pasteRequest('4/4 repaste');
  await new Promise((r) => setTimeout(r, 900));

  log('✅ Collage demandé');
  return true;
}

/* ================================================================
   scanSlides — find the first free slot across all slides
   ----------------------------------------------------------------
   Returns { slideNumber, slot, rect, isDuplicated } or null.
   'isDuplicated' means: this slide has audit-img-* copies from
   the previous slide — they must be wiped before placing.
================================================================ */
async function scanSlides() {
  return await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    /* Load all shapes for all slides */
    for (const slide of slides.items) {
      slide.shapes.load('items');
    }
    await context.sync();

    /* Get audit image names per slide */
    const slideAuditNames = [];
    for (const slide of slides.items) {
      const names = slide.shapes.items
        .filter((s) => (s.name || '').startsWith(LAYOUT.NAMESPACE))
        .map((s) => s.name);
      slideAuditNames.push(names);
    }

    /* Scan slides in order, find first with a free slot */
    for (let i = 0; i < slides.items.length; i++) {
      const names = slideAuditNames[i];
      const prevNames = i > 0 ? slideAuditNames[i - 1] : [];

      /* Detect duplicate: current slide has ≥1 audit name that exists on previous slide */
      const isDuplicated =
        names.length > 0 &&
        prevNames.some((n) => names.includes(n));

      if (isDuplicated) {
        /* Wipe the copies — the slide will be usable */
        return {
          slideNumber: i + 1,
          slot: 'left',
          rect: LAYOUT.SLOT_LEFT,
          isDuplicated: true,
        };
      }

      /* Count images overlapping left slot */
      const leftTaken = imagesOverlapSlot(slides.items[i].shapes.items, LAYOUT.SLOT_LEFT);
      if (!leftTaken) {
        return {
          slideNumber: i + 1,
          slot: 'left',
          rect: LAYOUT.SLOT_LEFT,
          isDuplicated: false,
        };
      }

      /* Count images overlapping right slot */
      const rightTaken = imagesOverlapSlot(slides.items[i].shapes.items, LAYOUT.SLOT_RIGHT);
      if (!rightTaken) {
        return {
          slideNumber: i + 1,
          slot: 'right',
          rect: LAYOUT.SLOT_RIGHT,
          isDuplicated: false,
        };
      }
    }

    return null;   /* all slides full */
  });
}

function imagesOverlapSlot(shapes, slot) {
  const PAD = 5;
  const sx1 = slot.x - PAD;
  const sy1 = slot.y - PAD;
  const sx2 = slot.x + slot.w + PAD;
  const sy2 = slot.y + slot.h + PAD;

  for (const s of shapes) {
    if (!s || s.type !== PowerPoint.ShapeType.image) continue;
    const l = s.left ?? 0;
    const t = s.top ?? 0;
    const r = l + (s.width ?? 0);
    const b = t + (s.height ?? 0);
    if (!(r < sx1 || l > sx2 || b < sy1 || t > sy2)) return true;
  }
  return false;
}

/* ================================================================
   wipeAuditImages — delete all audit-img-* from a slide
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
        if (!s || !(s.name || '').startsWith(LAYOUT.NAMESPACE)) continue;
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
   positionNewImage — find the newest image on a slide and move it
================================================================ */
async function positionNewImage(slideNumber, fitted, templateKey) {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < slideNumber) return { ok: false, reason: 'slide missing' };

      const slide = slides.items[slideNumber - 1];
      slide.shapes.load('items');
      await context.sync();

      const candidates = slide.shapes.items.filter(
        (s) =>
          s.type === PowerPoint.ShapeType.image &&
          !(s.name || '').startsWith(LAYOUT.NAMESPACE)
      );

      if (candidates.length === 0) return { ok: false, reason: 'no new image' };

      const newImg = candidates[candidates.length - 1];

      /* Delete any other untagged candidates (defensive) */
      for (let i = 0; i < candidates.length - 1; i++) {
        try { candidates[i].delete(); } catch (e) {}
      }

      try { newImg.name = LAYOUT.NAMESPACE + templateKey; } catch (e) {}

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
   focusSlide
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
    return true;
  } catch (e) {
    return false;
  }
}

/* ================================================================
   Process one image (used by the queue)
================================================================ */
async function processOneImage(dataUrl, templateKey) {
  const tpl = window.AuditTemplates.get(templateKey);
  if (!tpl) throw new Error('Template inconnu : ' + templateKey);

  /* 1. Decode */
  const dims = await decodeImageDims(dataUrl);
  const aspect = dims.w / dims.h;

  /* 2. Detect variant (for logging) */
  const variant = window.AuditTemplates.detectVariant(templateKey, dims.w, dims.h);
  log(`Catégorie : ${tpl.label} | Variante détectée : ${variant ? variant.label : '—'} (aspect ${aspect.toFixed(2)})`);

  /* 3. Find free slot */
  const slot = await scanSlides();
  if (!slot) {
    log('⚠️ Toutes les diapositives sont pleines', 'err');
    setStatus('⚠️ Diapositives pleines — dupliquez-en une', 'err');
    notify('⚠️ Toutes les diapositives sont pleines — dupliquez-en une');
    return { retry: true };
  }

  log(`Slot trouvé : slide ${slot.slideNumber} (${slot.slot}), duplicated=${slot.isDuplicated}`);

  /* 4. If duplicated slide — wipe copies */
  if (slot.isDuplicated) {
    const w = await wipeAuditImages(slot.slideNumber);
    if (w.ok && w.deleted > 0) log(`🗑 ${w.deleted} copie(s) supprimée(s) sur slide ${slot.slideNumber}`);
  }

  /* 5. Compute fit */
  const fitted = containFit(slot.rect, dims.w, dims.h);
  log(`Placement : x=${fitted.x.toFixed(0)} y=${fitted.y.toFixed(0)} w=${fitted.w.toFixed(0)} h=${fitted.h.toFixed(0)}`);

  /* 6. Focus slide */
  await focusSlide(slot.slideNumber);

  /* 7. Paste via CDP */
  const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
  const inserted = await insertViaPaste(base64);
  if (!inserted) throw new Error('Insertion échouée');

  /* 8. Position */
  await new Promise((r) => setTimeout(r, 1200));
  const pos = await positionNewImage(slot.slideNumber, fitted, templateKey);
  if (pos.ok) {
    log(`✅ Image placée sur slide ${slot.slideNumber} (${slot.slot})`, 'ok');
    placementState.imagesPlaced++;
    setStatus(`✅ Image placée (slide ${slot.slideNumber}, ${slot.slot})`, 'ok');
    return { retry: false };
  } else {
    log(`⚠️ Positionnement échoué : ${pos.reason}`, 'err');
    return { retry: false };
  }
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
      const result = await processOneImage(item.dataUrl, item.templateKey);
      if (result && result.retry) {
        /* Put it back at the head — retry later */
        __queue.unshift(item);
        log(`⏳ En attente — retry dans ${LAYOUT.RETRY_MS}ms`);
        await new Promise((r) => setTimeout(r, LAYOUT.RETRY_MS));
      }
    } catch (e) {
      log('❌ Erreur : ' + e.message, 'err');
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
  log(`📥 File d'attente : ${__queue.length} image(s)`);

  if (!__processing) processQueue();
}