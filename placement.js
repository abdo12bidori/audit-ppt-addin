/* ================================================================
   Audit Capture — Placement engine v1.4
   ================================================================
   All the image placement logic:
   - config (fractions, no absolute measurements)
   - decode image dims
   - scan slide, detect header/footer zones
   - compute grid, contain-fit
   - insert image via paste (CDP) with retry-after-focus
   - name & position

   v1.4: Chrome refuses `navigator.clipboard.write` when the document
         isn't focused. The taskpane runs in an iframe, so writing to
         the clipboard from there usually fails on the first try. The
         extension's CDP click (which focuses the PPT tab) happens
         AFTER the first write. So we now:
           1. Try to write the clipboard (likely fails silently)
           2. Ask the extension to fire CDP paste (focuses PPT tab)
           3. RETRY the clipboard write — this time the tab is focused
           4. Ask for a second CDP paste
           5. Poll for the shape
   ================================================================ */

/* ----------------------------------------------------------------
   Config — all fractions of the slide size
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
  CLOSED_FLAG: 'audit-closed',
};

/* ----------------------------------------------------------------
   Shared state (defined in taskpane.js, referenced here)
---------------------------------------------------------------- */
const placementState = {
  imagesPlaced: 0,
  masterWarningShown: false,
};

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
   Parse "audit-img-N" → N
================================================================ */
function auditImageIndex(name) {
  if (!name || !name.startsWith(CFG.NAMESPACE)) return Infinity;
  const n = parseInt(name.slice(CFG.NAMESPACE.length), 10);
  return Number.isFinite(n) ? n : Infinity;
}

/* ================================================================
   Scan slide — collect audit images, detect header/footer
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

  /* Sort by numeric part of name → stable order regardless of z-order */
  auditImages.sort((a, b) => auditImageIndex(a.name) - auditImageIndex(b.name));

  const headerBottom = header.length
    ? Math.max(...header.map((h) => h.bottom))
    : slideH * CFG.FALLBACK_HEADER_FRAC;

  const footerTop = footer.length
    ? Math.min(...footer.map((f) => f.top))
    : slideH * CFG.FALLBACK_FOOTER_FRAC;

  return { headerBottom, footerTop, auditImages, closed };
}

/* ================================================================
   Compute grid — dynamic N-way split
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
   fitContain — adapt with NO distortion, NO quality loss
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
   insertViaPaste — clipboard write + CDP paste, with retry
   ----------------------------------------------------------------
   The taskpane's clipboard write fails when the document isn't
   focused. The CDP click inside the PPT tab focuses it. So:
     1. Try clipboard write (may fail)
     2. Fire CDP paste (focuses the tab)
     3. Retry clipboard write (should succeed now)
     4. Fire CDP paste again (uses the fresh clipboard)
     5. Poll for the shape
================================================================ */
async function insertViaPaste(base64) {
  const dataUrl = 'data:image/png;base64,' + base64;

  /* Helper — ask the content script to write the clipboard.
     We don't wait for a reply (cross-frame replies are unreliable
     in this setup). Just fire and give it time. */
  async function requestClipboardWrite(label) {
    try {
      window.parent.postMessage({
        type: 'AUDIT_WRITE_CLIPBOARD',
        dataUrl,
        ts: Date.now(),
      }, '*');
      log(`→ [${label}] Demande d'écriture presse-papier envoyée`);
      await new Promise((r) => setTimeout(r, 500));
      return true;
    } catch (e) {
      log(`⚠️ [${label}] postMessage failed: ${e.message}`, 'err');
      return false;
    }
  }

  /* Helper — ask the extension to fire CDP paste */
  async function requestPaste(label) {
    try {
      window.parent.postMessage({
        type: 'AUDIT_PASTE_REQUEST',
        ts: Date.now(),
      }, '*');
      log(`→ [${label}] Demande de collage envoyée`);
      return true;
    } catch (e) {
      log(`⚠️ [${label}] postMessage failed: ${e.message}`, 'err');
      return false;
    }
  }

  /* ─── Round 1 ─── */
  await requestClipboardWrite('1/4 write');
  await requestPaste('2/4 paste');

  /* Wait for the CDP click to focus the PPT tab */
  await new Promise((r) => setTimeout(r, 900));

  /* ─── Round 2 — now the tab is focused ─── */
  await requestClipboardWrite('3/4 rewrite');
  await requestPaste('4/4 repaste');

  /* ─── Poll for the shape ─── */
  const start = Date.now();
  while (Date.now() - start < 5000) {
    await new Promise((r) => setTimeout(r, 300));
    const found = await didInsertWork();
    if (found) {
      log('✅ Collage détecté');
      return true;
    }
  }
  return false;
}

/* ================================================================
   Did addImage/paste actually insert a shape?
================================================================ */
async function didInsertWork() {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const last = slides.items[slides.items.length - 1];
      last.shapes.load('items');
      await context.sync();
      return last.shapes.items.some(
        (s) =>
          s.type === PowerPoint.ShapeType.image &&
          !(s.name || '').startsWith(CFG.NAMESPACE) &&
          !(s.name || '').startsWith(CFG.CLOSED_FLAG)
      );
    });
  } catch (e) {
    return false;
  }
}

/* ================================================================
   Core — runPlacement
================================================================ */
async function runPlacement(dataUrl, mode) {
  let dims;
  try {
    dims = await decodeImageDims(dataUrl);
  } catch (e) {
    log('⚠️ Cannot decode image dims — fallback 4:3', 'err');
    dims = { w: 800, h: 600 };
  }
  log(`Image : ${dims.w}×${dims.h}, mode=${mode}`);

  const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
  let totalAfter = 0;
  let slotForNew = null;

  await PowerPoint.run(async (context) => {
    const pageSetup = context.presentation.pageSetup;
    pageSetup.load(['slideWidth', 'slideHeight']);
    await context.sync();
    const slideW = pageSetup.slideWidth;
    const slideH = pageSetup.slideHeight;

    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    if (slides.items.length === 0) throw new Error('No slides in presentation');

    let target = slides.items[slides.items.length - 1];
    target.shapes.load('items');
    await context.sync();

    /* Master warning (once) */
    if (!placementState.masterWarningShown && slides.items.length === 1) {
      if (target.shapes.items.length <= 2) {
        log('⚠️ Le masque ne contient pas d\'en-tête.', 'err');
        placementState.masterWarningShown = true;
      }
    }

    /* Scan */
    let scan = scanSlide(target, slideW, slideH);
    log(`Slide ${slides.items.length} — ${scan.auditImages.length} image(s), header→${scan.headerBottom.toFixed(0)}, footer→${scan.footerTop.toFixed(0)}, closed=${scan.closed}`);

    /* New slide when full or closed */
    const needsNewSlide = scan.auditImages.length >= CFG.MAX_PER_ROW || scan.closed;
    if (needsNewSlide) {
      const reason = scan.closed ? 'clôturée (END)' : `pleine (${scan.auditImages.length}/${CFG.MAX_PER_ROW})`;
      log(`Slide ${reason} → nouvelle slide`);

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
      log(`Nouvelle slide : ${slides.items.length}`);
    }

    totalAfter = scan.auditImages.length + 1;
    const slots = computeGrid(scan, totalAfter, slideW, slideH);
    slotForNew = slots[totalAfter - 1];
    log(`Grille : ${totalAfter} slot(s), largeur ${slots[0].w.toFixed(0)}`);

    /* Reposition all existing audit images (overwrites manual edits) */
    for (let i = 0; i < scan.auditImages.length; i++) {
      const img = scan.auditImages[i];
      const slot = slots[i];
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
  });

  /* Insert the new image via paste (CDP) */
  log('→ Insertion via collage (CDP)');
  const inserted = await insertViaPaste(base64);
  if (!inserted) {
    throw new Error('Insertion par collage échouée');
  }

  /* Wait for shape */
  await new Promise((r) => setTimeout(r, 300));

  /* Final pass: name & position */
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    const lastSlide = slides.items[slides.items.length - 1];
    lastSlide.shapes.load('items');
    await context.sync();

    const candidates = lastSlide.shapes.items.filter(
      (s) =>
        s.type === PowerPoint.ShapeType.image &&
        !(s.name || '').startsWith(CFG.NAMESPACE) &&
        !(s.name || '').startsWith(CFG.CLOSED_FLAG)
    );

    if (candidates.length === 0) {
      log('⚠️ Nouvelle image introuvable après insertion', 'err');
      return;
    }

    const newImg = candidates[candidates.length - 1];
    const index = totalAfter;
    try { newImg.name = `${CFG.NAMESPACE}${index}`; } catch (e) {}

    const fitted = fitContain(slotForNew, dims.w, dims.h);
    newImg.left = fitted.x;
    newImg.top = fitted.y;
    newImg.width = fitted.w;
    newImg.height = fitted.h;

    await context.sync();
    log(`✅ Image ${index} placée — slide ${slides.items.length}, slot ${totalAfter}/${CFG.MAX_PER_ROW}`, 'ok');

    /* END → mark slide closed */
    if (mode === 'end') {
      try {
        const marker = lastSlide.shapes.addTextBox(' ');
        marker.name = `${CFG.CLOSED_FLAG}-${Date.now()}`;
        marker.left = -100;
        marker.top = -100;
        marker.width = 1;
        marker.height = 1;
        await context.sync();
        log('🔒 Slide clôturée (END)');
      } catch (e) {
        log('⚠️ Impossible de clôturer : ' + e.message, 'err');
      }
    }
  });

  placementState.imagesPlaced++;
  setStatus(`✅ Image placée (${totalAfter}/${CFG.MAX_PER_ROW})${mode === 'end' ? ' — clôturée' : ''}`, 'ok');
}
