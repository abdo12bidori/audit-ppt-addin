/* ================================================================
   Audit Capture — Placement engine v1.8
   ================================================================
   All the image placement logic.

   v1.8 changes:
     - HEADER_TOP_FRAC raised from 0.30 to 0.45, so slides whose
       header extends past 30% (like the audit template with the
       Analyse box) are correctly detected.
     - scanSlide now classifies a shape as header if EITHER its top
       OR its bottom falls within the top zone (catches boxes that
       start above the zone but end inside it).
     - Added group-shape handling — grouped shapes (type 'group')
       are considered in the header/footer detection too, since
       Office.js does not enumerate children.
     - Tracker record moved to right after paste succeeds.
     - Fallback chain for the "existing count" now uses
       max(live, tracked, effectiveCount).
   ================================================================ */

const CFG = {
  MAX_PER_ROW: 3,
  MARGIN_FRAC: 0.04,
  GAP_FRAC: 0.015,
  BODY_TOP_PAD_FRAC: 0.02,
  BODY_BOT_PAD_FRAC: 0.02,
  HEADER_TOP_FRAC: 0.45,        /* ★ was 0.30 — catches the Analyse box */
  FOOTER_BOT_FRAC: 0.85,
  FALLBACK_HEADER_FRAC: 0.15,
  FALLBACK_FOOTER_FRAC: 0.90,
  NAMESPACE: 'audit-img-',
  CLOSED_FLAG: 'audit-closed',
};

/* Fixed slide dimensions (16:9) */
const SLIDE_W = 960;
const SLIDE_H = 540;

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
   Scan slide — merge Office.js view with local tracking
================================================================ */
function scanSlide(slide, slideW, slideH, slideIndex) {
  const shapes = slide.shapes.items || [];
  const header = [];
  const footer = [];
  const auditImages = [];
  let closed = false;

  const headerTopBound = slideH * CFG.HEADER_TOP_FRAC;
  const footerBotBound = slideH * CFG.FOOTER_BOT_FRAC;

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

    /* ★ Header detection: shape belongs to header if its TOP is
       within the top zone OR its BOTTOM is still within it (catches
       boxes that start above the zone but end inside it — like the
       Analyse box in the audit template). */
    if (top < headerTopBound || (bottom > 0 && bottom < headerTopBound)) {
      header.push({ bottom });
      continue;
    }

    /* Footer detection: bottom edge is below the footer top bound */
    if (bottom > footerBotBound) {
      footer.push({ top });
    }
  }

  /* ★ Group shapes — Office.js returns them as a single shape
     with type 'group' and does not enumerate children via
     shapes.load('items'). Use the group's own top/height. */
  try {
    for (const s of shapes) {
      if (!s || s.type !== PowerPoint.ShapeType.group) continue;
      const top = s.top ?? 0;
      const bottom = top + (s.height ?? 0);
      if (top < headerTopBound || (bottom > 0 && bottom < headerTopBound)) {
        header.push({ bottom });
      } else if (bottom > footerBotBound) {
        footer.push({ top });
      }
    }
  } catch (e) {}

  /* Sort audit images by numeric name order → stable across z-order */
  auditImages.sort((a, b) => auditImageIndex(a.name) - auditImageIndex(b.name));

  const headerBottom = header.length
    ? Math.max(...header.map((h) => h.bottom))
    : slideH * CFG.FALLBACK_HEADER_FRAC;

  const footerTop = footer.length
    ? Math.min(...footer.map((f) => f.top))
    : slideH * CFG.FALLBACK_FOOTER_FRAC;

  /* ★ Merge with the local tracking DB (helpers.js) */
  let effectiveCount = auditImages.length;
  let phantomCount = 0;
  try {
    if (window.AuditHelpers && typeof window.AuditHelpers.mergeImageViews === 'function') {
      const merged = window.AuditHelpers.mergeImageViews(auditImages, slideIndex || 0);
      effectiveCount = Math.max(auditImages.length, merged.effectiveCount);
      phantomCount = merged.phantomCount;
      console.log(
        `[AuditCapture:helpers] slide ${slideIndex} — live=${merged.liveCount} tracked=${merged.trackedCount} phantom=${phantomCount} → effective=${effectiveCount}`
      );
    } else {
      console.warn('[AuditCapture:helpers] AuditHelpers not loaded — using live count only');
    }
  } catch (e) {
    console.warn('[AuditCapture:helpers] merge failed:', e.message);
  }

  return { headerBottom, footerTop, auditImages, closed, effectiveCount, phantomCount };
}

/* ================================================================
   Compute grid
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
   fitContain
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
   insertViaPaste — clipboard write + CDP paste + retry
================================================================ */
async function insertViaPaste(base64) {
  const dataUrl = 'data:image/png;base64,' + base64;

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

  await requestClipboardWrite('1/4 write');
  await requestPaste('2/4 paste');
  await new Promise((r) => setTimeout(r, 900));
  await requestClipboardWrite('3/4 rewrite');
  await requestPaste('4/4 repaste');
  await new Promise((r) => setTimeout(r, 900));

  log('✅ Demande de collage terminée');
  return true;
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

  /* Dedupe via helpers */
  try {
    if (window.AuditHelpers && typeof window.AuditHelpers.wasImageRecentlyPlaced === 'function') {
      if (window.AuditHelpers.wasImageRecentlyPlaced(dataUrl)) {
        log('⚠️ Image déjà placée récemment — ignorée');
        return;
      }
      window.AuditHelpers.markImageAsRecent(dataUrl);
    }
  } catch (e) {}

  const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
  let totalAfter = 0;
  let slotForNew = null;
  let currentSlideIndex = 0;

  /* ─── STEP 1: scan + reposition ─── */
  try {
    await Promise.race([
      PowerPoint.run(async (context) => {
        const slideW = SLIDE_W;
        const slideH = SLIDE_H;

        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        if (slides.items.length === 0) throw new Error('No slides in presentation');

        let target = slides.items[slides.items.length - 1];
        target.shapes.load('items');
        await context.sync();

        if (!placementState.masterWarningShown && slides.items.length === 1) {
          if (target.shapes.items.length <= 2) {
            log('⚠️ Le masque ne contient pas d\'en-tête.', 'err');
            placementState.masterWarningShown = true;
          }
        }

        let scan = scanSlide(target, slideW, slideH, slides.items.length);
        log(`Slide ${slides.items.length} — ${scan.auditImages.length} image(s) live, ${scan.effectiveCount} effective, header→${scan.headerBottom.toFixed(0)}, footer→${scan.footerTop.toFixed(0)}, closed=${scan.closed}`);

        const needsNewSlide = scan.effectiveCount >= CFG.MAX_PER_ROW || scan.closed;
        if (needsNewSlide) {
          const reason = scan.closed ? 'clôturée (END)' : `pleine (${scan.effectiveCount}/${CFG.MAX_PER_ROW})`;
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
          scan = scanSlide(target, slideW, slideH, slides.items.length);
          log(`Nouvelle slide : ${slides.items.length}`);
        }

        /* Fallback chain for existing count */
        let effectiveExisting = scan.auditImages.length;
        try {
          if (window.AuditHelpers && typeof window.AuditHelpers.getTrackedImages === 'function') {
            const tracked = window.AuditHelpers.getTrackedImages(slides.items.length);
            effectiveExisting = Math.max(effectiveExisting, tracked.length, scan.effectiveCount || 0);
          }
        } catch (e) {}
        totalAfter = effectiveExisting + 1;
        currentSlideIndex = slides.items.length;
        log(`Placement : ${effectiveExisting} image(s) détectée(s) → slot ${totalAfter}`);

        const slots = computeGrid(scan, totalAfter, slideW, slideH);
        slotForNew = slots[totalAfter - 1];
        log(`Grille : ${totalAfter} slot(s), largeur ${slots[0].w.toFixed(0)}, hauteur ${slots[0].h.toFixed(0)}`);

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
      }),
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error('scan timed out (6s)')), 6000)
      ),
    ]);
  } catch (e) {
    log('⚠️ Scan échoué : ' + e.message + ' — poursuite du collage', 'err');
  }

  /* ─── STEP 2: paste ─── */
  log('→ Insertion via collage (CDP)');
  const inserted = await insertViaPaste(base64);
  if (!inserted) {
    throw new Error('Insertion par collage échouée');
  }

  /* ★ Record in the local tracker IMMEDIATELY after paste succeeds —
       even if positioning later fails, the next send will know this
       image exists. */
  try {
    if (window.AuditHelpers && typeof window.AuditHelpers.recordPlacedImage === 'function') {
      window.AuditHelpers.recordPlacedImage(
        currentSlideIndex || 0,
        `${CFG.NAMESPACE}${totalAfter}`,
        { slotIndex: totalAfter }
      );
      log(`★ Enregistré dans le tracker : audit-img-${totalAfter} sur slide ${currentSlideIndex}`);
    } else {
      log('⚠️ AuditHelpers non chargé — pas de tracking', 'err');
    }
  } catch (e) {
    log('⚠️ Tracker error : ' + e.message, 'err');
  }

  /* ─── STEP 3: name & position — 2 attempts × 6 s ─── */
  await new Promise((r) => setTimeout(r, 1500));

  const localIndex = totalAfter;
  const localSlot = slotForNew;
  const localDims = dims;
  const localSlideIndex = currentSlideIndex;
  const localMode = mode;

  async function tryPositioning(timeoutMs) {
    return await Promise.race([
      PowerPoint.run(async (context) => {
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
          return { ok: false, reason: 'no candidate image' };
        }

        const newImg = candidates[candidates.length - 1];
        const index = localIndex;
        try { newImg.name = `${CFG.NAMESPACE}${index}`; } catch (e) {}

        const fitted = fitContain(localSlot, localDims.w, localDims.h);
        newImg.left = fitted.x;
        newImg.top = fitted.y;
        newImg.width = fitted.w;
        newImg.height = fitted.h;

        await context.sync();

        if (localMode === 'end') {
          try {
            const marker = lastSlide.shapes.addTextBox(' ');
            marker.name = `${CFG.CLOSED_FLAG}-${Date.now()}`;
            marker.left = -100;
            marker.top = -100;
            marker.width = 1;
            marker.height = 1;
            await context.sync();
          } catch (e) {}
        }

        return { ok: true };
      }),
      new Promise((resolve) =>
        setTimeout(() => resolve({ ok: false, reason: `timeout ${timeoutMs}ms` }), timeoutMs)
      ),
    ]);
  }

  let pos = await tryPositioning(6000);
  if (pos.ok) {
    log(`✅ Image ${localIndex} placée — slot ${localIndex}/${CFG.MAX_PER_ROW}`, 'ok');
  } else {
    log(`⚠️ Positionnement 1 échoué : ${pos.reason} — retry`, 'err');
    await new Promise((r) => setTimeout(r, 500));
    pos = await tryPositioning(6000);
    if (pos.ok) {
      log(`✅ Image ${localIndex} placée (retry) — slot ${localIndex}/${CFG.MAX_PER_ROW}`, 'ok');
    } else {
      log(`⚠️ Positionnement 2 échoué : ${pos.reason} — l'image reste visible mais non taguée`, 'err');
    }
  }

  placementState.imagesPlaced++;
  setStatus(`✅ Image placée (${localIndex}/${CFG.MAX_PER_ROW})${localMode === 'end' ? ' — clôturée' : ''}`, 'ok');
}