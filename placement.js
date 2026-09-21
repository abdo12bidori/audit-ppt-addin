/* ================================================================
   Audit Capture — Placement engine v2.4
   ================================================================
   v2.4:
     - ensureSlideExists uses slide 1's layout for new slides, so
       slide 2 inherits the same template (header/footer) as slide 1.
   ================================================================ */

const CFG = {
  SLIDE_W: 960,
  SLIDE_H: 540,
  NAMESPACE: 'audit-img-',
};

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
   fitContain
================================================================ */
function fitContain(rect, imgW, imgH) {
  const scale = Math.min(rect.w / imgW, rect.h / imgH);
  const w = imgW * scale;
  const h = imgH * scale;
  return {
    x: rect.x + (rect.w - w) / 2,
    y: rect.y + (rect.h - h) / 2,
    w,
    h,
  };
}

/* ================================================================
   insertViaPaste
================================================================ */
async function insertViaPaste(base64) {
  const dataUrl = 'data:image/png;base64,' + base64;

  async function requestClipboardWrite(label) {
    try {
      window.parent.postMessage({ type: 'AUDIT_WRITE_CLIPBOARD', dataUrl, ts: Date.now() }, '*');
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
      window.parent.postMessage({ type: 'AUDIT_PASTE_REQUEST', ts: Date.now() }, '*');
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
   ensureSlideExists
   ----------------------------------------------------------------
   Uses SLIDE 1's layout for any new slides. That way slide 2 (and
   any further slide) inherits the same template as slide 1.
================================================================ */
async function ensureSlideExists(targetSlideNumber) {
  try {
    return await Promise.race([
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        if (slides.items.length === 0) {
          return { ok: false, error: 'no slides at all' };
        }

        let added = 0;
        while (slides.items.length < targetSlideNumber) {
          /* Use slide 1 as the layout template */
          const templateSlide = slides.items[0];
          templateSlide.layout.load('id');
          await context.sync();
          const layoutId = templateSlide.layout.id;

          slides.add({ layoutId });
          await context.sync();
          slides.load('items');
          await context.sync();
          added++;
        }
        return { ok: true, added, total: slides.items.length };
      }),
      new Promise((resolve) =>
        setTimeout(() => resolve({ ok: false, error: 'ensure timeout (4s)' }), 4000)
      ),
    ]);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/* ================================================================
   deleteOverlappingImages
================================================================ */
async function deleteOverlappingImages(rect, slideNumber) {
  try {
    return await Promise.race([
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        if (slides.items.length < slideNumber) return { ok: true, deleted: 0 };

        const slide = slides.items[slideNumber - 1];
        slide.shapes.load('items');
        await context.sync();

        let deleted = 0;
        const PAD = 20;

        const tx1 = rect.x - PAD;
        const ty1 = rect.y - PAD;
        const tx2 = rect.x + rect.w + PAD;
        const ty2 = rect.y + rect.h + PAD;

        for (const s of slide.shapes.items) {
          if (!s || s.type !== PowerPoint.ShapeType.image) continue;

          const left = s.left ?? 0;
          const top = s.top ?? 0;
          const right = left + (s.width ?? 0);
          const bottom = top + (s.height ?? 0);

          const overlaps = !(right < tx1 || left > tx2 || bottom < ty1 || top > ty2);

          if (overlaps) {
            try {
              s.delete();
              deleted++;
            } catch (e) {}
          }
        }
        if (deleted > 0) await context.sync();
        return { ok: true, deleted };
      }),
      new Promise((resolve) =>
        setTimeout(() => resolve({ ok: false, error: 'delete timeout (4s)' }), 4000)
      ),
    ]);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/* ================================================================
   snapshotImages
================================================================ */
async function snapshotImages(slideNumber) {
  try {
    return await Promise.race([
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        if (slides.items.length < slideNumber) return { ok: true, ids: [] };

        const slide = slides.items[slideNumber - 1];
        slide.shapes.load('items');
        await context.sync();

        const ids = slide.shapes.items
          .filter((s) => s.type === PowerPoint.ShapeType.image)
          .map((s) => s.id);
        return { ok: true, ids };
      }),
      new Promise((resolve) =>
        setTimeout(() => resolve({ ok: false, error: 'snapshot timeout (3s)' }), 3000)
      ),
    ]);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/* ================================================================
   Core — runPlacement
================================================================ */
async function runPlacement(dataUrl, mode, templateKey) {
  const templates = window.AuditTemplates;
  if (!templates || typeof templates.get !== 'function') {
    throw new Error('AuditTemplates non chargé');
  }

  const tpl = templates.get(templateKey);
  if (!tpl) throw new Error('Template inconnu : ' + templateKey);

  log(`Template : ${tpl.label} → slide ${tpl.slide}, slot ${tpl.slot}`);

  let dims;
  try {
    dims = await decodeImageDims(dataUrl);
  } catch (e) {
    dims = { w: 800, h: 600 };
  }
  log(`Image : ${dims.w}×${dims.h}, mode=${mode}`);

  const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');

  /* Ensure the target slide exists */
  const ensure = await ensureSlideExists(tpl.slide);
  if (ensure.ok) {
    log(`Slides : ${ensure.total}${ensure.added ? ' (' + ensure.added + ' ajoutée(s))' : ''}`);
  }

  /* Delete overlapping images in the target slot */
  const del = await deleteOverlappingImages(tpl.rect, tpl.slide);
  if (del.ok && del.deleted > 0) {
    log(`🗑 ${del.deleted} image(s) supprimée(s) dans la zone cible`);
  } else if (del.ok) {
    log(`Zone cible vide`);
  } else {
    log(`⚠️ delete: ${del.error}`, 'err');
  }

  /* Snapshot remaining image IDs (for ID-diff detection) */
  const snap = await snapshotImages(tpl.slide);
  const beforeIds = new Set(snap.ok ? snap.ids : []);
  log(`Snapshot : ${beforeIds.size} image(s) existante(s)`);

  /* Focus the target slide */
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < tpl.slide) return;
      const target = slides.items[tpl.slide - 1];
      context.presentation.setSelectedSlides([target.id]);
      await context.sync();
    });
    log(`Slide ${tpl.slide} sélectionnée`);
  } catch (e) {
    log('⚠️ Focus slide échoué : ' + e.message, 'err');
  }

  /* Paste via CDP */
  log('→ Insertion via collage (CDP)');
  const inserted = await insertViaPaste(base64);
  if (!inserted) throw new Error('Insertion par collage échouée');

  await new Promise((r) => setTimeout(r, 1200));

  const fitted = fitContain(tpl.rect, dims.w, dims.h);
  log(`Position cible : x=${fitted.x.toFixed(0)} y=${fitted.y.toFixed(0)} w=${fitted.w.toFixed(0)} h=${fitted.h.toFixed(0)}`);

  async function tryPositioning(timeoutMs) {
    return await Promise.race([
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        if (slides.items.length < tpl.slide) return { ok: false, reason: 'slide missing' };

        const target = slides.items[tpl.slide - 1];
        target.shapes.load('items');
        await context.sync();

        const allImages = target.shapes.items.filter(
          (s) => s.type === PowerPoint.ShapeType.image
        );
        const newImages = allImages.filter((s) => !beforeIds.has(s.id));

        if (newImages.length === 0) {
          return { ok: false, reason: 'no new image found' };
        }

        const newImg = newImages[newImages.length - 1];

        /* Delete extra new images (defensive) */
        for (let i = 0; i < newImages.length - 1; i++) {
          try { newImages[i].delete(); } catch (e) {}
        }

        try { newImg.name = `${CFG.NAMESPACE}${tpl.key}`; } catch (e) {}

        newImg.left = fitted.x;
        newImg.top = fitted.y;
        newImg.width = fitted.w;
        newImg.height = fitted.h;

        await context.sync();

        try {
          if (window.AuditHelpers && typeof window.AuditHelpers.recordPlacedImage === 'function') {
            window.AuditHelpers.recordPlacedImage(
              tpl.slide,
              `${CFG.NAMESPACE}${tpl.key}`,
              { slotIndex: tpl.slot, w: fitted.w, h: fitted.h }
            );
          }
        } catch (e) {}

        return { ok: true };
      }),
      new Promise((resolve) =>
        setTimeout(() => resolve({ ok: false, reason: `timeout ${timeoutMs}ms` }), timeoutMs)
      ),
    ]);
  }

  let pos = await tryPositioning(4000);
  if (pos.ok) {
    log(`✅ Image ${tpl.label} placée sur slide ${tpl.slide}`, 'ok');
  } else {
    log(`⚠️ Positionnement 1 : ${pos.reason} — retry`, 'err');
    await new Promise((r) => setTimeout(r, 800));
    pos = await tryPositioning(4000);
    if (pos.ok) {
      log(`✅ Image ${tpl.label} placée (retry) sur slide ${tpl.slide}`, 'ok');
    } else {
      log(`⚠️ Positionnement échoué : ${pos.reason}`, 'err');
      try {
        if (window.AuditHelpers && typeof window.AuditHelpers.recordPlacedImage === 'function') {
          window.AuditHelpers.recordPlacedImage(
            tpl.slide,
            `${CFG.NAMESPACE}${tpl.key}__orphan_${Date.now()}`,
            { slotIndex: tpl.slot, w: fitted.w, h: fitted.h, orphanRect: tpl.rect }
          );
        }
      } catch (e) {}
    }
  }

  if (mode === 'end') {
    log('🔒 Mode END');
  }

  placementState.imagesPlaced++;
  setStatus(`✅ ${tpl.label} placé (slide ${tpl.slide})`, 'ok');
}