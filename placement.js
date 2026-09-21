/* ================================================================
   Audit Capture — Placement engine v2.1 (template-driven)
   ================================================================
   No more scanning. No more shape detection. No more timeouts.

   Flow:
     1. Look up template by key (from AuditTemplates)
     2. Ensure the required slide exists (add if missing)
     3. Focus the target slide
     4. If an image already exists in this slot → delete it
     5. Paste the new image via CDP
     6. Move it to the template's rect
     7. Name it audit-img-<key>

   User has full control afterwards — the Add-in never touches
   the image again.
   ================================================================ */

const CFG = {
  SLIDE_W: 960,      /* 16:9 */
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
   fitContain — scale image to fit rect, keep aspect, no distortion
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
   ensureSlideExists — ensures the presentation has at least N slides
================================================================ */
async function ensureSlideExists(targetSlideNumber) {
  try {
    return await Promise.race([
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        let added = 0;
        while (slides.items.length < targetSlideNumber) {
          const lastSlide = slides.items[slides.items.length - 1];
          lastSlide.layout.load('id');
          await context.sync();
          const layoutId = lastSlide.layout.id;

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
   deleteExistingInSlot — removes any image previously placed at
   this slot (identified by shape name audit-img-<templateKey>)
================================================================ */
async function deleteExistingInSlot(templateKey, slideNumber) {
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

        const targetName = `${CFG.NAMESPACE}${templateKey}`;
        let deleted = 0;
        for (const s of slide.shapes.items) {
          if ((s.name || '') === targetName) {
            try { s.delete(); deleted++; } catch (e) {}
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
   Core — runPlacement
================================================================ */
async function runPlacement(dataUrl, mode, templateKey) {
  /* ─── 1. Resolve template ─── */
  const templates = window.AuditTemplates;
  if (!templates || typeof templates.get !== 'function') {
    throw new Error('AuditTemplates non chargé');
  }

  const tpl = templates.get(templateKey);
  if (!tpl) {
    throw new Error('Template inconnu : ' + templateKey);
  }

  log(`Template : ${tpl.label} → slide ${tpl.slide}, slot ${tpl.slot}`);

  /* ─── 2. Decode dims ─── */
  let dims;
  try {
    dims = await decodeImageDims(dataUrl);
  } catch (e) {
    log('⚠️ Cannot decode image dims — fallback 4:3', 'err');
    dims = { w: 800, h: 600 };
  }
  log(`Image : ${dims.w}×${dims.h}, mode=${mode}`);

  const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');

  /* ─── 3. Ensure target slide exists ─── */
  const ensure = await ensureSlideExists(tpl.slide);
  if (ensure.ok) {
    log(`Slides : ${ensure.total}${ensure.added ? ' (' + ensure.added + ' ajoutée(s))' : ''}`);
  } else {
    log('⚠️ ensureSlideExists : ' + ensure.error, 'err');
  }

  /* ─── 4. Delete any existing image in the same slot ─── */
  const del = await deleteExistingInSlot(tpl.key, tpl.slide);
  if (del.ok && del.deleted > 0) {
    log(`🗑 ${del.deleted} ancienne image supprimée du slot`);
  }

  /* ─── 5. Focus the target slide before paste ─── */
  try {
    await Promise.race([
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        if (slides.items.length < tpl.slide) return;
        const target = slides.items[tpl.slide - 1];
        context.presentation.setSelectedSlides([target.id]);
        await context.sync();
      }),
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error('focus timeout (3s)')), 3000)
      ),
    ]);
    log(`Slide ${tpl.slide} sélectionnée`);
  } catch (e) {
    log('⚠️ Focus slide échoué : ' + e.message, 'err');
  }

  /* ─── 6. Paste the image via CDP ─── */
  log('→ Insertion via collage (CDP)');
  const inserted = await insertViaPaste(base64);
  if (!inserted) {
    throw new Error('Insertion par collage échouée');
  }

  /* ─── 7. Wait, then position the new image ─── */
  await new Promise((r) => setTimeout(r, 1500));

  const fitted = fitContain(tpl.rect, dims.w, dims.h);
  log(`Position cible : x=${fitted.x.toFixed(0)} y=${fitted.y.toFixed(0)} w=${fitted.w.toFixed(0)} h=${fitted.h.toFixed(0)}`);

  async function tryPositioning(timeoutMs) {
    return await Promise.race([
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        if (slides.items.length < tpl.slide) {
          return { ok: false, reason: 'slide missing' };
        }

        const target = slides.items[tpl.slide - 1];
        target.shapes.load('items');
        await context.sync();

        /* Find the newest untagged image on the target slide */
        const candidates = target.shapes.items.filter(
          (s) =>
            s.type === PowerPoint.ShapeType.image &&
            !(s.name || '').startsWith(CFG.NAMESPACE)
        );

        if (candidates.length === 0) {
          return { ok: false, reason: 'no candidate image' };
        }

        const newImg = candidates[candidates.length - 1];
        try { newImg.name = `${CFG.NAMESPACE}${tpl.key}`; } catch (e) {}

        newImg.left = fitted.x;
        newImg.top = fitted.y;
        newImg.width = fitted.w;
        newImg.height = fitted.h;

        await context.sync();

        /* Record in tracker (helpers.js) */
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

  let pos = await tryPositioning(6000);
  if (pos.ok) {
    log(`✅ Image ${tpl.label} placée sur slide ${tpl.slide}`, 'ok');
  } else {
    log(`⚠️ Positionnement échoué : ${pos.reason} — retry`, 'err');
    await new Promise((r) => setTimeout(r, 500));
    pos = await tryPositioning(6000);
    if (pos.ok) {
      log(`✅ Image ${tpl.label} placée (retry) sur slide ${tpl.slide}`, 'ok');
    } else {
      log(`⚠️ Positionnement définitif échoué : ${pos.reason}`, 'err');
    }
  }

  if (mode === 'end') {
    log('🔒 Mode END');
  }

  placementState.imagesPlaced++;
  setStatus(`✅ ${tpl.label} placé (slide ${tpl.slide})`, 'ok');
}