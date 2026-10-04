/* ================================================================
   Placement — FAST path (minimum Office.js round-trips)
   ----------------------------------------------------------------
   Why: every PowerPoint.run()/context.sync() is a slow round-trip in
   PowerPoint Online and blocks the browser while it runs. The old
   flow did ~10–40 of them per image (it scanned EVERY slide, then
   focus, count, 2 s wait, poll...). This file does it in:

     scanLast        1 run / 2 syncs  (slides + last slide shapes,
                                       also selects the last slide)
     insert          1 setSelectedDataAsync (exact position given)
     positionNewFast 1 sync per try, usually 1–2 tries (~150–400 ms)

   The new image is found by DIFFING shape ids (before / after), so
   nothing else on the slide is ever touched or deleted.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

const _sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/* ================================================================
   Slide baseline — detect slides the USER duplicated.
   When the user duplicates a slide in PowerPoint, the copy carries
   the audit-img-* images (and names) of the original, so the slide
   looks "full" / "already has this image". Slides that appear AFTER
   the add-in started (unknown id) and that the add-in never filled
   are therefore copies → their audit images are cleared and the
   slide is treated as empty.
================================================================ */
window.AuditPlacement.state = window.AuditPlacement.state || {};

window.AuditPlacement.initBaseline = async function () {
  const P = window.AuditPlacement;
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items/id');
      await context.sync();
      P.state.knownSlideIds = new Set(slides.items.map((x) => x.id));
    });
  } catch (e) { /* will be set lazily by the first scan */ }
};

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(() => { window.AuditPlacement.initBaseline(); });
}

/* Scan ONLY the last slide, select it, and return everything the
   flow needs: existing image of this type, free slot, shape ids. */
window.AuditPlacement.scanLast = async function (templateKey) {
  const P = window.AuditPlacement;
  const CFG = P.CFG;

  return await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items/id');
    await context.sync();

    const n = slides.items.length;
    if (n === 0) return null;
    const last = slides.items[n - 1];

    /* select + load in the SAME batch → a single sync */
    context.presentation.setSelectedSlides([last.id]);
    last.shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
    await context.sync();

    /* ⭐ user-duplicated slide? (new id, never filled by us) */
    const st = P.state;
    const firstScan = !st.knownSlideIds;
    if (firstScan) st.knownSlideIds = new Set();
    const isNewSlide = !firstScan && !st.knownSlideIds.has(last.id);
    slides.items.forEach((x) => st.knownSlideIds.add(x.id));

    let shapes = last.shapes.items;
    if (isNewSlide) {
      const copies = shapes.filter((s) => (s.name || '').startsWith(CFG.NAMESPACE));
      if (copies.length > 0) {
        copies.forEach((s) => s.delete());
        await context.sync();
        const gone = new Set(copies.map((s) => s.id));
        shapes = shapes.filter((s) => !gone.has(s.id));
        log(`📑 Slide dupliquée manuellement — ${copies.length} image(s) copiée(s) retirée(s)`);
      }
    }

    const audit = shapes.filter(
      (s) => s.type === PowerPoint.ShapeType.image && (s.name || '').startsWith(CFG.NAMESPACE)
    );

    const target = CFG.NAMESPACE + templateKey;
    const ex = audit.find((s) => s.name === target);

    let slot = null;
    if (!P.imagesOverlap(audit, CFG.SLOT_LEFT)) slot = { slot: 'left', rect: CFG.SLOT_LEFT };
    else if (!P.imagesOverlap(audit, CFG.SLOT_RIGHT)) slot = { slot: 'right', rect: CFG.SLOT_RIGHT };

    return {
      slideNumber: n,
      slideId: last.id,
      ids: shapes.map((s) => s.id),
      existing: ex
        ? { id: ex.id, rect: { x: ex.left ?? 0, y: ex.top ?? 0, w: ex.width ?? 0, h: ex.height ?? 0 } }
        : null,
      slot,
    };
  });
};

/* Delete one shape — no loading, a single sync. */
window.AuditPlacement.deleteShapeFast = async function (slideId, shapeId) {
  try {
    await PowerPoint.run(async (context) => {
      context.presentation.slides.getItem(slideId).shapes.getItem(shapeId).delete();
      await context.sync();
    });
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
};

/* Find the freshly inserted image (id not in beforeIds), name it,
   and set its exact rect. Short, increasing polling delays. */
window.AuditPlacement.positionNewFast = async function (slideId, beforeIds, fitted, templateKey) {
  const CFG = window.AuditPlacement.CFG;
  const delays = [120, 250, 400, 600, 800, 1000, 1200];

  for (const d of delays) {
    await _sleep(d);
    try {
      const res = await PowerPoint.run(async (context) => {
        const shapes = context.presentation.slides.getItem(slideId).shapes;
        shapes.load('items/id,items/type');
        await context.sync();

        const fresh = shapes.items.filter(
          (s) => s.type === PowerPoint.ShapeType.image && !beforeIds.includes(s.id)
        );
        if (fresh.length === 0) return null;

        const img = fresh[fresh.length - 1];
        img.name = CFG.NAMESPACE + templateKey;
        img.left = fitted.x;
        img.top = fitted.y;
        img.width = fitted.w;
        img.height = fitted.h;
        await context.sync();
        return { ok: true };
      });
      if (res) return res;
    } catch (e) { /* try again */ }
  }
  return { ok: false, reason: 'new image not found' };
};


/* ================================================================
   Slide full → wait for the USER to duplicate the slide.
   Shows a banner in the taskpane (with a Cancel button) and polls
   with a tiny call (1 sync) until a new last slide appears.
   Resolves true (new slide) / false (cancelled or timeout).
================================================================ */
window.AuditPlacement.waitForUserDuplicate = function (prevCount, prevLastId) {
  const CFG = window.AuditPlacement.CFG;
  return new Promise((resolve) => {
    const old = document.getElementById('audit-wait');
    if (old) old.remove();

    const box = document.createElement('div');
    box.id = 'audit-wait';
    box.innerHTML = `
      <div style="margin-bottom:10px">📑 <b>Slide pleine</b> (2 images)<br>
        Dupliquez la slide dans PowerPoint —<br>l'image sera placée automatiquement.</div>
      <button id="audit-wait-cancel" style="background:#374151;">Annuler</button>`;
    box.style.cssText = `position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);
      background:#1e293b;border:1px solid #f59e0b;border-radius:8px;padding:16px;z-index:99999;
      box-shadow:0 20px 60px rgba(0,0,0,.6);color:#e2e8f0;font-family:system-ui;width:300px;text-align:center;`;
    document.body.appendChild(box);

    let done = false;
    const finish = (v) => {
      if (done) return;
      done = true;
      try { box.remove(); } catch (e) {}
      resolve(v);
    };
    document.getElementById('audit-wait-cancel').onclick = () => finish(false);

    const t0 = Date.now();
    (async () => {
      while (!done) {
        if (Date.now() - t0 > CFG.WAIT_DUPLICATE_MS) return finish(false);
        await _sleep(1200);
        if (done) return;
        try {
          const changed = await PowerPoint.run(async (context) => {
            const slides = context.presentation.slides;
            slides.load('items/id');
            await context.sync();
            const n = slides.items.length;
            return n > prevCount || (n > 0 && slides.items[n - 1].id !== prevLastId);
          });
          if (changed) { await _sleep(400); return finish(true); }
        } catch (e) { /* keep polling */ }
      }
    })();
  });
};
