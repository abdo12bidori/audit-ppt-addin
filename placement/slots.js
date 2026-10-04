/* ================================================================
   Placement — Slot finding, wipe, focus
   ----------------------------------------------------------------
   - findFreeSlot: scans all slides, returns first free slot
   - wipeAuditImages: removes audit-img-* from a slide
   - cleanAllOrphans: removes images that have no audit-img-* name
   - focusSlide: selects a slide so CDP paste lands there
   - v2.0 (Fix A):
       • _auditNamesOnSlide(slideNumber) — list audit-img-* names
       • dedupeLastSlide() — wipes all audit-img-* from the last
         slide IF its names all also exist on an earlier slide
         (i.e. this slide is a freshly-duplicated copy).
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

/* ================================================================
   Does any image overlap the given rect?
================================================================ */
window.AuditPlacement.imagesOverlap = function (images, rect) {
  const PAD = window.AuditPlacement.CFG.OVERLAP_PAD;
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
};

/* ================================================================
   Find the first slide with a free slot
   Returns { slideNumber, slot, rect } or null
================================================================ */
window.AuditPlacement.findFreeSlot = async function () {
  const CFG = window.AuditPlacement.CFG;
  const overlap = window.AuditPlacement.imagesOverlap;

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

      if (!overlap(images, CFG.SLOT_LEFT)) {
        return { slideNumber: i + 1, slot: 'left', rect: CFG.SLOT_LEFT };
      }
      if (!overlap(images, CFG.SLOT_RIGHT)) {
        return { slideNumber: i + 1, slot: 'right', rect: CFG.SLOT_RIGHT };
      }
    }
    return null;
  });
};

/* ================================================================
   Wipe audit-img-* images on a slide
================================================================ */
window.AuditPlacement.wipeAuditImages = async function (slideNumber) {
  const CFG = window.AuditPlacement.CFG;
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
        if (!(s.name || '').startsWith(CFG.NAMESPACE)) continue;
        try { s.delete(); deleted++; } catch (e) {}
      }
      if (deleted > 0) await context.sync();
      return { ok: true, deleted };
    });
  } catch (e) {
    return { ok: false, error: e.message };
  }
};

/* ================================================================
   Clean orphan images on ALL slides
================================================================ */
window.AuditPlacement.cleanAllOrphans = async function () {
  const CFG = window.AuditPlacement.CFG;
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      for (const slide of slides.items) {
        slide.shapes.load('items');
      }
      await context.sync();

      let totalDeleted = 0;

      for (const slide of slides.items) {
        for (const s of slide.shapes.items) {
          if (s.type !== PowerPoint.ShapeType.image) continue;
          const name = s.name || '';
          if (!name.startsWith(CFG.NAMESPACE)) {
            try {
              s.delete();
              totalDeleted++;
            } catch (e) {}
          }
        }
      }

      if (totalDeleted > 0) await context.sync();
      return { ok: true, deleted: totalDeleted };
    });
  } catch (e) {
    return { ok: false, error: e.message };
  }
};

/* ================================================================
   Focus a slide (so the CDP paste lands on the right slide)
================================================================ */
window.AuditPlacement.focusSlide = async function (slideNumber) {
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
};

/* ================================================================
   Find an existing image of a given type on the LAST slide
================================================================ */
window.AuditPlacement.findExistingOfType = async function (templateKey) {
  const CFG = window.AuditPlacement.CFG;
  const targetName = CFG.NAMESPACE + templateKey;

  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length === 0) return null;

      const slide = slides.items[slides.items.length - 1];
      slide.shapes.load('items');
      await context.sync();

      for (const s of slide.shapes.items) {
        if (s.type !== PowerPoint.ShapeType.image) continue;
        if ((s.name || '') === targetName) {
          return {
            slideNumber: slides.items.length,
            id: s.id,
            rect: {
              x: s.left ?? 0,
              y: s.top ?? 0,
              w: s.width ?? 0,
              h: s.height ?? 0,
            },
          };
        }
      }
      return null;
    });
  } catch (e) {
    return null;
  }
};

/* ================================================================
   Delete a shape by ID on a specific slide
================================================================ */
window.AuditPlacement.deleteShapeById = async function (slideNumber, shapeId) {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < slideNumber) return { ok: false };

      const slide = slides.items[slideNumber - 1];
      slide.shapes.load('items');
      await context.sync();

      for (const s of slide.shapes.items) {
        if (s.id === shapeId) {
          try { s.delete(); } catch (e) {}
          await context.sync();
          return { ok: true };
        }
      }
      return { ok: false, reason: 'shape not found' };
    });
  } catch (e) {
    return { ok: false, error: e.message };
  }
};

/* ================================================================
   ⭐ FIX A — Detect and clear a freshly-duplicated slide.
   A duplicated slide contains audit-img-* shapes with the SAME
   names as the original. We detect this by finding the LAST slide
   whose audit-img-* names ALL also exist on another (earlier)
   slide. If found, we wipe all audit-img-* from the last slide
   so it becomes a fresh slot.
================================================================ */

/* Return the list of audit-img-* names on a given slide */
window.AuditPlacement._auditNamesOnSlide = async function (slideNumber) {
  const CFG = window.AuditPlacement.CFG;
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < slideNumber) return [];

      const slide = slides.items[slideNumber - 1];
      slide.shapes.load('items');
      await context.sync();

      return slide.shapes.items
        .filter((s) => s.type === PowerPoint.ShapeType.image)
        .map((s) => s.name || '')
        .filter((n) => n.startsWith(CFG.NAMESPACE));
    });
  } catch (e) {
    return [];
  }
};

/* Wipe ALL audit-img-* from the LAST slide if it looks duplicated. */
window.AuditPlacement.dedupeLastSlide = async function () {
  const CFG = window.AuditPlacement.CFG;
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const total = slides.items.length;
      if (total < 2) return { cleaned: 0 };

      for (const slide of slides.items) slide.shapes.load('items');
      await context.sync();

      const lastSlide = slides.items[total - 1];
      const lastNames = lastSlide.shapes.items
        .filter((s) => s.type === PowerPoint.ShapeType.image)
        .map((s) => s.name || '')
        .filter((n) => n.startsWith(CFG.NAMESPACE));

      if (lastNames.length === 0) return { cleaned: 0 };

      const earlierNames = new Set();
      for (let i = 0; i < total - 1; i++) {
        slides.items[i].shapes.items
          .filter((s) => s.type === PowerPoint.ShapeType.image)
          .forEach((s) => {
            const n = s.name || '';
            if (n.startsWith(CFG.NAMESPACE)) earlierNames.add(n);
          });
      }

      const isDuplicate = lastNames.every((n) => earlierNames.has(n));
      if (!isDuplicate) return { cleaned: 0 };

      let cleaned = 0;
      for (const s of lastSlide.shapes.items) {
        if (s.type !== PowerPoint.ShapeType.image) continue;
        const n = s.name || '';
        if (!n.startsWith(CFG.NAMESPACE)) continue;
        try { s.delete(); cleaned++; } catch (e) {}
      }
      if (cleaned > 0) await context.sync();

      return { cleaned };
    });
  } catch (e) {
    return { cleaned: 0, error: e.message };
  }
};

/* ================================================================
   Select the existing image on its slide (so the user SEES which
   image is about to be replaced). Best effort — silently ignored
   if the API is not available on this PowerPoint version.
================================================================ */
window.AuditPlacement.selectShape = async function (slideNumber, shapeId) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[slideNumber - 1];
      if (!slide) return;
      slide.setSelectedShapes([shapeId]);
      await context.sync();
    });
  } catch (e) { /* not supported — slide selection is enough */ }
};

/* ================================================================
   ⭐ Duplicate the LAST slide (used when every slide is full).
   1) exportAsBase64 + insertSlidesFromBase64 (true duplicate)
   2) fallback: slides.add() with the same layout/master
   Then all audit-img-* images are wiped from the new slide so it
   starts empty. Returns { ok, slideNumber }.
================================================================ */
window.AuditPlacement.duplicateLastSlide = async function () {
  const P = window.AuditPlacement;
  try {
    const before = await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      return slides.items.length;
    });
    if (before === 0) return { ok: false, reason: 'no slides' };

    let done = false;

    /* 1) True duplicate */
    try {
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const last = slides.items[slides.items.length - 1];
        const b64 = last.exportAsBase64();
        await context.sync();
        context.presentation.insertSlidesFromBase64(b64.value, {
          formatting: 'KeepSourceFormatting',
          targetSlideId: last.id,
        });
        await context.sync();
      });
      done = true;
    } catch (e) {
      console.warn('[slots] duplicate via base64 failed', e);
    }

    /* 2) Fallback: new slide with same layout */
    if (!done) {
      try {
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load('items');
          await context.sync();
          const last = slides.items[slides.items.length - 1];
          last.layout.load('id');
          last.slideMaster.load('id');
          await context.sync();
          context.presentation.slides.add({
            layoutId: last.layout.id,
            slideMasterId: last.slideMaster.id,
          });
          await context.sync();
        });
        done = true;
      } catch (e) {
        return { ok: false, reason: e.message };
      }
    }

    /* Wait until the new slide is visible to Office.js */
    let total = before;
    for (let i = 0; i < 10 && total <= before; i++) {
      await new Promise((r) => setTimeout(r, 500));
      total = await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        return slides.items.length;
      });
    }
    if (total <= before) return { ok: false, reason: 'slide not created' };

    /* Empty the new slide of audit images, then show it */
    await P.wipeAuditImages(total);
    await P.focusSlide(total);
    return { ok: true, slideNumber: total };
  } catch (e) {
    return { ok: false, reason: e.message };
  }
};
