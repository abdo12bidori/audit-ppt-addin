/* ================================================================
   Placement — Slot finding, wipe, focus
   ----------------------------------------------------------------
   - findFreeSlot: scans all slides, returns first free slot
   - wipeAuditImages: removes audit-img-* from a slide
   - cleanAllOrphans: removes images that have no audit-img-* name
       (these are unpositioned pastes left over from failed runs)
   - focusSlide: selects a slide so CDP paste lands there
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
   (leaves user's manual images alone)
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
   ----------------------------------------------------------------
   An "orphan" is any image on any slide that does NOT have a name
   starting with audit-img-*. These are the result of a paste that
   succeeded but whose positioning step failed — they occupy slots
   and block further placements.
   This function removes them, so the slides become usable again.
================================================================ */
window.AuditPlacement.cleanAllOrphans = async function () {
  const CFG = window.AuditPlacement.CFG;
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      /* Pre-load shapes for every slide */
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
   Returns { slideNumber, id, rect } or null
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

      /* Check the last slide only */
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