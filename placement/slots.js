/* ================================================================
   Placement — Slot finding, wipe, focus
   ================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

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