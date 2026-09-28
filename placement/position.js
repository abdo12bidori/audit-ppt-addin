/* ================================================================
   Placement — Position the newly pasted image
   ================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.positionNewImage = async function (slideNumber, fitted, templateKey) {
  const CFG = window.AuditPlacement.CFG;

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
          !(s.name || '').startsWith(CFG.NAMESPACE)
      );

      if (candidates.length === 0) return { ok: false, reason: 'no new image' };

      const newImg = candidates[candidates.length - 1];

      for (let i = 0; i < candidates.length - 1; i++) {
        try { candidates[i].delete(); } catch (e) {}
      }

      try { newImg.name = CFG.NAMESPACE + templateKey; } catch (e) {}

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
};
