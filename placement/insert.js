/* ================================================================
   Placement — Insert via Office.js setSelectedDataAsync
   ----------------------------------------------------------------
   v2.0:
     • Before inserting, counts the existing images on the target
       slide and returns that count so position.js can identify
       the NEW image by diffing.
     • Uses a longer settle delay after insert because PPT Online
       needs 1.5–3 s to reflect the shape in shapes.items.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.insertViaOfficeJs = async function (base64, slideNumber) {
  const P = window.AuditPlacement;

  try {
    /* 1) Pre-select the target slide WITHOUT bringing PPT to front */
    if (slideNumber && P.focusSlide) {
      await P.focusSlide(slideNumber);
    }

    /* 2) Count existing images on the target slide (before insert) */
    let imagesBefore = 0;
    try {
      imagesBefore = await P._countImagesOnSlide(slideNumber);
      log(`📸 Images avant insertion sur slide ${slideNumber} : ${imagesBefore}`);
    } catch (e) {
      console.warn('[insert] count before failed', e);
    }

    /* 3) Push the image via Office.js */
    const dataUrl = 'data:image/png;base64,' + base64;

    const result = await new Promise((resolve) => {
      try {
        Office.context.document.setSelectedDataAsync(
          dataUrl,
          { coercionType: Office.CoercionType.Image },
          function (asyncResult) {
            resolve(asyncResult);
          }
        );
      } catch (e) {
        resolve({ status: 'failed', error: { message: e.message } });
      }
    });

    if (result.status === Office.AsyncResultStatus.Succeeded) {
      log('✅ insertViaOfficeJs succeeded (no focus needed)');
      return { ok: true, imagesBefore };
    }

    const errMsg = result.error
      ? (result.error.message || JSON.stringify(result.error))
      : 'unknown error';

    log('⚠️ insertViaOfficeJs failed: ' + errMsg, 'err');
    return { ok: false, reason: errMsg, code: result.error?.code };
  } catch (e) {
    log('⚠️ insertViaOfficeJs threw: ' + e.message, 'err');
    return { ok: false, reason: e.message };
  }
};

/* ================================================================
   Helper — count images on a slide (fast, no side effects)
================================================================ */
window.AuditPlacement._countImagesOnSlide = async function (slideNumber) {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items.length < slideNumber) return 0;

      const slide = slides.items[slideNumber - 1];
      slide.shapes.load('items');
      await context.sync();

      return slide.shapes.items.filter(
        (s) => s.type === PowerPoint.ShapeType.image
      ).length;
    });
  } catch (e) {
    console.warn('[insert] _countImagesOnSlide failed', e);
    return 0;
  }
};