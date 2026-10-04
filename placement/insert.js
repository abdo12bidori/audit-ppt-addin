/* ================================================================
   Placement — Insert via Office.js setSelectedDataAsync
   ----------------------------------------------------------------
   v1.0:
     • Tries the cleanest path: pre-select the target slide via
       PowerPoint.run, then push the image via
       Office.context.document.setSelectedDataAsync with
       CoercionType.Image.
     • If PowerPoint Online accepts it → NO focus is needed and
       the PPT tab never activates.
     • Returns { ok: true } on success, { ok: false, reason }
       otherwise so paste.js can fall back to CDP.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.insertViaOfficeJs = async function (base64, slideNumber) {
  const P = window.AuditPlacement;

  try {
    /* 1) Pre-select the target slide WITHOUT bringing PPT to front */
    if (slideNumber && P.focusSlide) {
      await P.focusSlide(slideNumber);
    }

    /* 2) Push the image via Office.js */
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
      return { ok: true };
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