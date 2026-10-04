/* ================================================================
   Placement — Paste orchestrator
   ----------------------------------------------------------------
   v5.0:
     • Tries insertViaOfficeJs() FIRST (no focus, no flash).
     • If it fails → falls back to insertViaCdpPaste() (brief
       focus of the PPT tab, ~300 ms).
     • Reports the outcome to the extension so it can show a
       toast in the source page.

   v6.0:
     • Returns an object { ok, imagesBefore, method } instead of
       `true`, so process.js can pass imagesBefore to
       positionNewImage and reliably find the newly-inserted
       shape (Office.js sometimes takes 2–3 s to reflect it).
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.insertViaPaste = async function (base64, slideNumber) {
  const P = window.AuditPlacement;
  const sourceTabId = window.__auditSourceTabId || null;

  /* ── 1) Try the clean path (Office.js — no focus) ── */
  const officeResult = await P.insertViaOfficeJs(base64, slideNumber);

  if (officeResult && officeResult.ok) {
    /* Success — notify the extension so it can show a toast */
    try {
      window.parent.postMessage({
        type: 'AUDIT_SHOW_TOAST',
        text: '✅ Image insérée dans PowerPoint',
        sourceTabId: sourceTabId,
        ts: Date.now(),
      }, '*');
    } catch (e) {}

    /* ⭐ v6.0 — return imagesBefore so process.js can locate the
       newly-inserted shape reliably. */
    return {
      ok: true,
      method: 'officejs',
      imagesBefore:
        typeof officeResult.imagesBefore === 'number'
          ? officeResult.imagesBefore
          : null,
    };
  }

  /* ── 2) Fallback: CDP Ctrl+V ── */
  log('↩️ Fallback CDP (Office.js a échoué)', 'err');
  await P.insertViaCdpPaste(base64);

  /* Notify the extension of the fallback path */
  try {
    window.parent.postMessage({
      type: 'AUDIT_SHOW_TOAST',
      text: '✅ Image envoyée (méthode CDP)',
      sourceTabId: sourceTabId,
      ts: Date.now(),
    }, '*');
  } catch (e) {}

  /* CDP path doesn't provide a reliable imagesBefore count — the
     canvas-pasted image often appears with a random name. Return
     null so positionNewImage falls back to aspect-ratio matching. */
  return {
    ok: true,
    method: 'cdp',
    imagesBefore: null,
  };
};