/* ================================================================
   Placement — Paste orchestrator
   ----------------------------------------------------------------
   v5.0:
     • Tries insertViaOfficeJs() FIRST (no focus, no flash).
     • If it fails → falls back to insertViaCdpPaste() (brief
       focus of the PPT tab, ~300 ms).
     • Reports the outcome to the extension so it can show a
       toast in the source page.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.insertViaPaste = async function (base64, slideNumber) {
  const P = window.AuditPlacement;
  const sourceTabId = window.__auditSourceTabId || null;

  /* ── 1) Try the clean path (no focus) ── */
  const officeResult = await P.insertViaOfficeJs(base64, slideNumber);

  if (officeResult && officeResult.ok) {
    /* Success — notify the extension so it can show "✅ prêt" */
    try {
      window.parent.postMessage({
        type: 'AUDIT_SHOW_TOAST',
        text: '✅ Image insérée dans PowerPoint',
        sourceTabId: sourceTabId,
        ts: Date.now(),
      }, '*');
    } catch (e) {}
    return true;
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

  return true;
};