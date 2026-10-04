/* ================================================================
   Placement — Insert via CDP Ctrl+V (FALLBACK)
   ----------------------------------------------------------------
   v1.0:
     • This is the OLD paste.js behaviour, isolated.
     • Called ONLY when insertViaOfficeJs fails.
     • Requires the PPT tab to be briefly active (~300 ms).
     • Forwards dataUrl + sourceTabId to background.js.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.insertViaCdpPaste = async function (base64) {
  const dataUrl = 'data:image/png;base64,' + base64;
  const sourceTabId = window.__auditSourceTabId || null;

  /* Fallback clipboard write via the content script (rarely used) */
  try {
    window.parent.postMessage({ type: 'AUDIT_WRITE_CLIPBOARD', dataUrl, ts: Date.now() }, '*');
    await new Promise((r) => setTimeout(r, 300));
  } catch (e) {}

  /* Main CDP paste request */
  try {
    window.parent.postMessage({
      type: 'AUDIT_PASTE_REQUEST',
      dataUrl: dataUrl,
      sourceTabId: sourceTabId,
      ts: Date.now(),
    }, '*');
    log('→ CDP fallback: collage demandé');
  } catch (e) {
    log('⚠️ CDP fallback message failed: ' + e.message, 'err');
  }

  await new Promise((r) => setTimeout(r, 900));
  return true;
};