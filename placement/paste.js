/* ================================================================
   Placement — Paste via clipboard + CDP Ctrl+V
   ----------------------------------------------------------------
   v3.0:
     • Fix B — the paste request carries the dataUrl and
               sourceTabId so background.js can:
                 – do an atomic CDP clipboard write immediately
                   before Ctrl+V (Windows screenshots can no
                   longer steal the paste);
                 – restore focus to the source page after paste.
     • The sourceTabId is read from the last AUDIT_ADD_IMAGE
       message the taskpane received (stored on window).
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.insertViaPaste = async function (base64) {
  const dataUrl = 'data:image/png;base64,' + base64;

  /* Retrouve le sourceTabId stocké par taskpane.js au moment de
     la réception du message AUDIT_ADD_IMAGE. */
  const sourceTabId = window.__auditSourceTabId || null;

  /* ─────────── Fallback clipboard write (page context) ─────────── */
  const writeClipboard = async (label) => {
    try {
      window.parent.postMessage({ type: 'AUDIT_WRITE_CLIPBOARD', dataUrl, ts: Date.now() }, '*');
      log(`→ [${label}] presse-papier`);
      await new Promise((r) => setTimeout(r, 500));
    } catch (e) {
      log(`⚠️ ${label}: ${e.message}`, 'err');
    }
  };

  /* ─────────── Main paste request (atomic via background) ─────────── */
  const pasteRequest = async (label) => {
    try {
      window.parent.postMessage({
        type: 'AUDIT_PASTE_REQUEST',
        dataUrl: dataUrl,          /* Fix B */
        sourceTabId: sourceTabId,  /* Fix C */
        ts: Date.now(),
      }, '*');
      log(`→ [${label}] collage`);
    } catch (e) {
      log(`⚠️ ${label}: ${e.message}`, 'err');
    }
  };

  /* Fallback only — background.js does the atomic CDP write.
     We still call writeClipboard() once so that even if CDP
     fails, the clipboard holds the right content. */
  await writeClipboard('fallback');
  await pasteRequest('main');
  await new Promise((r) => setTimeout(r, 900));
  return true;
};