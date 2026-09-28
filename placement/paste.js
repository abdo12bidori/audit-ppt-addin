/* ================================================================
   Placement — Paste via clipboard + CDP Ctrl+V
   ================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.insertViaPaste = async function (base64) {
  const dataUrl = 'data:image/png;base64,' + base64;

  const writeClipboard = async (label) => {
    try {
      window.parent.postMessage({ type: 'AUDIT_WRITE_CLIPBOARD', dataUrl, ts: Date.now() }, '*');
      log(`→ [${label}] presse-papier`);
      await new Promise((r) => setTimeout(r, 500));
    } catch (e) {
      log(`⚠️ ${label}: ${e.message}`, 'err');
    }
  };

  const pasteRequest = async (label) => {
    try {
      window.parent.postMessage({ type: 'AUDIT_PASTE_REQUEST', ts: Date.now() }, '*');
      log(`→ [${label}] collage`);
    } catch (e) {
      log(`⚠️ ${label}: ${e.message}`, 'err');
    }
  };

  await writeClipboard('1/4');
  await pasteRequest('2/4');
  await new Promise((r) => setTimeout(r, 900));
  await writeClipboard('3/4');
  await pasteRequest('4/4');
  await new Promise((r) => setTimeout(r, 900));
  return true;
};