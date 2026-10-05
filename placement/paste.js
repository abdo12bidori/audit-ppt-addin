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

   v7.0 — SPEED:
     • NEW — skipOfficeJs option: when the user is ALREADY on the
       PPT tab (e.g. just clicked "Remplacer" or triggered a
       slide duplication), we go straight to CDP Ctrl+V. ~300 ms
       instead of 2–8 s, no Chrome freeze.
     • Auto-detects focus state via document.visibilityState and
       document.hasFocus() to decide which path to take.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

/* Detect if the taskpane's host page is currently focused /
   visible. When the user just clicked "Remplacer" in the taskpane,
   the PPT tab IS focused → we can safely use CDP. */
window.AuditPlacement._isPptFocused = function () {
  try {
    if (document.visibilityState !== 'visible') return false;
    if (typeof document.hasFocus === 'function' && !document.hasFocus()) return false;
    return true;
  } catch (e) {
    return false;
  }
};

window.AuditPlacement.insertViaPaste = async function (base64, slideNumber, fitted, options) {
  const P = window.AuditPlacement;
  const opts = options || {};
  const sourceTabId = window.__auditSourceTabId || null;

  /* ⭐ v7.0 — decide which path to take */
  const userOnPpt = P._isPptFocused();
  /* CDP is NEVER the first choice (it can freeze Chrome). It is only the
     last resort below, after Office.js failed twice. */
  const skipOfficeJs = P.CFG.PREFER_CDP_ALWAYS === true;

  if (skipOfficeJs) {
    log(`⚡ v7.0 — CDP direct path (userOnPpt=${userOnPpt}, forced=${opts.skipOfficeJs === true})`);
    await P.insertViaCdpPaste(base64);

    /* (toast now sent by process.js once the image is really placed) */

    return {
      ok: true,
      method: 'cdp',
      imagesBefore: null,
    };
  }

  /* ── 1) Try the clean path (Office.js — no focus) ── */
  let officeResult = await P.insertViaOfficeJs(base64, slideNumber, fitted, opts);
  if (!officeResult || !officeResult.ok) {
    log('↻ Office.js : 2e tentative…');
    await new Promise((r) => setTimeout(r, 400));
    officeResult = await P.insertViaOfficeJs(base64, slideNumber, fitted, opts);
  }

  if (officeResult && officeResult.ok) {
    /* (toast now sent by process.js once the image is really placed) */

    return {
      ok: true,
      method: 'officejs',
      imagesBefore:
        typeof officeResult.imagesBefore === 'number'
          ? officeResult.imagesBefore
          : null,
    };
  }

  /* ── 2) No-focus mode: do NOT touch the PPT tab ── */
  if (!P.CFG.ALLOW_CDP_FALLBACK) {
    log('❌ Insertion Office.js échouée — fallback CDP désactivé (pas de changement d\'onglet)', 'err');
    setStatus('❌ Insertion échouée : ' + (officeResult && officeResult.reason || '?'), 'err');
    return { ok: false, method: 'none', imagesBefore: null };
  }

  /* ── 3) Fallback: CDP Ctrl+V (switches tab briefly) ── */
  log('↩️ Fallback CDP (Office.js a échoué)', 'err');
  await P.insertViaCdpPaste(base64);

    /* (toast now sent by process.js once the image is really placed) */

  /* CDP path doesn't provide a reliable imagesBefore count — the
     canvas-pasted image often appears with a random name. Return
     null so positionNewImage falls back to aspect-ratio matching. */
  return {
    ok: true,
    method: 'cdp',
    imagesBefore: null,
  };
};