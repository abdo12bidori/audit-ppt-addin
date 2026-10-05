/* ================================================================
   Placement — Config
   ----------------------------------------------------------------
   All numbers in one place. Edit here to change the layout.
   Coordinates are in PPT points (1 pt = 1/72 inch).
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.CFG = {
  /* The two slot rectangles — measured from the PPT template
     (see the "Mesurer les formes" button output for slide 1) */
  SLOT_LEFT:  { x: 0,   y: 109, w: 311, h: 250 },
  SLOT_RIGHT: { x: 313, y: 153, w: 391, h: 206 },

  /* ⭐ v7 — CDP fallback settings.

     ALLOW_CDP_FALLBACK:
       true  = if insertViaOfficeJs fails, use CDP Ctrl+V.
       false = never touch the PPT tab; if Office.js fails, we
               simply give up.

     PREFER_CDP_ALWAYS:
       true  = always use CDP Ctrl+V, even when the user is NOT
               on the PPT tab (still works, just brings the tab
               forward briefly).
       false = let the orchestrator decide based on focus:
               if the user is already on PPT → CDP (~300 ms)
               if the user is on BAN/maps → Office.js first

     In all cases, the "Replace?" and "Duplicate slide" flows
     pass { skipOfficeJs: true } because the user has just been
     brought to PPT by requestPptFocus() — so CDP is the fast
     and safe path there. */
  ALLOW_CDP_FALLBACK: true,
  PREFER_CDP_ALWAYS: false,

  NAMESPACE: 'audit-img-',
  RETRY_MS: 3000,
  MAX_PER_SLIDE: 2,
  OVERLAP_PAD: 1,
};

window.AuditPlacement.state = { imagesPlaced: 0 };