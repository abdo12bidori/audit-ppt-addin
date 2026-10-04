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

  /* false = NEVER switch to the PPT tab (no CDP Ctrl+V fallback).
     Set to true only if you accept the tab flash as a last resort. */
  ALLOW_CDP_FALLBACK: false,

  NAMESPACE: 'audit-img-',
  RETRY_MS: 3000,
  MAX_PER_SLIDE: 2,
  OVERLAP_PAD: 1,
};

window.AuditPlacement.state = { imagesPlaced: 0 };
