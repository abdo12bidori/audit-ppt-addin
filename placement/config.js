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

  NAMESPACE: 'audit-img-',
  RETRY_MS: 3000,
  MAX_PER_SLIDE: 2,
  OVERLAP_PAD: 5,
};

window.AuditPlacement.state = { imagesPlaced: 0 };