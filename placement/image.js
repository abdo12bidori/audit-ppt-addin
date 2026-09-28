/* ================================================================
   Placement — Image utilities
   ================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.decodeImageDims = function (dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => reject(new Error('Cannot decode image'));
    img.src = dataUrl;
  });
};

window.AuditPlacement.containFit = function (slot, imgW, imgH) {
  let scale = Math.min(slot.w / imgW, slot.h / imgH);
  if (scale > 1) scale = 1;
  const w = imgW * scale;
  const h = imgH * scale;
  return {
    x: slot.x + (slot.w - w) / 2,
    y: slot.y + (slot.h - h) / 2,
    w,
    h,
  };
};
