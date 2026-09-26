/* ================================================================
   Templates index — 4 categories
================================================================ */
window.AuditTemplates = window.AuditTemplates || {};

window.AuditTemplatesList = ['street', 'obione', 'onb', 'ban'];

window.AuditTemplates.getList = function () {
  return window.AuditTemplatesList
    .map((k) => window.AuditTemplates[k])
    .filter(Boolean);
};

window.AuditTemplates.get = function (key) {
  return window.AuditTemplates[key] || null;
};

/* Detect the variant from the image aspect ratio */
window.AuditTemplates.detectVariant = function (categoryKey, imgW, imgH) {
  const tpl = window.AuditTemplates[categoryKey];
  if (!tpl || !tpl.variants || tpl.variants.length === 0) return null;

  const aspect = imgW / imgH;

  for (const v of tpl.variants) {
    if (aspect >= v.minAspect && aspect <= v.maxAspect) return v;
  }

  /* No match — pick the closest variant by midpoint */
  let best = tpl.variants[0];
  let bestDist = Infinity;
  for (const v of tpl.variants) {
    const mid = (v.minAspect + v.maxAspect) / 2;
    const d = Math.abs(aspect - mid);
    if (d < bestDist) { bestDist = d; best = v; }
  }
  return best;
};