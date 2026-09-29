/* ================================================================
   Placement — Process one image end to end
   ================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

window.AuditPlacement.processOneImage = async function (dataUrl, templateKey) {
  const P = window.AuditPlacement;

  const tpl = window.AuditTemplates.get(templateKey);
  if (!tpl) throw new Error('Template inconnu : ' + templateKey);

  const dims = await P.decodeImageDims(dataUrl);
  log(`Image : ${dims.w}×${dims.h} (${tpl.label})`);

  /* Clean orphans before scanning */
  const cleaned = await P.cleanAllOrphans();
  if (cleaned.ok && cleaned.deleted > 0) {
    log(`🧹 ${cleaned.deleted} image(s) orpheline(s) nettoyée(s)`);
  }

  const slot = await P.findFreeSlot();
  if (!slot) {
    log('⚠️ Toutes les slides sont pleines', 'err');
    setStatus('⚠️ Slides pleines — dupliquez-en une', 'err');
    return { retry: true };
  }
  log(`Slot libre : slide ${slot.slideNumber}, ${slot.slot}`);

  const w = await P.wipeAuditImages(slot.slideNumber);
  if (w.ok && w.deleted > 0) log(`🗑 ${w.deleted} image(s) audit supprimée(s)`);

  const fitted = P.containFit(slot.rect, dims.w, dims.h);
  log(`Fit : ${fitted.w.toFixed(0)}×${fitted.h.toFixed(0)} @ (${fitted.x.toFixed(0)},${fitted.y.toFixed(0)})`);

  await P.focusSlide(slot.slideNumber);

  const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
  await P.insertViaPaste(base64);

  await new Promise((r) => setTimeout(r, 1200));
  const pos = await P.positionNewImage(slot.slideNumber, fitted, templateKey);

  if (pos.ok) {
    log(`✅ Image placée (slide ${slot.slideNumber}, ${slot.slot})`, 'ok');
    P.state.imagesPlaced = (P.state.imagesPlaced || 0) + 1;
    setStatus(`✅ Image placée (slide ${slot.slideNumber})`, 'ok');
  } else {
    log(`⚠️ Positionnement échoué`, 'err');
  }

  return { retry: false };
};