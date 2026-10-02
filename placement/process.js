/* ================================================================
   Placement — Process one image end to end
   ----------------------------------------------------------------
   v2.0:
     • Fix A — dedupeLastSlide() before any placement (removes
               audit-img-* from a freshly-duplicated slide).
     • Fix C — ask the extension for focus ONLY when a replacement
               needs user confirmation. Normal placement never
               steals focus from the source page.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

/* Prompt user for confirmation. Returns a Promise<boolean>. */
window.AuditPlacement.confirmReplace = function (templateLabel) {
  return new Promise((resolve) => {
    const old = document.getElementById('audit-confirm');
    if (old) old.remove();

    const dlg = document.createElement('div');
    dlg.id = 'audit-confirm';
    dlg.innerHTML = `
      <div class="confirm-message">
        ⚠️ Une image <b>"${templateLabel}"</b> existe déjà.<br>
        Voulez-vous la remplacer ?
      </div>
      <div class="confirm-buttons">
        <button id="audit-confirm-cancel" style="background:#374151;">Ignorer</button>
        <button id="audit-confirm-ok" style="background:#dc2626;">Remplacer</button>
      </div>
    `;
    dlg.style.cssText = `
      position: fixed;
      top: 50%; left: 50%;
      transform: translate(-50%, -50%);
      background: #1e293b;
      border: 1px solid #dc2626;
      border-radius: 8px;
      padding: 16px;
      z-index: 99999;
      box-shadow: 0 20px 60px rgba(0,0,0,0.6);
      color: #e2e8f0;
      font-family: system-ui;
      width: 320px;
      text-align: center;
    `;
    document.body.appendChild(dlg);

    const cleanup = () => { try { dlg.remove(); } catch (e) {} };

    document.getElementById('audit-confirm-ok').onclick = () => { cleanup(); resolve(true); };
    document.getElementById('audit-confirm-cancel').onclick = () => { cleanup(); resolve(false); };
  });
};

window.AuditPlacement.processOneImage = async function (dataUrl, templateKey) {
  const P = window.AuditPlacement;
  const CFG = P.CFG;

  const tpl = window.AuditTemplates.get(templateKey);
  if (!tpl) throw new Error('Template inconnu : ' + templateKey);

  const dims = await P.decodeImageDims(dataUrl);
  log(`Image : ${dims.w}×${dims.h} (${tpl.label})`);

  /* Clean orphans */
  const cleaned = await P.cleanAllOrphans();
  if (cleaned.ok && cleaned.deleted > 0) {
    log(`🧹 ${cleaned.deleted} image(s) orpheline(s) nettoyée(s)`);
  }

  /* ⭐ FIX A — Strip audit images from a freshly-duplicated slide */
  const deduped = await P.dedupeLastSlide();
  if (deduped && deduped.cleaned > 0) {
    log(`🧹 Slide dupliquée détectée — ${deduped.cleaned} image(s) retirée(s)`);
  }

  /* Check for an existing image of the same type on the LAST slide */
  const existing = await P.findExistingOfType(templateKey);
  if (existing) {
    log(`⚠️ Une image "${tpl.label}" existe déjà (slide ${existing.slideNumber})`, 'err');

    /* ⭐ FIX C — ask the extension to bring the PPT tab to front
       ONLY now, because a user decision is required. */
    try {
      window.parent.postMessage({ type: 'AUDIT_REQUEST_FOCUS' }, '*');
    } catch (e) {}
    await new Promise((r) => setTimeout(r, 350));

    const ok = await P.confirmReplace(tpl.label);
    if (!ok) {
      log('Utilisateur a annulé le remplacement', 'err');
      setStatus('Image ignorée', 'err');
      return { retry: false };
    }

    const w = await P.deleteShapeById(existing.slideNumber, existing.id);
    if (w.ok) log(`🗑 Ancienne image supprimée`);

    const fitted = P.containFit(existing.rect, dims.w, dims.h);
    log(`Fit : ${fitted.w.toFixed(0)}×${fitted.h.toFixed(0)}`);

    await P.focusSlide(existing.slideNumber);

    const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
    await P.insertViaPaste(base64);

    await new Promise((r) => setTimeout(r, 1200));
    const pos = await P.positionNewImage(existing.slideNumber, fitted, templateKey);

    if (pos.ok) {
      log(`✅ Image remplacée (slide ${existing.slideNumber})`, 'ok');
      setStatus(`✅ Image remplacée`, 'ok');
    }
    return { retry: false };
  }

  /* No existing image of this type — normal placement */
  const slot = await P.findFreeSlot();
  if (!slot) {
    log('⚠️ Toutes les slides sont pleines', 'err');
    setStatus('⚠️ Slides pleines — dupliquez-en une', 'err');
    return { retry: true };
  }
  log(`Slot libre : slide ${slot.slideNumber}, ${slot.slot}`);

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