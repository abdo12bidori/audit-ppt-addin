/* ================================================================
   Placement — Process one image end to end
   ----------------------------------------------------------------
   v5.0:
     • insertViaPaste now takes the target slideNumber so the
       Office.js path can pre-select it.
     • Keeps Fix A (dedupeLastSlide) and the source-page replace
       dialog.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

/* Prompt user for confirmation — shown INSIDE the taskpane (still
   used as a fallback if the source-page dialog is unreachable). */
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

/* Ask the extension to show the replace dialog in the source page.
   Returns a Promise<boolean>. */
window.AuditPlacement.confirmReplaceInSource = function (templateLabel) {
  return new Promise((resolve) => {
    const requestId = 'r_' + Date.now() + '_' + Math.random().toString(36).slice(2);

    /* Listen for the answer once */
    const onAnswer = (e) => {
      if (!e.data || e.data.type !== 'AUDIT_REPLACE_ANSWER') return;
      if (e.data.requestId !== requestId) return;
      window.removeEventListener('message', onAnswer);
      resolve(!!e.data.ok);
    };
    window.addEventListener('message', onAnswer, false);

    /* Send the request */
    try {
      window.parent.postMessage({
        type: 'AUDIT_REPLACE_NEEDED',
        requestId: requestId,
        templateLabel: templateLabel,
        ts: Date.now(),
      }, '*');
    } catch (e) {
      window.removeEventListener('message', onAnswer);
      resolve(null);  /* null = couldn't reach the source page */
      return;
    }

    /* Timeout after 30s — fall back to taskpane dialog */
    setTimeout(() => {
      window.removeEventListener('message', onAnswer);
      resolve(null);
    }, 30000);
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

  /* Fix A — strip images from a freshly-duplicated slide */
  const deduped = await P.dedupeLastSlide();
  if (deduped && deduped.cleaned > 0) {
    log(`🧹 Slide dupliquée détectée — ${deduped.cleaned} image(s) retirée(s)`);
  }

  /* Check for an existing image of the same type on the LAST slide */
  const existing = await P.findExistingOfType(templateKey);
  if (existing) {
    log(`⚠️ Une image "${tpl.label}" existe déjà (slide ${existing.slideNumber})`, 'err');

    /* Try the source-page dialog first */
    let ok = await P.confirmReplaceInSource(tpl.label);

    /* Fallback to the in-taskpane dialog if the source page didn't reply */
    if (ok === null) {
      log('⚠️ Source dialog unreachable — using taskpane dialog');
      /* Ask the extension for focus so the user sees the dialog */
      try { window.parent.postMessage({ type: 'AUDIT_REQUEST_FOCUS' }, '*'); } catch (e) {}
      await new Promise((r) => setTimeout(r, 350));
      ok = await P.confirmReplace(tpl.label);
    }

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
    await P.insertViaPaste(base64, existing.slideNumber);

    await new Promise((r) => setTimeout(r, 1200));
    const pos = await P.positionNewImage(existing.slideNumber, fitted, templateKey);

    if (pos.ok) {
      log(`✅ Image remplacée (slide ${existing.slideNumber})`, 'ok');
      setStatus(`✅ Image remplacée`, 'ok');
    }
    return { retry: false };
  }

  /* Normal placement */
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
  await P.insertViaPaste(base64, slot.slideNumber);

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