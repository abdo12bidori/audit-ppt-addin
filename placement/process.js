/* ================================================================
   Placement — Process one image end to end
   ----------------------------------------------------------------
   v6.1:
     • Reads imagesBefore from insertViaPaste's result and passes
       it to positionNewImage.
     • Waits 2000 ms (instead of 1200 ms) before positioning.
     • If insertViaOfficeJs succeeded → tries positioning; if the
       position step still fails, we do a second attempt.
     • Logs the exact reason of any failure.

   v6.3:
     • NEW — requestPptFocus(): asks the extension to bring the
       PowerPoint tab to the front BEFORE showing the "Replace?"
       dialog and BEFORE duplicating a full slide, so the user can
       actually see what they are about to replace / which slide
       is being created.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

/* ================================================================
   Ask the extension to bring the PowerPoint tab to the front.
   Used before showing "Replace?" and before duplicating a slide
   so the user can SEE the slide in PPT while deciding.
   Resolves true if focus was granted (or after a 1.2s timeout).
================================================================ */
window.AuditPlacement.requestPptFocus = function () {
  return new Promise((resolve) => {
    const requestId = 'f_' + Date.now() + '_' + Math.random().toString(36).slice(2);

    const onAck = (e) => {
      if (!e.data || e.data.type !== 'AUDIT_FOCUS_ACK') return;
      if (e.data.requestId !== requestId) return;
      window.removeEventListener('message', onAck, false);
      resolve(true);
    };
    window.addEventListener('message', onAck, false);

    try {
      window.parent.postMessage({
        type: 'AUDIT_REQUEST_FOCUS',
        requestId: requestId,
        ts: Date.now(),
      }, '*');
    } catch (err) {
      window.removeEventListener('message', onAck, false);
      resolve(false);
      return;
    }

    /* If the relay never acks, don't block the flow forever */
    setTimeout(() => {
      window.removeEventListener('message', onAck, false);
      resolve(false);
    }, 1200);
  });
};

/* Prompt user for confirmation — shown INSIDE the taskpane. */
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

/* Ask the extension to show the replace dialog in the source page. */
window.AuditPlacement.confirmReplaceInSource = function (templateLabel) {
  return new Promise((resolve) => {
    const requestId = 'r_' + Date.now() + '_' + Math.random().toString(36).slice(2);

    const onAnswer = (e) => {
      if (!e.data || e.data.type !== 'AUDIT_REPLACE_ANSWER') return;
      if (e.data.requestId !== requestId) return;
      window.removeEventListener('message', onAnswer);
      resolve(!!e.data.ok);
    };
    window.addEventListener('message', onAnswer, false);

    try {
      window.parent.postMessage({
        type: 'AUDIT_REPLACE_NEEDED',
        requestId: requestId,
        templateLabel: templateLabel,
        ts: Date.now(),
      }, '*');
    } catch (e) {
      window.removeEventListener('message', onAnswer);
      resolve(null);
      return;
    }

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

    /* ⭐ v6.3 — bring PPT to the front BEFORE showing the dialog,
       so the user actually SEES the slide + selected image they
       are being asked about. */
    await P.requestPptFocus();

    /* Go to the slide that holds the image (inside PPT — no browser
       tab switch) and select it so the user sees what will be replaced */
    await P.focusSlide(existing.slideNumber);
    await P.selectShape(existing.slideNumber, existing.id);

    let ok = await P.confirmReplaceInSource(tpl.label);

    if (ok === null) {
      log('⚠️ Source dialog unreachable — using taskpane dialog');
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
    const insertResult = await P.insertViaPaste(base64, existing.slideNumber, fitted);

    if (!insertResult || !insertResult.ok) {
      return { retry: false };
    }

    /* ⭐ v6.0 — wait 2000 ms, then position with retry */
    await new Promise((r) => setTimeout(r, 2000));

    const imagesBefore = insertResult && typeof insertResult.imagesBefore === 'number'
      ? insertResult.imagesBefore
      : null;

    const pos = await P.positionNewImage(
      existing.slideNumber, fitted, templateKey, imagesBefore
    );

    if (pos.ok) {
      log(`✅ Image remplacée (slide ${existing.slideNumber})`, 'ok');
      setStatus(`✅ Image remplacée`, 'ok');
    } else {
      log(`⚠️ Positionnement échoué : ${pos.reason || pos.error || '?'}`, 'err');
    }
    return { retry: false };
  }

  /* Normal placement */
  let slot = await P.findFreeSlot();
  if (!slot) {
    log('📑 Toutes les slides sont pleines — duplication automatique…');
    setStatus('📑 Duplication de la slide…');

    /* ⭐ v6.3 — bring PPT to the front so the user sees the new
       slide appear, then the image land on it. */
    await P.requestPptFocus();

    const dup = await P.duplicateLastSlide();
    if (dup.ok) {
      log(`✅ Slide ${dup.slideNumber} dupliquée`, 'ok');
      slot = await P.findFreeSlot();
    } else {
      log('⚠️ Duplication échouée : ' + (dup.reason || '?'), 'err');
    }
  }
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
  const insertResult = await P.insertViaPaste(base64, slot.slideNumber, fitted);

  if (!insertResult || !insertResult.ok) {
    return { retry: false };
  }

  /* ⭐ v6.0 — wait 2000 ms, then position with retry */
  await new Promise((r) => setTimeout(r, 2000));

  const imagesBefore = insertResult && typeof insertResult.imagesBefore === 'number'
    ? insertResult.imagesBefore
    : null;

  const pos = await P.positionNewImage(
    slot.slideNumber, fitted, templateKey, imagesBefore
  );

  if (pos.ok) {
    log(`✅ Image placée (slide ${slot.slideNumber}, ${slot.slot})`, 'ok');
    P.state.imagesPlaced = (P.state.imagesPlaced || 0) + 1;
    setStatus(`✅ Image placée (slide ${slot.slideNumber})`, 'ok');
  } else {
    log(`⚠️ Positionnement échoué : ${pos.reason || pos.error || '?'}`, 'err');
  }

  return { retry: false };
};