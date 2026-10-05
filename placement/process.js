/* ================================================================
   Placement — Process one image end to end
   ----------------------------------------------------------------
   v6.0:
     • Reads imagesBefore from insertViaPaste's result and passes
       it to positionNewImage.
     • Waits 2000 ms (instead of 1200 ms) before positioning.
     • If insertViaOfficeJs succeeded → tries positioning; if the
       position step still fails, we do a second attempt.
     • Logs the exact reason of any failure.

   v6.3:
     • NEW — requestPptFocus(): asks the extension to bring the
       PowerPoint tab to the front BEFORE showing the "Replace?"
       dialog and BEFORE duplicating a full slide.

   v6.4 — SPEED:
     • requestPptFocus fallback timeout: 1200 → 400 ms.
     • Initial wait before positionNewImage: 2000 → 800 ms.

   v6.5 — SPEED / UX:
     • confirmReplaceInSource timeout: 30 s → 2 s (paramétrable).

   v6.6 — STABILITY / SPEED:
     • ⛔ cleanAllOrphans() and dedupeLastSlide() are NO LONGER
       called automatically in processOneImage. They were deleting
       the template images (logo, map, title cartouche) AND the
       image that had just been pasted. They remain available for
       manual invocation.

   v7.0 — CDP FAST PATH:
     • Replace flow: after requestPptFocus() + "Remplacer" click,
       the user is on PPT. We pass { skipOfficeJs: true } to
       insertViaPaste → CDP Ctrl+V directly (~300 ms).
     • Duplicate-slide flow: after requestPptFocus() +
       duplicateLastSlide(), same thing — the insertion uses CDP.
     • insertViaPaste also auto-detects focus and chooses CDP if
       the user happens to be on PPT already.
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

/* ================================================================
   Ask the extension to bring the PowerPoint tab to the front.
   Used before showing "Replace?" and before duplicating a slide
   so the user can SEE the slide in PPT while deciding.
   Resolves true if focus was granted (or after a 300 ms timeout).
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

    /* ⭐ v6.6 — fallback timeout 400 → 300 ms. */
    setTimeout(() => {
      window.removeEventListener('message', onAck, false);
      resolve(false);
    }, 1500);
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
window.AuditPlacement.confirmReplaceInSource = function (templateLabel, timeoutMs) {
  const TIMEOUT = typeof timeoutMs === 'number' ? timeoutMs : 2000;
  return new Promise((resolve) => {
    const requestId = 'r_' + Date.now() + '_' + Math.random().toString(36).slice(2);

    let settled = false;
    const settle = (value, reason) => {
      if (settled) return;
      settled = true;
      window.removeEventListener('message', onAnswer);
      clearTimeout(timer);
      console.log(`[AuditCapture:addin] confirmReplaceInSource → ${reason} (${value})`);
      resolve(value);
    };

    const onAnswer = (e) => {
      if (!e.data || e.data.type !== 'AUDIT_REPLACE_ANSWER') return;
      if (e.data.requestId !== requestId) return;
      settle(!!e.data.ok, 'answered');
    };
    window.addEventListener('message', onAnswer, false);

    try {
      window.parent.postMessage({
        type: 'AUDIT_REPLACE_NEEDED',
        requestId: requestId,
        templateLabel: templateLabel,
        ts: Date.now(),
      }, '*');
      console.log(`[AuditCapture:addin] confirmReplaceInSource → sent (requestId=${requestId}, timeout=${TIMEOUT}ms)`);
    } catch (e) {
      settle(null, 'postMessage failed');
      return;
    }

    const timer = setTimeout(() => {
      settle(null, `timeout after ${TIMEOUT}ms`);
    }, TIMEOUT);
  });
};

/* Toast in the source page — sent ONLY when the image is really in place
   (not when it is merely sent), so the message matches what you see. */
/* Which configured slot does this rectangle belong to? (largest overlap) */
window.AuditPlacement.slotRectFor = function (rect) {
  const CFG = window.AuditPlacement.CFG;
  const inter = (a, b) => {
    const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
    const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
    return w > 0 && h > 0 ? w * h : 0;
  };
  const l = inter(rect, CFG.SLOT_LEFT);
  const r = inter(rect, CFG.SLOT_RIGHT);
  if (l === 0 && r === 0) return rect;          /* outside both slots → keep old rect */
  return l >= r ? CFG.SLOT_LEFT : CFG.SLOT_RIGHT;
};

window.AuditPlacement.toast = function (text) {
  try {
    window.parent.postMessage({
      type: 'AUDIT_SHOW_TOAST',
      text: text,
      sourceTabId: window.__auditSourceTabId || null,
      ts: Date.now(),
    }, '*');
  } catch (e) {}
};

window.AuditPlacement.processOneImage = async function (dataUrl, templateKey) {
  const P = window.AuditPlacement;
  const CFG = P.CFG;

  const tpl = window.AuditTemplates.get(templateKey);
  if (!tpl) throw new Error('Template inconnu : ' + templateKey);

  const dims = await P.decodeImageDims(dataUrl);
  log(`Image : ${dims.w}×${dims.h} (${tpl.label})`);

  const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
  const t0 = Date.now();

  /* ⚡ ONE scan of the LAST slide (also selects it) */
  let scan = await P.scanLast(templateKey);
  if (!scan) { log('❌ Aucune slide', 'err'); return { retry: false }; }

  /* ───────────── Same type already there → replace ───────────── */
  if (scan.existing) {
    const ex = scan.existing;
    log(`⚠️ Une image "${tpl.label}" existe déjà (slide ${scan.slideNumber})`, 'err');

    const pptFront = await P.requestPptFocus();
    await P.selectShape(scan.slideNumber, ex.id);

    /* ⭐ The user is on the PPT tab (focus requested), so the dialog
       MUST appear there immediately. The old code waited up to 60 s for
       an answer from the source page the user could not see → the UI
       looked frozen. Now: taskpane dialog at once; the source-page
       dialog only runs in parallel if the ack failed or CONFIRM_IN_SOURCE
       is on, and the first answer wins. */
    let ok;
    const tp = P.confirmReplace(tpl.label);
    if (pptFront && !CFG.CONFIRM_IN_SOURCE) {
      ok = await tp;
    } else {
      /* The source-page dialog goes through a slow relay, so it used to
         pop up AFTER you had already answered in the taskpane. Now it is
         only requested if the taskpane has not been answered after 1.2 s. */
      let answered = false;
      tp.then(() => { answered = true; });
      const src = new Promise((resolve) => {
        setTimeout(() => {
          if (answered) return;                     /* never fire a stale dialog */
          P.confirmReplaceInSource(tpl.label, CFG.CONFIRM_TIMEOUT_MS)
            .then((v) => { if (v !== null) resolve(v); });
        }, 1200);
      });
      ok = await Promise.race([tp, src]);
      const dlg = document.getElementById('audit-confirm');
      if (dlg) dlg.remove();
      try { window.parent.postMessage({ type: 'AUDIT_REPLACE_CANCEL', ts: Date.now() }, '*'); } catch (e) {}
    }

    if (!ok) {
      log('Utilisateur a annulé le remplacement', 'err');
      setStatus('Image ignorée', 'err');
      return { retry: false };
    }

    /* ⭐ NOTHING IS LOST: insert the NEW image first, verify it, and only
       then delete the old one. If anything fails the old image stays. */
    /* ⭐ Fit into the SLOT the old image belongs to — NOT into the old
       image's own (already shrunk) rectangle, which made every replacement
       smaller and ignored the slot measurements. */
    const slotRect = P.slotRectFor(ex.rect);
    const fitted = P.containFit(slotRect, dims.w, dims.h);
    log(`Fit (slot ${slotRect === CFG.SLOT_LEFT ? 'gauche' : slotRect === CFG.SLOT_RIGHT ? 'droite' : 'inconnu'}) : ${fitted.w.toFixed(0)}×${fitted.h.toFixed(0)} @ (${fitted.x.toFixed(0)},${fitted.y.toFixed(0)})`);

    const ins = await P.insertViaPaste(base64, scan.slideNumber, fitted, { skipPrep: true });
    if (!ins || !ins.ok) {
      log('❌ Insertion échouée — ancienne image CONSERVÉE (image gardée pour réessayer)', 'err');
      setStatus('❌ Échec — ancienne image conservée', 'err');
      return { retry: false, keep: true };
    }

    const pos = await P.positionNewFast(scan.slideId, scan.ids, fitted, templateKey);
    if (!pos.ok) {
      log('⚠️ Nouvelle image introuvable — ancienne image CONSERVÉE, vérifiez la slide', 'err');
      setStatus('⚠️ Vérifiez la slide (ancienne image conservée)', 'err');
      return { retry: false };
    }

    const del = await P.deleteShapeFast(scan.slideId, ex.id);
    if (del.ok) log('🗑 Ancienne image supprimée (après succès)');
    log(`✅ Image remplacée (slide ${scan.slideNumber}) — ${Date.now() - t0} ms`, 'ok');
    P.toast('✅ Image envoyée — remplacée');
    setStatus('✅ Image remplacée', 'ok');
    return { retry: false };
  }

  /* ───────────── Normal placement ───────────── */
  if (!scan.slot) {
    if (CFG.AUTO_DUPLICATE) {
      log('📑 Slide pleine — duplication automatique…');
      setStatus('📑 Duplication de la slide…');
      const dup = await P.duplicateLastSlide();
      if (dup.ok) {
        log(`✅ Slide ${dup.slideNumber} dupliquée`, 'ok');
        scan = await P.scanLast(templateKey);
      } else {
        log('⚠️ Duplication échouée : ' + (dup.reason || '?'), 'err');
      }
    } else {
      /* ⭐ 2 images per slide → bring PPT to the front and wait for the
         user to duplicate; scanLast then clears the copied images. */
      log('📑 Slide pleine — en attente de la duplication par l\'utilisateur…');
      setStatus('📑 Slide pleine — dupliquez la slide', 'err');
      await P.requestPptFocus();

      const dupOk = await P.waitForUserDuplicate(scan.slideNumber, scan.slideId);
      if (!dupOk) {
        log('Duplication annulée — image GARDÉE (bouton Réessayer)', 'err');
        setStatus('Image gardée — cliquez Réessayer', 'err');
        return { retry: false, keep: true };
      }
      log('✅ Nouvelle slide détectée', 'ok');
      scan = await P.scanLast(templateKey);
    }
  }
  if (!scan || !scan.slot) {
    log('⚠️ Toutes les slides sont pleines', 'err');
    setStatus('⚠️ Slides pleines — dupliquez-en une', 'err');
    return { retry: true };
  }

  const fitted = P.containFit(scan.slot.rect, dims.w, dims.h);
  log(`Slot : slide ${scan.slideNumber}, ${scan.slot.slot} — fit ${fitted.w.toFixed(0)}×${fitted.h.toFixed(0)}`);

  const ins = await P.insertViaPaste(base64, scan.slideNumber, fitted, { skipPrep: true });
  if (!ins || !ins.ok) return { retry: true };   /* queue retries, then keeps the image */

  const pos = await P.positionNewFast(scan.slideId, scan.ids, fitted, templateKey);
  if (pos.ok) {
    log(`✅ Image placée (slide ${scan.slideNumber}, ${scan.slot.slot}) — ${Date.now() - t0} ms`, 'ok');
    P.toast(`✅ Image envoyée (slide ${scan.slideNumber})`);
    P.state.imagesPlaced = (P.state.imagesPlaced || 0) + 1;
    setStatus(`✅ Image placée (slide ${scan.slideNumber})`, 'ok');
  } else {
    log(`⚠️ Positionnement échoué : ${pos.reason || '?'}`, 'err');
  }
  return { retry: false };
};
