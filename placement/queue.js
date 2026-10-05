/* ================================================================
   Placement — Serialized queue (limited retries)
   ----------------------------------------------------------------
   - Only 1 image processed at a time
   - If a slot isn't found, retry up to 3 times then give up
   - This prevents the infinite retry loop when slides are full
================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

(function () {
  const P = window.AuditPlacement;
  const CFG = P.CFG;
  const MAX_ATTEMPTS = 3;

  /* ⭐ Images that could not be placed are KEPT here (never dropped) and
     can be re-sent with the "Réessayer" button. */
  P.state.failed = P.state.failed || [];

  function updateRetryButton() {
    let btn = document.getElementById('audit-retry');
    const n = P.state.failed.length;
    if (n === 0) { if (btn) btn.remove(); return; }
    if (!btn) {
      btn = document.createElement('button');
      btn.id = 'audit-retry';
      btn.style.cssText = 'position:fixed;left:10px;right:10px;bottom:10px;z-index:99998;' +
        'background:#f59e0b;color:#111;border:0;border-radius:8px;padding:10px;font-weight:600;cursor:pointer;';
      btn.onclick = () => P.retryFailed();
      document.body.appendChild(btn);
    }
    btn.textContent = `↻ Réessayer ${n} image(s) non placée(s)`;
  }

  function keepFailed(item, why) {
    if (!P.state.failed.some((f) => f.dataUrl === item.dataUrl)) {
      P.state.failed.push({ dataUrl: item.dataUrl, templateKey: item.templateKey });
    }
    log(`💾 Image conservée (${why}) — ${P.state.failed.length} en attente`, 'err');
    updateRetryButton();
  }

  P.retryFailed = function () {
    const list = P.state.failed.splice(0);
    updateRetryButton();
    list.forEach((f) => {
      __queue.push({ dataUrl: f.dataUrl, templateKey: f.templateKey, attempts: 0 });
    });
    log(`↻ Nouvelle tentative : ${list.length} image(s)`);
    if (!__processing) processQueue();
  };

  let __queue = [];
  let __processing = false;
  let __lastDataUrl = null;
  let __lastDataUrlTs = 0;

  async function processQueue() {
    if (__processing) return;
    __processing = true;

    while (__queue.length > 0) {
      const item = __queue.shift();
      item.attempts = (item.attempts || 0) + 1;

      try {
        const r = await P.processOneImage(item.dataUrl, item.templateKey);

        if (r && r.keep) {
          keepFailed(item, 'non placée');
        } else if (r && r.retry) {
          if (item.attempts < MAX_ATTEMPTS) {
            __queue.unshift(item);
            log(`⏳ Retry ${item.attempts}/${MAX_ATTEMPTS} dans ${CFG.RETRY_MS}ms`);
            await new Promise((res) => setTimeout(res, CFG.RETRY_MS));
          } else {
            log(`❌ ${MAX_ATTEMPTS} tentatives échouées`, 'err');
            setStatus('❌ Image gardée — cliquez Réessayer', 'err');
            keepFailed(item, `échec après ${MAX_ATTEMPTS} tentatives`);
          }
        }
      } catch (e) {
        log('❌ ' + e.message, 'err');
        keepFailed(item, 'erreur');
      }
    }

    __processing = false;
  }

  P.enqueueImage = function (dataUrl, templateKey) {
    const now = Date.now();
    if (dataUrl === __lastDataUrl && now - __lastDataUrlTs < 800) {
      log('Image ignorée (doublon)');
      return;
    }
    __lastDataUrl = dataUrl;
    __lastDataUrlTs = now;

    __queue.push({ dataUrl, templateKey: templateKey || 'street', attempts: 0 });
    log(`📥 File : ${__queue.length} image(s)`);

    if (!__processing) processQueue();
  };
})();