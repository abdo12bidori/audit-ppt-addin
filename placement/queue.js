/* ================================================================
   Placement — Serialized queue
   ================================================================ */
window.AuditPlacement = window.AuditPlacement || {};

(function () {
  const P = window.AuditPlacement;
  const CFG = P.CFG;

  let __queue = [];
  let __processing = false;
  let __lastDataUrl = null;
  let __lastDataUrlTs = 0;

  async function processQueue() {
    if (__processing) return;
    __processing = true;

    while (__queue.length > 0) {
      const item = __queue.shift();
      try {
        const r = await P.processOneImage(item.dataUrl, item.templateKey);
        if (r && r.retry) {
          __queue.unshift(item);
          log(`⏳ Retry dans ${CFG.RETRY_MS}ms`);
          await new Promise((res) => setTimeout(res, CFG.RETRY_MS));
        }
      } catch (e) {
        log('❌ ' + e.message, 'err');
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

    __queue.push({ dataUrl, templateKey: templateKey || 'street' });
    log(`📥 File : ${__queue.length} image(s)`);

    if (!__processing) processQueue();
  };
})();