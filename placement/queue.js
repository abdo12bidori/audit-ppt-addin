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

        if (r && r.retry) {
          if (item.attempts < MAX_ATTEMPTS) {
            __queue.unshift(item);
            log(`⏳ Retry ${item.attempts}/${MAX_ATTEMPTS} dans ${CFG.RETRY_MS}ms`);
            await new Promise((res) => setTimeout(res, CFG.RETRY_MS));
          } else {
            log(`❌ Abandon après ${MAX_ATTEMPTS} tentatives — dupliquez une slide`, 'err');
            setStatus('❌ Slides pleines — dupliquez une slide', 'err');
          }
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

    __queue.push({ dataUrl, templateKey: templateKey || 'street', attempts: 0 });
    log(`📥 File : ${__queue.length} image(s)`);

    if (!__processing) processQueue();
  };
})();