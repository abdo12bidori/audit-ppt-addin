/* ================================================================
   Audit Capture — Slide Query v1.0
   ================================================================
   Read-only query interface for the Add-in.

   The extension asks the taskpane "where will this next image go?"
   before the user captures. The Add-in replies with slide number,
   existing image count, next slot index, and whether a new slide
   will be created.

   Uses the same helpers.js tracker so the answer matches what the
   placement engine will actually do.
   ================================================================ */

window.AuditQuery = {
  /**
   * Query the current slide state.
   * Returns: {
   *   ok: true,
   *   slideNumber: 7,
   *   imagesOnSlide: 1,
   *   maxPerSlide: 3,
   *   nextSlot: 2,
   *   willCreateNewSlide: false,
   *   slideClosed: false,
   *   summary: "Slide 7 · slot 2/3"
   * }
   */
  async getSlideState() {
    const CFG_LOCAL = {
      MAX_PER_ROW: 3,
      HEADER_TOP_FRAC: 0.45,
      FOOTER_BOT_FRAC: 0.85,
      NAMESPACE: 'audit-img-',
      CLOSED_FLAG: 'audit-closed',
    };

    try {
      return await Promise.race([
        PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load('items');
          await context.sync();

          if (slides.items.length === 0) {
            return {
              ok: false,
              error: 'No slides in presentation',
              summary: 'Aucune slide',
            };
          }

          const lastSlide = slides.items[slides.items.length - 1];
          lastSlide.shapes.load('items');
          await context.sync();

          let auditCount = 0;
          let closed = false;

          for (const s of lastSlide.shapes.items) {
            const name = s.name || '';
            if (name.startsWith(CFG_LOCAL.CLOSED_FLAG)) {
              closed = true;
              continue;
            }
            if (name.startsWith(CFG_LOCAL.NAMESPACE)) {
              auditCount++;
              continue;
            }
          }

          /* Merge with the local tracker (helpers.js) */
          let trackedCount = 0;
          try {
            if (window.AuditHelpers && typeof window.AuditHelpers.getTrackedImages === 'function') {
              trackedCount = window.AuditHelpers.getTrackedImages(slides.items.length).length;
            }
          } catch (e) {}

          const effectiveCount = Math.max(auditCount, trackedCount);
          const nextSlot = effectiveCount + 1;
          const willCreateNewSlide = closed || effectiveCount >= CFG_LOCAL.MAX_PER_ROW;

          let summary;
          if (willCreateNewSlide) {
            const reason = closed ? 'slide clôturée' : 'slide pleine';
            summary = `⚠️ ${reason} → nouvelle slide ${slides.items.length + 1}`;
          } else {
            summary = `Slide ${slides.items.length} · slot ${nextSlot}/${CFG_LOCAL.MAX_PER_ROW}`;
          }

          return {
            ok: true,
            slideNumber: slides.items.length,
            imagesOnSlide: effectiveCount,
            liveImages: auditCount,
            trackedImages: trackedCount,
            maxPerSlide: CFG_LOCAL.MAX_PER_ROW,
            nextSlot,
            willCreateNewSlide,
            slideClosed: closed,
            summary,
          };
        }),
        new Promise((resolve) =>
          setTimeout(
            () => resolve({ ok: false, error: 'query timeout (3s)', summary: 'Timeout' }),
            3000
          )
        ),
      ]);
    } catch (e) {
      return { ok: false, error: e.message, summary: 'Erreur' };
    }
  },
};