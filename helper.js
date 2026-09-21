/* ================================================================
   Audit Capture — Helpers v1.0
   ================================================================
   Side-channel tracking for placed images.

   Problem: on some PowerPoint Online tenants, PowerPoint.run() can
   take 5–8 seconds to reflect a pasted shape. The scan in
   placement.js then sees 0 images, thinks the slide is empty, and
   places a new one at slot 1 — stacking two images on top of each
   other.

   Solution: after every successful placement, record the image in
   localStorage keyed by slide index. On the next placement, merge
   Office.js's view with our local record so we don't lose track of
   images we already placed.

   Also provides:
     - a naming retry helper
     - a "was this image recently placed" dedupe cache
     - a way to force-tag an image as audit-img-N when Office.js
       can't find it via filter
   ================================================================ */

const HELPERS_STORAGE_KEY = 'auditCapturePlacedImages_v1';

/* ================================================================
   Local tracking storage
================================================================ */
function _loadTracked() {
  try {
    const raw = localStorage.getItem(HELPERS_STORAGE_KEY);
    if (!raw) return { slides: {} };
    const obj = JSON.parse(raw);
    if (!obj || typeof obj !== 'object' || !obj.slides) return { slides: {} };
    return obj;
  } catch (e) {
    return { slides: {} };
  }
}

function _saveTracked(obj) {
  try {
    localStorage.setItem(HELPERS_STORAGE_KEY, JSON.stringify(obj));
  } catch (e) {
    console.warn('[AuditCapture:helpers] storage failed', e);
  }
}

/* ================================================================
   Public: record a placed image
   slideIndex = 1-based slide number
   imageName  = e.g. "audit-img-2"
   meta       = { slotIndex, w, h, ts }
================================================================ */
function recordPlacedImage(slideIndex, imageName, meta = {}) {
  const db = _loadTracked();
  const key = String(slideIndex);
  if (!db.slides[key]) db.slides[key] = [];
  /* Remove any existing entry with the same name (replacement) */
  db.slides[key] = db.slides[key].filter((x) => x.name !== imageName);
  db.slides[key].push({
    name: imageName,
    slotIndex: meta.slotIndex ?? null,
    w: meta.w ?? null,
    h: meta.h ?? null,
    ts: Date.now(),
  });
  _saveTracked(db);
  console.log(`[AuditCapture:helpers] recorded ${imageName} on slide ${slideIndex}`);
}

/* ================================================================
   Public: get tracked images for a slide
   Returns an array (may be empty)
================================================================ */
function getTrackedImages(slideIndex) {
  const db = _loadTracked();
  const key = String(slideIndex);
  return Array.isArray(db.slides[key]) ? db.slides[key].slice() : [];
}

/* ================================================================
   Public: clear tracked images for a slide (or all)
================================================================ */
function clearTrackedImages(slideIndex) {
  const db = _loadTracked();
  if (slideIndex == null) {
    _saveTracked({ slides: {} });
    console.log('[AuditCapture:helpers] cleared all tracked images');
    return;
  }
  const key = String(slideIndex);
  if (db.slides[key]) delete db.slides[key];
  _saveTracked(db);
  console.log(`[AuditCapture:helpers] cleared tracked images on slide ${slideIndex}`);
}

/* ================================================================
   Public: merge Office.js view with local tracking
   liveImages:  array returned by scanSlide (shapes with .name)
   slideIndex:  1-based slide number
   Returns the effective count and a list of "phantom" images
   (ones we placed but Office.js can't see anymore).
================================================================ */
function mergeImageViews(liveImages, slideIndex) {
  const liveNames = new Set((liveImages || []).map((s) => s.name).filter(Boolean));
  const tracked = getTrackedImages(slideIndex);
  const phantoms = tracked.filter((t) => !liveNames.has(t.name));
  const effectiveCount = liveImages.length + phantoms.length;
  return {
    effectiveCount,
    liveCount: liveImages.length,
    trackedCount: tracked.length,
    phantomCount: phantoms.length,
    phantoms,
    liveNames: Array.from(liveNames),
  };
}

/* ================================================================
   Public: dedupe cache — avoid placing the same image twice
   (protects against double postMessage arrivals)
================================================================ */
const HELPERS_RECENT_KEY = 'auditCaptureRecentHashes_v1';
const HELPERS_RECENT_MAX = 20;

function _hashString(s) {
  /* Fast non-crypto hash for dedupe only */
  let h = 0;
  if (!s) return '0';
  const len = Math.min(s.length, 10000); /* sample large strings */
  for (let i = 0; i < len; i++) {
    h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  }
  return String(h);
}

function markImageAsRecent(dataUrl) {
  const h = _hashString(dataUrl);
  try {
    const raw = localStorage.getItem(HELPERS_RECENT_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    const filtered = Array.isArray(arr) ? arr.filter((x) => x !== h) : [];
    filtered.push(h);
    const trimmed = filtered.slice(-HELPERS_RECENT_MAX);
    localStorage.setItem(HELPERS_RECENT_KEY, JSON.stringify(trimmed));
  } catch (e) {}
  return h;
}

function wasImageRecentlyPlaced(dataUrl) {
  const h = _hashString(dataUrl);
  try {
    const raw = localStorage.getItem(HELPERS_RECENT_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) && arr.includes(h);
  } catch (e) {
    return false;
  }
}

/* ================================================================
   Public: pick a candidate image that Office.js may have missed
   This is a fallback for the case where the newly-pasted image is
   not yet visible via shapes.load('items').
================================================================ */
async function findNewestUntaggedImage(slide, namespace, closedFlag) {
  try {
    slide.shapes.load('items');
    await slide.context.sync();
    const candidates = slide.shapes.items.filter(
      (s) =>
        s.type === PowerPoint.ShapeType.image &&
        !(s.name || '').startsWith(namespace) &&
        !(s.name || '').startsWith(closedFlag)
    );
    return candidates.length ? candidates[candidates.length - 1] : null;
  } catch (e) {
    console.warn('[AuditCapture:helpers] findNewestUntaggedImage failed', e);
    return null;
  }
}

/* ================================================================
   Public: reset helper (used by the Add-in's "Réinitialiser" button)
================================================================ */
function resetHelperState() {
  try {
    localStorage.removeItem(HELPERS_STORAGE_KEY);
    localStorage.removeItem(HELPERS_RECENT_KEY);
    console.log('[AuditCapture:helpers] reset complete');
  } catch (e) {}
}

/* ================================================================
   Expose helpers on window so placement.js can use them
   (both are plain scripts in the same page context)
================================================================ */
window.AuditHelpers = {
  recordPlacedImage,
  getTrackedImages,
  clearTrackedImages,
  mergeImageViews,
  markImageAsRecent,
  wasImageRecentlyPlaced,
  findNewestUntaggedImage,
  resetHelperState,
};