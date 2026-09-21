/* ================================================================
   Audit Capture — Templates Index
   ================================================================
   Registers all image-type templates. To add a new type later:
     1. Create templates/<yourtype>.js
     2. Add one line below: require via window.AuditTemplates.<key>
   That's it — the placement engine is fully generic.
================================================================ */

window.AuditTemplates = window.AuditTemplates || {};

/* Each template file self-registers on window.AuditTemplates.
   The order of these <script> tags in taskpane.html defines the
   order of the dropdown in the extension. */

window.AuditTemplatesList = [
  'street',
  'obione',
  'onb',
  'ban',
];

/* Helper: get list of templates as an array */
window.AuditTemplates.getList = function () {
  return window.AuditTemplatesList
    .map((key) => window.AuditTemplates[key])
    .filter(Boolean);
};

/* Helper: get one template by key */
window.AuditTemplates.get = function (key) {
  return window.AuditTemplates[key] || null;
};