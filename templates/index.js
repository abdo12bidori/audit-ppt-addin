window.AuditTemplates = window.AuditTemplates || {};

window.AuditTemplatesList = ['street', 'obione', 'onb', 'ban'];

window.AuditTemplates.getList = function () {
  return window.AuditTemplatesList
    .map((k) => window.AuditTemplates[k])
    .filter(Boolean);
};

window.AuditTemplates.get = function (key) {
  return window.AuditTemplates[key] || null;
};