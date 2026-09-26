window.AuditTemplates = window.AuditTemplates || {};
window.AuditTemplates.street = {
  key: 'street',
  label: 'Vue rue',
  variants: [
    { key: 'google', label: 'Google Street View', minAspect: 2.3, maxAspect: 2.7 },
    { key: 'bing',   label: 'Bing Streetside',    minAspect: 1.8, maxAspect: 2.2 },
    { key: 'apple',  label: 'Apple Look Around',  minAspect: 2.7, maxAspect: 3.1 },
  ],
  defaultVariant: 'google',
};