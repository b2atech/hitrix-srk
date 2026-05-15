/** @type {import('@docusaurus/plugin-content-docs').SidebarsConfig} */
const sidebars = {
  docs: [
    {
      type: 'doc',
      id: 'intro',
      label: 'Overview',
    },
    {
      type: 'category',
      label: 'HITRIX Domains',
      collapsed: false,
      items: [
        { type: 'doc', id: 'masters',              label: '01 · Masters & Configuration' },
        { type: 'doc', id: 'booking',              label: '02 · Booking & Order Management' },
        { type: 'doc', id: 'sales',                label: '03 · Sales' },
        { type: 'doc', id: 'purchase',             label: '04 · Purchase' },
        { type: 'doc', id: 'finance',              label: '05 · Finance & Accounting' },
        { type: 'doc', id: 'inventory-logistics',  label: '06 · Inventory & Logistics' },
        { type: 'doc', id: 'tax-compliance',       label: '07 · Tax Compliance' },
        { type: 'doc', id: 'reporting',            label: '08 · Reporting & Analytics' },
        { type: 'doc', id: 'communication',        label: '09 · Communication' },
        { type: 'doc', id: 'administration',       label: '10 · Administration & Security' },
      ],
    },
    {
      type: 'doc',
      id: 'feature-inventory',
      label: '📋 Feature Inventory (146 features)',
    },
    {
      type: 'doc',
      id: 'gap-analysis',
      label: '⚡ Gap Analysis',
    },
  ],
};

module.exports = sidebars;
