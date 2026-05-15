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
        { type: 'doc', id: 'masters',              label: '01 - Masters & Configuration' },
        { type: 'doc', id: 'booking',              label: '02 - Booking & Order Management' },
        { type: 'doc', id: 'sales',                label: '03 - Sales' },
        { type: 'doc', id: 'purchase',             label: '04 - Purchase' },
        { type: 'doc', id: 'finance',              label: '05 - Finance & Accounting' },
        { type: 'doc', id: 'inventory-logistics',  label: '06 - Inventory & Logistics' },
        { type: 'doc', id: 'tax-compliance',       label: '07 - Tax Compliance' },
        { type: 'doc', id: 'reporting',            label: '08 - Reporting & Analytics' },
        { type: 'doc', id: 'communication',        label: '09 - Communication' },
        { type: 'doc', id: 'administration',       label: '10 - Administration & Security' },
      ],
    },
    {
      type: 'category',
      label: 'Diagrams',
      collapsed: false,
      items: [
        { type: 'doc', id: 'diagrams/er-diagram',          label: 'ER Diagram' },
        { type: 'doc', id: 'diagrams/business-flows',      label: 'Business Process Flows' },
        { type: 'doc', id: 'diagrams/module-dependency',   label: 'Module Dependency Map' },
      ],
    },
    {
      type: 'category',
      label: 'HITRIX V2 (SQL Server)',
      collapsed: false,
      items: [
        { type: 'doc', id: 'v2/v2-overview',        label: '00 - System Overview' },
        { type: 'doc', id: 'v2/v2-data-model',      label: '01 - Data Model (ER)' },
        { type: 'doc', id: 'v2/v2-masters',         label: '02 - Master Data' },
        { type: 'doc', id: 'v2/v2-booking',         label: '03 - Booking & Orders' },
        { type: 'doc', id: 'v2/v2-sales',           label: '04 - Sales' },
        { type: 'doc', id: 'v2/v2-purchase',        label: '05 - Purchase' },
        { type: 'doc', id: 'v2/v2-finance',         label: '06 - Finance & Accounting' },
        { type: 'doc', id: 'v2/v2-gst',            label: '07 - GST Compliance' },
        { type: 'doc', id: 'v2/v2-reporting',            label: '08 - Reporting & Stored Procedures' },
        { type: 'doc', id: 'v2/v2-bank-reconciliation', label: '09 - Bank Reconciliation' },
        { type: 'doc', id: 'v2/v2-whatsapp',            label: '10 - WhatsApp Integration' },
        { type: 'doc', id: 'v2/v2-migration-plan',      label: '11 - Migration Plan to DhanMan' },
      ],
    },
    {
      type: 'doc',
      id: 'feature-inventory',
      label: 'Feature Inventory (146 features)',
    },
    {
      type: 'doc',
      id: 'gap-analysis',
      label: 'Gap Analysis',
    },
  ],
};

module.exports = sidebars;
