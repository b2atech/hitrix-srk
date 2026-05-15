// @ts-check
const { themes: prismThemes } = require('prism-react-renderer');

/** @type {import('@docusaurus/types').Config} */
const config = {
  title: 'SRK Legacy Docs',
  tagline: 'HITRIX VB6 system documentation & DhanMan migration gap analysis',
  favicon: 'img/favicon.ico',

  url: 'https://b2atech.github.io',
  baseUrl: '/hitrix-srk/',

  organizationName: 'b2atech',
  projectName: 'hitrix-srk',

  onBrokenLinks: 'warn',
  markdown: {
    hooks: {
      onBrokenMarkdownLinks: 'warn',
    },
  },

  i18n: {
    defaultLocale: 'en',
    locales: ['en'],
  },

  presets: [
    [
      'classic',
      /** @type {import('@docusaurus/preset-classic').Options} */
      ({
        docs: {
          path: '../docs',
          routeBasePath: 'docs',
          sidebarPath: require.resolve('./sidebars.js'),
        },
        blog: false,
        theme: {
          customCss: require.resolve('./src/css/custom.css'),
        },
      }),
    ],
  ],

  themeConfig:
    /** @type {import('@docusaurus/preset-classic').ThemeConfig} */
    ({
      image: 'img/og-card.png',
      navbar: {
        title: 'SRK Legacy Docs',
        items: [
          {
            type: 'docSidebar',
            sidebarId: 'docs',
            position: 'left',
            label: 'Documentation',
          },
          {
            href: 'https://github.com/b2atech/hitrix-srk',
            label: 'GitHub',
            position: 'right',
          },
        ],
      },
      footer: {
        style: 'dark',
        links: [
          {
            title: 'HITRIX Domains',
            items: [
              { label: 'Masters & Config', to: '/docs/masters' },
              { label: 'Sales', to: '/docs/sales' },
              { label: 'Purchase', to: '/docs/purchase' },
              { label: 'Finance', to: '/docs/finance' },
            ],
          },
          {
            title: 'Analysis',
            items: [
              { label: 'Feature Inventory', to: '/docs/feature-inventory' },
              { label: 'Gap Analysis', to: '/docs/gap-analysis' },
            ],
          },
        ],
        copyright: `Copyright © ${new Date().getFullYear()} SRK · B2A Technologies. Built with Docusaurus.`,
      },
      prism: {
        theme: prismThemes.github,
        darkTheme: prismThemes.dracula,
        additionalLanguages: ['sql', 'csharp', 'powershell', 'bash'],
      },
      colorMode: {
        defaultMode: 'light',
        disableSwitch: false,
        respectPrefersColorScheme: true,
      },
    }),
};

module.exports = config;
