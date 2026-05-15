import React from 'react';
import clsx from 'clsx';
import Link from '@docusaurus/Link';
import useDocusaurusContext from '@docusaurus/useDocusaurusContext';
import Layout from '@theme/Layout';
import styles from './index.module.css';

const features = [
  {
    emoji: '📦',
    title: '10 Business Domains',
    description:
      'Masters, Booking, Sales, Purchase, Finance, Inventory, Tax Compliance, Reporting, Communication, and Administration — all documented from source.',
    link: '/docs/masters',
    linkLabel: 'Browse domains →',
  },
  {
    emoji: '📋',
    title: '146 Features Catalogued',
    description:
      'Complete feature inventory extracted from the HITRIX VB6 codebase. Every transaction type, report, utility, and integration recorded.',
    link: '/docs/feature-inventory',
    linkLabel: 'View feature inventory →',
  },
  {
    emoji: '⚡',
    title: 'Gap Analysis vs DhanMan',
    description:
      'Side-by-side comparison of HITRIX capabilities against the DhanMan microservices platform. Critical, High, and Medium priority gaps identified.',
    link: '/docs/gap-analysis',
    linkLabel: 'View gap analysis →',
  },
];

function Feature({ emoji, title, description, link, linkLabel }) {
  return (
    <div className={clsx('col col--4', styles.feature)}>
      <div className="card padding--lg shadow--md">
        <div className={styles.featureEmoji}>{emoji}</div>
        <h3>{title}</h3>
        <p>{description}</p>
        <Link to={link}>{linkLabel}</Link>
      </div>
    </div>
  );
}

function HomepageHeader() {
  const { siteConfig } = useDocusaurusContext();
  return (
    <header className={clsx('hero hero--primary', styles.heroBanner)}>
      <div className="container">
        <h1 className="hero__title">{siteConfig.title}</h1>
        <p className="hero__subtitle">{siteConfig.tagline}</p>
        <div className={styles.buttons}>
          <Link className="button button--secondary button--lg margin-right--md" to="/docs/intro">
            Browse Documentation →
          </Link>
          <Link className="button button--outline button--secondary button--lg" to="/docs/gap-analysis">
            View Gap Analysis
          </Link>
        </div>
      </div>
    </header>
  );
}

export default function Home() {
  return (
    <Layout
      title="Home"
      description="HITRIX VB6 system documentation and DhanMan microservices migration gap analysis">
      <HomepageHeader />
      <main>
        <section className={styles.features}>
          <div className="container">
            <div className="row">
              {features.map((props, idx) => (
                <Feature key={idx} {...props} />
              ))}
            </div>
          </div>
        </section>

        <section className={clsx(styles.infoSection)}>
          <div className="container">
            <div className="row">
              <div className="col col--6">
                <h2>About HITRIX</h2>
                <p>
                  HITRIX is a VB6-based ERP system built by SRK for Indian textile and commodity
                  trading firms. It handles the complete trade cycle — from advance bookings and goods
                  inward to sales, purchase, GST compliance, bank reconciliation, and WhatsApp-based
                  customer communication.
                </p>
                <p>
                  This documentation was created to support the migration of HITRIX clients onto the
                  modern <strong>DhanMan</strong> platform built by B2A Technologies.
                </p>
              </div>
              <div className="col col--6">
                <h2>What's in Here</h2>
                <ul>
                  <li>Domain-by-domain breakdown of all HITRIX modules</li>
                  <li>Key data captured, business rules, and integration points per domain</li>
                  <li>Full list of Crystal Reports (100+ reports)</li>
                  <li>Flat feature inventory table (146 features)</li>
                  <li>Preliminary gap analysis against DhanMan's 7 microservices</li>
                  <li>Priority-ranked gaps: Critical → High → Medium</li>
                </ul>
              </div>
            </div>
          </div>
        </section>
      </main>
    </Layout>
  );
}
