import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './Newsletter.module.scss';          // ← .scss (was .css — now fixed)
import { INewsletterProps, INewsItem } from './INewsletterProps';

// ─────────────────────────────────────────────────────────────
// Shape of one SitePages item returned by the REST API
// ─────────────────────────────────────────────────────────────
interface ISitePageItem {
  Id: number;
  Title: string;
  FileRef: string;
  FirstPublishedDate?: string;
  Created: string;
  [column: string]: unknown;        // accepts dynamic category / spotlight columns
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** Formats ISO date → "25 Sep 2026" */
const formatDate = (iso: string): string => {
  if (!iso) return '';
  const d = new Date(iso);
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
};

// ─────────────────────────────────────────────────────────────
// REST call: get latest N spotlight news posts
//
// APPROACH:
//   SharePoint news posts live in the SitePages library.
//   PromotedState = 2  → published News Post (not 0=page, 1=draft)
//   Spotlight filter  → depends on how you marked them (see below):
//
//   OPTION A – you used the built-in "Promoted" / Spotlight toggle
//              in the News web part → SharePoint does NOT expose a
//              stable REST column for this, so we just get the
//              latest N published news posts ordered by date.
//              Set spotlightColumn = '' in property pane to use this.
//
//   OPTION B – you created a custom Yes/No column (e.g. "IsSpotlight")
//              and manually ticked it on each spotlight post.
//              Set spotlightColumn = 'IsSpotlight' (internal name).
//
// The code below supports BOTH: if spotlightColumn is empty it
// omits the filter; if it is set it adds "and <col> eq 1".
// ─────────────────────────────────────────────────────────────
const getSpotlightNews = async (props: INewsletterProps): Promise<INewsItem[]> => {
  const site = props.siteUrl.replace(/\/$/, '');
  const origin = new URL(site).origin;
  const serverRelative = new URL(site).pathname.replace(/\/$/, '');

  // Build $select – always need these base columns
  const baseColumns = ['Id', 'Title', 'FileRef', 'FirstPublishedDate', 'Created'];
  if (props.categoryColumn) baseColumns.push(props.categoryColumn);
  if (props.spotlightColumn) baseColumns.push(props.spotlightColumn);
  const select = baseColumns.join(',');

  // Build $filter
  let filter = 'PromotedState eq 2';                            // news posts only
  if (props.spotlightColumn) {
    filter += ` and ${props.spotlightColumn} eq 1`;            // spotlight filter
  }

  const url =
    `${site}/_api/web/GetList('${serverRelative}/SitePages')/items` +
    `?$select=${select}` +
    `&$filter=${encodeURIComponent(filter)}` +
    `&$orderby=FirstPublishedDate desc` +
    `&$top=${props.itemCount}`;

  const res: SPHttpClientResponse = await props.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1
  );

  if (!res.ok) {
    const body = await res.text();
    throw new Error(
      `Failed to load news (HTTP ${res.status}). ` +
      `Check the Site URL and column internal names. Details: ${body.substring(0, 200)}`
    );
  }

  const json: { value: ISitePageItem[] } = await res.json();

  return json.value.map((i): INewsItem => ({
    id: i.Id,
    title: i.Title,
    category: (i[props.categoryColumn] as string) || 'News',
    date: formatDate(i.FirstPublishedDate || i.Created),
    url: `${origin}${i.FileRef}`
  }));
};

// ─────────────────────────────────────────────────────────────
// Small SVG arrow icon used inside each card button
// ─────────────────────────────────────────────────────────────
const ArrowIcon = (): JSX.Element => (
  <svg width="11" height="11" viewBox="0 0 12 12" aria-hidden="true">
    <path d="M3 9L9 3M4 3h5v5" stroke="#fff" strokeWidth="1.8" fill="none" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

// ─────────────────────────────────────────────────────────────
// Skeleton card shown while loading
// ─────────────────────────────────────────────────────────────
const SkeletonCard = (): JSX.Element => (
  <div className={styles.card} style={{ cursor: 'default' }}>
    <div className={styles.skeletonMeta} />
    <div className={styles.skeletonTitle} />
    <div className={styles.skeletonTitle} style={{ width: '70%', marginTop: 8 }} />
    <div className={styles.footer}>
      <div className={styles.skeletonText} />
    </div>
  </div>
);

// ─────────────────────────────────────────────────────────────
// Main component
// ─────────────────────────────────────────────────────────────
const Newsletter: React.FC<INewsletterProps> = (props) => {
  const [items, setItems] = React.useState<INewsItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');

  React.useEffect(() => {
    let cancelled = false;

    const load = async (): Promise<void> => {
      setLoading(true);
      setError('');
      try {
        const news = await getSpotlightNews(props);
        if (!cancelled) setItems(news);
      } catch (e) {
        if (!cancelled) setError((e as Error).message);
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    void load();

    return () => { cancelled = true; };   // cleanup if props change mid-flight
  }, [props.siteUrl, props.spotlightColumn, props.categoryColumn, props.itemCount]);

  return (
    <section className={styles.spotlight} aria-label={props.title}>
      {/* Decorative blobs */}
      <div className={styles.blobOne} aria-hidden="true" />
      <div className={styles.blobTwo} aria-hidden="true" />

      {/* ── Header ── */}
      <header className={styles.header}>
        <div className={styles.headingGroup}>
          <h2 className={styles.title}>{props.title}</h2>
          <p className={styles.subtitle}>{props.subtitle}</p>
        </div>
        <div className={styles.line} aria-hidden="true" />
        <div className={styles.rightLabel}>{props.rightLabel}</div>
      </header>

      {/* ── Cards / states ── */}
      {loading && (
        <div className={styles.grid}>
          {Array.from({ length: props.itemCount }).map((_, i) => (
            <SkeletonCard key={i} />
          ))}
        </div>
      )}

      {!loading && error && (
        <div className={styles.message} role="alert">
          <strong>⚠ Could not load news.</strong><br />{error}
        </div>
      )}

      {!loading && !error && items.length === 0 && (
        <div className={styles.message}>No spotlight news posts found yet.</div>
      )}

      {!loading && !error && items.length > 0 && (
        <div className={styles.grid}>
          {items.map((n) => (
            <a
              key={n.id}
              href={n.url}
              className={styles.card}
              data-interception="off"   // open in full page, not SPFx intercepted
              aria-label={`Read the story: ${n.title}`}
            >
              <div className={styles.meta}>
                <span className={styles.category}>{n.category}</span>
                <span className={styles.date}>{n.date}</span>
              </div>
              <h3 className={styles.cardTitle}>{n.title}</h3>
              <div className={styles.footer}>
                <span className={styles.readMore}>Read the story</span>
                <span className={styles.arrow}><ArrowIcon /></span>
              </div>
            </a>
          ))}
        </div>
      )}
    </section>
  );
};

export default Newsletter;
