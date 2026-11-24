import { SPHttpClient } from '@microsoft/sp-http';

/** Props passed from the web part class to the React component */
export interface INewsletterProps {
  spHttpClient:   SPHttpClient;  // injected by the web part
  siteUrl:        string;        // absolute URL of the SharePoint site that holds news
  spotlightColumn: string;       // internal name of the Yes/No spotlight column → "DisplayAsSpotlight"
                                 // leave EMPTY to show just the latest N news posts (no spotlight filter)
  categoryColumn:  string;       // internal name of the Category column → "Category"
  itemCount:       number;       // how many cards to show (default 3)
  title:           string;       // heading  → "SPOTLIGHT"
  subtitle:        string;       // sub-line → "NEWS FROM AROUND SIRVA"
  rightLabel:      string;       // right label → "THIS WEEK"
}

/** Normalised news item used by the card renderer */
export interface INewsItem {
  id:       number;
  title:    string;
  category: string;   // from your category column, falls back to "News"
  date:     string;   // formatted "25 Sep 2026"
  url:      string;   // full absolute URL of the page
}
