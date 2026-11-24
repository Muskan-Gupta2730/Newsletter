import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import Newsletter from './components/Newsletter';
import { INewsletterProps } from './components/INewsletterProps';

export interface INewsletterWebPartProps {
  siteUrl:         string;   // leave empty → current site
  spotlightColumn: string;   // internal column name, leave empty to skip filter
  categoryColumn:  string;   // internal column name for category
  itemCount:       number;
  title:           string;
  subtitle:        string;
  rightLabel:      string;
}

export default class NewsletterWebPart extends BaseClientSideWebPart<INewsletterWebPartProps> {

  public render(): void {
    const p = this.properties;

    const element: React.ReactElement<INewsletterProps> = React.createElement(Newsletter, {
      spHttpClient:    this.context.spHttpClient,
      siteUrl:         p.siteUrl?.trim() || this.context.pageContext.web.absoluteUrl,
      spotlightColumn: p.spotlightColumn?.trim() ?? 'DisplayAsSpotlight',
      categoryColumn:  p.categoryColumn?.trim()  ?? 'Category',
      itemCount:       p.itemCount || 3,
      title:           p.title    || 'Spotlight',
      subtitle:        p.subtitle || 'News from around Sirva',
      rightLabel:      p.rightLabel || 'This week'
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Spotlight Newsletter settings' },
        groups: [
          {
            groupName: '📌 Data source',
            groupFields: [
              PropertyPaneTextField('siteUrl', {
                label: 'News site URL',
                description: 'Leave empty to use this site. Example: https://contoso.sharepoint.com/sites/Intranet',
                placeholder: 'https://contoso.sharepoint.com/sites/Intranet'
              }),
              PropertyPaneTextField('spotlightColumn', {
                label: 'Spotlight column (internal name)',
                description: 'Yes/No column in Site Pages. Your column is "DisplayAsSpotlight". Leave empty to show latest news without filtering.',
                placeholder: 'DisplayAsSpotlight'
              }),
              PropertyPaneTextField('categoryColumn', {
                label: 'Category column (internal name)',
                description: 'Choice column for the category label. Your column is "Category".',
                placeholder: 'Category'
              }),
              PropertyPaneSlider('itemCount', {
                label: 'Number of cards',
                min: 1,
                max: 6,
                value: 3,
                showValue: true
              })
            ]
          },
          {
            groupName: '✏️ Text labels',
            groupFields: [
              PropertyPaneTextField('title',      { label: 'Heading',      placeholder: 'Spotlight' }),
              PropertyPaneTextField('subtitle',   { label: 'Sub-heading',  placeholder: 'News from around Sirva' }),
              PropertyPaneTextField('rightLabel', { label: 'Right label',  placeholder: 'This week' })
            ]
          }
        ]
      }]
    };
  }
}
