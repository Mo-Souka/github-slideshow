import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MessageBar, MessageBarType } from '@fluentui/react';

import * as strings from 'GoodsReceivingWebPartStrings';
import { buildAppConfig, defaultSolutionConfig } from '../../config/appConfig';
import { AppError } from '../../models/AppError';
import { createSp } from '../../services/sp';
import { ReceivingRecordService } from '../../services/ReceivingRecordService';
import { AttachmentService } from '../../services/AttachmentService';
import { UserService } from '../../services/UserService';
import { App } from './components/App';

export interface IGoodsReceivingWebPartProps {
  /** Optional: absolute URL of another site that holds the list and library. */
  dataSiteUrl?: string;
  listTitle?: string;
  libraryTitle?: string;
  dateWindowDays?: number;
  pageSize?: number;
}

export default class GoodsReceivingWebPart extends BaseClientSideWebPart<IGoodsReceivingWebPartProps> {
  public render(): void {
    let element: React.ReactElement;
    try {
      const siteUrl = (this.properties.dataSiteUrl || '').trim() || this.context.pageContext.web.absoluteUrl;
      const config = buildAppConfig({
        siteUrl,
        listTitle: (this.properties.listTitle || '').trim() || undefined,
        libraryTitle: (this.properties.libraryTitle || '').trim() || undefined,
        dateWindowDays: this.properties.dateWindowDays,
        pageSize: this.properties.pageSize
      });
      const sp = createSp(this.context, config.siteUrl);
      const services = {
        records: new ReceivingRecordService(sp, config),
        attachments: new AttachmentService(sp, config),
        users: new UserService(sp, config)
      };
      element = React.createElement(App, { config, services });
    } catch (error) {
      const message = error instanceof AppError ? `${error.userMessage}\n${error.details || ''}` : String(error);
      element = React.createElement(
        MessageBar,
        { messageBarType: MessageBarType.error, isMultiline: true },
        React.createElement('div', { style: { whiteSpace: 'pre-wrap' } }, message)
      );
    }
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** Settings are applied with the "Apply" button, not on every keystroke. */
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleLabel,
                  placeholder: defaultSolutionConfig.list.title,
                  description: strings.ListTitleDescription
                }),
                PropertyPaneTextField('libraryTitle', {
                  label: strings.LibraryTitleLabel,
                  placeholder: defaultSolutionConfig.library.title
                }),
                PropertyPaneTextField('dataSiteUrl', {
                  label: strings.DataSiteUrlLabel,
                  placeholder: 'https://contoso.sharepoint.com/sites/warehouse',
                  description: strings.DataSiteUrlDescription
                })
              ]
            },
            {
              groupName: strings.ListViewGroupName,
              groupFields: [
                PropertyPaneSlider('dateWindowDays', {
                  label: strings.DateWindowLabel,
                  min: 1,
                  max: 365,
                  step: 1,
                  value: this.properties.dateWindowDays || defaultSolutionConfig.listView.defaultDateWindowDays
                }),
                PropertyPaneDropdown('pageSize', {
                  label: strings.PageSizeLabel,
                  options: [10, 25, 50, 100].map((n) => ({ key: n, text: String(n) })),
                  selectedKey: this.properties.pageSize || defaultSolutionConfig.listView.pageSize
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
