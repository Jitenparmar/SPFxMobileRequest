import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MobileRequestWebPartStrings';
import MobileRequest from './components/MobileRequest';
import { IMobileRequestProps } from './components/IMobileRequestProps';

export interface IMobileRequestWebPartProps {
  listName: string;
}

export default class MobileRequestWebPart extends BaseClientSideWebPart<IMobileRequestWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMobileRequestProps> = React.createElement(
      MobileRequest,
      {
        listName: this.properties.listName,
        spHttpClient: this.context.spHttpClient,  
        siteUrl: this.context.pageContext.web.absoluteUrl,
        context:this.context,  
      }
    );

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
      pages: [
        {
          header: {
            description: strings.ListNameFieldLabel
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
