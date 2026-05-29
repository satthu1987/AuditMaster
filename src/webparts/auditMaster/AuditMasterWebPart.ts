import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import AuditMaster from './components/AuditMaster';
import { IAuditMasterProps } from './models';

export interface IAuditMasterWebPartProps {
  description: string;
}

export default class AuditMasterWebPart extends BaseClientSideWebPart<IAuditMasterWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAuditMasterProps> = React.createElement(
      AuditMaster,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl
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
            description: 'Audit Master Settings'
          },
          groups: [
            {
              groupName: 'General',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
