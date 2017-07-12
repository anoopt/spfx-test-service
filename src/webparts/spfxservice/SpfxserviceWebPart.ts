import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'spfxserviceStrings';
import Spfxservice from './components/Spfxservice';
import { ISpfxserviceProps } from './components/ISpfxserviceProps';
import { ISpfxserviceWebPartProps } from './ISpfxserviceWebPartProps';

export default class SpfxserviceWebPart extends BaseClientSideWebPart<ISpfxserviceWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxserviceProps > = React.createElement(
      Spfxservice,
      {
        serviceScope: this.context.serviceScope,
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
