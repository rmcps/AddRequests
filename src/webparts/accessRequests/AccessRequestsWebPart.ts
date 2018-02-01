import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'AccessRequestsWebPartStrings';
import DefaultPage from './components/DefaultPage';
import { IAccessRequestsProps } from './components/IAccessRequestsProps';

export interface IAccessRequestsWebPartProps {
  description: string;
}

export default class AccessRequestsWebPart extends BaseClientSideWebPart<IAccessRequestsWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit();
  }
  public render(): void {
    const element: React.ReactElement<IAccessRequestsProps > = React.createElement(
      DefaultPage,
      {
        description: this.properties.description,
        context:this.context,
        dom: this.domElement,
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
                }
              )
          ]
            }
          ]
        }
      ]
    };
  }
}
