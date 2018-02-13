import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'AccessRequestsWebPartStrings';
import IDefaultProps from './components/IDefaultProps'
import DefaultPage from './components/DefaultPage';

export interface IAccessRequestsWebPartProps {
  requestsList: string;
  membersList: string;
  committeesList: string;
  membersCommitteesList: string;
}

export default class AccessRequestsWebPart extends BaseClientSideWebPart<IAccessRequestsWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit();
  }
  public render(): void {
    const element: React.ReactElement<IDefaultProps > = React.createElement(
      DefaultPage,
      {
        requestsList: this.properties.requestsList,
        membersList: this.properties.membersList,
        committeesList: this.properties.committeesList,
        membersCommitteesList: this.properties.membersCommitteesList,
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
                PropertyPaneTextField('requestsList', {
                  label: strings.RequestListFieldLabel
                }
              ),
              PropertyPaneTextField('membersList', {
                label: strings.MembersListFieldLabel
              }
            ),
            PropertyPaneTextField('committeesList', {
              label: strings.CommitteesList
            }
          ),
          PropertyPaneTextField('membersCommitteesList', {
            label: strings.MembersListFieldLabel
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
