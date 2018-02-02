import {
    SPHttpClient,
    SPHttpClientBatch,
    SPHttpClientResponse,
    ISPHttpClientOptions
  } from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import INewAccessRequest from '../models/INewAccessRequest';
import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';
import IModifyAccessRequest from "../models/IModifyAccessRequest";
import SPHttpClientConfiguration from '@microsoft/sp-http/lib/spHttpClient/SPHttpClientConfiguration';
import SPHttpClientBatchConfiguration from '@microsoft/sp-http/lib/spHttpClient/SPHttpClientBatchConfiguration';

  export default class SharePointDataProvider implements IAccessRequestsDataProvider {
    private _accessListTitle: string;
    private _listsUrl: string;
    private _accessListItemsUrl: string;
    private _webPartContext: IWebPartContext;
    private _lastListName: string = undefined;
    private _listItemEntityTypeName:string = undefined;
    private _committeesListTitle:string = 'UPHP Committees';
  
    public set accessListTitle(value:string) {
      this._accessListTitle = value;
    }

    public get accessListTitle() {
      return this._accessListTitle;
    }    

    public get accessListItemsUrl() {
      return this._accessListItemsUrl = `${this._listsUrl}/GetByTitle('${this._accessListTitle}')/items`;
    }
    
    public set webPartContext(value: IWebPartContext) {
      this._webPartContext = value;
      this._listsUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists`;
    }
  
    public get webPartContext(): IWebPartContext {
      return this._webPartContext;
    }

  public getItem(listId: number): Promise<IModifyAccessRequest> {
    return null;
  }
    public getMembers(): Promise<any> {
      return this._getMembers(this.webPartContext.spHttpClient);
    }

    private _getMembers(requester: SPHttpClient): Promise<any> {
      return null;
    }
    public getMemberCommittees(Id: any): Promise<any[]> {
      return null;
  }
    public getCommittees():Promise<any> {
      return this._getCommittees(this.webPartContext.spHttpClient);
    }
    private _getCommittees(requester:SPHttpClient):Promise<any> {
      const queryString: string = '?$select=Id,Title';
      let options: ISPHttpClientOptions = { 
        headers: { "accept": "application/json;odata=nometadata",
        "content-type": "application/json;odata=verbose" }
      };
      const queryUrl: string = `${this._listsUrl}/GetByTitle('${this._committeesListTitle}')/items` + queryString;
  
      return requester.get(queryUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    }  
    public saveNewItem(newItem:INewAccessRequest):Promise<boolean> {
      return this._saveNewItem(newItem,this.webPartContext.spHttpClient);
    }
    private _saveNewItem(newItem:INewAccessRequest, requester:SPHttpClient):Promise<any> {
      let restUrl = this.accessListItemsUrl.replace("/items","");
      const queryUrl: string = this.accessListItemsUrl;
      
      return this._getListItemEntityTypeName(this._accessListTitle,this.webPartContext.spHttpClient)
        .then ((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
            const body: string = JSON.stringify({
              '__metadata': {
                'type': listItemEntityTypeName
              },
              'FirstName': newItem.FirstName,
              'LastName': newItem.LastName,
              'EMail': newItem.EMail,
              'JobTitle': newItem.JobTitle,
              'Company': newItem.Company,
              'Office': newItem.Office,
              'Comments': newItem.Comments,
              'Title': `${newItem.FirstName} ${newItem.LastName}`,
              'RequestReason': 'New member',
              'RequestStatus': 'New',
              'AddCommitteesId': {
                'results': newItem.Committees
              }
            });       
          return requester.post(queryUrl,SPHttpClient.configurations.v1,
              {
                headers: {
                  'Accept': 'application/json;odata=nometadata',
                  'Content-type': 'application/json;odata=verbose',
                  'odata-version': ''
                },
                body: body
              })
            .then((postResponse: SPHttpClientResponse) => {
              return(postResponse);
            })
            .catch((error) => {
              return error;
            });
        });
  }
    private _getListItemEntityTypeName(listName:string, requester:SPHttpClient): Promise<string> {
      return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
        if (listName == this._lastListName && this._listItemEntityTypeName) {
          resolve(this._listItemEntityTypeName);
          return;
        }
        requester.get(`${this._listsUrl}/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          })
          .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
            return response.json();
          }, (error: any): void => {
            reject(error);
          })
          .then((response: { ListItemEntityTypeFullName: string }): void => {
            this._listItemEntityTypeName = response.ListItemEntityTypeFullName;
            this._lastListName = listName;
            resolve(this._listItemEntityTypeName);
          });
      });
    }
}