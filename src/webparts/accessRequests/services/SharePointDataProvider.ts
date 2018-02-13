import {
  SPHttpClient,
  SPHttpClientBatch,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import INewAccessRequest from '../models/INewAccessRequest';
import IAccessRequest from '../models/IAccessRequest';
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
  private _listItemEntityTypeName: string = undefined;
  // private _committeesListTitle: string = 'UPHP Committees';
  // private _membersList: string = 'UPHP Members';
  // private _membersCommList: string = 'UPHP Member Committees';

  constructor(webPartContext: IWebPartContext) {
    this._webPartContext = webPartContext;
    this._listsUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists`;    
  }

  public set accessListTitle(value: string) {
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

  public getCurrentUser(): Promise<any> {
    return this._getCurrentUser(this.webPartContext.spHttpClient);
  }
  private _getCurrentUser(requester: SPHttpClient): Promise<any> {
    let options: ISPHttpClientOptions = {
      headers: {
        "accept": "application/json;odata=nometadata",
        "content-type": "application/json;odata=verbose"
      }
    };
    const queryUrl: string = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/currentuser`;
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
  public getItem(requestId: string): Promise<IAccessRequest> {
    return this._getItem(requestId, this.webPartContext.spHttpClient);
  }
  public _getItem(requestId: string, requester: SPHttpClient): Promise<IAccessRequest> {
    return new Promise<IAccessRequest>((resolve, reject) => {
        const queryUrl: string = `${this._listsUrl}/GetByTitle('${this._accessListTitle}')/items${requestId}`;
        requester.get(queryUrl, SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response: SPHttpClientResponse) => {
            if (response.ok) {
              response.json().then((data: any) => {
                const reqItem: IAccessRequest = {
                  Id: data.Id,
                    Title: data.Title,
                    Comments: data.Comments,
                    AddCommittees: data.AddCommittees.map(c => c.Title).join(","),
                    RemoveCommittees: data.RemoveCommittees.map(c => c.Title).join(","),
                    Created: data.Created,
                    EMail: data.Email,
                    FirstName: data.FirstName,
                    JobTitle: data.JobTitle,
                    LastName: data.LastName,
                    Modified: data.Modified,
                    Company: data.Company,
                    Office: data.Office,
                    RequestReason: data.RequestReason,
                    RequestStatus: data.RequestStatus,
                    AuthorId: data.AuthorId,
                    CreatedBy: data.Author.Title,
                    EditorId: data.EditorId
                };
                resolve(reqItem);
              })
                .catch((error) => { reject(error); });
            }
            else {
              reject(response);
            }
          })
          .catch((error) => { reject(error); });
    });
  }
  public getItemsForCurrentUser():Promise<IAccessRequest[]> {
    return this._getItemsForCurrentUser(this.webPartContext.spHttpClient);
  }
  private _getItemsForCurrentUser(requester: SPHttpClient):Promise<IAccessRequest[]> {
    return new Promise<IAccessRequest[]>((resolve, reject) => {
      this.getCurrentUser().then((result) => {
        let filterString: string = `spLoginName eq '${result.LoginName}' or EMail eq '${result.Email}' or AuthorId eq ${result.Id}`;
        filterString = "&$filter=" + encodeURIComponent(filterString);
        const queryString: string = `?$orderby=Id desc&$select=*,Author/Title,AddCommittees/Title,RemoveCommittees/Title&$expand=Author,AddCommittees,RemoveCommittees${filterString}`;
        const queryUrl: string = `${this._listsUrl}/GetByTitle('${this._accessListTitle}')/items${queryString}`;
        return requester.get(queryUrl, SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response: SPHttpClientResponse) => {
            if (response.ok) {
              response.json().then((data: any) => {
                const reqItems: IAccessRequest[] = data.value.map((item) => {
                  return {
                    Id: item.Id,
                    Title: item.Title,
                    Comments: item.Comments,
                    AddCommittees: item.AddCommittees.map(c => c.Title),
                    RemoveCommittees: item.RemoveCommittees.map(c => c.Title),
                    Created: new Date(item.Created).toLocaleDateString('en-US'),
                    EMail: item.EMail,
                    FirstName: item.FirstName,
                    JobTitle: item.JobTitle,
                    LastName: item.LastName,
                    Modified: item.Modified,
                    Company: item.Company,
                    Office: item.Office,
                    RequestReason: item.RequestReason,
                    RequestStatus: item.RequestStatus,
                    AuthorId: item.AuthorId,
                    CreatedBy: item.Author.Title,
                    EditorId: item.EditorId
                  };
                });
                resolve(reqItems);
              })
                .catch((error) => { reject(error); });
            }
            else {
              reject(response);
            }
          })
          .catch((error) => { reject(error); });
      });
    });
  }
  public getMembers(membersList: string, ): Promise<IModifyAccessRequest[]> {
    return this._getMembers(membersList, this.webPartContext.spHttpClient);
  }
  private _getMembers(membersList: string, requester: SPHttpClient): Promise<IModifyAccessRequest[]> {
    const queryString: string = '?$select=Id,spLoginName,Title,EMail';
    let options: ISPHttpClientOptions = {
      headers: {
        "accept": "application/json;odata=nometadata",
        "content-type": "application/json;odata=verbose"
      }
    };
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${membersList}')/items` + queryString;
    return new Promise<IModifyAccessRequest[]>((resolve, reject) => {
      requester.get(queryUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            response.json().then((data: any) => {
              const members: IModifyAccessRequest[] = data.value.map((item) => {
                return {
                  Id: item.Id,
                  spLoginName: item.spLoginName,
                  Title: item.Title,
                  EMail: item.EMail
                };
              });
              resolve(members);
            });
          }
          else {
            reject(response);
          }
        })
        .catch((error) => {reject(error);});
    });
  }
  public getMemberCommittees(membersCommList: string, loginName: any): Promise<any> {
    return this._getMemberCommittees(membersCommList, loginName, this.webPartContext.spHttpClient);
  }
  private _getMemberCommittees(membersCommList: string, loginName: string, requester: SPHttpClient): Promise<any> {
    let decodedName = loginName.replace(/#/g, "%23");
    const queryString: string = `?$select=Id,Title,CommitteeId&$filter=Title%20eq%20'${decodedName}'`;
    let options: ISPHttpClientOptions = {
      headers: {
        "accept": "application/json;odata=nometadata",
        "content-type": "application/json;odata=verbose"
      }
    };
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${membersCommList}')/items` + queryString;
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
  public getCommittees(committeesListTitle: string): Promise<any> {
    return this._getCommittees(committeesListTitle, this.webPartContext.spHttpClient);
  }
  private _getCommittees(committeesListTitle: string, requester: SPHttpClient): Promise<any> {
    const queryString: string = '?$select=Id,Title';
    let options: ISPHttpClientOptions = {
      headers: {
        "accept": "application/json;odata=nometadata",
        "content-type": "application/json;odata=verbose"
      }
    };
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${committeesListTitle}')/items` + queryString;

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
  public saveNewItem(newItem: INewAccessRequest): Promise<boolean> {
    return this._saveNewItem(newItem, this.webPartContext.spHttpClient);
  }
  private _saveNewItem(newItem: INewAccessRequest, requester: SPHttpClient): Promise<any> {
    let restUrl = this.accessListItemsUrl.replace("/items", "");
    const queryUrl: string = this.accessListItemsUrl;

    return this._getListItemEntityTypeName(this._accessListTitle, this.webPartContext.spHttpClient)
      .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
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
        return requester.post(queryUrl, SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: body
          })
          .then((postResponse: SPHttpClientResponse) => {
            return (postResponse);
          })
          .catch((error) => {
            return error;
          });
      });
  }
  public saveModifyRequest(item: IModifyAccessRequest): Promise<any> {
    return this._saveModifyRequest(item, this.webPartContext.spHttpClient);
  }
  private _saveModifyRequest(item: IModifyAccessRequest, requester: SPHttpClient): Promise<any> {
    const requestReason: string = item.RequestReason == 'Terminate' ? 'Terminate' : 'Change';
    let restUrl = this.accessListItemsUrl.replace("/items", "");
    const queryUrl: string = this.accessListItemsUrl;

    return this._getListItemEntityTypeName(this._accessListTitle, this.webPartContext.spHttpClient)
      .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
        const body: string = JSON.stringify({
          '__metadata': {
            'type': listItemEntityTypeName
          },
          'FirstName': item.FirstName,
          'LastName': item.LastName,
          'Comments': item.Comments,
          'Title': item.Title,
          'EMail': item.EMail,
          'RequestReason': requestReason,
          'RequestStatus': 'New',
          'spLoginName': item.spLoginName,
          'AddCommitteesId': {
            'results': item.AddCommittees ? item.AddCommittees : [],
          },
          'RemoveCommitteesId': {
            'results': item.RemoveCommittees ? item.RemoveCommittees : [],
          }
        });
        return requester.post(queryUrl, SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: body
          })
          .then((postResponse: SPHttpClientResponse) => {
            return (postResponse);
          })
          .catch((error) => {
            return error;
          });
      });

  }
  private _getListItemEntityTypeName(listName: string, requester: SPHttpClient): Promise<string> {
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