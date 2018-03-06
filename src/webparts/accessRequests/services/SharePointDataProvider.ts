import {
  SPHttpClient,
  SPHttpClientBatch,
  SPHttpClientResponse,
  ISPHttpClientOptions,
  SPHttpClientConfiguration,
  SPHttpClientBatchConfiguration
} from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import INewAccessRequest from '../models/INewAccessRequest';
import IAccessRequest from '../models/IAccessRequest';
import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';
import IModifyAccessRequest from "../models/IModifyAccessRequest";
import ITask from '../models/ITask';
// import SPHttpClientConfiguration from '@microsoft/sp-http/lib/spHttpClient/SPHttpClientConfiguration';
// import SPHttpClientBatchConfiguration from '@microsoft/sp-http/lib/spHttpClient/SPHttpClientBatchConfiguration';
import IFinalTask from '../models/IFinalTask';

export default class SharePointDataProvider implements IAccessRequestsDataProvider {
  private _accessListTitle: string;
  private _listsUrl: string;
  private _accessListItemsUrl: string;
  private _webPartContext: IWebPartContext;
  private _lastListName: string = undefined;
  private _listItemEntityTypeName: string = undefined;

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
  private async _getCurrentUser(requester: SPHttpClient): Promise<any> {
    let options: ISPHttpClientOptions = {
      headers: {
        "accept": "application/json;odata=nometadata",
        "content-type": "application/json;odata=verbose"
      }
    };
    const queryUrl: string = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/currentuser`;
    try {
      const response: SPHttpClientResponse = await requester.get(queryUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
      if (!response.ok && response.status !== 200) {
        throw new Error(response.statusMessage);
      }
      else {
        const results = await response.json();
        return results;
      }
    }
    catch (error) {
      throw new Error(error.message);
    }
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
  public getItemsForCurrentUser(currentUser?: any): Promise<IAccessRequest[]> {
    return this._getItemsForCurrentUser(this.webPartContext.spHttpClient);
  }
  private async _getItemsForCurrentUser(requester: SPHttpClient, currentUser?: any): Promise<IAccessRequest[]> {
    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;

    // const response: Promise<any> = await this.getCurrentUser();
    // const user = await response;
    let filterString: string = `spLoginName eq '${user.LoginName}' or EMail eq '${user.Email}' or AuthorId eq ${user.Id}`;
    filterString = "&$filter=" + encodeURIComponent(filterString);
    const queryString: string = `?$orderby=Id desc&$select=*,Author/Title,AddCommittees/Title,RemoveCommittees/Title&$expand=Author,AddCommittees,RemoveCommittees${filterString}`;
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${this._accessListTitle}')/items${queryString}`;
    try {
      const qryResponse: SPHttpClientResponse = await requester.get(queryUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
      if (!qryResponse.ok) {
        throw new Error(qryResponse.statusText + ": " + qryResponse.statusMessage);
      }
      const data = await qryResponse.json();
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
      return reqItems;
    }
    catch (error) {
      throw new Error(error);
    }
  }
  public getMembers(membersList: string, ): Promise<IModifyAccessRequest[]> {
    return this._getMembers(membersList, this.webPartContext.spHttpClient);
  }
  private async _getMembers(membersList: string, requester: SPHttpClient): Promise<IModifyAccessRequest[]> {
    const queryString: string = '?$select=Id,spLoginName,Title,EMail';
    let options: ISPHttpClientOptions = {
      headers: {
        "accept": "application/json;odata=nometadata",
        "content-type": "application/json;odata=verbose"
      }
    };
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${membersList}')/items` + queryString;
    try {
      const response: SPHttpClientResponse = await requester.get(queryUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
      if (!response.ok) {
        throw new Error(response.statusText + ": " + response.statusMessage);
      }
      const data = await response.json();
      const members: IModifyAccessRequest[] = data.value.map((item) => {
        return {
          Id: item.Id,
          spLoginName: item.spLoginName,
          Title: item.Title,
          EMail: item.EMail
        };
      });
      return members;
    }
    catch (error) {
      throw new Error(error);
    }
  }
  public getMemberCommittees(membersCommList: string, loginName: any): Promise<any> {
    return this._getMemberCommittees(membersCommList, loginName, this.webPartContext.spHttpClient);
  }
  private async _getMemberCommittees(membersCommList: string, loginName: string, requester: SPHttpClient): Promise<any> {
    let decodedName = loginName.replace(/#/g, "%23");
    const queryString: string = `?$select=Id,Title,CommitteeId&$filter=Title%20eq%20'${decodedName}'`;
    let options: ISPHttpClientOptions = {
      headers: {
        "accept": "application/json;odata=nometadata",
        "content-type": "application/json;odata=verbose"
      }
    };
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${membersCommList}')/items` + queryString;
    try {
      const response: SPHttpClientResponse = await requester.get(queryUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
      return await response.json();
    }
    catch (error) {
      throw new Error(error);
    }
  }
  public getCommittees(committeesListTitle: string): Promise<any> {
    return this._getCommittees(committeesListTitle, this.webPartContext.spHttpClient);
  }
  private async _getCommittees(committeesListTitle: string, requester: SPHttpClient): Promise<any> {
    const queryString: string = '?$select=Id,Title';
    let options: ISPHttpClientOptions = {
      headers: {
        "accept": "application/json;odata=nometadata",
        "content-type": "application/json;odata=verbose"
      }
    };
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${committeesListTitle}')/items` + queryString;

    const response: SPHttpClientResponse = await requester.get(queryUrl, SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      });
    return await response.json();
  }
  public saveNewItem(newItem: INewAccessRequest): Promise<boolean> {
    return this._saveNewItem(newItem, this.webPartContext.spHttpClient);
  }
  private async _saveNewItem(newItem: INewAccessRequest, requester: SPHttpClient): Promise<any> {
    let restUrl = this.accessListItemsUrl.replace("/items", "");
    const queryUrl: string = this.accessListItemsUrl;
    try {
      const listItemEntityTypeName = await this._getListItemEntityTypeName(this._accessListTitle, this.webPartContext.spHttpClient);
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
        'RequestReason': 'New access',
        'RequestStatus': `${this._getFormattedDate(new Date())} New request`,
        'AddCommitteesId': {
          'results': newItem.Committees
        }
      });
      const response: SPHttpClientResponse = await requester.post(queryUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: body
        });
      if (!response.ok && response.status !== 204) {
        throw new Error(response.statusMessage);
      }
      return Promise.resolve("ok");
    }
    catch (error) {
      throw new Error(error);
    }
  }
  public saveChangeRequest(item: IModifyAccessRequest): Promise<any> {
    return this._saveChangeRequest(item, this.webPartContext.spHttpClient);
  }
  private async _saveChangeRequest(item: IModifyAccessRequest, requester: SPHttpClient): Promise<any> {
    const queryUrl: string = this.accessListItemsUrl;
    let requests = [];
    if (item.RequestReason == "Terminate") {
      requests.push("Terminate");
    }
    else {
      if (item.AddCommittees || (item.RemoveCommittees == null || item.RemoveCommittees.length == 0)) {
        requests.push("Add access");
      }
      if (item.RemoveCommittees != null && item.RemoveCommittees.length > 0) {
        requests.push("Remove access");
      }
    }

    const spBatch: SPHttpClientBatch = requester.beginBatch();
    const postResponses: Promise<SPHttpClientResponse>[] = [];

    const entityTypeName = await this._getListItemEntityTypeName(this._accessListTitle, this.webPartContext.spHttpClient);

    const postHeaders = {
      //'Accept': 'application/json;odata=verbose',
      'Content-type': 'application/json;odata=verbose',
      'odata-version': ''
    };
    for (const req of requests) {
      let body: any = {
        '@data.type': `${entityTypeName}`,
        'FirstName': item.FirstName,
        'LastName': item.LastName,
        'Comments': item.Comments,
        'Title': item.Title,
        'EMail': item.EMail,
        'RequestReason': req,
        'RequestStatus': `${this._getFormattedDate(new Date())} New request`,
        'spLoginName': item.spLoginName,
        'AddCommitteesId': [],
        'RemoveCommitteesId': [],
      };
      switch (req) {
        case 'Add access':
          body.AddCommitteesId = item.AddCommittees ? item.AddCommittees : [];
          break;
        case 'Remove access':
          body.RemoveCommitteesId = item.RemoveCommittees ? item.RemoveCommittees : [];
          break;
        default:
          break;
      }
      const postResponse: Promise<SPHttpClientResponse> = spBatch.post(queryUrl, SPHttpClientBatch.configurations.v1,
        { body: JSON.stringify(body) });
      postResponses.push(postResponse);
    }
    try {
      await spBatch.execute();
      for (let response of postResponses) {
        let itemResponse: SPHttpClientResponse = await response;
        if (!itemResponse.ok && itemResponse.status !== 201) {
          throw new Error(itemResponse.statusMessage);
        }
      }
      return Promise.resolve("ok");
    }
    catch (error) {
      console.log(error);
      throw new Error(error.message);
    }
  }
  public getTasksForCurrentUser(requestsByCommList: string, currentUser?: any): Promise<ITask[]> {
    return this._getTasksForCurrentUser(requestsByCommList, this.webPartContext.spHttpClient);
  }
  private async _getTasksForCurrentUser(requestsByCommList: string, requester: SPHttpClient, currentUser?: any): Promise<ITask[]> {
    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;
    let filterString: string = `substringof('${user.LoginName}',Approvers) and Outcome ne 'Approved' and Outcome ne 'Rejected'`;
    filterString = "&$filter=" + encodeURIComponent(filterString);
    const queryString: string = `?$orderby=Id desc&$select=Id,RequestStatus,RequestId/Title,CompletionStatus,Outcome,Created,Modified,Committee/Title&$expand=Committee/Title,RequestId/Title${filterString}`;
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${requestsByCommList}')/items${queryString}`;
    try {
      const qryResponse: SPHttpClientResponse = await requester.get(queryUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
      if (!qryResponse.ok) {
        throw new Error(qryResponse.statusText + ": " + qryResponse.statusMessage);
      }
      const data = await qryResponse.json();
      const reqItems: ITask[] = data.value.map((item) => {
        return {
          Id: item.Id,
          Name: item.RequestId.Title,
          Committee: item.Committee.Title,
          RequestStatus: item.RequestStatus,
          CompletionStatus: item.CompletionStatus,
          Outcome: item.Outcome,
          Created: new Date(item.Created).toLocaleDateString('en-US'),
          Modified: new Date(item.Modified).toLocaleDateString('en-US'),
          Updating: false,
          CurrentUser: user.Id
        };
      });
      return reqItems;
    }
    catch (error) {
      throw new Error(error);
    }
  }
  public updateForCommittee(item: ITask, requestsByCommList: string, currentUser?: any): Promise<boolean> {
    return this._updateForCommittee(item, requestsByCommList, this.webPartContext.spHttpClient, currentUser);
  }
  private async _updateForCommittee(item: ITask, requestsByCommList: string, requester: SPHttpClient, currentUser?: any): Promise<boolean> {
    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;

    const approvalsUrl = `https://prod-42.westus.logic.azure.com/workflows/13e4c8e31da946cba3e82c96d67446a0/triggers/manual/paths/invoke/items/${item.Id}?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TMc_96taHAQpZQhCBJ6Vg_YIGAfNtXbyxZHlBmGYJOo`;
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${requestsByCommList}')/items(${item.Id})`;
    const entityTypeName = await this._getListItemEntityTypeName(requestsByCommList, requester);
    const body: string = JSON.stringify({
      '@data.type': entityTypeName,
      'Outcome': item.Outcome,
      'RequestStatus': `${this._getFormattedDate(new Date)} ${item.Outcome} by ${user.Title}\n${item.RequestStatus}`,
      'CompletionStatus': 'Completed',
      'ApprovalComments': item.ApprovalComments
    });
    const headers: Headers = new Headers();
    headers.append('If-Match', '*');
    try {
      const fetchResponse = await requester.fetch(queryUrl, SPHttpClient.configurations.v1,
        {
          body: body,
          headers,
          method: 'PATCH'
        });
      if (!fetchResponse.ok || fetchResponse.status !== 204) {
        throw new Error(fetchResponse.statusMessage);
      }
      else {
        const response = await requester.get(approvalsUrl, SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
        if (!response.ok || response.status !== 202) {
          throw new Error(response.statusMessage);
        }
        else {
          return Promise.resolve(true);
        }
      }
    }
    catch (error) {
      console.log(error);
      throw new Error(error.message);
    }
  }
  public updateForRequest(item: IFinalTask, currentUser?: any) {
    return this._updateForRequest(item, this.webPartContext.spHttpClient, currentUser);
  }
  private async _updateForRequest(item: IFinalTask, requester: SPHttpClient, currentUser?: any) {
    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;

    const approvalsUrl = `https://prod-61.westus.logic.azure.com/workflows/a957592735af43f5bec601631b68632d/triggers/manual/paths/invoke/request/${item.Id}?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=FP27ZBs5kmEUD0rAFVmz_vFxoHUf5mkEwlc-5C0nRcI`;
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${this._accessListTitle}')/items(${item.Id})`;
    const entityTypeName = await this._getListItemEntityTypeName(this._accessListTitle, requester);
    const body: string = JSON.stringify({
      '@data.type': entityTypeName,
      'CompletionStatus': item.CompletionStatus,
      'RequestStatus': `${this._getFormattedDate(new Date)} ${item.CompletionStatus} by ${user.Title}\n${item.RequestStatus}`,
      'ApprovalComments': item.ApprovalComments
    });
    const headers: Headers = new Headers();
    headers.append('If-Match', '*');
    try {
      const postResponse = await requester.fetch(queryUrl, SPHttpClient.configurations.v1,
        {
          body: body,
          headers,
          method: 'PATCH'
        });
      if (!postResponse.ok || postResponse.status !== 204) {
        throw new Error(postResponse.statusMessage);
      }
      else {
        const response = await requester.get(approvalsUrl, SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
        if (!response.ok || response.status !== 202) {
          throw new Error(response.statusMessage);
        }
        else {
          return Promise.resolve(true);
        }
      }
    }
    catch (error) {
      console.log(error);
      throw new Error(error.message);
    }
  }
  public getFinalTasks(requestsByCommList: string, currentUser?: any): Promise<IFinalTask[]> {
    return this._getFinalTasks(requestsByCommList, this.webPartContext.spHttpClient);
  }
  public async _getFinalTasks(requestsByCommList: string, requester: SPHttpClient, currentUser?: any): Promise<IFinalTask[]> {
    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;

    const reqItems: IFinalTask[] = [];
    const commTasks: ITask[] = [];
    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'odata-version': ''
    };

    let filterString: string = `CompletionStatus eq 'Pending final Approval'`;
    filterString = "&$filter=" + encodeURIComponent(filterString);
    let queryString: string = `?$select=Id,Title,Company,JobTitle,RequestReason,Comments,RequestStatus${filterString}`;
    const ReqsUrl: string = `${this._listsUrl}/GetByTitle('${this._accessListTitle}')/items${queryString}`;
    try {
      const reqTasks = await requester.fetch(ReqsUrl, SPHttpClient.configurations.v1, { headers: headers });
      if (!reqTasks.ok && reqTasks.status !== 200) {
        throw new Error(reqTasks.statusMessage);
      }
      const requestsResults = await reqTasks.json();
      const requests = requestsResults.value;

      if (requests != null && requests.length > 0) {
        const spBatch: SPHttpClientBatch = requester.beginBatch();
        const commsResponses: Promise<SPHttpClientResponse>[] = [];
        for (let request of requests) {
          filterString = `RequestId eq ${request.Id} and CompletionStatus eq 'Completed'`;
          filterString = "&$filter=" + encodeURIComponent(filterString);
          queryString = `?$select=Id,Title,RequestId/Id,Committee/Title,Outcome,ApprovalComments,RequestStatus&$expand=Committee/Title,RequestId${filterString}`;
          let commUrl: string = `${this._listsUrl}/GetByTitle('${requestsByCommList}')/items${queryString}`;
          commsResponses.push(spBatch.get(commUrl, SPHttpClientBatch.configurations.v1));
        }
        try {
          await spBatch.execute();
          for (let response of commsResponses) {
            const result: SPHttpClientResponse = await response;
            if (!result.ok && result.status !== 200) {
              throw new Error(result.statusMessage);
            }
            const commReq = await result.json();
            for (let task of commReq.value) {
              commTasks.push({
                Id: task.Id,
                RequestId: task.RequestId.Id,
                Name: task.Title,
                Committee: task.Committee.Title,
                RequestStatus: task.RequestStatus,
                Outcome: task.Outcome,
                CompletionStatus: task.CompletionStatus,
                Created: task.Created,
                Modified: task.Modified,
                CurrentUser: user.Id
              });
            }
          }
        }
        catch (error) {
          throw new Error(error.message);
        }
      }
      for (let request of requests) {
        const reqItem: IFinalTask = {
          Id: request.Id,
          Title: request.Title,
          Comments: request.Comments,
          JobTitle: request.JobTitle,
          Office: request.Company,
          RequestReason: request.RequestReason,
          RequestStatus: request.RequestStatus,
          CompletionStatus: request.CompletionStatus,
          CommitteeTasks: commTasks.filter((t) => t.RequestId == request.Id)
        };
        reqItems.push(reqItem);
      }

    }
    catch (error) {
      console.log(error);
      throw new Error(error.message);
    }
    return reqItems;
  }
  private async _getListItemEntityTypeName(listName: string, requester: SPHttpClient): Promise<string> {

    if (listName == this._lastListName && this._listItemEntityTypeName) {
      return Promise.resolve(this._listItemEntityTypeName);
    }
    try {
      const response: SPHttpClientResponse = await requester.get(`${this._listsUrl}/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
      if (!response.ok && response.status !== 200) {
        throw new Error(response.statusMessage);
      }
      else {
        const results = await response.json();
        this._listItemEntityTypeName = results.ListItemEntityTypeFullName;
        this._lastListName = listName;
        Promise.resolve(this._listItemEntityTypeName);
      }
    }
    catch (error) {
      throw new Error(error);
    }
  }
  private _getFormattedDate(d) {
    let thisDay = (d.getDate() < 10 ? '0' : '') + d.getDate();
    let thisMonth = (d.getMonth() < 10 ? '0' : '') + (d.getMonth() + 1);
    return d.getFullYear() + '-' + thisMonth + '-' + thisDay;
  }
}