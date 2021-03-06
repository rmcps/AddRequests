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
  public getItem(requestId: string, requestsByCommList: string): Promise<IAccessRequest> {
    return this._getItem(requestId, requestsByCommList, this.webPartContext.spHttpClient);
  }
  public async _getItem(requestId: string, requestsByCommList: string, requester: SPHttpClient): Promise<IAccessRequest> {
    let commArr: ITask[] = [];
    const queryReqUrl: string = `${this._listsUrl}/GetByTitle('${this._accessListTitle}')/items(${requestId})?$select=*,Author/Title,ApprovedBy/Title,AddCommittees/Title,RemoveCommittees/Title&$expand=Author,ApprovedBy,AddCommittees,RemoveCommittees`;
    let filterString: string = `RequestId eq ${requestId}`;
    filterString = "$filter=" + encodeURIComponent(filterString);
    const queryCommUrl: string = `${this._listsUrl}/GetByTitle('${requestsByCommList}')/items?${filterString}&$select=*,Committee/Title,ApprovedBy/Title&$expand=Committee/Title,ApprovedBy/Title`;
    const spBatch: SPHttpClientBatch = requester.beginBatch();
    const batchArr: Promise<SPHttpClientResponse>[] = [];
    try {
      batchArr.push(spBatch.get(queryReqUrl, SPHttpClientBatch.configurations.v1));
      batchArr.push(spBatch.get(queryCommUrl, SPHttpClientBatch.configurations.v1));

      await spBatch.execute();
      let response: SPHttpClientResponse = await batchArr[0];
      if (!response.ok && response.status !== 200) {
        throw new Error(response.statusMessage);
      }
      const reqData = await response.json();
      response = await batchArr[1];
      if (!response.ok && response.status !== 200) {
        throw new Error(response.statusMessage);
      }
      const commData = await response.json();

      for (let item of commData.value) {
        commArr.push({
          Id: item.Id,
          RequestId: item.RequestIdId,
          Name: reqData.Title,
          RequestType: item.Title,
          Committee: item.Committee.Title,
          RequestStatus: item.RequestStatus,
          Outcome: item.Outcome,
          ApprovedBy: item.ApprovedBy ? item.ApprovedBy.Title : '',
          ApprovalComments: item.ApprovalComments,
          CompletionStatus: item.CompletionStatus,
          Created: item.Created,
          Modified: item.Modified,
          CurrentUser: null
        });
      }
      const reqItem: IAccessRequest = {
        Id: reqData.Id,
        Title: reqData.Title,
        Comments: reqData.Comments,
        AddCommittees: reqData.AddCommittees.map(c => c.Title),
        RemoveCommittees: reqData.RemoveCommittees.map(c => c.Title),
        Created: reqData.Created,
        EMail: reqData.EMail,
        FirstName: reqData.FirstName,
        JobTitle: reqData.JobTitle,
        LastName: reqData.LastName,
        Modified: reqData.Modified,
        Company: reqData.Company,
        Office: reqData.Office,
        RequestReason: reqData.RequestReason,
        RequestStatus: reqData.RequestStatus,
        CompletionStatus: reqData.CompletionStatus,
        Outcome: reqData.Outcome,
        ApprovedBy: reqData.ApprovedBy ? reqData.ApprovedBy.Title : '',
        ApprovalComments: reqData.ApprovalComments,
        AuthorId: reqData.AuthorId,
        CreatedBy: reqData.Author.Title,
        EditorId: reqData.EditorId,
        CommitteeApprovals: commArr
      };
      return reqItem;
    }
    catch (error) {
      throw new Error(error.message);
    }
  }
  public getItemsForCurrentUser(currentUser?: any): Promise<IAccessRequest[]> {
    return this._getItemsForCurrentUser(this.webPartContext.spHttpClient);
  }
  private async _getItemsForCurrentUser(requester: SPHttpClient, currentUser?: any): Promise<IAccessRequest[]> {
    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;

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
          CompletionStatus: item.CompletionStatus,
          Outcome: item.Outcome,
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
    const queryString: string = '?$top=1000&$select=Id,spLoginName,Title,EMail';
    let options: ISPHttpClientOptions = {
      headers: {
        "accept": "application/json;odata=nometadata" //,
        //"content-type": "application/json;odata=verbose"
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
      const listItemEntityTypeName = await this._getListItemEntityTypeName(this._accessListTitle, requester);
      if (listItemEntityTypeName === '' || listItemEntityTypeName.length === 0) {
        throw new Error('Failed to retrieve ListItemEntityTypeFullName');
      }
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
        'RequestStatus': `${this._getFormattedDateTime(new Date())} New request`,
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
    if (item.RequestReason === "Terminate") {
      requests.push("Terminate");
    }
    else {
      if (item.AddCommittees !== null && item.AddCommittees.length > 0) {
        requests.push("Add access");
      }
      if (item.RemoveCommittees !== null && item.RemoveCommittees.length > 0) {
        requests.push("Remove access");
      }
      if (requests.length == 0) {
        throw new Error("No reason for request chosen.");
      }
    }

    const spBatch: SPHttpClientBatch = requester.beginBatch();
    const postResponses: Promise<SPHttpClientResponse>[] = [];

    const entityTypeName = await this._getListItemEntityTypeName(this._accessListTitle, requester);

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
        'RequestStatus': `${this._getFormattedDateTime(new Date())} New request`,
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
  public getTasksForCurrentUser(requestsByCommList: string, alltasks: boolean, currentUser?: any): Promise<ITask[]> {
    return this._getTasksForCurrentUser(requestsByCommList, this.webPartContext.spHttpClient, alltasks);
  }
  private async _getTasksForCurrentUser(requestsByCommList: string, requester: SPHttpClient, alltasks: boolean, currentUser?: any): Promise<ITask[]> {
    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;
    let filterString: string = "";
    if (!alltasks) {
      filterString = ` and substringof('${user.LoginName}',Approvers)`;
    }
    filterString = `Outcome ne 'Approved' and Outcome ne 'Rejected' and Outcome ne 'Canceled'${filterString}`;
    filterString = "&$filter=" + encodeURIComponent(filterString);
    const queryString: string = `?$orderby=Id desc&$select=Id,Title,RequestStatus,RequestIdId,RequestId/Title,CompletionStatus,Outcome,Created,Modified,Committee/Title&$expand=Committee/Title,RequestId/Title${filterString}`;
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
          RequestId: item.RequestIdId,
          Name: item.RequestId.Title,
          Committee: item.Committee.Title,
          RequestType: item.Title,
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
  public updateCommitteeTaskItem(item: ITask, requestsByCommList: string, currentUser?: any): Promise<boolean> {
    return this._updateCommitteeTaskItem(item, requestsByCommList, this.webPartContext.spHttpClient, currentUser);
  }
  private async _updateCommitteeTaskItem(item: ITask, requestsByCommList: string, requester: SPHttpClient, currentUser?: any): Promise<boolean> {
    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;
    const approvalsUrl = `https://prod-42.westus.logic.azure.com/workflows/13e4c8e31da946cba3e82c96d67446a0/triggers/manual/paths/invoke/items/${item.Id}?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TMc_96taHAQpZQhCBJ6Vg_YIGAfNtXbyxZHlBmGYJOo`;
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${requestsByCommList}')/items(${item.Id})`;
    const entityTypeName = await this._getListItemEntityTypeName(requestsByCommList, requester);
    const body: string = JSON.stringify({
      '@data.type': entityTypeName,
      'Outcome': item.Outcome,
      'RequestStatus': `${this._getFormattedDateTime(new Date)} ${item.Outcome} by ${user.Title}\n${item.RequestStatus}`,
      'CompletionStatus': 'Completed',
      'ApprovalComments': item.ApprovalComments,
      'ApprovedById': user.Id
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
        return Promise.resolve(true);
        // const response = await requester.get(approvalsUrl, SPHttpClient.configurations.v1,
        //   {
        //     headers: {
        //       'Accept': 'application/json;odata=nometadata',
        //       'odata-version': ''
        //     }
        //   });
        // if (!response.ok || response.status !== 202) {
        //   throw new Error(response.statusMessage);
        // }
        // else {
        //   return Promise.resolve(true);
        // }
      }
    }
    catch (error) {
      console.log(error);
      throw new Error(error.message);
    }
  }
  public updateAllCommitteeTaskItems(items: ITask[], requestsByCommList: string, currentUser?: any): Promise<boolean> {
    return this._updateAllCommitteeTaskItems(items, requestsByCommList, this.webPartContext.spHttpClient, currentUser);
  }
  private async _updateAllCommitteeTaskItems(items: ITask[], requestsByCommList: string, requester: SPHttpClient, currentUser?: any): Promise<boolean> {
    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;

    const queryUrl: string = `${this._listsUrl}/GetByTitle('${requestsByCommList}')/items`;
    const headers: Headers = new Headers();
    headers.append('If-Match', '*');
    const entityTypeName = await this._getListItemEntityTypeName(requestsByCommList, requester);

    for (let item of items) {
      const body: string = JSON.stringify({
        '@data.type': entityTypeName,
        'Outcome': item.Outcome,
        'RequestStatus': `${this._getFormattedDateTime(new Date)} ${item.Outcome} by ${user.Title}\n${item.RequestStatus}`,
        'CompletionStatus': 'Completed',
        'ApprovalComments': item.ApprovalComments,
        'ApprovedById': user.Id
      });
      let itemQueryUrl = `${queryUrl}(${item.Id})`;
      try {
        const fetchResponse = await requester.fetch(itemQueryUrl, SPHttpClient.configurations.v1,
          {
            body: body,
            headers,
            method: 'PATCH'
          });
        if (!fetchResponse.ok || fetchResponse.status !== 204) {
          throw new Error(fetchResponse.statusMessage);
        }
        // else {
        //   return Promise.resolve(true);

        // }
      }
      catch (error) {
        console.log(error);
        throw new Error(error.message);
      }
    }
    return Promise.resolve(true);
  }
  public updateForRequest(items: IFinalTask[], currentUser?: any) {
    return this._updateForRequest(items, this.webPartContext.spHttpClient, currentUser);
  }
  private async _updateForRequest(items: IFinalTask[], requester: SPHttpClient, currentUser?: any) {
    // sleep
    //await this._sleep(5000);

    let user: any = currentUser == null ? await this.getCurrentUser() : currentUser;
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${this._accessListTitle}')/items`;
    const entityTypeName = await this._getListItemEntityTypeName(this._accessListTitle, requester);
    const headers: Headers = new Headers();
    headers.append('If-Match', '*');
    for (let item of items) {
      let itemQueryUrl: string = `${queryUrl}(${item.Id})`;
      const body: string = JSON.stringify({
        '@data.type': entityTypeName,
        'CompletionStatus': item.CompletionStatus,
        'Outcome': item.Outcome,
        'RequestStatus': `${this._getFormattedDateTime(new Date)} ${item.Outcome} by ${user.Title}\n${item.RequestStatus}`,
        'ApprovalComments': item.ApprovalComments,
        'ApprovedById': user.Id
      });
      try {
        const fetchResponse = await requester.fetch(itemQueryUrl, SPHttpClient.configurations.v1,
          {
            body: body,
            headers,
            method: 'PATCH'
          });
        debugger
        if (!fetchResponse.ok || fetchResponse.status !== 204) {
          throw new Error(fetchResponse.statusMessage);
        }
        // else {
        //   //return Promise.resolve(true);
        // }
      }
      catch (error) {
        console.log(error);
        throw new Error(error.message);
      }
    }
    return Promise.resolve(true);
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
          queryString = `?$select=Id,Title,RequestId/Id,Committee/Title,Outcome,ApprovedBy/Title,ApprovalComments,RequestStatus&$expand=Committee/Title,ApprovedBy/Title,RequestId${filterString}`;
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
                ApprovedBy: task.ApprovedBy ? task.ApprovedBy.Title : '',
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
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${listName}')?$select=ListItemEntityTypeFullName`;
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
        this._listItemEntityTypeName = results.ListItemEntityTypeFullName;
        this._lastListName = listName;
        return Promise.resolve(this._listItemEntityTypeName);
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
  private _getFormattedDateTime(d) {
    let thisDay = this._addZero(d.getDate());
    let thisMonth = this._addZero(d.getMonth() + 1);
    let thisTime = this._addZero(d.getHours()) + ":" + this._addZero(d.getMinutes());
    return d.getFullYear() + '-' + thisMonth + '-' + thisDay + ' ' + thisTime;
  }
  private _addZero(n) {
    n = n < 10 ? '0' + n : n;
    return n;
  }
  private _sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}