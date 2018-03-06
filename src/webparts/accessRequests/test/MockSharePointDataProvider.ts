import { IWebPartContext } from '@microsoft/sp-webpart-base';
import * as lodash from '@microsoft/sp-lodash-subset';
import IAccessRequest from '../models/IAccessRequest';
import INewAccessRequest from "../models/INewAccessRequest";
import IModifyAccessRequest from "../models/IModifyAccessRequest";
import SharePointDataProvider from '../models/IAccessRequestsDataProvider';
import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';
import ITask from '../models/ITask';
import IFinalTask from '../models/IFinalTask';

export default class MockNewAccessRequest implements IAccessRequestsDataProvider {

    private _accessListTitle: string;
    private _listsUrl: string;
    private _accessListItemsUrl: string;
    private _webPartContext: IWebPartContext;
    private _items: any;
    private _committees: any;
    private _memberCommittees: any;

    constructor() {
        this._items = {
            value: [
                {
                    Id: "1", spLoginName: "1", Title: "Mary May", FirstName: 'Mary', LastName: 'May', EMail: 'mmary@rmcps.com', Company: '',
                    Committees: [],
                    RequestReason: 'New member'
                },
                {
                    Id: "2", spLoginName: "2", Title: "Judy Jones", FirstName: 'Judy', LastName: 'Jones', EMail: 'jjones@rmcps.com', Company: '',
                    Committees: [],
                    RequestReason: 'New member'
                },
                {
                    Id: "3", spLoginName: "3", Title: "Tiny Thomas", FirstName: 'Tiny', LastName: 'Thomas', EMail: 'ttom@rmcps.com', Company: '',
                    Committees: [],
                    RequestReason: 'New member'
                },
                {
                    Id: "4", spLoginName: "4", Title: "Walt Disney", FirstName: 'Walt', LastName: 'Disney', EMail: 'wdisney@rmcps.com', Company: '',
                    Committees: [],
                    RequestReason: 'New member'
                }
            ]
        };
        this._committees = {
            value: [
                { Id: 1, Title: 'Board', ID: 1 },
                { Id: 2, Title: 'Executive', ID: 2 },
                { Id: 3, Title: 'Legal', ID: 3 },
                { Id: 4, Title: 'HIT', ID: 4 },
                { Id: 5, Title: 'More', ID: 5 },
                { Id: 6, Title: 'Something', ID: 6 }
            ]
        };
        this._memberCommittees = {
            "value": [
                { "loginName": "1", "Committee": "Board" },
                { "loginName": "1", "Committee": "Legal" },
                { "loginName": "2", "Committee": ["Board"] },
                { "loginName": "3", "Committee": "Executive" },
                { "loginName": "3", "Committee": "HIT" },
                { "loginName": "4", "Committee": "HIT" },
            ]
        };
    }

    public set accessListTitle(value: string) {
        this._accessListTitle = value;
    }

    public get accessListTitle() {
        return this._accessListTitle;
    }

    public set accessListItemsUrl(value: string) {
        this._accessListItemsUrl = value;
    }

    public set webPartContext(value: IWebPartContext) {
        this._webPartContext = value;
    }

    public get webPartContext(): IWebPartContext {
        return this._webPartContext;
    }
    public getItem(requestId: string): Promise<IAccessRequest> {
        const item: IAccessRequest = lodash.clone(this._items[1]);
        return new Promise<IAccessRequest>((resolve) => {
            setTimeout(() => resolve(item), 500);
        });
    }
    public getCurrentUser(): Promise<any> {
        return null;
    }

    public getItemsForCurrentUser(): Promise<IAccessRequest[]> {
        return null;
    }
    public getMembers(membersList: string): Promise<any> {
        const items: any = lodash.clone(this._items);

        return new Promise<any>((resolve) => {
            setTimeout(() => resolve(items), 500);
        });
    }
    public getMemberCommittees(membersCommList: string, loginName: string): Promise<any> {
        let selected = {
            "value": this._memberCommittees.value.filter((member) => member.loginName == loginName)
        };
        const items: any = lodash.clone(selected);
        return new Promise<any>((resolve) => {
            setTimeout(() => resolve(items), 500);
        });
    }
    public getCommittees(committeesListTitle: string): Promise<any> {
        let comms = this._committees;
        const items: any = lodash.clone(comms);

        return new Promise<any>((resolve) => {
            setTimeout(() => resolve(items), 500);
        });
    }
    public saveNewItem(newItem: INewAccessRequest): Promise<any> {
        const result = {
            ok: true
        };
        return new Promise<any>((resolve) => {
            setTimeout(() => resolve(result), 500);
        });
    }
    public saveChangeRequest(item: IModifyAccessRequest): Promise<any> {
        return null;
    }
    public getTasksForCurrentUser(requestsByCommList: string): Promise<ITask[]> {
        return null;
    }
    public updateForCommittee(item:ITask, requestsByCommList: string, action: "approve" | "reject"): Promise<boolean> {
        return Promise.resolve(true);
    }
    public updateForRequest(item:IFinalTask) {
        return null;
    }
    public getFinalTasks(requestsByCommList: string):Promise<IFinalTask[]> {
        return null;
    }
}
