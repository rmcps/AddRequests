import { IWebPartContext } from '@microsoft/sp-webpart-base';
import * as lodash from '@microsoft/sp-lodash-subset';
import INewAccessRequest from "../models/INewAccessRequest";
import IModifyAccessRequest from "../models/IModifyAccessRequest";
import SharePointDataProvider from '../models/IAccessRequestsDataProvider';
import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';

export default class MockNewAccessRequest implements IAccessRequestsDataProvider {

    private _accessListTitle: string;
    private _listsUrl: string;
    private _accessListItemsUrl: string;
    private _webPartContext: IWebPartContext;
    private _items: any;
    private _committees: any;
    private _memberCommittees: any;

    constructor() {
        this._items = {value: [
            {Id:"1", FirstName: 'Mary', LastName:'May',EMail:'mmary@rmcps.com',Company:'', 
                Committees:[],
                RequestReason:'New member'
            },
            {Id:"2", FirstName: 'Judy', LastName:'Jones',EMail:'jjones@rmcps.com',Company:'', 
                Committees:[],
                RequestReason:'New member'
            },
            {Id:"3", FirstName: 'Tany', LastName:'Thomas',EMail:'ttom@rmcps.com',Company:'', 
                Committees:[],
                RequestReason:'New member'
            },
            {Id:"4", FirstName: 'Walt', LastName:'Whitman',EMail:'wwhitman@rmcps.com',Company:'', 
                Committees:[],
                RequestReason:'New member'
            }            
        ]};
        this._committees = {value: [
            {Id:1,Title:'Board',ID:1},
            {Id:2,Title:'Executive',ID:2},
            {Id:3,Title:'Legal',ID:3},
            {Id:4,Title:'HIT',ID:4},
            {Id:5,Title:'More',ID:5},
            {Id:6,Title:'Something',ID:6}
            ]};
        this._memberCommittees = {value: [
            {memberId: 1, committeeIds:[1,3]},
            {memberId: 2, committeeIds:[1]},
            {memberId: 3, committeeIds:[2,4]},
            {memberId: 4, committeeIds:[4]},
            ]};
    }
  
    public set accessListTitle(value:string) {
        this._accessListTitle = value;
    }
  
    public get accessListTitle() {
        return this._accessListTitle;
    }    

    public set accessListItemsUrl(value:string) {
        this._accessListItemsUrl = value;
    }

    public set webPartContext(value: IWebPartContext) {
        this._webPartContext = value;
    }

    public get webPartContext(): IWebPartContext {
        return this._webPartContext;
    }
    public getItem(listId: number): Promise<IModifyAccessRequest> {
        const item: IModifyAccessRequest = lodash.clone(this._items[1]);
        return new Promise<IModifyAccessRequest>((resolve) => {
            setTimeout(() => resolve(item), 500);
        });
    }
    public getMembers(): Promise<any> {
        const items: IModifyAccessRequest[] = lodash.clone(this._items);

        return new Promise<any>((resolve) => {
            setTimeout(() => resolve(items), 500);
        });
    }
    public getMemberCommittees(Id: any): Promise<any> {
        let selected = this._memberCommittees.value.filter((member) => member.memberId == Id );              
        const items: any = lodash.clone(selected[0].committeeIds);
        return new Promise<any>((resolve) => {
            setTimeout(() => resolve(items), 500);
        });
    }
    public getCommittees(): Promise<any> {
        let comms = this._committees;
        const items: any = lodash.clone(comms);

        return new Promise<any>((resolve) => {
            setTimeout(() => resolve(items), 500);
        });
    }
    public saveNewItem(newItem:INewAccessRequest):Promise<any> {        
        const result = {
            ok: true
        };
        return new Promise<any>((resolve) => {
            setTimeout(() => resolve(result), 500);
        });
    }
}
