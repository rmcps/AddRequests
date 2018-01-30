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
    private _items: IModifyAccessRequest[];
    private _committees: any;
    //private _listItemEntityTypeFullName:string;
    constructor() {
        //this._listItemEntityTypeFullName = "listName";
        this._items = [
            {Id:1, FirstName: 'Mary', LastName:'May',EMail:'mmary@rmcps.com',Company:'', 
                Committees:[],
                RequestReason:'New member'
            },
            {Id:1, FirstName: 'Judy', LastName:'Jones',EMail:'jjones@rmcps.com',Company:'', 
                Committees:[],
                RequestReason:'New member'
            }
            
        ];
        this._committees = [
            {Id:1,Title:'Board'},
            {Id:2,Title:'Executive'},
            {Id:3,Title:'Legal'},
            {Id:4,Title:'HIT'},
        ];
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
  
    public getCommittees(): Promise<any> {
        const items: any = lodash.clone(this._committees);

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
