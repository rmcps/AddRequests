import { IWebPartContext } from '@microsoft/sp-webpart-base';
import INewAccessRequest from '../models/INewAccessRequest';
import IModifyAccessRequest from './IModifyAccessRequest';

interface INewAccessRequestsDataProvider {

    accessListTitle: string;
    accessListItemsUrl;
    webPartContext: IWebPartContext;   
    getItem(listId:number): Promise<IModifyAccessRequest>;
    getMembers(): Promise<any>;
    getMemberCommittees(loginName: any):Promise<any>;
    getCommittees(): Promise<any>;
    saveNewItem(INewAccessRequest): Promise<any>;
    saveModifyRequest(IModifyAccessRequest): Promise<any>;
}

export default INewAccessRequestsDataProvider;