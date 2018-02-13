import { IWebPartContext } from '@microsoft/sp-webpart-base';
import INewAccessRequest from '../models/INewAccessRequest';
import IModifyAccessRequest from './IModifyAccessRequest';
import IAccessRequest from './IAccessRequest';

interface INewAccessRequestsDataProvider {

    accessListTitle: string;
    accessListItemsUrl;
    webPartContext: IWebPartContext;   
    getCurrentUser():Promise<any>;
    getItem(requestId:string): Promise<IAccessRequest>;
    getItemsForCurrentUser():Promise<IAccessRequest[]>;
    getMembers(membersList: string): Promise<any>;
    getMemberCommittees(membersCommList: string, loginName: any):Promise<any>;
    getCommittees(committeesListTitle: string): Promise<any>;
    saveNewItem(INewAccessRequest): Promise<any>;
    saveModifyRequest(IModifyAccessRequest): Promise<any>;
}

export default INewAccessRequestsDataProvider;