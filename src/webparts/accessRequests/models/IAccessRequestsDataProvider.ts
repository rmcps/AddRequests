import { IWebPartContext } from '@microsoft/sp-webpart-base';
import INewAccessRequest from '../models/INewAccessRequest';
import IModifyAccessRequest from './IModifyAccessRequest';
import IAccessRequest from './IAccessRequest';
import ITask from '../models/ITask';

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
    saveChangeRequest(IModifyAccessRequest): Promise<any>;
    getTasksForCurrentUser(requestsByCommList: string):Promise<ITask[]>;
    updateForCommittee(itemId: string, action: "Approved" | "Rejected", requestsByCommList: string):Promise<boolean>;
    updateForRequest(itemId: string, action: "Approved" | "Rejected"):Promise<any>;
}

export default INewAccessRequestsDataProvider;