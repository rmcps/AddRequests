import { IWebPartContext } from '@microsoft/sp-webpart-base';
import INewAccessRequest from './INewAccessRequest';
import IModifyAccessRequest from './IModifyAccessRequest';
import IAccessRequest from './IAccessRequest';
import ITask from './ITask';
import IFinalTask from './IFinalTask';
interface INewAccessRequestsDataProvider {

    accessListTitle: string;
    accessListItemsUrl;
    webPartContext: IWebPartContext;   
    getCurrentUser():Promise<any>;
    getItem(requestId:string, requestsByCommList: string): Promise<IAccessRequest>;
    getItemsForCurrentUser(currentUser?: any):Promise<IAccessRequest[]>;
    getMembers(membersList: string): Promise<any>;
    getMemberCommittees(membersCommList: string, loginName: any):Promise<any>;
    getCommittees(committeesListTitle: string): Promise<any>;
    saveNewItem(INewAccessRequest): Promise<any>;
    saveChangeRequest(IModifyAccessRequest): Promise<any>;
    getTasksForCurrentUser(requestsByCommList: string, alltasks: boolean, currentUser?: any):Promise<ITask[]>;
    updateCommitteeTaskItem(item: ITask, requestsByCommList: string, currentUser?: any):Promise<boolean>;
    updateAllCommitteeTaskItems(items: ITask[], requestsByCommList: string, currentUser?: any): Promise<boolean>;
    updateForRequest(items: IFinalTask[], currentUser?: any):Promise<any>;
    getFinalTasks(requestsByCommList: string, currentUser?: any):Promise<IFinalTask[]>;
}

export default INewAccessRequestsDataProvider;