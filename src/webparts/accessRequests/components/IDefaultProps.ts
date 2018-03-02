import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';

export default interface IDefaultProps {
    requestsList: string;
    membersList: string;
    committeesList: string;
    membersCommitteesList: string;
    requestsByCommitteeList: string;
    context:any;
    dom:any;  
    dataProvider?: IAccessRequestsDataProvider;
    
}
