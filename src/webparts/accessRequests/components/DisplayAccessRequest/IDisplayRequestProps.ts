import IDefaultProps from '../IDefaultProps';
import IAccessRequest from '../../models/IAccessRequest';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
export default interface IDisplayRequestProps {
    dataProvider: IAccessRequestsDataProvider;
    recordType: "New" | "Change" | "Display";
    //item: IAccessRequest;
    requestId: string;
    requestsByCommList: string;
    additionalInfo?: string;
    onReturnClick: any;
}