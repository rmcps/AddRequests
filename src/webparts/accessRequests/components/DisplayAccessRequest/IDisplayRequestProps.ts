import IDefaultProps from '../DefaultPage//IDefaultProps';
import IAccessRequest from '../../models/IAccessRequest';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import {IDisplayView} from '../../utilities/types';
export default interface IDisplayRequestProps {
    dataProvider: IAccessRequestsDataProvider;
    recordType: "New" | "Change" | "Display";
    //item: IAccessRequest;
    requestId: string;
    requestsByCommList: string;
    additionalInfo?: string;
    callingView: IDisplayView;
    onReturnClick: any;
}
