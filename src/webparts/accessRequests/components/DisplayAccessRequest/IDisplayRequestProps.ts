import IDefaultProps from '../IDefaultProps';
import IAccessRequest from '../../models/IAccessRequest';

export default interface IDisplayRequestProps {
    recordType: "New" | "Change" | "Display";
    item: IAccessRequest;
    additionalInfo?: string;
}