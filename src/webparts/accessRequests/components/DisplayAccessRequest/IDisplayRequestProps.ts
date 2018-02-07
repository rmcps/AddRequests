import IDefaultProps from '../IDefaultProps';
import IAccessRequest from '../../models/IAccessRequest';

export default interface IDisplayRequestProps extends IAccessRequest, IDefaultProps {
    recordType: "New" | "Modified" | "Display";
    addtionalInfo: string;
}