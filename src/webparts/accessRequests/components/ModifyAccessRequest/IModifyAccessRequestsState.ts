import IModifyAccessRequest from "../../models/IModifyAccessRequest";

export interface IModifyAccessRequestsState {
  status: string;
  Item: IModifyAccessRequest;
  errors: string[];
  isLoadingData: boolean;
  members:IModifyAccessRequest[];
  committees:any;
  selectedCommittees:any[];
  originalCommittees:any[];
  dropDownErrorMsg:string;
  enableSave:boolean;
  hideDialog:boolean;
}