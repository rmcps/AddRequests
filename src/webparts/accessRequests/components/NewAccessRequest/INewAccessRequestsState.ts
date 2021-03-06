import INewAccessRequest from "../../models/INewAccessRequest";

export interface INewAccessRequestsState {
  status: string;
  newItem: INewAccessRequest;
  errors: string[];
  isLoadingData: boolean;
  isSaving: boolean;
  committees:any;
  selectedCommittees:any[];
  CommitteeAccess: boolean;
  dropDownErrorMsg:string;
  hideDialog:boolean;
  enableSave:boolean;
}