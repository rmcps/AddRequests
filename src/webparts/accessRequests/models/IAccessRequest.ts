import ITask from "./ITask";

interface IAccessRequest {
    Id?: string;
    Title?: string;
    Comments?: string;
    AddCommittees?: any[];
    RemoveCommittees?: any[];
    Created?: string;
    EMail?: string;
    FirstName?: string;
    JobTitle?: string;
    LastName?: string;
    Modified?: string;
    Company?: string;
    Office?: string;
    RequestReason?: string;
    RequestStatus?: string;
    CompletionStatus?: string;
    Outcome?: string;
    ApprovedBy?: string;
    ApprovalComments?: string;
    AuthorId?: number;
    CreatedBy?: string;
    EditorId?: number;
    ModifiedBy?: string;
    CommitteeApprovals?: ITask[];
  }
  
  export default IAccessRequest;

  export interface Requests {
      value: IAccessRequest[];
  }