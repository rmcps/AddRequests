interface IModifyAccessRequest {
    Id?: string;
    spLoginName?: any;
    Comments?: string;
    AddCommittees?: any[];
    RemoveCommittees?: any[];
    EMail?: string;
    FirstName?: string;
    JobTitle?: string;
    LastName?: string;
    Title?: string;    
    Office?: string;
    RequestReason?: string;
  }

  export default IModifyAccessRequest;
  