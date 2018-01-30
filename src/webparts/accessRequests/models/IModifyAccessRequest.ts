interface IModifyAccessRequest {
    Id?: number;
    Comments?: string;
    Committees?: any[];
    EMail?: string;
    FirstName?: string;
    JobTitle?: string;
    LastName?: string;
    Company?: string;
    Office?: string;
    RequestReason?: string;
  }

  export default IModifyAccessRequest;
  