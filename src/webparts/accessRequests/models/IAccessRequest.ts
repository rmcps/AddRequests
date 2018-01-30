interface IAccessRequest {
    Id?: number;
    Title?: string;
    Comments?: string;
    Committees?: any[];
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
    AuthorId?: number;
    EditorId?: number;
  }
  
  export default IAccessRequest;

  export interface Requests {
      value: IAccessRequest[];
  }