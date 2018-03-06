import ITask from "./ITask";

interface IFinalTask {
    Id?: string;
    Comments?: string;
    JobTitle?: string;
    Title?: string;    
    Office?: string;
    RequestReason?: string;
    RequestStatus?: string;
    CompletionStatus: string;
    ApprovalComments?: string;
    CommitteeTasks?: ITask[];
  }

  export default IFinalTask;
  