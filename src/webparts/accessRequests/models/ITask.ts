export default interface ITask {
    Id: string;
    RequestId?: string;
    Name: string;
    Committee: string;
    RequestStatus: string;
    RequestType?: string;
    Outcome: string;
    CompletionStatus: string;
    ApprovalComments?: string;
    Created: string;
    Modified: string;
    Updating?: boolean;
    CurrentUser: string;
}