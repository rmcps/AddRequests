export default interface ITask {
    Id: string;
    RequestId?: string;
    Name: string;
    Committee: string;
    RequestStatus: string;
    Outcome: string;
    CompletionStatus: string;    
    Created: string;
    Modified: string;
    Updating?: boolean;
}