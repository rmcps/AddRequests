declare interface IAccessRequestsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  RequestListFieldLabel: string;
  MembersListFieldLabel: string;
  CommitteesList: string;
  MembersCommitteesListFieldLabel: string;
  RequestsByCommitteeListFieldLabel: string;
  FinalApproverIdFieldLabel: string;
}

declare module 'AccessRequestsWebPartStrings' {
  const strings: IAccessRequestsWebPartStrings;
  export = strings;
}

declare interface IAccessRequestsFormStrings {
  SubmitButtonText:string;
  TextInputRequired:string;
  SelectionRequired:string;
}