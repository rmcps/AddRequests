import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';

export interface IAccessRequestsProps {
  description: string;
  context:any;
  dom:any;
  dataProvider?: IAccessRequestsDataProvider;
}
