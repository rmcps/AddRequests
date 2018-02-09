import * as React from 'react';
import * as ReactDom from 'react-dom';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from './AccessRequests.module.scss';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import SharePointDataProvider from '../services/SharePointDataProvider';
import MockSharePointDataProvider from '../test/MockSharePointDataProvider';
import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';
import NewAccessRequest from './NewAccessRequest/NewAccessRequest';
import ModifyAccessRequest from './ModifyAccessRequest/ModifyAccessRequest';
import { IAccessRequestsProps } from './IAccessRequestsProps';
import IDefaultProps from './IDefaultProps';
import AccessRequestList from './AccessRequestsList/AccessRequestList';
import IAccessRequest from '../models/IAccessRequest';
import DisplayRequest from './DisplayAccessRequest/DisplayRequest';
import IDisplayRequestProps from './DisplayAccessRequest/IDisplayRequestProps';
import TopNav from './Navigation/TopNav'

export interface IDefaultState {
  listItems: IAccessRequest[];
  show: "List" |"New" | "Change" | "Display";
}

export default class DefaultPage extends React.Component<IDefaultProps, IDefaultState> {
  private _dataProvider: IAccessRequestsDataProvider;

  constructor(props: IDefaultProps, state: IDefaultState) {
    super(props);
    // set initial state   
    this.state = {
      listItems: [],
      show: "List"
    }; 
    
    if (DEBUG && Environment.type === EnvironmentType.Local) {
      this._dataProvider = new MockSharePointDataProvider();

    } else {
      this._dataProvider = new SharePointDataProvider();
      this._dataProvider.webPartContext = this.props.context;
      this._dataProvider.accessListTitle = "Site Access Requests";
    }
    
  }
  public componentWillReceiveProps(nextProps: IDefaultProps) {
    this._dataProvider.getItemsForCurrentUser().then((items: IAccessRequest[]) => {
      this.setState({listItems: items});     
    });
    
    // this.setState({
    //   status: this.listNotConfigured(nextProps) ? 'Please configure list in Web Part properties' : '',
    // });
  }
  public componentDidMount() {
    this._dataProvider.getItemsForCurrentUser().then((items: IAccessRequest[]) => {
      this.setState({listItems: items});     
    });
    
  }
  public render(): React.ReactElement<IDefaultProps> {
    return (
      <div className={styles.accessRequests}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <h2 className={styles.headerBar}>Member Access Request Submission</h2>
            </div>
          </div>
              <TopNav onItemSelected={this.handleMenuItemSelected} />
    {this.state.show == "List" &&  <AccessRequestList items= {this.state.listItems} onItemSelected={this.handleItemSelected} /> }
        </div>
      </div>
    );
  }

@autobind
private handleMenuItemSelected(selectedItem) {
  switch(selectedItem) {
    case "addNew":
      this._onAddNew();
      break;
      case "change":
      this._onChange();
      break;

    }
}

@autobind
  private _onAddNew() {
    const element: React.ReactElement<IAccessRequestsProps > = React.createElement(
        NewAccessRequest,
        {
          description: this.props.description,
          context:this.props.context,
          dom: this.props.dom,
        }
      );
      ReactDom.unmountComponentAtNode(this.props.dom);          
      ReactDom.render(element, this.props.dom);
  }
  @autobind
  private _onChange() {
    const element: React.ReactElement<IAccessRequestsProps > = React.createElement(
      ModifyAccessRequest,
        {
          description: this.props.description,
          context:this.props.context,
          dom: this.props.dom,
        }
      );
      ReactDom.unmountComponentAtNode(this.props.dom);          
      ReactDom.render(element, this.props.dom);
  }
@autobind
private _onCancel():void {
  window.location.href = "https://uphpcin.sharepoint.com";
}

@autobind
private handleItemSelected(requestId) {
  const rItems = this.state.listItems;
  const requestItem: IAccessRequest = rItems.filter((i) => i.Id == requestId).pop();  // should only return 1 element so take the last.
  const element: React.ReactElement<IDisplayRequestProps> = React.createElement(
    DisplayRequest, {
      context: this.props.context,
      dom: this.props.dom,
      recordType: "Display",
      Id: requestId,
      Title: requestItem.Title,
      EMail: requestItem.EMail,
      JobTitle: requestItem.JobTitle,
      Company: requestItem.Company,
      Comments: requestItem.Comments,
      RequestReason: requestItem.RequestReason,
      AddCommittees: requestItem.AddCommittees,
      RemoveCommittees: requestItem.RemoveCommittees,
      additionalInfo: "",
      description: ""
    }
  );
  ReactDom.unmountComponentAtNode(this.props.dom);
  ReactDom.render(element, this.props.dom);  
}
  // private listNotConfigured(props: IAccessRequestsProps): boolean {
  //   return props.listName === undefined ||
  //     props.listName === null ||
  //     props.listName.length === 0;
  // }

    
}
