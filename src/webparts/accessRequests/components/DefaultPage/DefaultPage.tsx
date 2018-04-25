import * as React from 'react';
import * as ReactDom from 'react-dom';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import styles from '../AccessRequests.module.scss';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import SharePointDataProvider from '../../services/SharePointDataProvider';
import MockSharePointDataProvider from '../../test/MockSharePointDataProvider';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import NewAccessRequest from '../NewAccessRequest/NewAccessRequest';
import ModifyAccessRequest from '../ModifyAccessRequest/ModifyAccessRequest';
import IDefaultProps from '../DefaultPage/IDefaultProps';
import AccessRequestList from '../AccessRequestsList/AccessRequestList';
import IAccessRequest from '../../models/IAccessRequest';
import DisplayRequest from '../DisplayAccessRequest/DisplayRequest';
import IDisplayRequestProps from '../DisplayAccessRequest/IDisplayRequestProps';
import TaskList from '../TaskList/TaskList';
import FinalTaskList from '../FinalTaskList/FinalTaskList';
import TopNav from '../Navigation/TopNav';
import {IDisplayView} from '../../utilities/types';

export interface IDefaultState {
  show: IDisplayView;
  //selectedItem: IAccessRequest;
  selectedRequestId: string;
  listNotConfigured: boolean;
  currentUser: any;
  callingView: IDisplayView;
}
export interface IViewRequest {
  view: IDisplayView;
  requestId?: string;
}
export default class DefaultPage extends React.Component<IDefaultProps, IDefaultState> {
  private _dataProvider: IAccessRequestsDataProvider;

  constructor(props: IDefaultProps, state: IDefaultState) {
    super(props);
    let showView:IViewRequest = this._getParams();
    // set initial state   
    this.state = {
      show: showView.view,
      selectedRequestId: showView.requestId,
      listNotConfigured: false,
      currentUser: null,
      callingView: "List"
    };

    if (DEBUG && Environment.type === EnvironmentType.Local) {
      this._dataProvider = new MockSharePointDataProvider();

    } else {
      this._dataProvider = new SharePointDataProvider(this.props.context);
      this._dataProvider.accessListTitle = this.props.requestsList;
    }

  }
  public async componentWillReceiveProps(nextProps: IDefaultProps) {
    this.setState({
      listNotConfigured: this.listNotConfigured(nextProps),
    });
    if (!this.state.currentUser) {
      const results = await this._dataProvider.getCurrentUser();
      this.setState({ currentUser : results });
    }
  }
  public async componentDidMount() {
    if (!this.state.currentUser) {
      const results = await this._dataProvider.getCurrentUser();
      this.setState({ currentUser : results });
    }
  }
  public render(): React.ReactElement<IDefaultProps> {
    return (
      <div className={styles.accessRequests}>
        <div className={styles.outerContainer}>
          <div className={styles.row}>
            <div className={styles.column2}>
              <div className={styles.headerBar}>
                {this.state.listNotConfigured ?
                  <MessageBar messageBarType={MessageBarType.warning}>Please configure the lists for this component first.</MessageBar>
                  : <TopNav isApprover={this.state.currentUser && Number(this.state.currentUser.Id) === Number(this.props.finalApproverId)} 
                            onItemSelected={this.handleViewSelected} show={this.state.show} />
                }
              </div>
            </div>
          </div>
          <div className={styles.innerContent}>
            {(this.state.listNotConfigured == false && this.state.show == "List") && <AccessRequestList dataProvider={this._dataProvider}
              onItemSelected={this.handleItemSelected} currentUser={this.state.currentUser} />}
            {(this.state.listNotConfigured == false && this.state.show == "Display") && <DisplayRequest dataProvider={this._dataProvider} recordType="Display" 
              requestId={this.state.selectedRequestId} requestsByCommList={this.props.requestsByCommitteeList} 
              callingView={this.state.callingView} onReturnClick={this.handleViewSelected} />}
            {(this.state.listNotConfigured == false && this.state.show == "New") && <NewAccessRequest
              dataProvider={this._dataProvider} committeesListTitle={this.props.committeesList} onRecordAdded={this.handleViewSelected} />}
            {(this.state.listNotConfigured == false && this.state.show == "Change") && <ModifyAccessRequest
              dataProvider={this._dataProvider} membersList={this.props.membersList} membersCommList={this.props.membersCommitteesList}
              committeesListTitle={this.props.committeesList} onRecordAdded={this.handleViewSelected} />}
            {(this.state.listNotConfigured == false && this.state.show == "Tasks") && <TaskList dataProvider={this._dataProvider} 
                  requestsByCommList = {this.props.requestsByCommitteeList} currentUser={this.state.currentUser} 
                  isApprover={this.state.currentUser && Number(this.state.currentUser.Id) === Number(this.props.finalApproverId)} 
                  onTaskItemSelected={this.handleItemSelected} /> }
            {(this.state.listNotConfigured == false && this.state.show === "FinalTasks") && <FinalTaskList dataProvider={this._dataProvider} 
                  requestsByCommList = {this.props.requestsByCommitteeList} currentUser={this.state.currentUser}  onTaskItemSelected={this.handleItemSelected} /> }
          </div>
        </div>
      </div>
    );
  }

  @autobind
  private handleViewSelected(selectedView: IDisplayView) {
    this.setState({
      show: selectedView
    });
  }
  @autobind
  private handleItemSelected(requestId, callingView?: IDisplayView) {
    this.setState({
      //selectedItem: item,
      selectedRequestId: requestId,
      show: "Display",
      callingView: callingView ? callingView: "List"
    });

  }
  private listNotConfigured(props: IDefaultProps): boolean {

    return props.requestsList === undefined ||
      props.requestsList === null ||
      props.requestsList.length === 0 ||
      props.membersList === undefined ||
      props.membersList === null ||
      props.membersList.length === 0 ||
      props.committeesList === undefined ||
      props.committeesList === null ||
      props.committeesList.length === 0 ||
      props.membersCommitteesList === undefined ||
      props.membersCommitteesList === null ||
      props.membersCommitteesList.length === 0;

  }
  private _getParams(): IViewRequest{
    const urlParams = new URLSearchParams(window.location.search);
    const view = urlParams.get("view");
    switch(view) {
      case "tasks":
        return {view: "Tasks"};
      case "finaltasks":
        return {view: "FinalTasks"};
      case "display":
      return {view: "Display", requestId: urlParams.get("requestid")};
      default:
      return {view: "List"};
    }
  }
}
