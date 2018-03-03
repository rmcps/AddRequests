import * as React from 'react';
import * as ReactDom from 'react-dom';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import styles from './AccessRequests.module.scss';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import SharePointDataProvider from '../services/SharePointDataProvider';
import MockSharePointDataProvider from '../test/MockSharePointDataProvider';
import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';
import NewAccessRequest from './NewAccessRequest/NewAccessRequest';
import ModifyAccessRequest from './ModifyAccessRequest/ModifyAccessRequest';
import IDefaultProps from './IDefaultProps';
import AccessRequestList from './AccessRequestsList/AccessRequestList';
import IAccessRequest from '../models/IAccessRequest';
import DisplayRequest from './DisplayAccessRequest/DisplayRequest';
import IDisplayRequestProps from './DisplayAccessRequest/IDisplayRequestProps';
import TaskList from './TaskList/TaskList';
import FinalTaskList from './FinalTaskList/FinalTaskList';
import TopNav from './Navigation/TopNav';

export interface IDefaultState {
  show: "List" | "New" | "Change" | "Display" | "Tasks" | "FinalTasks";
  selectedItem: IAccessRequest;
  listNotConfigured: boolean;
}

export default class DefaultPage extends React.Component<IDefaultProps, IDefaultState> {
  private _dataProvider: IAccessRequestsDataProvider;

  constructor(props: IDefaultProps, state: IDefaultState) {
    const urlParams = new URLSearchParams(window.location.search);

    super(props);
    // set initial state   
    this.state = {
      show: urlParams.get("view") === "tasks" ? "Tasks" : (urlParams.get("view") === "finaltasks" ? "FinalTasks" : "List"),
      selectedItem: null,
      listNotConfigured: false,
    };

    if (DEBUG && Environment.type === EnvironmentType.Local) {
      this._dataProvider = new MockSharePointDataProvider();

    } else {
      this._dataProvider = new SharePointDataProvider(this.props.context);
      this._dataProvider.accessListTitle = this.props.requestsList;
    }

  }
  public componentWillReceiveProps(nextProps: IDefaultProps) {
    this.setState({
      listNotConfigured: this.listNotConfigured(nextProps),
    });
  }
  public componentDidMount() {

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
                  : <TopNav onItemSelected={this.handleViewSelected} show={this.state.show} />
                }
              </div>
            </div>
          </div>
          <div className={styles.innerContent}>
            {(this.state.listNotConfigured == false && this.state.show == "List") && <AccessRequestList dataProvider={this._dataProvider}
              onItemSelected={this.handleItemSelected} />}
            {(this.state.listNotConfigured == false && this.state.show == "Display") && <DisplayRequest item={this.state.selectedItem}
              recordType="Display" onReturnClick={this.handleViewSelected} />}
            {(this.state.listNotConfigured == false && this.state.show == "New") && <NewAccessRequest
              dataProvider={this._dataProvider} committeesListTitle={this.props.committeesList} onRecordAdded={this.handleViewSelected} />}
            {(this.state.listNotConfigured == false && this.state.show == "Change") && <ModifyAccessRequest
              dataProvider={this._dataProvider} membersList={this.props.membersList} membersCommList={this.props.membersCommitteesList}
              committeesListTitle={this.props.committeesList} onRecordAdded={this.handleViewSelected} />}
            {(this.state.listNotConfigured == false && this.state.show == "Tasks") && <TaskList dataProvider={this._dataProvider} requestsByCommList = {this.props.requestsByCommitteeList} /> }
            {(this.state.listNotConfigured == false && this.state.show == "FinalTasks") && <FinalTaskList dataProvider={this._dataProvider} requestsByCommList = {this.props.requestsByCommitteeList} /> }
          </div>
        </div>
      </div>
    );
  }

  @autobind
  private handleViewSelected(selectedView) {
    switch (selectedView) {
      case "addNew":
        this.setState({
          show: "New"
        });
        break;
      case "change":
        this.setState({
          show: "Change"
        });
        break;
      case "display":
        this.setState({
          show: "Display"
        });
        break;
      case "list":
        this.setState({
          show: "List"
        });
        break;
        case "tasks":
        this.setState({
          show: "Tasks"
        });
        break;
        case "finaltasks":
        this.setState({
          show: "FinalTasks"
        });
        break;
      }
  }
  @autobind
  private handleItemSelected(item) {
    this.setState({
      selectedItem: item,
      show: "Display"
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

}
