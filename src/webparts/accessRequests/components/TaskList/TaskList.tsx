import * as React from 'react';
import styles from '../AccessRequests.module.scss';
import taskStyles from '../TaskList/TaskList.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import ITask from '../../models/ITask';
import TaskItem from '../TaskItem/TaskItem';

export interface ITaskListProps {
  //onItemSelected: any;
  dataProvider: IAccessRequestsDataProvider;
  requestsByCommList: string;
  currentUser: any;
  onTaskItemSelected: any;
}
export interface ITaskListState {
  taskItems: ITask[];
  dataIsLoading: boolean;
  errorMsg: string;
  hideDialog: boolean;
}
export default class TaskList extends React.Component<ITaskListProps, ITaskListState> {
  constructor(props: ITaskListProps) {
    super(props);
    this.state = {
      taskItems: [],
      dataIsLoading: true,
      errorMsg: null,
      hideDialog: true
    };
  }
  public async componentWillReceiveProps(nextProps: ITaskListProps) {
    try {
      let results = await this.props.dataProvider.getTasksForCurrentUser(this.props.requestsByCommList, this.props.currentUser);
      this.setState({
        taskItems: results,
        dataIsLoading: false
      });
    }
    catch (error) {
      console.log(error);
      this.setState({ dataIsLoading: false });
    }
  }
  public async componentDidMount() {
    try {
      const results = await this.props.dataProvider.getTasksForCurrentUser(this.props.requestsByCommList, this.props.currentUser);
      this.setState({
        taskItems: results,
        dataIsLoading: false
      });
    }
    catch (error) {
      console.log(error);
      this.setState({ dataIsLoading: false });
    }
  }
  public render() {
    return (
      <div className={styles.row}>
        <div className={styles.column2}>
          <h3>My Tasks</h3>
          {this.state.dataIsLoading ? <Spinner size={SpinnerSize.medium} /> : null}
          {this.state.errorMsg ? <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}>
              {this.state.errorMsg}
            </MessageBar>
            : null
            }
          <Fabric>
            <FocusZone direction={FocusZoneDirection.vertical}>
              <List
                items={this.state.taskItems} className={taskStyles.TaskList}
                onRenderCell={this._onRenderCell}
              />
            </FocusZone>
          </Fabric>
        </div>
        <Dialog
          hidden={this.state.hideDialog}
          dialogContentProps={{
            type: DialogType.normal,
            subText: "Updating..."
          }}
          modalProps={{
            isBlocking: true,
          }}
        >
        </Dialog>        
      </div>
    );
  }
  @autobind
  private _onRenderCell(item: ITask, index: number | undefined): JSX.Element {
    return (
      <TaskItem item={item} onApprovalAction={this._handleApprovalAction} onError={this._handleErrors} onShowRequest={this._handleShowRequest} />
    );
  }
  @autobind
  private async _handleApprovalAction(item: ITask) {
    this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
      prevState.errorMsg = null;
      prevState.hideDialog = false;
      return prevState;
    });
    
    try {
      const result = await this.props.dataProvider.updateForCommittee(item, this.props.requestsByCommList);
        if (result) {
          let newItems = this.state.taskItems.filter((i) => i.Id !== item.Id);
          this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
            prevState.taskItems = newItems;
            prevState.hideDialog = true;
            return prevState;
          });
        }
    }
    catch (error) {
      console.log(error);
      this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
        prevState.errorMsg = "Error updating approval";
        prevState.hideDialog = true;
        return prevState;
      });
    }
  }
  @autobind
  private async _handleErrors(errorMessage: string) {
    this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
      prevState.errorMsg = errorMessage;
      return prevState;
    });  
  }
  @autobind
  private _handleShowRequest(requestId) {
    this.props.onTaskItemSelected(requestId, "Tasks");
  }
}