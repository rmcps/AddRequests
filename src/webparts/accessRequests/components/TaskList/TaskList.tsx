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
import { Link } from 'office-ui-fabric-react/lib/Link';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import ITask from '../../models/ITask';
import TaskItem from '../TaskItem/TaskItem';

export interface ITaskListProps {
  //onItemSelected: any;
  dataProvider: IAccessRequestsDataProvider;
  requestsByCommList: string;
  currentUser: any;
  isApprover: boolean;
  onTaskItemSelected: any;
}
export interface ITaskListState {
  taskItems: ITask[];
  dataIsLoading: boolean;
  showAllLink: boolean;
  errorMsg: string;
  hideDialog: boolean;
}
export default class TaskList extends React.Component<ITaskListProps, ITaskListState> {
  constructor(props: ITaskListProps) {
    super(props);
    this.state = {
      taskItems: [],
      dataIsLoading: true,
      showAllLink: true,
      errorMsg: null,
      hideDialog: true
    };
  }
  public async componentWillReceiveProps(nextProps: ITaskListProps) {
    this._getTasks(false);
  }
  public async componentDidMount() {
    this._getTasks(false);
  }
  public render() {
    return (
      <div>
        <div className={styles.row}>
          <div className={styles.column2}>
            {this.props.isApprover && this.state.showAllLink && <DefaultButton disabled={false} text='Show All' onClick={this._onShowAllTasks} />}  
            {this.state.taskItems.length > 1 && <DefaultButton disabled={false} text='Approve All' onClick={this._onApproveAllTasks} />}
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column2}>
            {this.state.taskItems.length > 0 && <h3>Your approval is requested on the items below</h3>}
            {this.state.taskItems.length == 0 && <h3>No pending approvals</h3>}
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
      <TaskItem item={item} isApprover={this.props.isApprover} onApprovalAction={this._handleApprovalAction} onError={this._handleErrors}
        onShowRequest={this._handleShowRequest} onApprovalCommentsChanged={this._handleApprovalCommentsChanged}
      />
    );
  }
  private async _getTasks(allTasks: boolean) {
    try {
      const results = await this.props.dataProvider.getTasksForCurrentUser(this.props.requestsByCommList, allTasks, this.props.currentUser);
      this.setState({
        taskItems: results,
        dataIsLoading: false,
        showAllLink: !allTasks
      });
    }
    catch (error) {
      console.log(error);
      this.setState({ dataIsLoading: false });
    }
  }
  @autobind
  private _handleApprovalCommentsChanged(item: ITask) {
    this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
      prevState.taskItems = this.state.taskItems.map((el, index) => {
        if (el.Id == item.Id) {
          el.ApprovalComments = item.ApprovalComments;
        }
        return el;
      });
      return prevState;
    });

  }
  @autobind
  private async _handleApprovalAction(item: ITask) {
    this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
      prevState.errorMsg = null;
      prevState.hideDialog = false;
      return prevState;
    });

    try {
      const result = await this.props.dataProvider.updateCommitteeTaskItem(item, this.props.requestsByCommList);
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
  @autobind
  private async _onShowAllTasks() {
    this._getTasks(true);
  }
  @autobind
  private async _onApproveAllTasks() {
    this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
      prevState.errorMsg = null;
      prevState.hideDialog = false;
      return prevState;
    });

    try {
      let items: ITask[] = this.state.taskItems.map(el => {
        return { ...el, Outcome: 'Approved' };
      });
      const result = await this.props.dataProvider.updateAllCommitteeTaskItems(items, this.props.requestsByCommList);
      if (result) {
        this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
          prevState.taskItems = [];
          prevState.showAllLink = true;
          prevState.hideDialog = true;
          return prevState;
        });
      }
    }
    catch (error) {
      console.log(error);
      this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
        prevState.errorMsg = "Error. Not all items approved.";
        prevState.hideDialog = true;
        return prevState;
      });
    }
  }
}