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
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import ITask from '../../models/ITask';
import TaskItem from '../TaskItem/TaskItem';

export interface ITaskListProps {
  //onItemSelected: any;
  dataProvider: IAccessRequestsDataProvider;
  requestsByCommList: string;
  currentUser: any;
}
export interface ITaskListState {
  taskItems: ITask[];
  dataIsLoading: boolean;
  errors: string[];
}
export default class TaskList extends React.Component<ITaskListProps, ITaskListState> {
  constructor(props: ITaskListProps) {
    super(props);
    this.state = {
      taskItems: [],
      dataIsLoading: true,
      errors: [],
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
          {this.renderErrors()}          
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
    );
  }
  @autobind
  private _onRenderCell(item: ITask, index: number | undefined): JSX.Element {
    return (
      <TaskItem item={item} onApprovalAction={this.handleApprovalAction} onError={this.handleErrors} />
    );
  }
  @autobind
  private async handleApprovalAction(item: ITask) {

    this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
      prevState.errors.length = 0;
      return prevState;
    });
    
    try {
      const result = await this.props.dataProvider.updateForCommittee(item, this.props.requestsByCommList);
        if (result) {
          let newItems = this.state.taskItems.filter((i) => i.Id !== item.Id);
          this.setState({ taskItems: newItems });
        }
    }
    catch (error) {
      console.log(error);
      this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
        prevState.errors.push("Error updating approval");
        return prevState;
      });
    }
  }
  @autobind
  private async handleErrors(errorMessage: string) {
    this.setState((prevState: ITaskListState, props: ITaskListProps): ITaskListState => {
      prevState.errors.push(errorMessage);
      return prevState;
    });
  }
  private renderErrors() {
    return this.state.errors.length > 0
      ?
      <div>
        {
          this.state.errors.map((item, idx) =>
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}
            >
              {item}
            </MessageBar>
          )
        }
      </div>
      : null;
  }
  private clearError(idx: number) {
    this.setState((prevState, props) => {
      return { ...prevState, errors: prevState.errors.splice(idx, 1) };
    });
  }
}