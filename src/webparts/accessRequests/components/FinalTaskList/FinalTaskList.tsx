import * as React from 'react';
import styles from '../AccessRequests.module.scss';
import taskStyles from '../FinalTaskList/FinalTaskList.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import IFinalTask from '../../models/IFinalTask';
import FinalTaskItem from '../FinalTaskItem/FinalTaskItem';

export interface IFinalTaskProps {
  //onItemSelected: any;
  dataProvider: IAccessRequestsDataProvider;
  requestsByCommList: string;
  currentUser: any;
  onTaskItemSelected: any;
}
export interface IFinalTaskState {
  taskItems: IFinalTask[];
  dataIsLoading: boolean;
  errorMsg: string;
  hideDialog: boolean;
}
export default class FinalTaskList extends React.Component<IFinalTaskProps, IFinalTaskState> {
  constructor(props: IFinalTaskProps) {
    super(props);
    this.state = {
      taskItems: [],
      dataIsLoading: true,
      errorMsg: null,
      hideDialog: true
    };
  }
  public async componentWillReceiveProps(nextProps: IFinalTaskProps) {
    try {
      let result = await this.props.dataProvider.getFinalTasks(this.props.requestsByCommList, this.props.currentUser);
      this.setState({
        taskItems: result,
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
      const result = await this.props.dataProvider.getFinalTasks(this.props.requestsByCommList, this.props.currentUser);
      this.setState({
        taskItems: result,
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
      <div>
        <div className={styles.row}>
          <div className={styles.column2}>
            {this.state.taskItems.length > 1 &&
                <DefaultButton disabled={false} text='Approve All' onClick={this._onApproveAllTasks} />
            }
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column2}>
            {this.state.taskItems.length > 0 && <h3>Your approval is requested on the items below</h3>}
            {this.state.taskItems.length == 0 && <h3>No pending final approvals</h3>}
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
      </div>
    );
  }
  @autobind
  private _onRenderCell(item: IFinalTask, index: number | undefined): JSX.Element {
    return (
      <FinalTaskItem item={item} onApprovalAction={this._handleApprovalAction} onError={this._handleItemError} onShowRequest={this._handleShowRequest} />
    );
  }
  @autobind
  private async _handleApprovalAction(item: IFinalTask) {

    this.setState((prevState: IFinalTaskState, props: IFinalTaskProps): IFinalTaskState => {
      prevState.errorMsg = null;
      prevState.hideDialog = false;
      return prevState;
    });
    let itemArr: IFinalTask[] = new Array(item);
    try {
      const result = await this.props.dataProvider.updateForRequest(itemArr);
      if (result) {
        let newItems = this.state.taskItems.filter((i) => i.Id !== item.Id);
        this.setState((prevState: IFinalTaskState, props: IFinalTaskProps): IFinalTaskState => {
          prevState.taskItems = newItems;
          prevState.hideDialog = true;
          return prevState;
        });
      }
    }
    catch (error) {
      console.log(error);
      this.setState((prevState: IFinalTaskState, props: IFinalTaskProps): IFinalTaskState => {
        prevState.errorMsg = "Error updating approval";
        prevState.hideDialog = true;
        return prevState;
      });
    }
  }
  @autobind
  private _handleItemError(errorMessage: string) {
    this.setState((prevState: IFinalTaskState, props: IFinalTaskProps): IFinalTaskState => {
      prevState.errorMsg = errorMessage;
      return prevState;
    });
  }
  @autobind
  private _handleShowRequest(requestId) {
    this.props.onTaskItemSelected(requestId, "FinalTasks");
  }
  @autobind
  private async _onApproveAllTasks() {
    this.setState((prevState: IFinalTaskState, props: IFinalTaskProps): IFinalTaskState => {
      prevState.errorMsg = null;
      prevState.hideDialog = false;
      return prevState;
    });
    try {
      let items: IFinalTask[] = this.state.taskItems.map(el => {
        return { ...el, Outcome: 'Approved', CompletionStatus: 'Pending', };
      });
      
      const result = await this.props.dataProvider.updateForRequest(items);
      if (result) {
        this.setState((prevState: IFinalTaskState, props: IFinalTaskProps): IFinalTaskState => {
          prevState.taskItems = [];
          prevState.hideDialog = true;
          return prevState;
        });
      }
    }
    catch (error) {
      console.log(error);
      this.setState((prevState: IFinalTaskState, props: IFinalTaskProps): IFinalTaskState => {
        prevState.errorMsg = "Error. Not all items approved.";
        prevState.hideDialog = true;
        return prevState;
      });
    }
  }
  
}