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
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import IFinalTask from '../../models/IFinalTask';
import FinalTaskItem from '../FinalTaskItem/FinalTaskItem';

export interface IFinalTaskProps {
  //onItemSelected: any;
  dataProvider: IAccessRequestsDataProvider;
  requestsByCommList: string;
}
export interface IFinalTaskState {
  taskItems: IFinalTask[];
  dataIsLoading: boolean;
  errors: string[];
}
export default class FinalTaskList extends React.Component<IFinalTaskProps, IFinalTaskState> {
  constructor(props: IFinalTaskProps) {
    super(props);
    this.state = {
      taskItems: [],
      dataIsLoading: true,
      errors: [],
    };
  }
  public async componentWillReceiveProps(nextProps: IFinalTaskProps) {
    try {
      let result = await this.props.dataProvider.getFinalTasks(this.props.requestsByCommList);
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
      const result = await this.props.dataProvider.getFinalTasks(this.props.requestsByCommList);
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
      <div className={styles.row}>
        <div className={styles.column2}>
          <h3>Approvals</h3>
          {this.state.dataIsLoading ? <Spinner size={SpinnerSize.medium} /> : null}          
          <Fabric>
            <FocusZone direction={FocusZoneDirection.vertical}>
              <List
                items={this.state.taskItems} className={taskStyles.TaskList}
                onRenderCell={this._onRenderCell}
              />
            </FocusZone>
          </Fabric>
          {this.renderErrors()}
        </div>
      </div>
    );
  }
  @autobind
  private _onRenderCell(item: IFinalTask, index: number | undefined): JSX.Element {
    return (
      <FinalTaskItem item={item} onApprovalAction={this._handleApprovalAction} />
    );
  }
  @autobind
  private async _handleApprovalAction(item: IFinalTask) {

    this.setState((prevState: IFinalTaskState, props: IFinalTaskProps): IFinalTaskState => {
      prevState.errors.length = 0;
      return prevState;
    });
    try {
      const result = await this.props.dataProvider.updateForRequest(item.Id,
            item.CompletionStatus == "Approved" ? "Approved" : "Rejected");
        if (result) {
          let newItems = this.state.taskItems.filter((i) => i.Id !== item.Id);
          this.setState({ taskItems: newItems });
        }
    }
    catch (error) {
      console.log(error);
      this.setState((prevState: IFinalTaskState, props: IFinalTaskProps): IFinalTaskState => {
        prevState.errors.push("Error updating approval");
        return prevState;
      });
    }
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