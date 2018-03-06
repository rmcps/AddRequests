import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from '../AccessRequests.module.scss';
import taskStyles from '../TaskItem/TaskItem.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import ITask from '../../models/ITask';

export interface ITaskItemProps {
    item: ITask;
    onApprovalAction: any;
    onError: any;
}
export interface ITaskItemState {
    updating:boolean; 
    approvalComments:string;
    errorMsg: string;
}
export default class TaskList extends React.Component<ITaskItemProps, ITaskItemState> {
    constructor(props) {
        super(props);
        this.state = {
            updating: false,
            approvalComments: null,
            errorMsg: '',
        };
    }
    public componentWillReceiveProps(nextProps: ITaskItemProps) {
        this.setState({ updating: false});
    }
    public shouldComponentUpdate(newProps: ITaskItemProps) {
        return (
            this.props.item !== newProps.item ||
            this.props.item.Outcome != newProps.item.Outcome
        );
    }
    public render() {
        return (
            <div className={taskStyles.itemCell} data-is-focusable={true}>
                <div className={taskStyles.itemContent}>
                    <div><span className={taskStyles.itemLabel}>Name:</span> {this.props.item.Name}</div>
                    <div><span className={taskStyles.itemLabel}>Committee:</span> {this.props.item.Committee}</div>
                    <div><span className={taskStyles.itemLabel}>Status:</span> {this.props.item.RequestStatus.split('\n').map((item, key) => {return <span key={key}>{item}<br/></span>})}</div>
                    <div><span className={taskStyles.itemLabel}>Submitted:</span> {this.props.item.Created} </div>
                    <div><TextField placeholder='Comments' name='ApprovalComments' 
                        value={this.state.approvalComments} multiline onChanged={this._onApprovalCommentsChanged} /> 
                    </div>
                </div>
                <div className={taskStyles.actionIconsContainer}>
                    <IconButton
                        data-action='Approved'
                        className={taskStyles.actionIcons}
                        disabled={this.state.updating}
                        iconProps={{ iconName: 'Accept' }}
                        title='Approve'
                        ariaLabel='Approve Item'
                        onClick={this._onItemApproved}
                    />
                    <IconButton
                        data-action='Rejected'
                        className={taskStyles.actionIcons}
                        disabled={this.state.updating}
                        iconProps={{ iconName: 'Cancel' }}
                        title='Reject'
                        ariaLabel='Reject Item'
                        onClick={this._onItemRejected}
                    />
                </div>
            </div>
        );
    }
    @autobind
    private _onItemApproved(event: React.MouseEvent<HTMLButtonElement>) {
        this.setState((prevState: ITaskItemState, props: ITaskItemProps): ITaskItemState => {
            prevState.updating = true;
            return prevState;
          });
        const newItem: ITask = {...this.props.item, Outcome:'Approved', ApprovalComments: this.state.approvalComments};
        this.props.onApprovalAction(newItem);
    }
    @autobind
    private _onItemRejected(event: React.MouseEvent<HTMLButtonElement>) {
        if (this.state.approvalComments === null || this.state.approvalComments.length < 1) {
            this.props.onError("Please enter a reason for rejecting this item.");
            return null;
        }
        const newItem: ITask = {...this.props.item, Outcome:'Rejected',ApprovalComments: this.state.approvalComments};
        this.props.onApprovalAction(newItem);
    }    
    @autobind
    private _onApprovalCommentsChanged(value:string) {
        this.setState((prevState: ITaskItemState, props: ITaskItemProps): ITaskItemState => {
            prevState.approvalComments = value;
            return prevState;
          });
      
    }
}
