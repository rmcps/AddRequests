import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from '../AccessRequests.module.scss';
import taskStyles from '../TaskItem/TaskItem.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import ITask from '../../models/ITask';

export interface ITaskItemProps {
    item: ITask;
    onApprovalAction: any;
    onError: any;
}
export interface ITaskItemState {
    approvalComments: string;
}
export default class TaskList extends React.Component<ITaskItemProps, ITaskItemState> {
    constructor(props) {
        super(props);
        this.state = {
            approvalComments: null
        };
    }
    public render() {
        return (
            <div className={taskStyles.itemCell} data-is-focusable={true}>
                <div className={taskStyles.itemContent}>
                    <div><span className={taskStyles.itemLabel}>Name:</span> {this.props.item.Name}</div>
                    <div><span className={taskStyles.itemLabel}>Committee:</span> {this.props.item.Committee}</div>
                    <div><span className={taskStyles.itemLabel}>Request:</span> {this.props.item.RequestType}</div>
                    <div><span className={taskStyles.itemLabel}>Status:</span>
                        <ul>
                            {this.props.item.RequestStatus.split('\n').map((item, key) => { return <li key={key}>{item}</li> })}
                        </ul>
                    </div>                    
                    <div><TextField placeholder='Comments' name='ApprovalComments'
                        value={this.state.approvalComments} multiline onChanged={this._onApprovalCommentsChanged} />
                    </div>
                </div>
                <div className={taskStyles.actionIconsContainer}>
                    <IconButton
                        data-action='Approved'
                        className={taskStyles.approveButton}
                        iconProps={{ iconName: 'Accept' }}
                        disabled={false}
                        title='Approve'
                        ariaLabel='Approve Item'
                        onClick={this._onItemApproved}
                    />
                    <IconButton
                        data-action='Rejected'
                        className={taskStyles.rejectButton}
                        disabled={false}
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
        const newItem: ITask = { ...this.props.item, Outcome: 'Approved', ApprovalComments: this.state.approvalComments };
        this.props.onApprovalAction(newItem);
    }
    @autobind
    private _onItemRejected(event: React.MouseEvent<HTMLButtonElement>) {
        if (this.state.approvalComments === null || this.state.approvalComments.length < 1) {
            this.props.onError("Please enter a reason for rejecting this item.");
            return null;
        }
        const newItem: ITask = { ...this.props.item, Outcome: 'Rejected', ApprovalComments: this.state.approvalComments };
        this.props.onApprovalAction(newItem);
    }
    @autobind
    private _onApprovalCommentsChanged(value: string) {
        this.setState((prevState: ITaskItemState, props: ITaskItemProps): ITaskItemState => {
            prevState.approvalComments = value;
            return prevState;
        });

    }
}
