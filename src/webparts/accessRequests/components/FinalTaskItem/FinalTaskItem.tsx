import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from '../AccessRequests.module.scss';
import taskStyles from '../FinalTaskItem/FinalTaskItem.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Link } from 'office-ui-fabric-react/lib/Link';
import IFinalTask from '../../models/IFinalTask';

export interface IFinalTaskItemProps {
    item: IFinalTask;
    onApprovalAction: any;
    onError: any;
    onShowRequest;
}
export interface IFinalTaskItemState {
    approvalComments: string;
}
export default class TaskList extends React.Component<IFinalTaskItemProps, IFinalTaskItemState> {
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
                    <div>
                        <span className={taskStyles.itemLabel}>Name: </span>{this.props.item.Title}
                    </div>
                    <div><span className={taskStyles.itemLabel}>Reason: </span>
                        <Link href="#" onClick={this._onShowRequest} data-requestId={this.props.item.Id}>{this.props.item.RequestReason}</Link>
                    </div>
                    <div><TextField placeholder='Your Comments' name='ApprovalComments'
                        value={this.state.approvalComments} multiline onChanged={this._onApprovalCommentsChanged} />
                    </div>                    
                    <div><span className={taskStyles.itemLabel}>Approval History:</span>
                        <ul>
                            {this.props.item.RequestStatus ? this.props.item.RequestStatus.split('\n').map((item, key) => { return <li key={key}>{item}</li>; }) : ""}
                        </ul>
                    </div>
                    {this.props.item.CommitteeTasks.map(c => (
                        <CommitteeItem item={c} />
                    ))}
                </div>
                <div className={taskStyles.actionIconsContainer}>
                    <IconButton
                        data-action='Approved'
                        className={taskStyles.approveButton}
                        disabled={false}
                        iconProps={{ iconName: 'Accept' }}
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
                    <IconButton
                        data-action='Canceled'
                        className={taskStyles.cancelButton}
                        disabled={false}
                        iconProps={{ iconName: 'Blocked' }}
                        title='Cancel'
                        ariaLabel='Cancel Item'
                        onClick={this._onItemCanceled}
                    />
                </div>
            </div>
        );
    }
    @autobind
    private _onItemApproved(event: React.MouseEvent<HTMLButtonElement>) {
        const newItem: IFinalTask = { ...this.props.item, Outcome: 'Approved', CompletionStatus: 'Pending', ApprovalComments: this.state.approvalComments };
        this.props.onApprovalAction(newItem);
    }
    @autobind
    private _onItemRejected(event: React.MouseEvent<HTMLButtonElement>) {
        if (this.state.approvalComments === null || this.state.approvalComments.length < 1) {
            this.props.onError("Please enter a reason for rejecting this item.");
            return null;
        }
        const newItem: IFinalTask = { ...this.props.item, Outcome: 'Rejected', CompletionStatus: 'Pending', ApprovalComments: this.state.approvalComments };
        this.props.onApprovalAction(newItem);
    }
    @autobind
    private _onItemCanceled(event: React.MouseEvent<HTMLButtonElement>) {
        if (this.state.approvalComments === null || this.state.approvalComments.length < 1) {
            this.props.onError("Please enter a reason for canceling this item.");
            return null;
        }
        const newItem: IFinalTask = { ...this.props.item, Outcome: 'Canceled', CompletionStatus: 'Pending', ApprovalComments: this.state.approvalComments };
        this.props.onApprovalAction(newItem);
    }    
    @autobind
    private _onApprovalCommentsChanged(value: string) {
        this.setState((prevState: IFinalTaskItemState, props: IFinalTaskItemProps): IFinalTaskItemState => {
            prevState.approvalComments = value;
            return prevState;
        });

    }
    @autobind
    private _onShowRequest(event: React.MouseEvent<HTMLButtonElement>) {
        const attributes: NamedNodeMap = event.currentTarget.attributes;
        const requestId = attributes.getNamedItem("data-requestId").value;
        this.props.onShowRequest(requestId);
    }
}
function CommitteeItem(props) {
    return (
        <div className={styles.row}>
            <div className={taskStyles.itemContainer}>
                <div className={styles.column1}><span className={taskStyles.itemLabel}>Committee:</span> {props.item.Committee}</div>
                <div className={styles.column1}><span className={taskStyles.itemLabel}>Outcome:</span> {props.item.Outcome}</div>
                {props.item.ApprovedBy && <div className={styles.column1}><span className={taskStyles.itemLabel}>Approver:</span> {props.item.ApprovedBy}</div>}
            </div>
        </div>
    );
}
