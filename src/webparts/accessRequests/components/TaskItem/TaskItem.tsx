import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from '../AccessRequests.module.scss';
import taskStyles from '../TaskItem/TaskItem.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import ITask from '../../models/ITask';

export interface ITaskItemProps {
    item: ITask;
    isApprover: boolean;
    onApprovalAction: any;
    onError: any;
    onShowRequest: any;
    onApprovalCommentsChanged: any;
}

export default class TaskItem extends React.Component<ITaskItemProps, null> {

    public render() {
        return (
            <div className={taskStyles.itemCell} data-is-focusable={true}>
                <div className={taskStyles.itemContent}>
                    <div><span className={taskStyles.itemLabel}>Name: </span>{this.props.item.Name}
                    </div>
                    <div><span className={taskStyles.itemLabel}>Committee:</span> {this.props.item.Committee}</div>
                    <div><span className={taskStyles.itemLabel}>Request: </span>
                        <Link href="#" onClick={this._onShowRequest} data-requestId={this.props.item.RequestId}>{this.props.item.RequestType}</Link>
                    </div>
                    <div><span className={taskStyles.itemLabel}>Status:</span>
                        <ul>
                            {this.props.item.RequestStatus.split('\n').map((item, key) => { return <li key={key}>{item}</li> })}
                        </ul>
                    </div>
                    <div><TextField placeholder='Comments' name='ApprovalComments'
                        value={this.props.item.ApprovalComments} multiline onChanged={this._onApprovalCommentsChanged} />
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
        const newItem: ITask = { ...this.props.item, Outcome: 'Approved' };
        this.props.onApprovalAction(newItem);
    }
    @autobind
    private _onItemRejected(event: React.MouseEvent<HTMLButtonElement>) {
        if (this.props.item.ApprovalComments === null || this.props.item.ApprovalComments.length < 1) {
            this.props.onError("Please enter a reason for rejecting this item.");
            return null;
        }
        const newItem: ITask = { ...this.props.item, Outcome: 'Rejected' };
        this.props.onApprovalAction(newItem);
    }
    @autobind
    private _onApprovalCommentsChanged(value: string) {
        const item: ITask = { ...this.props.item, ApprovalComments: value }
        this.props.onApprovalCommentsChanged(item);
    }
    @autobind
    private _onShowRequest(event: React.MouseEvent<HTMLButtonElement>) {
        const attributes: NamedNodeMap = event.currentTarget.attributes;
        const requestId = attributes.getNamedItem("data-requestId").value;
        this.props.onShowRequest(requestId);
    }
}
