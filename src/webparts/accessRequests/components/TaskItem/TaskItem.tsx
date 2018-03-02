import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from '../AccessRequests.module.scss';
import taskStyles from '../TaskItem/TaskItem.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import ITask from '../../models/ITask';

export interface ITaskItemProps {
    item: ITask;
    onApprovalAction: any;
}
export default class TaskList extends React.Component<ITaskItemProps, {updating:boolean}> {
    constructor(props) {
        super(props);
        this.state = {
            updating: false
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
                    <div><span className={taskStyles.itemLabel}>Status:</span> {this.props.item.RequestStatus}</div>
                    <div><span className={taskStyles.itemLabel}>Submitted:</span> {this.props.item.Created} </div>
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
        this.setState({updating: true});
        const newItem: ITask = {...this.props.item, Outcome:'Approved'};
        this.props.onApprovalAction(newItem);
    }
    @autobind
    private _onItemRejected(event: React.MouseEvent<HTMLButtonElement>) {
        this.setState({updating: true});
        const newItem: ITask = {...this.props.item, Outcome:'Rejected'};
        this.props.onApprovalAction(newItem);
    }    
}
