import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from '../AccessRequests.module.scss';
import taskStyles from '../FinalTaskItem/FinalTaskItem.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import IFinalTask from '../../models/IFinalTask';

export interface IFinalTaskItemProps {
    item: IFinalTask;
    onApprovalAction: any;
    onError: any;
}
export interface IFinalTaskItemState {
    updating:boolean; 
    approvalComments:string;
    errorMsg: string;
}
export default class TaskList extends React.Component<IFinalTaskItemProps, IFinalTaskItemState> {
    constructor(props) {
        super(props);
        this.state = {
            updating: false,
            approvalComments: null,
            errorMsg:null
        };
    }
    public componentWillReceiveProps(nextProps: IFinalTaskItemProps) {
        this.setState({ updating: false});
    }
    public shouldComponentUpdate(newProps: IFinalTaskItemProps) {
        return (
            this.props.item !== newProps.item //||
            //this.props.item.Outcome != newProps.item.Outcome
        );
    }
    public render() {
        return (
            <div className={taskStyles.itemCell} data-is-focusable={true}>
                <div className={taskStyles.itemContent}>
                    <div><span className={taskStyles.itemLabel}>Name:</span> {this.props.item.Title}</div>
                    <div><span className={taskStyles.itemLabel}>Reason:</span> {this.props.item.RequestReason}</div>
                    <div><span className={taskStyles.itemLabel}>Status:</span> 
                        {this.props.item.RequestStatus ? this.props.item.RequestStatus.split('\n').map((item, key) => {return <span key={key}>{item}<br/></span>}): ""}
                    </div>
                    {this.props.item.CommitteeTasks.map(c => (
                    <CommitteeItem item={c} />
                    ))}
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
        this.setState({updating: true});
        const newItem: IFinalTask = {...this.props.item, CompletionStatus:'Approved', ApprovalComments: this.state.approvalComments};
        this.props.onApprovalAction(newItem);
    }
    @autobind
    private _onItemRejected(event: React.MouseEvent<HTMLButtonElement>) {
        if (this.state.approvalComments === null || this.state.approvalComments.length < 1) {
            this.props.onError("Please enter a reason for rejecting this item.");
            return null;
        }        
        this.setState({updating: true});
        const newItem: IFinalTask = {...this.props.item, CompletionStatus:'Rejected', ApprovalComments: this.state.approvalComments};
        this.props.onApprovalAction(newItem);
    }    
    @autobind
    private _onApprovalCommentsChanged(value:string) {
        this.setState((prevState: IFinalTaskItemState, props: IFinalTaskItemProps): IFinalTaskItemState => {
            prevState.approvalComments = value;
            return prevState;
          });
      
    }    
}
function CommitteeItem(props) {
    return (
        <div className={styles.row}>
        <div className={styles.column1}><span className={taskStyles.itemLabel}>Committee:</span> {props.item.Committee}</div>
        <div className={styles.column1}><span className={taskStyles.itemLabel}>Outcome:</span> {props.item.Outcome}</div>
        <div className={styles.column1}><span className={taskStyles.itemLabel}>Status:</span> 
            {props.item.RequestStatus ? props.item.RequestStatus.split('\n').map((item, key) => {return <span key={key}>{item}<br/></span>}) : ""}
        </div>
        </div>
    );
}
