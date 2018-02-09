import * as React from 'react';
import * as ReactDom from 'react-dom';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import styles from '../AccessRequests.module.scss';
import NewAccessRequest from '../NewAccessRequest/NewAccessRequest';
import ModifyAccessRequest from '../ModifyAccessRequest/ModifyAccessRequest';
import { IAccessRequestsProps } from '../IAccessRequestsProps';
import IDisplayRequestProps from './IDisplayRequestProps';
import DefaultPage from '../DefaultPage'

export interface IDisplayRequestState {
    message: string;
    messageBarType: MessageBarType;
}
export default class DisplayRequest extends React.Component<IDisplayRequestProps, IDisplayRequestState> {

    constructor(props: IDisplayRequestProps) {
        super(props);

        this.state = {
            message: '',
            messageBarType: MessageBarType.info,
        };
        this._onAddNew = this._onAddNew.bind(this);
        this._onAddExisting = this._onAddExisting.bind(this);
    }
    public componentWillReceiveProps() {
    }

    public componentWillMount() {
        if (this.props.recordType == "New") {
            this.setState({
                message: "Your new access request was created.  You will receive email updates with the status of your request.",
                messageBarType: MessageBarType.success
            });
        }
        else if (this.props.recordType == "Change") {
            this.setState({
                message: "Your change request was created.  You will receive email updates with the status of your request.",
                messageBarType: MessageBarType.success
            });
        }
    }
    public render(): React.ReactElement<IDisplayRequestProps> {
        debugger
        return (
            <div className={styles.accessRequests}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <h2 className={styles.headerBar}>Member Access Request Submission</h2>

                            {this.state.message && <MessageBar messageBarType={this.state.messageBarType}>{this.state.message}</MessageBar>}

                            <div><Link onClick={this._onAddNew}>Add a new member access request</Link> </div>
                            <div><Link onClick={this._onAddExisting}>Add a requet to modify an existing member</Link></div>
                            <div><Link onClick={this._onDisplayAll}>View All Requests</Link></div>
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            {this.props.RequestReason && <div className={styles.fieldContainer}>
                                <TextField label='Reason for request' disabled={true} value={this.props.RequestReason} />
                            </div>}
                            {this.props.Title && <div className={styles.fieldContainer}>
                                <TextField label='Name' disabled={true} value={this.props.Title} />
                            </div>}
                            {this.props.EMail && <div className={styles.fieldContainer}>
                                <TextField label='Email' disabled={true} value={this.props.EMail} />
                            </div>}
                            {this.props.JobTitle && <div className={styles.fieldContainer}>
                                <TextField label='Title' disabled={true} value={this.props.JobTitle} />
                            </div>}
                            {this.props.Company && <div className={styles.fieldContainer}>
                                <TextField label='Organization' disabled={true} value={this.props.Company} />
                            </div>}
                            {this.props.Office && <div className={styles.fieldContainer}>
                                <TextField label='Phone' disabled={true} value={this.props.Office} />
                            </div>}
                            {this.props.Comments && <div className={styles.fieldContainer}>
                                <TextField label='Comments' disabled={true} multiline value={this.props.Comments} />
                            </div>}
                            {this.props.AddCommittees.length>0 && <div className={styles.fieldContainer}>
                                <TextField label='Add Committees' disabled={true} multiline value={this.props.AddCommittees.join(", ")} />
                            </div>}
                            
                            {this.props.RemoveCommittees.length>0 && <div className={styles.fieldContainer}>
                                <TextField label='Remove Committees' disabled={true} multiline value={this.props.RemoveCommittees.join(", ")} />
                            </div>}
                            {this.props.additionalInfo && <div className={styles.fieldContainer}>
                                <TextField disabled={true} multiline value={this.props.additionalInfo} />
                            </div>}
                        </div>
                    </div>
                </div>
            </div>
        );
    }
    private _onAddNew():void {
        const element: React.ReactElement<IAccessRequestsProps> = React.createElement(
            NewAccessRequest,
            {
                description: this.props.description,
                context: this.props.context,
                dom: this.props.dom,
            }
        );
        ReactDom.unmountComponentAtNode(this.props.dom);
        ReactDom.render(element, this.props.dom);
    }
    private _onAddExisting():void {
        const element: React.ReactElement<IAccessRequestsProps> = React.createElement(
            ModifyAccessRequest,
            {
                description: this.props.description,
                context: this.props.context,
                dom: this.props.dom,
            }
        );
        ReactDom.unmountComponentAtNode(this.props.dom);
        ReactDom.render(element, this.props.dom);
    }
    @autobind
    private _onDisplayAll():void {
        const element: React.ReactElement<IAccessRequestsProps> = React.createElement(
            DefaultPage,
            {
                description: this.props.description,
                context: this.props.context,
                dom: this.props.dom,
            }
        );
        ReactDom.unmountComponentAtNode(this.props.dom);
        ReactDom.render(element, this.props.dom);
    }
    @autobind
    private _onCancel():void {
      window.location.href = "https://uphpcin.sharepoint.com";
    }
    
}
