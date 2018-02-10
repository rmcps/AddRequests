import * as React from 'react';
import * as ReactDom from 'react-dom';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import styles from '../AccessRequests.module.scss';
// import NewAccessRequest from '../NewAccessRequest/NewAccessRequest';
// import ModifyAccessRequest from '../ModifyAccessRequest/ModifyAccessRequest';
// import { IAccessRequestsProps } from '../IAccessRequestsProps';
import IDisplayRequestProps from './IDisplayRequestProps';
// import DefaultPage from '../DefaultPage'

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
        // if (this.props.recordType == "New") {
        //     this.setState({
        //         message: "Your new access request was created.  You will receive email updates with the status of your request.",
        //         messageBarType: MessageBarType.success
        //     });
        // }
        // else if (this.props.recordType == "Change") {
        //     this.setState({
        //         message: "Your change request was created.  You will receive email updates with the status of your request.",
        //         messageBarType: MessageBarType.success
        //     });
        // }
    }
    public render(): React.ReactElement<IDisplayRequestProps> {
        return (
                    <div className={styles.row}>
                        <div className={styles.column}>
                            {this.props.item.RequestReason && <div className={styles.fieldContainer}>
                                <TextField label='Reason for request' disabled={true} value={this.props.item.RequestReason} />
                            </div>}
                            {this.props.item.Title && <div className={styles.fieldContainer}>
                                <TextField label='Name' disabled={true} value={this.props.item.Title} />
                            </div>}
                            {this.props.item.EMail && <div className={styles.fieldContainer}>
                                <TextField label='Email' disabled={true} value={this.props.item.EMail} />
                            </div>}
                            {this.props.item.JobTitle && <div className={styles.fieldContainer}>
                                <TextField label='Title' disabled={true} value={this.props.item.JobTitle} />
                            </div>}
                            {this.props.item.Company && <div className={styles.fieldContainer}>
                                <TextField label='Organization' disabled={true} value={this.props.item.Company} />
                            </div>}
                            {this.props.item.Office && <div className={styles.fieldContainer}>
                                <TextField label='Phone' disabled={true} value={this.props.item.Office} />
                            </div>}
                            {this.props.item.Comments && <div className={styles.fieldContainer}>
                                <TextField label='Comments' disabled={true} multiline value={this.props.item.Comments} />
                            </div>}
                            {this.props.item.AddCommittees.length>0 && <div className={styles.fieldContainer}>
                                <TextField label='Add Committees' disabled={true} multiline value={this.props.item.AddCommittees.join(", ")} />
                            </div>}
                            
                            {this.props.item.RemoveCommittees.length>0 && <div className={styles.fieldContainer}>
                                <TextField label='Remove Committees' disabled={true} multiline value={this.props.item.RemoveCommittees.join(", ")} />
                            </div>}
                            {this.props.additionalInfo && <div className={styles.fieldContainer}>
                                <TextField disabled={true} multiline value={this.props.additionalInfo} />
                            </div>}
                        </div>
                    </div>
        );
    }
    private _onAddNew():void {
        // const element: React.ReactElement<IAccessRequestsProps> = React.createElement(
        //     NewAccessRequest,
        //     {
        //         description: this.props.description,
        //         context: this.props.context,
        //         dom: this.props.dom,
        //     }
        // );
        // ReactDom.unmountComponentAtNode(this.props.dom);
        // ReactDom.render(element, this.props.dom);
    }
    private _onAddExisting():void {
        // const element: React.ReactElement<IAccessRequestsProps> = React.createElement(
        //     ModifyAccessRequest,
        //     {
        //         description: this.props.description,
        //         context: this.props.context,
        //         dom: this.props.dom,
        //     }
        // );
        // ReactDom.unmountComponentAtNode(this.props.dom);
        // ReactDom.render(element, this.props.dom);
    }
    @autobind
    private _onDisplayAll():void {
        // const element: React.ReactElement<IAccessRequestsProps> = React.createElement(
        //     DefaultPage,
        //     {
        //         description: this.props.description,
        //         context: this.props.context,
        //         dom: this.props.dom,
        //     }
        // );
        // ReactDom.unmountComponentAtNode(this.props.dom);
        // ReactDom.render(element, this.props.dom);
    }
    @autobind
    private _onCancel():void {
    //   window.location.href = "https://uphpcin.sharepoint.com";
    }
    
}
