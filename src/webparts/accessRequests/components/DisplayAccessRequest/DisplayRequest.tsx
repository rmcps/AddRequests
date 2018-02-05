import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import styles from '../AccessRequests.module.scss';
import NewAccessRequest from '../NewAccessRequest/NewAccessRequest';
import ModifyAccessRequest from '../ModifyAccessRequest/ModifyAccessRequest';
import { IAccessRequestsProps } from '../IAccessRequestsProps';
import IDisplayRequestProps from './IDisplayRequestProps';

export interface IDisplayRequestState {
    message: string;
    messageBarType: MessageBarType;
    committeeString: string;
}
export default class DisplayRequest extends React.Component<IDisplayRequestProps, IDisplayRequestState> {

    constructor(props: IDisplayRequestProps) {
        super(props);

        this.state = {
            message: '',
            messageBarType: MessageBarType.info,
            committeeString: ''
        };
        this._onAddNew = this._onAddNew.bind(this);
        this._onAddExisting = this._onAddExisting.bind(this);        
    }
    public componentWillMount() {
        if (this.props.recordType == "New") {
            this.setState({
                message: "Your new access request was created.  You will receive email updates with the status of your request.",
                messageBarType: MessageBarType.success
        });
        }
        else if (this.props.recordType == "Modified") {
            this.setState({
                message: "Your access request changed request was created.  You will receive email updates with the status of your request.",
                messageBarType: MessageBarType.success
        });
        }
        this.setState({
            committeeString: this.props.Committees.map(c => c.Title).join(",")
        });
    }
    public render(): React.ReactElement<IDisplayRequestProps> {
        return (
            <div className={styles.accessRequests }>
                <div className={ styles.container }>
                <div className= {styles.row}>
                <MessageBar messageBarType={this.state.messageBarType}>{this.state.message}</MessageBar>
                </div>
                    <div className= {styles.row}>
                    <Link onClick={this._onAddNew}>Add a new member access request</Link> &nbsp;&nbsp;&nbsp;
                    <Link onClick={this._onAddExisting}>Add a requet to modify an existing member</Link>
                    </div>
                     <div className= {styles.row}>
                       <div className={ styles.column }>  
                        <TextField label='First Name' disabled={ true } value={this.props.FirstName} />
                        <TextField label='Last Name' disabled={ true } value={this.props.LastName} />
                        <TextField label='Email' disabled={ true } value={this.props.EMail} />
                        {this.props.JobTitle && <TextField label='Title' disabled={ true } value={this.props.JobTitle} />}
                        <TextField label='Organization' disabled={ true } value={this.props.Company} />
                        {this.props.Office && <TextField label='Phone' disabled={ true } value={this.props.Office} /> }
                        {this.props.Comments && <TextField label='Comments' disabled={ true } multiline value={this.props.Comments} /> }
                        <TextField label='Committees' disabled={ true } multiline
                            value={this.state.committeeString}
                        />
                        </div>
                    </div>
                </div>
            </div>
        );
    }   
    private _onAddNew() {
        const element: React.ReactElement< IAccessRequestsProps > = React.createElement(
            NewAccessRequest,
            {
              description: this.props.description,
              context:this.props.context,
              dom: this.props.dom,
            }
          );
          ReactDom.unmountComponentAtNode(this.props.dom);          
          ReactDom.render(element, this.props.dom);
      }
      private _onAddExisting() {
        const element: React.ReactElement<IAccessRequestsProps > = React.createElement(
          ModifyAccessRequest,
            {
              description: this.props.description,
              context:this.props.context,
              dom: this.props.dom,
            }
          );
          ReactDom.unmountComponentAtNode(this.props.dom);          
          ReactDom.render(element, this.props.dom);
      }
         
}
