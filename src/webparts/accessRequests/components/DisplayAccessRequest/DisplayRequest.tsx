import * as React from 'react';
import * as ReactDom from 'react-dom';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import styles from '../AccessRequests.module.scss';
import IDisplayRequestProps from './IDisplayRequestProps';

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
    }
    public componentWillReceiveProps() {
    }

    public componentWillMount() {
    }
    public render(): React.ReactElement<IDisplayRequestProps> {
        return (
            <div>
        <IconButton
          className={styles.chevronReturn}
          disabled={false}
          iconProps={{ iconName: 'ChevronLeftMed' }}
          title='Return to list'
          ariaLabel='Return to list'
          onClick={this._onReturn}
        />                
                <div className={styles.row}>
                    <div className={styles.column1}>

                        {this.props.item.RequestReason && <div className={styles.fieldContainer}>
                            <TextField label='Reason for request' disabled={true} value={this.props.item.RequestReason} />
                        </div>}
                        {this.props.item.Title && <div className={styles.fieldContainer}>
                            <TextField label='Name' disabled={true} value={this.props.item.Title} />
                        </div>}
                        {this.props.item.EMail && <div className={styles.fieldContainer}>
                            <TextField label='Email' disabled={true} value={this.props.item.EMail} />
                        </div>}
                    </div>
                    <div className={styles.column1}>
                        {this.props.item.JobTitle && <div className={styles.fieldContainer}>
                            <TextField label='Title' disabled={true} value={this.props.item.JobTitle} />
                        </div>}
                        {this.props.item.Company && <div className={styles.fieldContainer}>
                            <TextField label='Organization' disabled={true} value={this.props.item.Company} />
                        </div>}
                        {this.props.item.Office && <div className={styles.fieldContainer}>
                            <TextField label='Phone' disabled={true} value={this.props.item.Office} />
                        </div>}
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column2}>

                        {this.props.item.Comments && <div className={styles.fieldContainer}>
                            <TextField label='Comments' disabled={true} multiline value={this.props.item.Comments} />
                        </div>}
                        {this.props.item.AddCommittees.length > 0 && <div className={styles.fieldContainer}>
                            <TextField label='Add Committees' disabled={true} multiline value={this.props.item.AddCommittees.join(", ")} />
                        </div>}

                        {this.props.item.RemoveCommittees.length > 0 && <div className={styles.fieldContainer}>
                            <TextField label='Remove Committees' disabled={true} multiline value={this.props.item.RemoveCommittees.join(", ")} />
                        </div>}
                        {this.props.item.RequestStatus && <div className={styles.fieldContainer}>
                            <TextField label='Status' disabled={true} multiline value={this.props.item.RequestStatus} />
                        </div>}
                    </div>
                </div>
            </div>
        );
    }
    @autobind
    private _onReturn() {
        this.props.onReturnClick("list");
    }
}
