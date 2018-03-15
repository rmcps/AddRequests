import * as React from 'react';
import * as ReactDom from 'react-dom';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import styles from '../AccessRequests.module.scss';
import styles2 from './DisplayRequest.module.scss';
import IDisplayRequestProps from './IDisplayRequestProps';
import IAccessRequest from '../../models/IAccessRequest';

export interface IDisplayRequestState {
    message: string;
    messageBarType: MessageBarType;
    item: IAccessRequest;
    dataIsLoading: boolean;
    errorMsg: string;
}
export default class DisplayRequest extends React.Component<IDisplayRequestProps, IDisplayRequestState> {

    constructor(props: IDisplayRequestProps) {
        super(props);

        this.state = {
            message: '',
            messageBarType: MessageBarType.info,
            item: {},
            dataIsLoading: true,
            errorMsg: null
        };
    }
    public async componentWillReceiveProps(nextProps: IDisplayRequestProps) {
        try {
            const result = await this.props.dataProvider.getItem(this.props.requestId, this.props.requestsByCommList);
            this.setState({
                item: result,
                dataIsLoading: false,
                errorMsg: null
            });
        }
        catch (error) {
            console.log(error);
            this.setState({
                dataIsLoading: false,
                errorMsg: 'An error occured loading this request.'
            });
        }
    }
    public async componentDidMount() {
        if (undefined == this.state.item || this.state.item == null
            || undefined == this.state.item.Id || this.state.item.Id == null) {
            try {
                const result = await this.props.dataProvider.getItem(this.props.requestId, this.props.requestsByCommList);
                this.setState({
                    item: result,
                    dataIsLoading: false,
                    errorMsg: null
                });
            }
            catch (error) {
                console.log(error);
                this.setState({
                    dataIsLoading: false,
                    errorMsg: 'An error occured loading this request.'
                });
            }
        }
    }
    public render(): React.ReactElement<IDisplayRequestProps> {
        return (
            <div>
                {this.state.errorMsg ? <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={true}>
                    {this.state.errorMsg}
                </MessageBar>
                    : null
                }
                {this.state.dataIsLoading ? <Spinner size={SpinnerSize.medium} /> : null}
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
                        {this.state.item.RequestReason && <div className={styles.fieldContainer}>
                            <TextField label='Reason for request' disabled={true} value={this.state.item.RequestReason} />
                        </div>}
                        {this.state.item.Title && <div className={styles.fieldContainer}>
                            <TextField label='Name' disabled={true} value={this.state.item.Title} />
                        </div>}
                        {this.state.item.JobTitle && <div className={styles.fieldContainer}>
                            <TextField label='Title' disabled={true} value={this.state.item.JobTitle} />
                        </div>}
                    </div>
                    <div className={styles.column1}>
                        {this.state.item.Company && <div className={styles.fieldContainer}>
                            <TextField label='Organization' disabled={true} value={this.state.item.Company} />
                        </div>}
                        {this.state.item.EMail && <div className={styles.fieldContainer}>
                            <TextField label='Email' disabled={true} value={this.state.item.EMail} />
                        </div>}
                        {this.state.item.Office && <div className={styles.fieldContainer}>
                            <TextField label='Phone' disabled={true} value={this.state.item.Office} />
                        </div>}
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column2}>

                        {this.state.item.Comments && <div className={styles.fieldContainer}>
                            <TextField label='Comments' disabled={true} multiline value={this.state.item.Comments} />
                        </div>}
                        {this.state.item.AddCommittees && this.state.item.AddCommittees.length > 0 && <div className={styles.fieldContainer}>
                            <TextField label='Add Committees' disabled={true} multiline value={this.state.item.AddCommittees.join(", ")} />
                        </div>}

                        {this.state.item.RemoveCommittees && this.state.item.RemoveCommittees.length > 0 && <div className={styles.fieldContainer}>
                            <TextField label='Remove Committees' disabled={true} multiline value={this.state.item.RemoveCommittees.join(", ")} />
                        </div>}
                        {this.state.item.RequestStatus && <div className={styles.fieldContainer}>
                            <TextField label='Status' disabled={true} multiline value={this.state.item.RequestStatus} />
                        </div>}
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column2}>
                        {this.state.item && this.state.item.CommitteeApprovals && this.state.item.CommitteeApprovals.length > 0 &&
                            <div className={styles2.committeesList}>
                                <span className={styles2.header}>Committees</span>
                                <ul>
                                    {this.state.item.CommitteeApprovals.map((item, key) => {
                                        return <li key={key}>{item.Committee}<span>{item.RequestType}</span>
                                            {item.RequestStatus ? item.RequestStatus.split('\n').map((item, key) => { return <span key={key}>{item}</span> }) : ""}
                                        </li>
                                    })
                                    }
                                </ul>
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
