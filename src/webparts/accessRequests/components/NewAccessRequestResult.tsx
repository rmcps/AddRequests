import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import styles from './AccessRequests.module.scss';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import INewAccessRequest from '../models/INewAccessRequest';

export interface INewAccessRequestResultProps {
    status:string;
    newAccessRequest: INewAccessRequest;
}

export default class INewAccessRequestResult extends React.Component<INewAccessRequestResultProps, null> {

    constructor(props: INewAccessRequestResultProps) {
        super(props);
        this._redirectToHome = this._redirectToHome.bind(this);
    }
    public render(): React.ReactElement<INewAccessRequestResultProps> {
        return (
            <div className={ styles.accessRequests }>
                <div className={ styles.container }>
                    <div className= {styles.row}>
                        <MessageBar messageBarType={MessageBarType.warning}>{this.props.status}</MessageBar>
                    </div>
                    <div className= {styles.row}>
                        <div>Would you like to add a new access request?</div>
                    </div>
                    <div className= {styles.row}>
                    <div className={ styles.formButtonsContainer}>
                <PrimaryButton
                  disabled={ false }
                  text='Yes'
                  onClick= {this._redirectToHome}
                />
                <DefaultButton
                  disabled={ false }
                  text='No'
                  onClick= {this._redirectToHome}
              />
              </div>
                    </div>
                </div>
            </div>
        );        
    }
    private _redirectToHome() {
        window.location.href = "https://uphpcin.sharepoint.com";
    }
}