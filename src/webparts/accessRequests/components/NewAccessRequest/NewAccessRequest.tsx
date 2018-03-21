import * as React from 'react';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import styles from '../AccessRequests.module.scss';
import { INewAccessRequestsState } from './INewAccessRequestsState';
import { escape } from '@microsoft/sp-lodash-subset';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import INewAccessRequest from '../../models/INewAccessRequest';

export interface NewAccessRequestProps {
  dataProvider: IAccessRequestsDataProvider;
  committeesListTitle: string;
  onRecordAdded: any;
}

export default class NewAccessRequest extends React.Component<NewAccessRequestProps, INewAccessRequestsState> {
  private _savingMessage:string = "Saving record...";

  constructor(props: NewAccessRequestProps, state: INewAccessRequestsState) {
    super(props);
    // set initial state
    this.state = this.setCleanState(true);
  }
  public componentWillMount() {
    // FOR TESTING ONLY.  Remove after:
    // let access: INewAccessRequest = {
    //   FirstName: "Sheila",
    //   LastName: "Allen",
    //   EMail: 'sheila.allen@rmcps.com',
    //   Company: "RMC"
    // };
    // this.setState({
    //   newItem: access
    // });
  }
  public componentWillReceiveProps(nextProps: NewAccessRequestProps): void {
  }
  public async componentDidMount() {
    if (this.state.committees.length < 1) {
      try {
        const response = await this.props.dataProvider.getCommittees(this.props.committeesListTitle);
        this.setState({
          committees: response.value,
        });
      }
      catch(error) {
        console.log(error);
      }
    }
  }
  public async componentDidUpdate() {
    if (this.state.status === this._savingMessage) {
      try {
        const response = await this.props.dataProvider.saveNewItem(this.state.newItem);
        if (response == 'ok') {
          this.setState({
            hideDialog: false,
            isSaving: false,
            status: '',
          });
        }
        else {
          throw new Error('Error: Failed to save record.');
        }
      }
      catch (error) {
        console.log(error);
        this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
          prevState.errors.push('Error: Failed to save record.');
          prevState.status = '';
          prevState.isSaving = false;
          return prevState;
        });
      }
    }
  }
  public render(): React.ReactElement<NewAccessRequestProps> {
    return (
      <form>
        <div className={styles.row}>
          <div className={styles.column1}>
            <div className={styles.fieldContainer}>
              <TextField placeholder='First Name' name='FirstName' required={true} value={this.state.newItem.FirstName}
                onChanged={this._onFirstNameChanged}
                validateOnFocusIn validateOnFocusOut underlined
              />
            </div>
            <div className={styles.fieldContainer}>
              <TextField placeholder='Last Name' name='LastName' required={true} value={this.state.newItem.LastName}
                onChanged={this._onLastNameChanged}
                validateOnFocusIn validateOnFocusOut underlined
              />
            </div>
            <div className={styles.fieldContainer}>
              <TextField placeholder='Email' name='EMail' required={true} value={this.state.newItem.EMail}
                onChanged={this._onEmailChanged}
                validateOnFocusIn validateOnFocusOut underlined
              />
            </div>
          </div>
          <div className={styles.column1}>
            <div className={styles.fieldContainer}>
              <TextField placeholder='Company' name='Company' required={true} value={this.state.newItem.Company}
                onChanged={this._onCompanyChanged}
                validateOnFocusIn validateOnFocusOut underlined
              />
            </div>
            <div className={styles.fieldContainer}>
              <TextField placeholder='Job Title' name='JobTitle'
                onChanged={this._onJobTitleChanged} underlined
              />
            </div>
            <div className={styles.fieldContainer}>
              <TextField placeholder='Phone Number' name='Office'
                onChanged={this._onOfficeChanged} underlined
              />
            </div>
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column2}>
            <div className={styles.fieldContainer}>
              {this.state.CommitteeAccess && <Dropdown
                onChanged={this._onChangeMultiSelect}
                placeHolder='Select committee(s)'
                selectedKeys={this.state.newItem.Committees}
                errorMessage={this.state.dropDownErrorMsg}
                multiSelect options={this.state.committees.map((item) => ({ key: item.ID, text: item.Title }))}
              />}
            </div>
            <div className={styles.fieldContainer}>
              <Toggle
                checked={this.state.CommitteeAccess}
                label='Committee Access Requested? Select No for top level site access only (no committees).'
                onText='Yes'
                offText='No'
                onChanged={this._onToggleNoCommittees}
              />
            </div>
            <div className={styles.fieldContainer}>
              <TextField name='Comments' multiline rows={2} placeholder='Enter any special instructions'
                onChanged={this._onCommentsChanged}
              />
            </div>
            {this.state.isSaving ? <Spinner size={SpinnerSize.small} /> : null}
            <div className={styles.formButtonsContainer}>
              <PrimaryButton
                disabled={
                  !this.state.enableSave || !this.state.newItem || this.state.isSaving
                }
                text='Save' onClick={this._saveItem}
              />
              <DefaultButton disabled={false} text='Reset' onClick={this._resetItem} />
            </div>
            {this.renderErrors()}
            <Dialog
              hidden={this.state.hideDialog}
              onDismiss={this._closeDialog}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Your request was saved.',
                subText: 'Your new access request was created.  You will receive email updates with the status of your request.'
              }}
              modalProps={{
                isBlocking: true,
                containerClassName: 'accessRequests'
              }}
            >
              <DialogFooter>
                <PrimaryButton onClick={this._closeDialog} text='OK' />
              </DialogFooter>

            </Dialog>
          </div>
        </div >
      </form >
    );
  }
  private setCleanState(initial: boolean): any {
    const newItem = {
      FirstName: null,
      LastName: null,
      EMail: null,
      Comments: null,
      JobTitle: null,
      Company: null,
      Office: null,
      Committees: []
    };

    let cleanState: any = {
      newItem: newItem,
      dropDownErrorMsg: '',
      CommitteeAccess: true,
      errors: [],
    };

    if (initial) {
      cleanState.status = '';
      cleanState.isLoadingData = false;
      cleanState.isSaving = false;
      cleanState.committees = [];
      cleanState.selectedCommittees = [];
      cleanState.hideDialog = true;
      cleanState.enableSave = true;
    }
    return cleanState;
  }
  private renderErrors() {
    return this.state.errors.length > 0
      ?
      <div>
        {
          this.state.errors.map((item, idx) =>
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}
            >
              {item}
            </MessageBar>
          )
        }
      </div>
      : null;
  }
  private clearError(idx: number) {
    this.setState((prevState, props) => {
      return { ...prevState, errors: prevState.errors.splice(idx, 1) };
    });
  }
  private updateStateWithFieldValue(fieldName: string, value: string) {
    this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
      prevState.newItem[fieldName] = value;
      return prevState;
    });

  }
  @autobind
  private _onFirstNameChanged(value: string) {
    this.updateStateWithFieldValue('FirstName', value);
  }
  @autobind
  private _onLastNameChanged(value: string) {
    this.updateStateWithFieldValue('LastName', value);

  }
  @autobind
  private _onEmailChanged(value: string) {
    this.updateStateWithFieldValue('EMail', value);

  }
  @autobind
  private _onJobTitleChanged(value: string) {
    this.updateStateWithFieldValue('JobTitle', value);
  }
  @autobind
  private _onCompanyChanged(value: string) {
    this.updateStateWithFieldValue('Company', value);
  }
  @autobind
  private _onOfficeChanged(value: string) {
    this.updateStateWithFieldValue('Office', value);
  }
  @autobind
  private _onCommentsChanged(value: string) {
    this.updateStateWithFieldValue('Comments', value);

  }
  @autobind
  private _onChangeMultiSelect(item: IDropdownOption) {
    let updatedSelectedItems = (this.state.newItem.Committees && this.state.newItem.Committees.length > 0) ? this.copyArray(this.state.newItem.Committees) : [];
    if (item.selected) {
      // add the option if it's checked
      updatedSelectedItems.push(item.key);
    } else {
      // remove the option if it's unchecked
      let currIndex = updatedSelectedItems.indexOf(item.key);
      if (currIndex > -1) {
        updatedSelectedItems.splice(currIndex, 1);
      }
    }
    this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
      prevState.newItem.Committees = updatedSelectedItems;
      prevState.dropDownErrorMsg = (this.state.CommitteeAccess || updatedSelectedItems.length > 0) ? '' : 'Select one or more committee(s)';
      return prevState;
    });
  }
  @autobind
  private _onToggleNoCommittees(checked: boolean) {
    this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
      prevState.CommitteeAccess = checked;
      prevState.newItem.Committees = checked ? prevState.newItem.Committees : [];
      return prevState;
    });
  }
  @autobind
  private _getErrorMessage(value: string): string {
    return (value.length > 0 || value != "")
      ? ''
      : `A value is required.`;
  }
  public copyArray(array: any[]): any[] {
    let newArray: any[] = [];
    for (let i = 0; i < array.length; i++) {
      newArray[i] = array[i];
    }
    return newArray;
  }
  private _validateEmail(value: string) {
    // regex from http://stackoverflow.com/questions/46155/validate-email-address-in-javascript
    let re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(value);
  }
  @autobind
  private _resetItem(): void {
    this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
      prevState = this.setCleanState(false);
      return prevState;
    });
  }
  @autobind
  private _saveItem(): Promise<void> {
    this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
      prevState.errors.length = 0;
      return prevState;
    });
    let errMsg = '';
    if (this.state.newItem.FirstName === null || this.state.newItem.FirstName.length === 0) { errMsg = 'First name is required.'; }
    else if (this.state.newItem.LastName === null || this.state.newItem.LastName.length === 0) { errMsg = 'Last name is required.'; }
    else if (this.state.newItem.EMail === null || this.state.newItem.EMail.length === 0) { errMsg = 'Email is required.'; }
    else if (this.state.newItem.Company === null || this.state.newItem.Company.length === 0) { errMsg = 'Company is required.'; }

    if (errMsg.length > 0) {
      this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
        prevState.errors.push(errMsg);
        return prevState;
      });
      return null;
    }
    if (!this._validateEmail(this.state.newItem.EMail)) {
      this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
        prevState.errors.push('Email address is invalid.');
        return prevState;
      });
      return null;
    }
    if (this.state.CommitteeAccess && (!this.state.newItem.Committees || this.state.newItem.Committees.length < 1)) {
      this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
        prevState.dropDownErrorMsg = 'Select one or more committee(s)';
        return prevState;
      });
      return null;
    }
    this.setState({
      status: this._savingMessage,
      isSaving: true,
    });
  }
  @autobind
  private _closeDialog() {
    this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
      prevState.hideDialog = true;
      return prevState;
    });
    this.props.onRecordAdded("List");
  }
}
