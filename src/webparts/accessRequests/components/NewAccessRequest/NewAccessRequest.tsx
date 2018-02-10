import * as React from 'react';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Dialog, DialogType, DialogFooter} from 'office-ui-fabric-react/lib/Dialog';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import styles from '../AccessRequests.module.scss';
import { INewAccessRequestsState } from './INewAccessRequestsState';
import { escape } from '@microsoft/sp-lodash-subset';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import INewAccessRequest from '../../models/INewAccessRequest';

export interface NewAccessRequestProps {
  dataProvider: IAccessRequestsDataProvider;
  onRecordAdded: any;
}

export default class NewAccessRequest extends React.Component<NewAccessRequestProps, INewAccessRequestsState> {
  // private _dataProvider: IAccessRequestsDataProvider;

  constructor(props: NewAccessRequestProps, state: INewAccessRequestsState) {
    super(props);
    // set initial state
    this.state = {
      status: '', 
      isLoadingData: false,
      isSaving: false,
      newItem: {},
      errors: [],
      committees: [],
      selectedCommittees: [],
      dropDownErrorMsg: 'Select one or more committee(s)',
      hideDialog: true,
      enableSave: true
    };

  }
  public componentWillMount() {
    // FOR TESTING ONLY.  Remove after:
    let access: INewAccessRequest = {
      FirstName: "Sheila",
      LastName: "Allen",
      EMail: 'sheila.allen@rmcps.com',
      Company: "RMC"
    };
    this.setState({
      newItem: access
    });
  }
  public componentWillReceiveProps(nextProps: NewAccessRequestProps): void {
    // this.setState({
    //   status: this.listNotConfigured(nextProps) ? 'Please configure list in Web Part properties' : '',
    // });
  }
  public componentDidMount() {
    if (this.state.committees.length < 1) {
      this.props.dataProvider.getCommittees().then(response => {
        this.setState({
          committees: response.value,
        });
      });
    }
  }
  public render(): React.ReactElement<NewAccessRequestProps> {
    return (
      <div className={styles.row}>
        <div className={styles.column}>
          <form>
            <div className={styles.formFieldsContainer}>
              <div className={styles.fieldContainer}>
                <TextField placeholder='First Name' name='FirstName' required={true} value={this.state.newItem.FirstName}
                  onChanged={this._onFirstNameChanged} onGetErrorMessage={this._getErrorMessage}
                  validateOnFocusIn validateOnFocusOut underlined
                />
              </div>
              <div className={styles.fieldContainer}>
                <TextField placeholder='Last Name' name='LastName' required={true} value={this.state.newItem.LastName}
                  onChanged={this._onLastNameChanged} onGetErrorMessage={this._getErrorMessage}
                  validateOnFocusIn validateOnFocusOut underlined
                />
              </div>
              <div className={styles.fieldContainer}>
                <TextField placeholder='Email' name='EMail' required={true} value={this.state.newItem.EMail}
                  onChanged={this._onEmailChanged} onGetErrorMessage={this._getErrorMessage}
                  validateOnFocusIn validateOnFocusOut underlined
                />
              </div>
              <div className={styles.fieldContainer}>
                <TextField placeholder='Job Title' name='JobTitle'
                  onChanged={this._onJobTitleChanged} underlined
                />
              </div>
              <div className={styles.fieldContainer}>
                <TextField placeholder='Company' name='Company' required={true} value={this.state.newItem.Company}
                  onChanged={this._onCompanyChanged} onGetErrorMessage={this._getErrorMessage}
                  validateOnFocusIn validateOnFocusOut underlined
                />
              </div>
              <div className={styles.fieldContainer}>
                <TextField placeholder='Phone Number' name='Office'
                  onChanged={this._onOfficeChanged} underlined
                />
              </div>
              <div className={styles.fieldContainer}>
                <TextField name='Comments' multiline rows={2} placeholder='Enter any special instructions'
                  onChanged={this._onCommentsChanged}
                />
              </div>
              <div className={styles.fieldContainer}>
                <Dropdown
                  onChanged={this._onChangeMultiSelect}
                  placeHolder='Select committee(s)'
                  selectedKeys={this.state.selectedCommittees}
                  errorMessage={this.state.dropDownErrorMsg}
                  multiSelect options={this.state.committees.map((item) => ({ key: item.ID, text: item.Title }))}
                />
              </div>
            </div>

            {this.state.isSaving ? <Spinner size={SpinnerSize.small} /> : null}
            <div className={styles.formButtonsContainer}>
              <PrimaryButton
                disabled={
                  !this.state.enableSave || !this.state.newItem || this.state.isSaving
                }
                text='Save' onClick={this._saveItem}
              />
              <DefaultButton disabled={false} text='Cancel' onClick={this._cancelItem} />
            </div>
            {this.renderErrors()}
          </form>
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
          <PrimaryButton onClick={ this._closeDialog } text='OK' />
          </DialogFooter>
          
          </Dialog>
        </div>
      </div>
    );
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
    //if(!value == '' || value.length < 1) {}

    //this.setState({newItem: {...this.state.newItem, FirstName: value}});
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
    let updatedSelectedItems = this.state.selectedCommittees.length > 0 ? this.copyArray(this.state.selectedCommittees) : [];
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
      prevState.selectedCommittees = updatedSelectedItems;
      prevState.dropDownErrorMsg = updatedSelectedItems.length > 0 ? '' : 'Select one or more committee(s)';
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
  private _cancelItem(): void {
    window.location.href = "https://uphpcin.sharepoint.com";
  }
  @autobind
  private _saveItem(): Promise<void> {
    this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
      prevState.errors.length = 0;
      return prevState;
    });
    if (
      this.state.newItem.FirstName == '' ||
      this.state.newItem.LastName == '' ||
      this.state.newItem.EMail == '' ||
      this.state.newItem.Company == '' ||
      (!this.state.newItem.Committees || this.state.newItem.Committees.length < 1)
    ) {
      //this.setState({errors: [...this.state.errors, 'Some required fields are missing.'],});
      this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
        prevState.errors.push('Some required fields are missing.');
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
    this.setState({
      status: 'Saving record...',
      isSaving: true,
    });
    this.props.dataProvider.saveNewItem(this.state.newItem).then((result) => {
      if (result.ok) {
        this.setState({ 
          hideDialog: false,
          isSaving: false
        });        
      }
      else {
        this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
          prevState.errors.push('Error: Failed to save record.');
          prevState.status = '';
          prevState.isSaving = false;
          return prevState;
        });
      }
    });

  }
  @autobind
  private _closeDialog() {
    this.setState((prevState: INewAccessRequestsState, props: NewAccessRequestProps): INewAccessRequestsState => {
      prevState.hideDialog = true;
      return prevState;
    });
    this.props.onRecordAdded("list");
  }  
}
