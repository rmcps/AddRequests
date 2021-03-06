import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { ComboBox, IComboBoxProps, IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { Dropdown, IDropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import styles from '../AccessRequests.module.scss';
import { IModifyAccessRequestsState } from './IModifyAccessRequestsState';
import { escape } from '@microsoft/sp-lodash-subset';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import IModifyAccessRequest from '../../models/IModifyAccessRequest';

export interface IModifyAccessRequestProps {
  dataProvider: IAccessRequestsDataProvider;
  membersList: string;
  committeesListTitle: string;
  membersCommList: string;
  onRecordAdded: any;
}

export default class ModifyAccessRequest extends React.Component<IModifyAccessRequestProps, IModifyAccessRequestsState> {
  private _savingMessage:string = "Saving record...";

  constructor(props: IModifyAccessRequestProps, state: IModifyAccessRequestsState) {
    super(props);
    this.state = this.setCleanState(true);
  }
  public componentWillReceiveProps(nextProps: IModifyAccessRequestProps): void {
  }
  public async componentDidMount() {
    if (this.state.committees.length < 1) {
      try {
        const response = await this.props.dataProvider.getCommittees(this.props.committeesListTitle);
        this.setState({
          committees: response.value,
        });
      }
      catch (error) {
        console.log(error);
      }
      // this.props.dataProvider.getCommittees(this.props.committeesListTitle).then(response => {
      //   this.setState({
      //     committees: response.value
      //   });
      // });
    }
    if (this.state.members.length < 1) {
      try {
        let response = await this.props.dataProvider.getMembers(this.props.membersList);
        this.setState({ members: response });
      }
      catch (error) {
        console.log(error);
      }
      // this.props.dataProvider.getMembers(this.props.membersList).then(response => {
      //   this.setState({ members: response });
      // });
    }
  }
  public async componentDidUpdate() {
    if (this.state.status === this._savingMessage) {
      const response = await this.props.dataProvider.saveChangeRequest(this.state.Item);
        if (response === 'ok') {
          this.setState({
            hideDialog: false,
            status: ""
          });
  
        }
        else {
          this.setState((prevState: IModifyAccessRequestsState, props: IModifyAccessRequestProps): IModifyAccessRequestsState => {
            prevState.errors.push('Error: Failed to save record.');
            prevState.status = '';
            prevState.enableSave = true;
            return prevState;
          });
        }      
    }
  }
  public render(): React.ReactElement<IModifyAccessRequestProps> {
    return (
      <form>
        <div className={styles.row}>
          <div className={styles.column2}>
            <div className={styles.fieldContainer}>
              <ComboBox
                selectedKey={this.state.Item.spLoginName}
                className='MemberCombo'
                label='Select a Member:'
                id='MemberCombo'
                ariaLabel='Member List'
                allowFreeform={false}
                autoComplete='on'
                options={this.state.members.map((item) => ({ key: item.spLoginName, value: item.spLoginName, text: item.Title }))}
                onChanged={this._onMemberChanged}
              />
            </div>
            <div className={styles.fieldContainer}>
              <Toggle
                checked={this.state.Item.RequestReason == 'Terminate'}
                label="To completely REMOVE this user's access to the UPHPCIN site, select Yes"
                onText='Yes'
                offText='No'
                onChanged={this._onToggleRemoveUser}
              />
            </div>
            <div className={styles.fieldContainer}>
              <TextField label='Comments' name='Comments' multiline rows={2} placeholder='Enter any special instructions'
                onChanged={this._onCommentsChanged}
              />
            </div>
            {(this.state.Item.spLoginName && this.state.Item.RequestReason != 'Terminate') && <div className={styles.fieldContainer}>
              <Dropdown
                onChanged={this._onChangeMultiSelect}
                placeHolder='Select committee(s)'
                label="User's committees.  Check a name to add access, uncheck a name to remove access:"
                selectedKeys={this.state.selectedCommittees}
                errorMessage={this.state.dropDownErrorMsg}
                multiSelect options={this.state.committees.map((item) => ({ key: item.Id, text: item.Title }))}
              />
            </div>}
            <div className={styles.formButtonsContainer}>
              <PrimaryButton
                disabled={
                  !this.state.enableSave || !this.state.Item || this.state.status == 'Saving record...'
                }
                text='Save'
                onClick={this._saveItem}
              />
              <DefaultButton
                disabled={false}
                text='Reset' onClick={this._resetItem}
              />
            </div>
            {this.renderErrors()}
            <Dialog
              hidden={this.state.hideDialog}
              onDismiss={this._closeDialog}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Your request was saved.',
                subText: 'Your access request change was created.  You will receive email updates with the status of your request.'
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
        </div>
      </form>

    );
  }
  private setCleanState(initial: boolean): any {
    const objItem = {
      spLoginName: null,
      FirstName: null,
      LastName: null,
      EMail: null,
      Title: null,
      JobTitle: null,
      Office: null,
      Comments: null,
      AddCommittees: [],
      RemoveCommittees: [],
      RequestReason: null
    };

    let cleanState:any = {
      Item: objItem,
      errors: [],
      selectedCommittees: [],
      originalCommittees: [],
      dropDownErrorMsg: '',
      enableSave: false,
      hideDialog: true  
    };

    if (initial) {
      cleanState.status = '';
      cleanState.members = [];
      cleanState.committees = [];
      cleanState.isLoadingData = false;
      cleanState.isSaving = false;
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
    this.setState((prevState: IModifyAccessRequestsState, props: IModifyAccessRequestProps): IModifyAccessRequestsState => {
      prevState.Item[fieldName] = value;
      return prevState;
    });

  }
  @autobind
  private _onCommentsChanged(value: string) {
    this.updateStateWithFieldValue('Comments', value);

  }
  @autobind
  private _onToggleRemoveUser(checked: boolean) {
    this.setState((prevState: IModifyAccessRequestsState, props: IModifyAccessRequestProps): IModifyAccessRequestsState => {
      prevState.Item.RequestReason = checked ? 'Terminate' : '';
      return prevState;
    });
  }
  @autobind
  private async _onMemberChanged(option: IComboBoxOption, index: number, value: string) {
    const response = await this.props.dataProvider.getMemberCommittees(this.props.membersCommList, option.key);

    this.setState((prevState: IModifyAccessRequestsState, props: IModifyAccessRequestProps): IModifyAccessRequestsState => {
      prevState.Item.spLoginName = option.key;
      prevState.Item.Title = option.text;
      const selMember = this.state.members.filter((mem) => {
        return mem.spLoginName == option.key;
      });
      prevState.Item.EMail = selMember[0].EMail;
      prevState.originalCommittees = response.value.map(c => c.CommitteeId);
      prevState.selectedCommittees = response.value.map(c => c.CommitteeId);
      prevState.enableSave = option.key ? true : false;
      return prevState;
    });
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
    this.setState((prevState: IModifyAccessRequestsState, props: IModifyAccessRequestProps): IModifyAccessRequestsState => {
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
  @autobind
  private _resetItem(): void {
    this.setState((prevState: IModifyAccessRequestsState, props: IModifyAccessRequestProps): IModifyAccessRequestsState => {
      prevState = this.setCleanState(false);
      return prevState;
    });

  }
  @autobind
  private async _saveItem(): Promise<void> {
      let arrAdd = [];
      let arrRemove = [];
    if (this.state.Item.RequestReason != "Terminate") {
      if (this.state.selectedCommittees.length < 1) {
        this.setState({ dropDownErrorMsg: "Select one or more committee(s)" });
        return null;
      }
      this.state.selectedCommittees.forEach((item => {
        if (this.state.originalCommittees.indexOf(item) == -1) {
          arrAdd.push(item);
        }
      }));
      this.state.originalCommittees.forEach((item => {
        if (this.state.selectedCommittees.indexOf(item) == -1) {
          arrRemove.push(item);
        }
      }));
    }    
    this.setState((prevState: IModifyAccessRequestsState, props: IModifyAccessRequestProps): IModifyAccessRequestsState => {
      prevState.Item.AddCommittees = arrAdd;
      prevState.Item.RemoveCommittees = arrRemove;
      prevState.status = this._savingMessage;
      prevState.enableSave = false;
      return prevState;
    });
  }
  @autobind
  private _closeDialog() {
    this.setState((prevState: IModifyAccessRequestsState, props: IModifyAccessRequestProps): IModifyAccessRequestsState => {
      prevState.hideDialog = true;
      return prevState;
    });
    this.props.onRecordAdded("List");
  }

}
