import * as React from 'react';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import {
  ComboBox,
  IComboBoxProps,
  IComboBoxOption
} from 'office-ui-fabric-react/lib/ComboBox';
import { Dropdown, IDropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import styles from '../AccessRequests.module.scss';
import { IAccessRequestsProps } from '../IAccessRequestsProps';
import { IModifyAccessRequestsState } from './IModifyAccessRequestsState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import SharePointDataProvider from '../../services/SharePointDataProvider';
import MockSharePointDataProvider from '../../test/MockSharePointDataProvider';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import IModifyAccessRequest from '../../models/IModifyAccessRequest';

export default class ModifyAccessRequest extends React.Component<IAccessRequestsProps, IModifyAccessRequestsState> {
  private _dataProvider: IAccessRequestsDataProvider;

  constructor(props: IAccessRequestsProps, state: IModifyAccessRequestsState) {
    super(props);
    // set initial state
    this.state = {
      status: '', //this.listNotConfigured(this.props) ? 'Please configure list in Web Part properties' : '',
      isLoadingData: false,
      Item:{},
      errors:[],
      members:[],
      committees: [],
      selectedCommittees:[],
      dropDownErrorMsg:'',
      hideDialog:true,
      enableSave:true
    };
    /*
    Create the appropriate data provider depending on where the web part is running.
    The DEBUG flag will ensure the mock data provider is not bundled with the web part when you package the solution for distribution, that is, using the --ship flag with the package-solution gulp command.
    */
    if (DEBUG && Environment.type === EnvironmentType.Local) {
      this._dataProvider = new MockSharePointDataProvider();
      
    } else {
      this._dataProvider = new SharePointDataProvider();
      this._dataProvider.webPartContext = this.props.context;
      this._dataProvider.accessListTitle = "Site Access Requests";
    }
    
  }
  public componentWillReceiveProps(nextProps: IAccessRequestsProps): void {
    // this.setState({
    //   status: this.listNotConfigured(nextProps) ? 'Please configure list in Web Part properties' : '',
    // });
  }
  public componentDidMount() {
    if (this.state.committees.length < 1) {
      this._dataProvider.getCommittees().then(response => {
        this.setState({
          committees: response.value
        }); 
      });
    }
    if (this.state.members.length < 1) {
      this._dataProvider.getMembers().then(response => {
        this.setState({members: response.value});
      });
    }
  }
  public render(): React.ReactElement<IAccessRequestsProps> {
    return (
      <div className={ styles.accessRequests }>
        <div className={ styles.container }>
        <div className= {styles.row}>
        {this.renderErrors()}
        <MessageBar messageBarType={MessageBarType.warning}>{this.state.status}</MessageBar>
          <span className={ styles.title }>Access Requests</span>
        </div>
          <div className={ styles.row }>
          <form>
            <div className={ styles.column }>      
              <div className={ styles.formFieldsContainer}>
                <ComboBox
                  selectedKey= {this.state.Item.spLoginName}
                  className='MemberCombo'
                  label='Member:'
                  id='MemberCombo'
                  ariaLabel='Member List'
                  allowFreeform={ false }
                  autoComplete='on'
                  options={this.state.members.map((item) => ({key:item.spLoginName, value:item.spLoginName, text:item.Title}) )}
                  onChanged={ this._onMemberChanged }
                />
                <TextField label='Comments' name='Comments' multiline rows={2} placeholder='Enter any special instructions'
                  onChanged={this._onCommentsChanged}
                />
                <Dropdown
                    onChanged={ this._onChangeMultiSelect }
                    placeHolder='Select committee(s)'
                    label='Commitees:'
                    selectedKeys={ this.state.selectedCommittees }
                    errorMessage={this.state.dropDownErrorMsg }
                    multiSelect options={this.state.committees.map((item) => ({key:item.Title, text:item.Title}) )}
                />                   
          </div>
              <div className={ styles.formButtonsContainer}>
                <PrimaryButton
                  disabled={ 
                    !this.state.enableSave || !this.state.Item || this.state.status == 'Saving record...' 
                }
                  text='Save'
                  onClick= {this._saveItem}
                />
                <DefaultButton
                  disabled={ false }
                  text='Reset'
              />
              </div>

            </div>
            </form>
          </div>
          <div className={ styles.row }>            
          <Dialog 
            hidden={ this.state.hideDialog }
            onDismiss={ this._closeDialog }
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Request created',
              subText: "Your new access request was created.  You will receive email updates with the status of your request."
            }}
            modalProps={{
              isBlocking: true,
              containerClassName: 'ms-dialogMainOverride'
            }}
          >
          <DialogFooter>
            <PrimaryButton onClick={ this._closeDialog } text='OK' />
          </DialogFooter>
          </Dialog>
          </div>                    
        </div>
      </div>
    );
  }

  // private listNotConfigured(props: IAccessRequestsProps): boolean {
  //   return props.listName === undefined ||
  //     props.listName === null ||
  //     props.listName.length === 0;
  // }
  private renderErrors() {
    return this.state.errors.length > 0
      ?
      <div>
        {
          this.state.errors.map( (item, idx) =>
           <MessageBar
             messageBarType={ MessageBarType.error }
             isMultiline={ true }
           >
             {item}
           </MessageBar>
          )
        }
      </div>
      : null;
   }
   private clearError(idx: number) {
    this.setState( (prevState, props) => {
      return {...prevState, errors: prevState.errors.splice( idx, 1 )};
    } );
  }
  private updateStateWithFieldValue(fieldName:string, value:string){
    this.setState((prevState: IModifyAccessRequestsState ,props:IAccessRequestsProps): IModifyAccessRequestsState => {
        prevState.Item[fieldName] = value;
      return prevState;
    });

  }
  @autobind
  private _onCommentsChanged(value:string) {
    this.updateStateWithFieldValue('Comments',value);

  }
  @autobind
  private _onMemberChanged(option: IComboBoxOption, index: number, value: string) {
    this._dataProvider.getMemberCommittees(option.key).then(response => {      
      // let sel = [];
      // response.value.forEach(function(valObj, idx) {
      //     valObj.Committee.forEach(function(comm) {
      //       sel.push(comm);
      //     });
      // });
      // debugger;
      
      this.setState((prevState: IModifyAccessRequestsState ,props:IAccessRequestsProps): IModifyAccessRequestsState => {
        prevState.Item.spLoginName = option.key;
        prevState.selectedCommittees = response.value.map(c => c.Committee);
        return prevState;
      });
    });
   
   }
  
  @autobind
  private _onChangeMultiSelect(item: IDropdownOption) {
    debugger
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
      this.setState((prevState: IModifyAccessRequestsState ,props:IAccessRequestsProps): IModifyAccessRequestsState => {
        prevState.Item.Committees = updatedSelectedItems;
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
  private _showDialog() {
    this.setState({ hideDialog: false });
  }

  @autobind
  private _closeDialog() {
    this.setState({ hideDialog: true });
    window.location.href = "https://uphpcin.sharepoint.com";
  }
  @autobind
  private async _saveItem(): Promise<void> {
    
    if(this.state.selectedCommittees.length < 1) {
      this.setState({dropDownErrorMsg: "Select one or more committee(s)"});
      return null;
    }
    this.setState( {
      status: 'Saving record...',
    });
  
    this._dataProvider.saveNewItem(this.state.Item).then((result) => {

    });
    
  }
}
