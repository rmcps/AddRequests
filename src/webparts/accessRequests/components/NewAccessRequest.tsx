import * as React from 'react';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import styles from './AccessRequests.module.scss';
import { IAccessRequestsProps } from './IAccessRequestsProps';
import { INewAccessRequestsState } from './INewAccessRequestsState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import SharePointDataProvider from '../services/SharePointDataProvider';
import MockSharePointDataProvider from '../test/MockSharePointDataProvider';
import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';
import CommitteesList from '../components/CommitteesList';
import INewAccessRequest from '../models/INewAccessRequest';

export default class NewAccessRequest extends React.Component<IAccessRequestsProps, INewAccessRequestsState> {
  private _dataProvider: IAccessRequestsDataProvider;

  constructor(props: IAccessRequestsProps, state: INewAccessRequestsState) {
    super(props);
    // set initial state
    this.state = {
      status: '', //this.listNotConfigured(this.props) ? 'Please configure list in Web Part properties' : '',
      isLoadingData: false,
      newItem:{},
      errors:[],
      committees: [],
      selectedCommittees:[],
      dropDownErrorMsg:'Select one or more committee(s)',
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
      this._dataProvider.accessListTitle = "New Access Requests";
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
          committees: response.value,
        }); 
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
                <TextField label='First Name' name='FirstName' required={ true }
                  onChanged={this._onFirstNameChanged} onGetErrorMessage={ this._getErrorMessage }
                  validateOnFocusIn validateOnFocusOut
                />
                <TextField label='Last Name' name='LastName' required={ true }
                  onChanged={this._onLastNameChanged} onGetErrorMessage={ this._getErrorMessage }
                  validateOnFocusIn validateOnFocusOut
                />
                <TextField label='Email' name='EMail' required={ true }
                  onChanged={this._onEmailChanged} onGetErrorMessage={ this._getErrorMessage }
                  validateOnFocusIn validateOnFocusOut
                />
                <TextField label='Job Title' name='JobTitle' 
                  onChanged={this._onJobTitleChanged}
                />
                <TextField label='Company' name='Company' required={ true }
                  onChanged={this._onCompanyChanged} onGetErrorMessage={ this._getErrorMessage }
                  validateOnFocusIn validateOnFocusOut
                />
                <TextField label='Phone Number' name='Office' 
                  onChanged={this._onOfficeChanged}
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
                    multiSelect options={this.state.committees.map((item) => ({key:item.ID, text:item.Title}) )}
                />                   
          </div>
              <div className={ styles.formButtonsContainer}>
                <PrimaryButton
                  disabled={ 
                    !this.state.enableSave || !this.state.newItem || this.state.status == 'Saving record...' 
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
    this.setState((prevState: INewAccessRequestsState ,props:IAccessRequestsProps): INewAccessRequestsState => {
        prevState.newItem[fieldName] = value;
      return prevState;
    });

  }
  @autobind
  private _onFirstNameChanged(value:string) {
    this.updateStateWithFieldValue('FirstName',value);
    //if(!value == '' || value.length < 1) {}

    //this.setState({newItem: {...this.state.newItem, FirstName: value}});
  }
  @autobind
  private _onLastNameChanged(value:string) {
    this.updateStateWithFieldValue('LastName',value);

  }
  @autobind
  private _onEmailChanged(value:string) {
    this.updateStateWithFieldValue('EMail',value);

  }
  @autobind
  private _onJobTitleChanged(value:string) {
    this.updateStateWithFieldValue('JobTitle',value);
  }
  @autobind
  private _onCompanyChanged(value:string) {
    this.updateStateWithFieldValue('Company',value);
  }
  @autobind
  private _onOfficeChanged(value:string) {
    this.updateStateWithFieldValue('Office',value);
  }
  @autobind
  private _onCommentsChanged(value:string) {
    this.updateStateWithFieldValue('Comments',value);

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
      this.setState((prevState: INewAccessRequestsState ,props:IAccessRequestsProps): INewAccessRequestsState => {
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
  private _validateEmail (value:string) {
    // regex from http://stackoverflow.com/questions/46155/validate-email-address-in-javascript
    let re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(value);
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
    this.setState((prevState: INewAccessRequestsState ,props:IAccessRequestsProps): INewAccessRequestsState => {
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
      this.setState((prevState: INewAccessRequestsState ,props:IAccessRequestsProps): INewAccessRequestsState => {
        prevState.errors.push('Some required fields are missing.');
        return prevState;
      });    
      return null;
    }
    if (!this._validateEmail(this.state.newItem.EMail)) {
      //this.setState({errors: [...this.state.errors, 'Email address is invalid.'],});
      this.setState((prevState: INewAccessRequestsState ,props:IAccessRequestsProps): INewAccessRequestsState => {
        prevState.errors.push('Email address is invalid.');
        return prevState;
      });
      return null;
    }
    this.setState( {
      status: 'Saving record...',
    });
  
    this._dataProvider.saveNewItem(this.state.newItem).then((result) => {
      if(result.ok) {
        this._showDialog();
        this.setState( {
          enableSave:false,
          status: 'New Access Request Created',
        });
      }
      else {
        //this.setState( {
        //  status: '',
        //});        
        this.setState((prevState: INewAccessRequestsState ,props:IAccessRequestsProps): INewAccessRequestsState => {
          prevState.errors.push('Error: Failed to save record.');
          prevState.status = '';
          return prevState;
        });          
      }
    });
    
  }
}
