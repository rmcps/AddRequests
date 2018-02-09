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
import { IAccessRequestsProps } from '../IAccessRequestsProps';
import { IModifyAccessRequestsState } from './IModifyAccessRequestsState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import SharePointDataProvider from '../../services/SharePointDataProvider';
import MockSharePointDataProvider from '../../test/MockSharePointDataProvider';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';
import IModifyAccessRequest from '../../models/IModifyAccessRequest';
import DisplayRequest from '../DisplayAccessRequest/DisplayRequest';
import IDisplayRequestProps from '../DisplayAccessRequest/IDisplayRequestProps';

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
      originalCommittees:[],
      dropDownErrorMsg:'',
      enableSave:false
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
        this.setState({members: response});
      });
    }
  }
  public render(): React.ReactElement<IAccessRequestsProps> {
    return (
      <div className={styles.accessRequests}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={styles.headerBar}> <h2>Change Access Requests</h2></div>
              <div className={styles.subTitle}>Request for changes to member access</div>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <form>
                <div className={styles.formFieldsContainer}>
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
                      label='Remove user access'
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
                  {this.state.Item.RequestReason != 'Terminate' && <div className={styles.fieldContainer}>
                    <Dropdown
                      onChanged={this._onChangeMultiSelect}
                      placeHolder='Select committee(s)'
                      label='Add or Remove Commitees:'
                      selectedKeys={this.state.selectedCommittees}
                      errorMessage={this.state.dropDownErrorMsg}
                      multiSelect options={this.state.committees.map((item) => ({ key: item.Id, text: item.Title }))}
                    />
                  </div>}
                </div>
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
                    text='Cancel' onClick={this._cancelItem}
                  />
                </div>
                {this.renderErrors()}
              </form>
            </div>
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
  private _onToggleRemoveUser(checked:boolean) {
    this.setState((prevState: IModifyAccessRequestsState ,props:IAccessRequestsProps): IModifyAccessRequestsState => {
      prevState.Item.RequestReason = checked ? 'Terminate' : '';
      return prevState;
    });
  }
  @autobind
  private _onMemberChanged(option: IComboBoxOption, index: number, value: string) {
    this._dataProvider.getMemberCommittees(option.key).then(response => {      
    this.setState((prevState: IModifyAccessRequestsState ,props:IAccessRequestsProps): IModifyAccessRequestsState => {
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
      this.setState((prevState: IModifyAccessRequestsState ,props:IAccessRequestsProps): IModifyAccessRequestsState => {
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
  private _cancelItem(): void {
    window.location.href = "https://uphpcin.sharepoint.com";
  }
  @autobind
  private async _saveItem(): Promise<void> {    
    if(this.state.Item.RequestReason != "Terminate") {
      if (this.state.selectedCommittees.length < 1) {
        this.setState({dropDownErrorMsg: "Select one or more committee(s)"});
        return null;
      }
      let arrAdd = [];
      this.state.selectedCommittees.forEach((item => {
        if (this.state.originalCommittees.indexOf(item) == -1) {
          arrAdd.push(item);
        }
      }));
      this.setState((prevState: IModifyAccessRequestsState ,props:IAccessRequestsProps): IModifyAccessRequestsState => {
        prevState.Item.AddCommittees = arrAdd;
        return prevState;
      });              
      let arrRemove = [];
      this.state.originalCommittees.forEach((item => {
        if (this.state.selectedCommittees.indexOf(item) == -1) {          
          arrRemove.push(item);
        }
      }));
      this.setState((prevState: IModifyAccessRequestsState ,props:IAccessRequestsProps): IModifyAccessRequestsState => {
        prevState.Item.RemoveCommittees = arrRemove;
        return prevState;
      });              
    }
    this.setState( {
      status: 'Saving record...',
    });
  
    this._dataProvider.saveModifyRequest(this.state.Item).then((result) => {
      if(result.ok) {
        let additionalInfo: string;
        if (this.state.Item.RequestReason != 'Terminate') {
          if (this.state.Item.AddCommittees) {
            let comm = this.state.committees.filter((item) => 
                    this.state.Item.AddCommittees.indexOf(item.ID) !== -1);
            additionalInfo = `Add committees: ${comm.map(c => c.Title).join(",")}\r\n`;
          }
          if (this.state.Item.RemoveCommittees) {
            let comm = this.state.committees.filter((item) => 
                    this.state.Item.AddCommittees.indexOf(item.ID) !== -1);            
            additionalInfo = `${additionalInfo}Remove committees: ${comm.map(c => c.Title).join(",")}\r\n`;
          }
        }
        const element: React.ReactElement<IDisplayRequestProps > = React.createElement(
          DisplayRequest, {
            description: this.props.description,
            context:this.props.context,
            dom: this.props.dom,      
            recordType: "Change",
            RequestReason: this.state.Item.RequestReason,
            Title: this.state.Item.Title,
            EMail: this.state.Item.EMail,
            Comments: this.state.Item.Comments,
            additionalInfo: additionalInfo,
            }
        );      
        ReactDom.unmountComponentAtNode(this.props.dom);          
        ReactDom.render(element, this.props.dom);
      }
      else {
        this.setState((prevState: IModifyAccessRequestsState ,props:IAccessRequestsProps): IModifyAccessRequestsState => {
          prevState.errors.push('Error: Failed to save record.');
          prevState.status = '';
          return prevState;
        });          
      }

    });
    
  }
}
