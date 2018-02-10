import * as React from 'react';
import * as ReactDom from 'react-dom';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from './AccessRequests.module.scss';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import SharePointDataProvider from '../services/SharePointDataProvider';
import MockSharePointDataProvider from '../test/MockSharePointDataProvider';
import IAccessRequestsDataProvider from '../models/IAccessRequestsDataProvider';
import NewAccessRequest from './NewAccessRequest/NewAccessRequest';
import ModifyAccessRequest from './ModifyAccessRequest/ModifyAccessRequest';
import { IAccessRequestsProps } from './IAccessRequestsProps';
import IDefaultProps from './IDefaultProps';
import AccessRequestList from './AccessRequestsList/AccessRequestList';
import IAccessRequest from '../models/IAccessRequest';
import DisplayRequest from './DisplayAccessRequest/DisplayRequest';
import IDisplayRequestProps from './DisplayAccessRequest/IDisplayRequestProps';
import TopNav from './Navigation/TopNav';

export interface IDefaultState {
  show: "List" | "New" | "Change" | "Display";
  selectedItem: IAccessRequest;
}

export default class DefaultPage extends React.Component<IDefaultProps, IDefaultState> {
  private _dataProvider: IAccessRequestsDataProvider;

  constructor(props: IDefaultProps, state: IDefaultState) {
    super(props);
    // set initial state   
    this.state = {
      show: "List",
      selectedItem: null,
    };

    if (DEBUG && Environment.type === EnvironmentType.Local) {
      this._dataProvider = new MockSharePointDataProvider();

    } else {
      this._dataProvider = new SharePointDataProvider();
      this._dataProvider.webPartContext = this.props.context;
      this._dataProvider.accessListTitle = "Site Access Requests";
    }

  }
  public componentWillReceiveProps(nextProps: IDefaultProps) {

    // this.setState({
    //   status: this.listNotConfigured(nextProps) ? 'Please configure list in Web Part properties' : '',
    // });
  }
  public componentDidMount() {

  }
  public render(): React.ReactElement<IDefaultProps> {
    return (
      <div className={styles.accessRequests}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={styles.headerBar}>
              <h2 className={styles.title}>Member Access Request Submission</h2>
              <TopNav onItemSelected={this.handleViewSelected} />
              </div>
            </div>
          </div>
          {this.state.show == "List" && <AccessRequestList dataProvider={this._dataProvider} onItemSelected={this.handleItemSelected} />}
          {this.state.show == "Display" && <DisplayRequest item={this.state.selectedItem} recordType="Display" />}
          {this.state.show == "New" && <NewAccessRequest dataProvider={this._dataProvider} onRecordAdded={this.handleViewSelected} />}
          {this.state.show == "Change" && <ModifyAccessRequest dataProvider={this._dataProvider} onRecordAdded={this.handleViewSelected}/>}
        </div>
      </div>
    );
  }

  @autobind
  private handleViewSelected(selectedItem) {
    switch (selectedItem) {
      case "addNew":
        this.setState({
          show: "New"
        });
        break;
      case "change":
        this.setState({
          show: "Change"
        });
        break;
      case "display":
        this.setState({
          show: "Display"
        });
        break;
      case "list":
        this.setState({
          show: "List"
        });
        break;
      case "cancel":
        this._onCancel();
        break;
    }
  }
  @autobind
  private _onCancel(): void {
    window.location.href = "https://uphpcin.sharepoint.com";
  }

  @autobind
  private handleItemSelected(item) {
    this.setState({
      selectedItem: item,
      show: "Display"
    });

  }
  // private listNotConfigured(props: IAccessRequestsProps): boolean {
  //   return props.listName === undefined ||
  //     props.listName === null ||
  //     props.listName.length === 0;
  // }


}
