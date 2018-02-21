import * as React from 'react';
import styles from '../AccessRequests.module.scss';
import aLstyles from './AccessRequestList.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { List } from 'office-ui-fabric-react/lib/List';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import IAccessRequest from '../../models/IAccessRequest';
import IAccessRequestsDataProvider from '../../models/IAccessRequestsDataProvider';

export interface IAccessRequestListProps {
  onItemSelected: any;
  dataProvider: IAccessRequestsDataProvider;
}

export interface IAccessRequestListState {
  listItems: IAccessRequest[];
  dataIsLoading: boolean;
}

export default class AccessRequestList extends React.Component<IAccessRequestListProps, IAccessRequestListState> {
  constructor(props: IAccessRequestListProps) {
    super(props);
    // set initial state   
    this.state = {
      listItems: [],
      dataIsLoading: true
    };
  }
  public componentWillReceiveProps(nextProps: IAccessRequestListProps) {
    this.props.dataProvider.getItemsForCurrentUser().then((items: IAccessRequest[]) => {
      this.setState({
        listItems: items,
        dataIsLoading: false
      });
    });
  }
  public componentDidMount() {
    this.props.dataProvider.getItemsForCurrentUser().then((items: IAccessRequest[]) => {
      this.setState({
        listItems: items,
        dataIsLoading: false
      });
    });
  }
  public render() {
    return (
      <div className={styles.row}>
        <div className={styles.column2}>
          <h3>My Requests</h3>
          {this.state.dataIsLoading ? <Spinner size={SpinnerSize.medium} /> : null}
          <Fabric>
            <FocusZone direction={FocusZoneDirection.vertical}>
              <List
                className={aLstyles.accessRequestsList}
                items={this.state.listItems}
                onRenderCell={this._onRenderCell}
              />
            </FocusZone>
          </Fabric>
        </div>
      </div>
    );
  }

  @autobind
  private _onRenderCell(item: IAccessRequest, index: number | undefined): JSX.Element {
    return (
      <div className={aLstyles.itemCell} data-is-focusable={true}>
        <div className={aLstyles.itemContent}>
          <div className={aLstyles.itemName}><span className={aLstyles.itemLabel}>Requested For:</span> {item.Title}</div>
          <div><span className={aLstyles.itemLabel}>Email:</span> {item.EMail}</div>
          <div><span className={aLstyles.itemLabel}>Reason for Request:</span> {item.RequestReason}</div>
          <div><span className={aLstyles.itemLabel}>Status:</span> {item.RequestStatus}</div>
          <div><span className={aLstyles.itemLabel}>Created By:</span> {item.CreatedBy}</div>
          <div><span className={aLstyles.itemLabel}>Submitted:</span> {item.Created} </div>
          {item.AddCommittees.length > 0 && <div><span className={aLstyles.itemLabel}>Add Committees:</span> {item.AddCommittees.join(", ")}</div>}
          {item.RemoveCommittees.length > 0 && <div><span className={aLstyles.itemLabel}>Remove Committees:</span> {item.RemoveCommittees.join(",")}</div>}
        </div>
        <IconButton
          data-requestId={item.Id}
          className={aLstyles.chevron}
          disabled={false}
          iconProps={{ iconName: 'ChevronRight' }}
          title='Show Item'
          ariaLabel='Show Item'
          onClick={this._onItemClick}
        />
      </div>
    );
  }

  @autobind
  private _onItemClick(event?: React.MouseEvent<HTMLButtonElement>) {
    const attributes: NamedNodeMap = event.currentTarget.attributes;
    const requestId = attributes.getNamedItem("data-requestId").value;
    const item = this.state.listItems.filter((i) => i.Id == requestId)[0];
    this.props.onItemSelected(item);
  }
}
