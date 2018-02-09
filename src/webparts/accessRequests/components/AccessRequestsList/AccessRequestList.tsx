import * as React from 'react';
import styles from './AccessRequestList.module.scss';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { List } from 'office-ui-fabric-react/lib/List';
import IAccessRequest from '../../models/IAccessRequest';

export interface IAccessRequestProps {
  items: IAccessRequest[];
  onItemSelected: any;
}

export default class AccessRequestList extends React.Component<IAccessRequestProps, {}> {
  constructor(props: IAccessRequestProps) {
    super(props);
  }
  public render() {
    return (
      <div className={styles.row}>
        <div className={styles.column}>
          <h3>My Requests</h3>

          <FocusZone direction={FocusZoneDirection.vertical}>
            <List
              className={styles.accessRequestsList}
              items={this.props.items}
              onRenderCell={this._onRenderCell}
            />
          </FocusZone>
        </div>
      </div>
    );
  }

  @autobind
  private _onRenderCell(item: IAccessRequest, index: number | undefined): JSX.Element {
    return (
      <div className={styles.itemCell} data-is-focusable={true}>
        <div className={styles.itemContent}>
          <div className={styles.itemName}><span className={styles.itemLabel}>Requested For:</span> {item.Title}</div>
          <div><span className={styles.itemLabel}>Email:</span> {item.EMail}</div>
          <div><span className={styles.itemLabel}>Reason for Request:</span> {item.RequestReason}</div>
          <div><span className={styles.itemLabel}>Status:</span> {item.RequestStatus}</div>
          <div><span className={styles.itemLabel}>Created By:</span> {item.CreatedBy}</div>
          <div><span className={styles.itemLabel}>Submitted:</span> {item.Created} </div>
          {item.AddCommittees.length > 0 && <div><span className={styles.itemLabel}>Add Committees:</span> {item.AddCommittees.join(", ")}</div>}
          {item.RemoveCommittees.length > 0 && <div><span className={styles.itemLabel}>Remove Committees:</span> {item.RemoveCommittees.join(",")}</div>}
        </div>
        <IconButton
          data-requestId={item.Id}
          className={styles.chevron}
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
    this.props.onItemSelected(requestId);
  }
}
