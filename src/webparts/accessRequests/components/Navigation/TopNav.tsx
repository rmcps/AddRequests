import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import styles from '../AccessRequests.module.scss';
import {IDisplayView} from '../../utilities/types';

export interface TopNavProps {
    isApprover: boolean;
    onItemSelected: any;
    show: IDisplayView;
}

export default class TopNav extends React.Component<TopNavProps, {}> {
    constructor(props: TopNavProps) {
        super(props);
    }

    public render() {
        return (
            <div className={styles.pageNav}>
                <Link data-target-name="New" onClick={this._onItemSelected} className={"New" == this.props.show ? styles.btnSelected : null}>New member</Link>
                <Link data-target-name="Change" onClick={this._onItemSelected } className={"Change" == this.props.show ? styles.btnSelected : null}>Modify member</Link>
                <Link data-target-name="List" onClick={this._onItemSelected} className={"List" == this.props.show ? styles.btnSelected : null}>My Requests</Link>
                <Link data-target-name="Tasks" onClick={this._onItemSelected} className={"Tasks" == this.props.show ? styles.btnSelected : null}>My Tasks</Link>
                {this.props.isApprover && <Link data-target-name="FinalTasks" onClick={this._onItemSelected} className={"FinalTasks" == this.props.show ? styles.btnSelected : null}>Final Approvals</Link>}
            </div>
        );
    }
    @autobind
    private _onItemSelected(event: React.MouseEvent<HTMLAnchorElement>): void {
        const attributes: NamedNodeMap = event.currentTarget.attributes;
        const selItem:IDisplayView = attributes.getNamedItem("data-target-name").value as IDisplayView;
        this.props.onItemSelected(selItem);
    }
}