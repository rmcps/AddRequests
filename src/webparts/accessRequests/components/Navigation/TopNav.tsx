import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import styles from '../AccessRequests.module.scss';

export interface TopNavProps {
    isApprover: boolean;
    onItemSelected: any;
    show: "List" | "New" | "Change" | "Display" | "Tasks" | "FinalTasks";
}

export default class TopNav extends React.Component<TopNavProps, {}> {
    constructor(props: TopNavProps) {
        super(props);
    }

    public render() {
        return (
            <div className={styles.pageNav}>
                {"New" !== this.props.show && <Link data-target-name="addNew" onClick={this._onItemSelected}>New member access</Link>}
                {"Change" !== this.props.show && <Link data-target-name="change" onClick={this._onItemSelected}>Modify existing member</Link>}
                {"List" !== this.props.show && <Link data-target-name="list" onClick={this._onItemSelected}>My Requests</Link>}
                {"Tasks" !== this.props.show && <Link data-target-name="tasks" onClick={this._onItemSelected}>My Tasks</Link>}
                {"FinalTasks" !== this.props.show && this.props.isApprover && <Link data-target-name="finaltasks" onClick={this._onItemSelected}>Final Approvals</Link>}
            </div>
        );
    }
    @autobind
    private _onItemSelected(event: React.MouseEvent<HTMLAnchorElement>): void {
        const attributes: NamedNodeMap = event.currentTarget.attributes;
        const selItem = attributes.getNamedItem("data-target-name").value;
        this.props.onItemSelected(selItem);
    }
}