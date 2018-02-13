import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import styles from '../AccessRequests.module.scss';

export interface TopNavProps {
    onItemSelected: any;
}

export default class TopNav extends React.Component<TopNavProps, {}> {
    constructor(props: TopNavProps) {
        super(props);
    }

    public render() {
        return (
            <div className={styles.pageNav}>
                <Link data-target-name="addNew" onClick={this._onItemSelected}>New member access</Link>
                <Link data-target-name="change" onClick={this._onItemSelected}>Modify existing member</Link>
                <Link data-target-name="list" onClick={this._onItemSelected}>My Requests</Link>
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