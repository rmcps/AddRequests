import * as React from 'react';
import { Dropdown, IDropdown,IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import styles from './AccessRequests.module.scss';
import { autobind, BaseComponent } from 'office-ui-fabric-react/lib/Utilities';

class CommitteesList extends React.Component<any, any> {
    public render() {
        let selectedCommittees = this.state.selectedCommittees;
        let committees = this.props.committees;
        return (
            <div className='CommitteesListControl'>
                <Dropdown
                    onChanged={ () => this.props.onChanged(this) }
                    placeHolder='Select committee(s)'
                    label='Commitees:'
                    selectedKeys={ selectedCommittees }
                    multiSelect options={committees.map((item) => ({key:item.Id, text:item.Title}) )}
                />                   
            </div>
        );
    }
}
export default CommitteesList;