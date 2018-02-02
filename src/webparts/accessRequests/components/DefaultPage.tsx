import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Link } from 'office-ui-fabric-react/lib/Link';
import styles from './AccessRequests.module.scss';
import NewAccessRequest from './NewAccessRequest';
import ModifyAccessRequest from './ModifyAccessRequest';
import { IAccessRequestsProps } from './IAccessRequestsProps';
import IDefaultProps from './IDefaultProps';

export default class DefaultPage extends React.Component<IDefaultProps, null> {

  constructor(props) {
    super(props);
    // set initial state    
    this._onAddNew = this._onAddNew.bind(this);
    this._onAddExisting = this._onAddExisting.bind(this);
  }
  public componentWillReceiveProps(nextProps: IDefaultProps): void {
    // this.setState({
    //   status: this.listNotConfigured(nextProps) ? 'Please configure list in Web Part properties' : '',
    // });
  }
  public componentDidMount() {

  }
  public render(): React.ReactElement<IDefaultProps> {
    return (
      <div className={ styles.accessRequests }>
        <div className={ styles.container }>
        <div className= {styles.row}>
        <h2>Member Access Request Submission</h2>
        <Link onClick={this._onAddNew}>Add a new member access request</Link> &nbsp;&nbsp;&nbsp;
        <Link onClick={this._onAddExisting}>Add a requet to modify an existing member</Link>
          </div>                    
        </div>
      </div>
    );
  }

  private _onAddNew() {
    const element: React.ReactElement<IAccessRequestsProps > = React.createElement(
        NewAccessRequest,
        {
          description: this.props.description,
          context:this.props.context,
          dom: this.props.dom,
        }
      );
      ReactDom.unmountComponentAtNode(this.props.dom);          
      ReactDom.render(element, this.props.dom);
  }
  private _onAddExisting() {
    const element: React.ReactElement<IAccessRequestsProps > = React.createElement(
      ModifyAccessRequest,
        {
          description: this.props.description,
          context:this.props.context,
          dom: this.props.dom,
        }
      );
      ReactDom.unmountComponentAtNode(this.props.dom);          
      ReactDom.render(element, this.props.dom);
  }
  // private listNotConfigured(props: IAccessRequestsProps): boolean {
  //   return props.listName === undefined ||
  //     props.listName === null ||
  //     props.listName.length === 0;
  // }

    
}
