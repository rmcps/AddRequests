import * as React from 'react';
import * as ReactDom from 'react-dom';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import styles from './AccessRequests.module.scss';
import NewAccessRequest from './NewAccessRequest/NewAccessRequest';
import ModifyAccessRequest from './ModifyAccessRequest/ModifyAccessRequest';
import { IAccessRequestsProps } from './IAccessRequestsProps';
import IDefaultProps from './IDefaultProps';

export default class DefaultPage extends React.Component<IDefaultProps, null> {

  constructor(props) {
    super(props);
    // set initial state    
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
            <div className={styles.sectionDivider}> <h2>Member Access Request Submission</h2></div>
          </div>        
          <div className= {styles.row}>
            <div><Link onClick={this._onAddNew}>Add a new member access request</Link></div>
            <div><Link onClick={this._onAddExisting}>Add a requet to modify an existing member</Link></div>  
            <div><Link onClick={this._onCancel}>Cancel and return</Link></div>  
          </div>
      </div>
      </div>
    );
  }
@autobind
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
  @autobind
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
@autobind
private _onCancel():void {
  window.location.href = "https://uphpcin.sharepoint.com";
}
  // private listNotConfigured(props: IAccessRequestsProps): boolean {
  //   return props.listName === undefined ||
  //     props.listName === null ||
  //     props.listName.length === 0;
  // }

    
}
