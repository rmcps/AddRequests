import * as React from 'react';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import styles from './AccessRequests.module.scss';

export interface IDefaultProps {
    context:any;
}

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
        <p>Proin auctor, libero eget ornare mattis, mi neque suscipit erat, mollis faucibus quam ligula quis eros. Praesent rutrum maximus ante et fringilla. Duis scelerisque eleifend sem, sed egestas sem maximus in. Quisque quis semper ligula. Duis eros neque, luctus id dui eu, pharetra maximus nibh. Mauris in tortor tortor. Nam malesuada nunc vitae ligula rhoncus volutpat. Phasellus sed ligula neque. Etiam sed euismod nisi. Etiam ac orci sed ante efficitur fringilla. Ut laoreet et nunc et finibus. Duis non bibendum mi, vel maximus ex. Mauris at ornare arcu, id suscipit urna. Aliquam erat volutpat.</p>
            <PrimaryButton onClick={ this._onAddNew } text='Add a new member access request' />
            <DefaultButton text='Add a request for an existing member' />                  
          </div>                    
        </div>
      </div>
    );
  }

  private _onAddNew() {

  }
  private _onAddExisting() {

  }
  // private listNotConfigured(props: IAccessRequestsProps): boolean {
  //   return props.listName === undefined ||
  //     props.listName === null ||
  //     props.listName.length === 0;
  // }

    
}
