import * as React from 'react';
import styles from './MobileRequest.module.scss';
import { IMobileRequestProps } from './IMobileRequestProps';
import { IReactCrudState } from "./IReactCrudState";
import { escape } from '@microsoft/sp-lodash-subset';
import { Log } from "@microsoft/sp-core-library";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from './IListItem';

const Log_SOURCE: string = "Mobile Request Approval";

export default class MobileRequest extends React.Component<IMobileRequestProps, IReactCrudState> {

  constructor(props: IMobileRequestProps, state: IReactCrudState) {
    super(props);
    this.state = {
      status: 'Ready',
      items: []
    };
  }

  public render(): React.ReactElement<IMobileRequestProps> {
    return (
      <div className={styles.mobileRequest}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.listName)}</p>
              <a href="#" className={`${styles.button}`} onClick={() => this.createItem()}>
                <span className={styles.label}>Create item</span>
              </a>
            </div>
            {this.state.status}
          </div>
        </div>
      </div>
    );
  }
  private createItem(): void {
    this.setState({
      status: 'Creating item...',
      items: []
    });

    const body: string = JSON.stringify({
      'Title': `Item ${new Date()}`,
      'Cost':1,
      'Comments':`Item ${new Date()}`,
      'Justification':`Item ${new Date()}`
    });
    try {
    this.props.context.spHttpClient.post(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('MobileRequestApproval')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then((item: IListItem): void => {
        this.setState({
          status: `Item '${item.Title}' (ID: ${item.Cost}) successfully created`,
          items: []
        });
      }, (error: any): void => {
        this.setState({
          status: 'Error while creating the item: ' + error,
          items: []
        });
      });
    } catch (error) {
      Log.error(Log_SOURCE, error);
    }
  }
}
