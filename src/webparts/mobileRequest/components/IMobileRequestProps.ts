import { SPHttpClient } from '@microsoft/sp-http'; 
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMobileRequestProps {
  listName: string;
  spHttpClient: SPHttpClient;  
  context: WebPartContext;
  siteUrl: string;  
}
