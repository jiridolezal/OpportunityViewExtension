import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IOpportunity } from '../../IOpportunity';


export interface IViewApplicationCustomizerProperties {
  testMessage: string;
}

export default class ViewApplicationCustomizer
  extends BaseApplicationCustomizer<IViewApplicationCustomizerProperties> {

  private spHttpClient: SPHttpClient;

  public onInit(): Promise<void> {
    // Obtain SPHttpClient instance from context
    this.spHttpClient = this.context.spHttpClient;

    
    
    // Make a GET request to fetch items from the "Temporary" list
    this.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('oneSfaRecordsList')/items?$filter=sfaLeadId eq 'D22_CZ-TMCZ~00Q3N000007AFRJUA4'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          console.log(`Error getting data: ${response.statusText}`);
          return Promise.reject(response.statusText);
        }
      })
      .then((responseData: any) => {
        console.log(responseData);
        const data: IOpportunity = responseData.value[0];
        console.log(`Successfully loaded Opportunity: ${data.sfaLeadId}`);

        // Fetch user information for each ID
        Promise.all([
          this.getUserInfo(data.sfaSalerStringId),
          this.getUserInfo(data.sfaBidManagerStringId),
          this.getUserInfo(data.sfaGarantStringId),
          this.getUserInfo(data.sfaLegalStringId)
        ])
        .then((usersData: any[]) => {
          const salerName = usersData[0].Title;
          const managerName = usersData[1].Title;
          const garantName = usersData[2].Title;
          const legalName = usersData[3].Title;

          // Display the dynamic content
          const targetElement = document.querySelector('.od-TopBar-item.od-TopBar-commandBar.od-TopBar-commandBar--suiteNavSearch');
          if (targetElement && (!document.getElementById('InjectedExtensionDiv'))) {
            targetElement.insertAdjacentHTML('afterend', `
              <div id="InjectedExtensionDiv">
                <p>Title: ${data.Title}</p>
                <p>Sfa Lead Id: ${data.sfaLeadId}</p>
                <p>Sfa Customer: ${data.sfaCustomer}</p>
                <p>Sfa Saler: ${salerName}</p>
                <p>Sfa Manager: ${managerName}</p>
                <p>Sfa Garant: ${garantName}</p>
                <p>Sfa Legal: ${legalName}</p>
              </div>
            `);
          } else {
            console.error("Target element not found.");
          }
        })
        .catch((error: any) => {
          console.error(`Error: ${error}`);
        });
      })
      .catch((error: any) => {
        console.error(`Error: ${error}`);
      });
    
    return Promise.resolve();
  }

  // Method to fetch user information by ID
  private getUserInfo(userId: string): Promise<any> {
    return this.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/getuserbyid(${userId})`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          console.log(`Error getting user data: ${response.statusText}`);
          return Promise.reject(response.statusText);
        }
      });
  }
}