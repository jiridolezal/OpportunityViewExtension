import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IOpportunity } from '../../IOpportunity';

export interface IViewApplicationCustomizerProperties {
  testMessage: string;
}

export default class ViewApplicationCustomizer
  extends BaseApplicationCustomizer<IViewApplicationCustomizerProperties> {

  private spHttpClient: SPHttpClient;
  //private _counter: number = 1;
  private previousUrl: string;

  public onInit(): Promise<void> {
    // Obtain SPHttpClient instance from context
    this.spHttpClient = this.context.spHttpClient;

    // Save the initial URL
    this.previousUrl = window.location.href;

    // Start polling for URL changes
    this.startUrlPolling();
 
    // Render the custom div initially
    this.renderCustomDiv();
 
    return Promise.resolve();
  }

  private startUrlPolling(): void {
    setInterval(() => {
        const currentUrl = window.location.href;
        if (currentUrl !== this.previousUrl) {
            // If URL has changed, rerender the custom div
            this.renderCustomDiv();
            // Update the previous URL
            this.previousUrl = currentUrl;
        }
        console.log("Timer hit!")
    }, 500); // Poll every 1 second (adjust interval as needed)
  }

  private renderCustomDiv(): void {
    // Find URL, parse it and call the correct endpoint with REST API
    const url = window.location.href;
    const decodedUrl = decodeURIComponent(url);
    const parts = decodedUrl.split('/');
    let opportunity: string;
    if (parts.length < 16) {
      let divToRemove = document.getElementById("InjectedExtensionDiv");
      if (divToRemove) {
        if (divToRemove.parentNode) {
          divToRemove.parentNode.removeChild(divToRemove);
        }
      }
      return;
    }else if (parts.length == 16) {
      const last = parts[15];
      const lastSplit = last.split('&');
      opportunity = lastSplit[0];
    }else{
      opportunity = parts[15];
    }
  
    
    // Make a GET request to fetch items from the "Temporary" list
    this.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('oneSfaRecordsList')/items?$filter=sfaLeadId eq '${opportunity}'`, SPHttpClient.configurations.v1)
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

          // Create or update the dynamic content
          let injectedDiv = document.getElementById('InjectedExtensionDiv');
          if (!injectedDiv) {
            // If the div doesn't exist, create a new one
            const targetElement = document.querySelector('.od-TopBar-item.od-TopBar-commandBar.od-TopBar-commandBar--suiteNavSearch');
            if (targetElement) {
              targetElement.insertAdjacentHTML('afterend', `
                <div id="InjectedExtensionDiv">
                  <p>${opportunity}</p>
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
          } else {
            // If the div exists, update its content
            injectedDiv.innerHTML = `
              <p>${opportunity}</p>
              <p>Title: ${data.Title}</p>
              <p>Sfa Lead Id: ${data.sfaLeadId}</p>
              <p>Sfa Customer: ${data.sfaCustomer}</p>
              <p>Sfa Saler: ${salerName}</p>
              <p>Sfa Manager: ${managerName}</p>
              <p>Sfa Garant: ${garantName}</p>
              <p>Sfa Legal: ${legalName}</p>
            `;
          }
        })
        .catch((error: any) => {
          console.error(`Error: ${error}`);
        });
      })
      .catch((error: any) => {
        console.error(`Error: ${error}`);
      });
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
