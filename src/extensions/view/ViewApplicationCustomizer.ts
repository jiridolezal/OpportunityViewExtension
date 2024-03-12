/* eslint-disable */

import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IOpportunity } from '../../IOpportunity';
import styles from './ViewApplicationCustomizer.module.scss';

export interface IViewApplicationCustomizerProperties {
  testMessage: string;
}

export interface IConfig {
  tenantId: string,
  opportunityUrl: string,
  leadUrl: string,
}

export default class ViewApplicationCustomizer
  extends BaseApplicationCustomizer<IViewApplicationCustomizerProperties> {

  private spHttpClient: SPHttpClient;
  private config: IConfig = {tenantId: "b213b057-1008-4204-8c53-8147bc602a29", //TMobile specific tenant ID
                             opportunityUrl: "https://tmobileczsk--situat.sandbox.lightning.force.com/lightning/cmp/coredt__NavigateTo?c__objectName=Opportunity&c__externalId=", 
                             leadUrl: "https://tmobileczsk--situat.sandbox.lightning.force.com/lightning/cmp/coredt__NavigateTo?c__objectName=Lead&c__externalId="};
  private previousUrl: string;

  public onInit(): Promise<void> {
    // Obtain SPHttpClient instance from context
    this.spHttpClient = this.context.spHttpClient;

    // Check if the current page is "Verejne_zakazky"
    if (window.location.href.toLowerCase().indexOf("/sites/tmozakazky/verejne_zakazky") !== -1) {
      // Render the custom div only if on "Verejne_zakazky" page
      this.renderCustomDiv();
    }

    // Save the initial URL
    this.previousUrl = window.location.href;

    // Start polling for URL changes
    this.startUrlPolling();
    return Promise.resolve();
  }

  // private async loadConfig(): Promise<void> {
  //   try {
  //       const configUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/getfilebyserverrelativeurl('/sites/tmoZakazky/scripts/spfx.json')/$value`;
  //       const response: SPHttpClientResponse = await this.context.spHttpClient.get(configUrl, SPHttpClient.configurations.v1);
  //       const configJson: any = await response.json();
  //       this.config = configJson;
  //       console.log("Config loaded", this.config); 
  //   } catch (error) {
  //       console.error('Error loading config:', error);
  //   }
  // }

  private startUrlPolling(): void {
    setInterval(() => {
        const currentUrl = window.location.href;
        if (currentUrl !== this.previousUrl) {
            // Update the previous URL
            this.previousUrl = currentUrl;

            // Check if the current page is "Verejne_zakazky"
            if (currentUrl.toLowerCase().indexOf("/sites/tmozakazky/verejne_zakazky") !== -1) {
                // If URL has changed and on "Verejne_zakazky" page, rerender the custom div
                this.renderCustomDiv();
            } else {
                // If not on "Verejne_zakazky" page, remove the custom div
                let divToRemove = document.getElementById("InjectedExtensionDiv");
                if (divToRemove && divToRemove.parentNode) {
                    divToRemove.parentNode.removeChild(divToRemove);
                }
            }
        }
        console.log("Polling for URL change...");
    }, 500); // Poll every half second (adjust interval as needed)
  }

  private removeInjectedExtensionDiv(): void {
    let divToRemove = document.getElementById("InjectedExtensionDiv");
    if (divToRemove && divToRemove.parentNode) {
      divToRemove.parentNode.removeChild(divToRemove);
    }
  }

  private renderCustomDiv(): void {
    // Find URL, parse it and call the correct endpoint with REST API
    const url = window.location.href;
    const decodedUrl = decodeURIComponent(url);
    // Find the index of the part that starts with 'id='
    const idIndex = decodedUrl.indexOf('id=');
    if (idIndex === -1) {
      this.removeInjectedExtensionDiv();
      return;
    }
    // Get the parts after 'id='
    const partsAfterId = decodedUrl.substring(idIndex + 4).split('/');
    let opportunity: string;
    if (partsAfterId.length < 4) {
      this.removeInjectedExtensionDiv();
      return;
    } else if (partsAfterId.length === 4) {
      opportunity = partsAfterId[3].split('&')[0];
    } else {
      opportunity = partsAfterId[3];
    }

    console.log(`Opportunity ID: ${opportunity}`);
  
    // Make a GET request to fetch items from the "Temporary" list
    this.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('oneSfaRecordsList')/items?$filter=sfaLeadId eq '${opportunity}'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          throw Error("Failed to fetch data");
        }
      })
      .then((responseData: any) => {
        if (responseData.length === 0) {
          // If the data is empty, throw an error
          throw Error('No matching sfaLeadId found in the database.');
        }

        console.log("Data from the REST Api call loaded: ", responseData);
        const data: IOpportunity = responseData.value[0];
        console.log(`Successfully loaded Opportunity: ${data.sfaLeadId}`, opportunity);

        // Create or update the dynamic content
        let injectedDiv = document.getElementById("InjectedExtensionDiv");

        if (!injectedDiv) {
          // If the div doesn't exist, create a new one
          const targetElement = document.querySelector('.od-TopBar-item.od-TopBar-commandBar.od-TopBar-commandBar--suiteNavSearch');
          if (targetElement) {
            targetElement.insertAdjacentElement('afterend', this.generateInjectedDiv(data));
          } else {
            console.error("Target element not found. Cannot insert the custom div.");
          }
        } else {
          // If the div exists, update its content
          injectedDiv = this.generateInjectedDiv(data);
        }
      })
      .catch((error: any) => {
        this.removeInjectedExtensionDiv();
        console.error(`Get request was not successful: ${error}`);
      });
  }

  // Method to fetch user information by ID
  private getUserInfo(userId: string): Promise<any> {
    return this.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/getuserbyid(${userId})`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          console.error(`Error getting user data: ${response.statusText}`);
          return Promise.reject(response.statusText);
        }
      });
  }

  // Method to generate Opportunity items
  private generateOpportunityItem(parameterName: string, parameterValue: string): HTMLElement {
    if (parameterValue === null || parameterValue === undefined) {
      parameterValue = 'N/A';
    }
    
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityItemView;

    let parameterNamePar = document.createElement('p');
    parameterNamePar.className = styles.opportunityItemParamName;
    parameterNamePar.innerHTML = parameterName;

    let parameterValuePar = document.createElement('p');
    parameterValuePar.className = styles.opportunityItemParamValue;
    parameterValuePar.innerHTML = parameterValue;

    divElem.appendChild(parameterNamePar);
    divElem.appendChild(parameterValuePar);

    return divElem;
  }

  private generateButtons(data: IOpportunity): HTMLElement {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityButtonContainer;


    let teamsButton : HTMLButtonElement | null = document.createElement('button');
    if ((data.sfaTeamId !== null && data.sfaTeamId !== undefined) &&
        (data.sfaGenChannel !== null && data.sfaGenChannel !== undefined) &&
        (this.config.tenantId !== null && this.config.tenantId !== undefined)) {
      teamsButton.className = styles.opportunityLinkButton;
      teamsButton.innerHTML = 'Teams';
      teamsButton.addEventListener('click', () => {
        const sfaGenChannel = data.sfaGenChannel;
        const sfaTeamId = data.sfaTeamId;
        const teamsUrl = `https://teams.microsoft.com/v2/l/channel/${sfaGenChannel}/General?groupId=${sfaTeamId}&tenantId=${this.config.tenantId}`;
        window.open(teamsUrl, '_blank');
      });
    }else{
      teamsButton = null;
    }

    let salesForceButton: HTMLButtonElement | null = document.createElement('button');
    let salesForceUrl: string;
    if (data.sfaOpportunityId !== null && data.sfaOpportunityId !== undefined &&
        this.config.opportunityUrl !== null && this.config.opportunityUrl !== undefined) {
      salesForceUrl = `${this.config.opportunityUrl}${data.sfaOpportunityId}`;
    }else if (data.sfaLeadId !== null && data.sfaLeadId !== undefined &&
              this.config.leadUrl !== null && this.config.leadUrl !== undefined) {
      salesForceUrl = `${this.config.leadUrl}${data.sfaLeadId}`;
    }else{
      salesForceButton = null;
    }

    if (teamsButton !== null) {
      divElem.appendChild(teamsButton);
    }
    if (salesForceButton !== null) {
      salesForceButton.className = styles.opportunityLinkButton;
      salesForceButton.innerHTML = 'SalesForce';
      salesForceButton.addEventListener('click', () => {
        window.open(salesForceUrl, '_blank');
      });
      divElem.appendChild(salesForceButton);
    }
    
    return divElem;
  }

  private generateTitle(title: string): HTMLElement {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityTitleContainer;

    let titleParam = document.createElement('p');
    titleParam.className = styles.opportunityTitleParam;
    titleParam.innerHTML = 'Název zakázky';

    let titleValue = document.createElement('p');
    titleValue.className = styles.opportunityTitleValue;
    titleValue.innerHTML = title;

    divElem.appendChild(titleParam);
    divElem.appendChild(titleValue);

    return divElem;
  }

  private generateItems(data: IOpportunity): HTMLElement {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityItems;

    // Fetch user information for each ID
    Promise.all([
      (data.sfaSalerStringId === null || data.sfaSalerStringId == undefined) 
       ? null 
       : this.getUserInfo(data.sfaSalerStringId),
      (data.sfaBidManagerStringId === null || data.sfaBidManagerStringId == undefined) 
       ? null 
       : this.getUserInfo(data.sfaBidManagerStringId),
      (data.sfaGarantStringId === null || data.sfaGarantStringId == undefined) 
       ? null 
       : this.getUserInfo(data.sfaGarantStringId),
      (data.sfaLegalStringId === null || data.sfaLegalStringId == undefined) 
       ? null 
       : this.getUserInfo(data.sfaLegalStringId)
    ])
    .then((usersData: any[]) => {
      const salerName = (usersData[0] === null || usersData[0] == undefined)
      ? null
      : usersData[0].Title;
      const managerName = (usersData[1] === null || usersData[1] == undefined)
      ? null
      : usersData[1].Title;
      const garantName = (usersData[2] === null || usersData[2] == undefined)
      ? null
      : usersData[2].Title;
      const legalName = (usersData[3] === null || usersData[3] == undefined)
      ? null
      : usersData[3].Title;
      
      divElem.appendChild(this.generateOpportunityItem('Zadavatel', data.sfaCustomer));
      divElem.appendChild(this.generateOpportunityItem('Název VZ', data.sfaLeadName));
      divElem.appendChild(this.generateOpportunityItem('RFP Day', data.sfaRfpDay));
      divElem.appendChild(this.generateOpportunityItem('Bid. M', managerName));
      divElem.appendChild(this.generateOpportunityItem('Garant nabídky', garantName));
      divElem.appendChild(this.generateOpportunityItem('Legal', legalName));
      divElem.appendChild(this.generateOpportunityItem('Obchodník', salerName));
      divElem.appendChild(this.generateOpportunityItem('Go/NoGo', data.sfaGoNoGo));
    })

    return divElem;
  }

  private generateContent(data: IOpportunity): HTMLElement {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityViewContent;
    
    let titleDiv = this.generateTitle(data.Title);
    let itemsDiv = this.generateItems(data);

    divElem.appendChild(titleDiv);
    divElem.appendChild(itemsDiv);

    return divElem;
  }

  private generateInjectedDiv(data: IOpportunity): HTMLElement {
    const wholeDiv = document.createElement("div");
    wholeDiv.className = styles.wholeDiv;

    const baseDiv = document.createElement("div");

    baseDiv.setAttribute("id", "InjectedExtensionDiv");
    baseDiv.className = styles.baseInjectedDiv

    wholeDiv.appendChild(this.generateContent(data));
    wholeDiv.appendChild(this.generateButtons(data));

    baseDiv.appendChild(wholeDiv);

    return baseDiv;
  }
}

/* eslint-enable */