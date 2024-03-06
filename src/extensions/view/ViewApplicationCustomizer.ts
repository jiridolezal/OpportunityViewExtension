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
  rootFolderIndex: number
}

export default class ViewApplicationCustomizer
  extends BaseApplicationCustomizer<IViewApplicationCustomizerProperties> {

  private spHttpClient: SPHttpClient;
  private config: IConfig;
  private previousUrl: string;

  public onInit(): Promise<void> {
    // Obtain SPHttpClient instance from context
    this.spHttpClient = this.context.spHttpClient;

    // Load config
    this.loadConfig();

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

  private async loadConfig(): Promise<void> {
    try {
        const configUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/getfilebyserverrelativeurl('/sites/tmoZakazky/scripts/spfx.json')/$value`;
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(configUrl, SPHttpClient.configurations.v1);
        const configJson: any = await response.json();
        this.config = configJson;
        console.log("Config loaded", this.config); 
    } catch (error) {
        console.error('Error loading config:', error);
    }
  }

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
        console.log("Timer hit!");
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

        // Create or update the dynamic content
        let injectedDiv = document.getElementById("InjectedExtensionDiv");

        if (!injectedDiv) {
          // If the div doesn't exist, create a new one
          const targetElement = document.querySelector('.od-TopBar-item.od-TopBar-commandBar.od-TopBar-commandBar--suiteNavSearch');
          if (targetElement) {
            targetElement.insertAdjacentElement('afterend', this.generateInjectedDiv(data));
          } else {
            console.error("Target element not found.");
          }
        } else {
          // If the div exists, update its content
          injectedDiv = this.generateInjectedDiv(data);
        }
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

  // Method to generate Opportunity items
  private generateOpportunityItem(parameterName: string, parameterValue: string): HTMLElement {
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

    let teamsButton = document.createElement('button');
    teamsButton.className = styles.opportunityLinkButton;
    teamsButton.innerHTML = 'Teams';
    teamsButton.addEventListener('click', () => {
      const sfaGenChannel = data.sfaGenChannel;
      const sfaTeamId = data.sfaTeamId;
      const teamsUrl = `https://teams.microsoft.com/v2/l/channel/${sfaGenChannel}/General?groupId=${sfaTeamId}&tenantId=${this.config.tenantId}`;
      window.open(teamsUrl, '_blank');
    });

    let salesForceButton = document.createElement('button');
    salesForceButton.className = styles.opportunityLinkButton;
    salesForceButton.innerHTML = 'SalesForce';
    let salesForceUrl: string;
    if (data.sfaOpportunityId === null || data.sfaOpportunityId === undefined) {
      salesForceUrl = `${this.config.opportunityUrl}${data.sfaOpportunityId}`;
    }else{
      salesForceUrl = `${this.config.leadUrl}${data.sfaLeadId}`;
    }
    salesForceButton.addEventListener('click', () => {
      window.open(salesForceUrl, '_blank');
    });

    divElem.appendChild(teamsButton);
    divElem.appendChild(salesForceButton);

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