import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'LmAppCustomizerApplicationCustomizerStrings';
import styles from './LmAppCustomizerApplicationCustomizer.module.scss';
import { graph } from '@pnp/graph'
// import '@pnp/graph/Users'
import "@pnp/graph/teams"
import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
import { Web } from '@pnp/sp/webs';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

const LOG_SOURCE: string = 'LmAppCustomizerApplicationCustomizer';
let myVar;
export interface ILmAppCustomizerApplicationCustomizerProperties {
  testMessage: string;
  Bottom: string;
}

// let bellHtml = `<div style="width: 48px"><a href='/sites/LibertyMutual/SitePages/User-NO.aspx'><img src='data:image/png;base64,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' 
//       style="margin-left:auto; margin-right:auto; display:block; height:16px; margin-top:34%;"></div>`;
let bellHtml = `<div class="lm-bell-icon">
                  <a href='/sites/LibertyMutual/SitePages/User-NO.aspx'>
                    <img src='https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SourceCode/img/bell.svg'/>
                  </a>
                </div>`;

let searchIconHtml = `<div class="lm-bell-icon" id="lm-bell">
                <a data-interception="off" target="_blank" href='https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/_layouts/15/search.aspx'>
                  <img src='https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SourceCode/img/lm-search.svg'/>
                </a>
              </div>`;

let companyLogo = `<div class="lm-logo">
                  <img src="https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SourceCode/img/OnDemandLogo.png"/>
                  </div>`

export default class LmAppCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<ILmAppCustomizerApplicationCustomizerProperties> {
    private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    // this.FetchUsers();
    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }
    sp.setup({
      spfxContext: this.context
    })
    // graph.setup({
    //   spfxContext: this.context
    // });
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);


    this.context.application.navigatedEvent.add(this, () => {
      // setTimeout(() => {
      // myTimer = window.setInterval(this.bindCustomHtml, 500);
      // }, 1000);
      var myVar = setInterval(myTimer, 1000);

      async function myTimer() {
        //var centerRegion = document.getElementById("centerRegion");
        var leftRegion = document.getElementById("leftRegion");
        if (leftRegion != null) {
          clearInterval(myVar);
          if (location.href.indexOf("_layouts/15/search") == -1) {
            var x = document.getElementById("O365_SearchBoxContainer_container");
            x.style.display = "none";
            //Adding Bell & Search Icon in suite bar
            //centerRegion.innerHTML += searchIconHtml + bellHtml;

            //Adding Liberty OnDemand Logo in suite bar
            //Adding Welcome <userName> in suitbar
            await sp.web.currentUser().then(user => {
              let userNameHtml = `<div class="custom-username">Welcome ${user.Title}</div>`;
              if (document.getElementsByClassName("lm-bell-icon").length == 0) {
                document.getElementById("leftRegion").innerHTML = userNameHtml + companyLogo;
                document.getElementById("centerRegion").innerHTML += searchIconHtml + bellHtml;
              }
            });
            //document.getElementById("lm-bell").addEventListener("click", showHideSearchbox);
          }
        }
      }

      function showHideSearchbox() {
        var x = document.getElementById("O365_SearchBoxContainer_container");
        if (window.getComputedStyle(x).visibility === "hidden") {
          x.style.visibility = "visible";
        }
        else {
          x.style.visibility = "hidden";
        }
      }
    });



    //Injecting custom CSS
    const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
    let customStyle: HTMLLinkElement = document.createElement("link");
    customStyle.href = "https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SourceCode/css/customCSS.css?v=1.1";
    customStyle.rel = "stylesheet";
    customStyle.type = "text/css";
    head.insertAdjacentElement("beforeEnd", customStyle);

    // let jqueryTag: HTMLScriptElement = document.createElement("script");
    // jqueryTag.src = "https://code.jquery.com/jquery-3.6.0.min.js";
    // jqueryTag.type = "text/javascript";
    // document.getElementsByTagName("head")[0].appendChild(jqueryTag);

    // let scriptTag: HTMLScriptElement = document.createElement("script");
    // scriptTag.src = "https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SourceCode/js/customJS.js";
    // scriptTag.type = "text/javascript";
    // document.getElementsByTagName("head")[0].appendChild(scriptTag);

    return Promise.resolve();
  }



  protected FetchUsers(): void {
    this.context.aadHttpClientFactory
      .getClient("https://graph.microsoft.com")
      .then((client: AadHttpClient) => {
        return client
          .get(
            "https://graph.microsoft.com/v1.0/teams/78b2fdab-65df-4a96-acf3-20e9ff56955f/channels/19:t6eyPcyzT4AkqxKTEsfs6fO2JK3Njw-xsskTmOnMPo41@thread.tacv2/messages",
            AadHttpClient.configurations.v1
          );
      })
      .then(response => {
        return response.json();
      })
      .then(json => {
        console.log(json);
      }
      )
  }



  private async loadUserDetails(): Promise<string> {
    try {
      let user = await sp.web.currentUser();
      return user.Title;
    } catch (error) {
      console.log("Error in loadUserDetails : ", error);
    }
  }

  private _renderPlaceHolders(): void {
      
    // Handling the bottom placeholder
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );
  
      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }
  
      if (this.properties) {
        let bottomString: string = this.properties.Bottom;
        if (!bottomString) {
          bottomString = "(Bottom property was not defined.)";
        }
  
        if (this._bottomPlaceholder.domElement) {
          this._bottomPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.bottom}">
            <div class="${styles.footerLogo}">
              <a href="#!"><img src="https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SiteAssets/img/footer-logo.jpg" /></a>
            </div>
            <div class = "${styles.copyright}">
              <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> 
                © 2020 Liberty Insurace Pte Ltd. 51 club street #3-00 Liberty House Singapore 069428
            </div>
            <div class="${styles.footerSocial}">
                <a href="#!">
                <img src="https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SiteAssets/img/facebook-logo.svg" width="24" />
                </a>
                <a href="#!">
                  <img src="https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SiteAssets/img/twitter.svg" width="24" />
                </a>
                <a href="#!">
                  <img src="https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SiteAssets/img/linkedin.svg" width="24" />
                </a>
                <a href="#!">
                  <img src="https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/SiteAssets/img/youtube.svg" width="24" />
                </a>
            </div>
          </div>`;
        }
      }
    }
  }

  private _onDispose(): void {
    console.log('[NotificationApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
