import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestCamlQueryWebPart.module.scss';
import * as strings from 'TestCamlQueryWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; 
//import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { LogLevel, PnPLogging } from "@pnp/logging";

require("bootstrap");

export interface ITestCamlQueryWebPartProps {
  description: string;
  division : string;
  teamTermID : string;
  parentTermID : string;
  libraryName: string[];
}

export default class TestCamlQueryWebPart extends BaseClientSideWebPart<ITestCamlQueryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public async render(): Promise<void> {

    this.domElement.innerHTML = `
    <section class="${styles.testCamlQuery} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div class="${styles.row}">     
        <div class="${styles.row}" id="libraryName"></div>
      </div>
    </section>`;

    //*** check data exists for libraries and display tab buttons */
    this._getDataAsync(false,"",""); 
  }

  private async _getDataAsync(flag:boolean,library:string,category:string): Promise<void> {
    console.log('renderlistasync');
    const dcDivisions : string[] = ["asm","cen","cnn","emp","hea"];

    if(!flag){
      this.properties.libraryName = ["Policies", "Procedures","Guides", "Forms", "General"];
    }else{
      this.properties.libraryName = [library];
    }

    for (let x = 0; x < this.properties.libraryName.length; x++) {
      dcDivisions.forEach(async (site,index)=>{
        this._getData(flag,site,this.properties.libraryName[x],"IPES Wales",category)
          .then(async (response) => {
            console.log("renderlistasync",response);
            if(response.length>0){
              if(!flag){
                await this._renderLibraryTabs(this.properties.libraryName[x]).then( ()=> {
                  this._setLibraryListeners();
                });
              }else{
                console.log('running other functions');
              }
            }
          })
          .catch(() => {});
      });
    }
    return;
  }

  private async _getData(flag:boolean,site:string,library:string,team:string,category:string): Promise<any> {      
    console.log('checkdata');

    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));  
    const tenant_uri = this.context.pageContext.web.absoluteUrl.split('/',3)[2];
    const dcTitle = site+"_dc";
    const webDC = Web([sp.web,`https://${tenant_uri}/sites/${dcTitle}/`]); 
    let rowLimitString : string;
    let view: string = "";
    
    if(!flag){
      rowLimitString="<RowLimit>10</RowLimit>";
    }else{
      rowLimitString="";
    }

    if (category === "") {
      view =
        `<View>
          <Query>
            <Where>
              <Or>
                <Eq>
                  <FieldRef Name="DC_Team"/>
                  <Value Type="TaxonomyFieldType">${team}</Value>
                </Eq>
                <Contains>
                  <FieldRef Name="DC_SharedWith"/>
                  <Value Type="TaxonomyFieldTypeMulti">${team}</Value>
                </Contains>
              </Or>
            </Where>
            <OrderBy>
              <FieldRef Name="DC_Division" Ascending="TRUE" />
              <FieldRef Name="DC_Folder" Ascending="TRUE" />
              <FieldRef Name="DC_SubFolder01" Ascending="TRUE" />
              <FieldRef Name="DC_SubFolder02" Ascending="TRUE" />
              <FieldRef Name="DC_SubFolder03" Ascending="TRUE" />
              <FieldRef Name="LinkFilename" Ascending="TRUE" />
            </OrderBy>          
          </Query>
          ${rowLimitString}
        </View>`;
    } else {
      view =
        `<View>
          <Query>
            <Where>
              <Or>
                <Eq>
                  <FieldRef Name="DC_Category"/>
                  <Value Type="TaxonomyFieldType">${category}</Value>
                </Eq>
                <Contains>
                  <FieldRef Name="DC_SharedWith"/>
                  <Value Type="TaxonomyFieldTypeMulti">${category}</Value>
                </Contains>
              </Or>
            </Where>
            <OrderBy>
              <FieldRef Name="DC_Division" Ascending="TRUE" />
              <FieldRef Name="DC_Folder" Ascending="TRUE" />
              <FieldRef Name="DC_SubFolder01" Ascending="TRUE" />
              <FieldRef Name="DC_SubFolder02" Ascending="TRUE" />
              <FieldRef Name="DC_SubFolder03" Ascending="TRUE" />
              <FieldRef Name="LinkFilename" Ascending="TRUE" />
            </OrderBy>           
          </Query>
          ${rowLimitString}
        </View>`;
    }

    //console.log(view);

    return webDC.lists.getByTitle(library)
      .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
      .then(async (response) => {
        //console.log("checkdata",response);
        return response;  
      })
      .catch(() => {});    
  }

  private async _renderLibraryTabs(library:string): Promise<void> {
    console.log('renderlist');

    //const dataTarget:string=library.toLowerCase();
    let html: string = '';

    //console.log("renderlist",library);
    html = `<button class="btn btn-primary text-center mb-1" id="${library}_btn" type="button"><h6 class="libraryText">${library}</h6></button>`;

    if(this.domElement.querySelector('#libraryName') !== null){
      this.domElement.querySelector('#libraryName')!.innerHTML += html;
    }
    return;
  }

  private _setLibraryListeners() : void {
    console.log("librarylisteners");

    // Appending the `!` operator here
    //const container = document.querySelector('#Guides_btn')!

    let timer:any;
    //const policyElem=document.getElementById("Policies_btn")!
    const procedureElem=document.getElementById("Procedures_btn")!
    const guideElem=document.getElementById("Guides_btn")!
    const formElem=document.getElementById("Forms_btn")!
    ///const generalElem=document.getElementById("General_btn")!

    // TypeScript will not complain about the container being possibly `null`
    //container.addEventListener('click', () => alert('Guides Button clicked')) 

    // *** event listeners for main document libraries

    procedureElem.addEventListener("click", event => { 
      if(event.detail===1){
        timer=setTimeout(async () => {
          alert(`Procedure Button clicked`)
          this._getDataAsync(true,"Procedures",""); 
          //await this.getData("Policies",1,"");
        },100);
      }
    });
    procedureElem.addEventListener("dblclick",event => {
      clearTimeout(timer);
    });

    guideElem.addEventListener("click", event => {
      if(event.detail===1){
        timer=setTimeout(async () => {
          alert(`Guides Button clicked`)
          this._getDataAsync(true,"Guides",""); 
          //await this.getData("Guides",3,"");
        },100);
      }
    });
    guideElem.addEventListener("dblclick",event => {
      clearTimeout(timer);
    });

    formElem.addEventListener("click", event => { 
      if(event.detail===1){
        timer=setTimeout(async () => {
          alert(`Form Button clicked`)
          //await this.getData("Policies",1,"");
        },100);
      }
    });
    formElem.addEventListener("dblclick",event => {
      clearTimeout(timer);
    });               

  }

  public async onInit(): Promise<void> {
    await super.onInit();
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}