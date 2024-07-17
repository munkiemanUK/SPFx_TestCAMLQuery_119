import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule
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
import {SPHttpClient,SPHttpClientResponse} from "@microsoft/sp-http";  

require("bootstrap");
//let libCount : number = 0;
const groupID = "4660ef58-779c-4970-bcd7-51773916e8dd";  // Prod Document Centre Terms

export interface ITestCamlQueryWebPartProps {
  description: string;
  URL : string;
  tenantURL: string[];
  division : string;
  divisionDC: string;
  dcURL : string;
  teamName : string;
  libraryName: string[];
  libraries : string[];
  libraryNamePrev : string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id : string;
  Title : string;
  DC_Team: string;
  TermGuid : string;
}
export default class TestCamlQueryWebPart extends BaseClientSideWebPart<ITestCamlQueryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public async render(): Promise<void> {

    this.properties.libraryNamePrev = "";
    this.properties.URL = this.context.pageContext.web.absoluteUrl;
    this.properties.tenantURL = this.properties.URL.split('/',5);

    this.domElement.innerHTML = `
    <section class="${styles.testCamlQuery} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>      
      <div class="${styles.row}">
        <h4>${this.properties.divisionDC}</h4>
        <h4>${this.properties.division}</h4>
        <h4>${this.properties.teamName}</h4>
        <div class="${styles.row}" id="libraryName"></div>
      </div>
    </section>`;

    this._getCustomTermsAsync();
  }

  private async _getCustomTermsAsync(): Promise<void> {
    const setID = "be84d0a6-e641-4f6d-830e-11e81f13e2f1";  // Prod Custom Set Terms
    let termID: string = "";
    let label: string = "";
    let termName: string = "";

    try {

      switch (this.properties.division) {
        case "Assessments":
          termID = "90a0a9eb-bbcc-4693-9674-e56c4d41375f";  // Prod Assessment Custom Terms
          this.properties.dcURL ="https://" + this.properties.tenantURL[2] + `/sites/asm_dc`;
          break;
        case "Central":
          termID = "471a563b-a4d9-4ce7-a8e6-4124562b3ace";  // Prod Central Custom Terms
          this.properties.dcURL ="https://" + this.properties.tenantURL[2] + `/sites/cen_dc`;
          break;
        case "Connect":
          termID = "3532f8fc-4ad2-415c-94ff-c5c7af559996";  // Prod Connect Custom Terms
          break;
        case "Employability":
          termID = "feb3d3c8-d948-4d3e-b997-a2ea74653b3e";  // Prod Employability Custom Terms
          break;
        case "Health":
          termID = "c9dfa3b6-c7c6-4e74-a738-0ffe54e1ff5c";  // Prod Health Custom Terms
          break;
      }
    
      await this._getCustomTerms(setID, termID)
      .then(async (terms) => {
        //terms.forEach(async (item:any,index:number)=>{
        for (let x = 0; x < terms.value.length; x++) {
          label = terms.value[x].labels[0].name;
          termID = terms.value[x].id;
  
          if (label === this.properties.teamName) {
            await this._getCustomTerms(setID, termID).then(async (response) => {
              console.log("getCustomTabs",response);
              for(let x=0;x<response.value.length;x++){
                termName = response.value[x].labels[0].name;   
                console.log("termName",termName); 
                await this._getDataAsync(false,"Custom",termName);
                //console.log("customFlag",this.properties.customDataFlag);
              }
           });
          }
        }
      })
      .catch((err) => {console.log('renderCustomTabs ERROR:',err)});

    } catch (err) {
      //await this.addError(this.properties.teamName,"_renderCustomTabsAsync",err);
      //Log.error('DocumentCentre', new Error('_renderCustomTabsAsync Error message'), err);
    }

    //fetchCustomTerms= true;
    return;
  }

  private async _getCustomTerms(setID: string, termID: string): Promise<any> {
    try {
      const url: string = `https://${this.properties.tenantURL[2]}/_api/v2.1/termStore/groups/${groupID}/sets/${setID}/terms/${termID}/children?select=id,labels`;

      return this.context.spHttpClient
        .get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            console.log("getCustomTerms response ok",response);
            return response.json();
          }
        });        
    } catch (err) {
      //await this.addError(this.properties.teamName,"getCustomTabs",err);
      console.log("getCustomTerms error",err);
      //Log.error('DocumentCentre', new Error('getCustomTabs Error message'), err);
    }
  }

/*

                await this.checkData("asm_dc","Custom",this.properties.teamName,termName)
                .then( (response:any)=> {
                  if(response.value.length>0){
                    console.log("checkData",response);
                  }
                })

  private async checkData(dcName:string,library:string,team:string,category:string):Promise<ISPLists> {
    
    let requestUrl = '';

    if(category === ''){
      requestUrl=`https://${this.properties.tenantURL[2]}/sites/${dcName}/_api/web/lists/GetByTitle('${library}')/items?$select=*,FieldValuesAsText/OData__x005f_ModerationStatus,TaxCatchAll/Term&$expand=FieldValuesAsText/OData__x005f_ModerationStatus,TaxCatchAll/Term&$filter=TaxCatchAll/DC_Team eq '${team}'&$top=10`;
    }else{
      requestUrl=`https://${this.properties.tenantURL[2]}/sites/${dcName}/_api/web/lists/GetByTitle('${library}')/items?$select=*,FieldValuesAsText/OData__x005f_ModerationStatus,TaxCatchAll/Term&$expand=FieldValuesAsText/OData__x005f_ModerationStatus,TaxCatchAll/Term&$filter=TaxCatchAll/DC_SharedWith eq '${category}'&$top=10`;
    }
    
    return this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response : SPHttpClientResponse) => {
        return response.json();
      //}).catch(err => {
      //  console.log("checkdata ERROR:",err.response.data);
      });   
  }
*/

  private async _getDataAsync(flag:boolean,library:string,category:string): Promise<void> {
    console.log('getDataAsync',flag,library,category);

    const dcDivisions : string[] = ["asm","cen","cnn","emp","hea"];

    if(library===""){
      this.properties.libraryName = ["Policies", "Procedures","Guides", "Forms", "General"];
    }else{
      this.properties.libraryName = [library];
    }

    for (let x = 0; x < this.properties.libraryName.length; x++) {
      dcDivisions.forEach(async (site,index)=>{
        await this._getData(flag,site,this.properties.libraryName[x],this.properties.teamName,category)
          .then(async (response) => {
            if(response.length>0){
              console.log("getDataAsync response",category,response);
              if(!flag){
                //if(category===""){
                  //await this._setLibraryTabs(this.properties.libraryName[x]);
                //}else{
                //  await this._setLibraryTabs(category);
                  console.log("renderTab",category);
                //}   
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
    console.log('getdata info',library,category);

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
    }

    return webDC.lists.getByTitle(library)
      .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
      .then(async (response) => {
        return response;  
      })
      .catch(() => {});    
  }

/*
  private async _setLibraryTabs(library: string): Promise<void>{
    console.log("setLibrary",library,this.properties.libraryNamePrev,libCount);
    
    if(this.properties.libraryNamePrev !== library){
      this.properties.libraryNamePrev = library;
      this.properties.libraries[libCount] = library;
      libCount++;
    }

    return;
  }
*/

  public async onInit(): Promise<void> {
    await super.onInit();
    this.properties.libraries = [];
    this.properties.libraryName = [];

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
                }),
                PropertyPaneDropdown("divisionDC", {
                  label: "Please Choose Document Centre",
                  options: [
                    { key: "All", text: "All Document Centres" },
                    { key: "ASM", text: "Assessments DC Only" },
                    { key: "CEN", text: "Central DC Only" },
                    { key: "CNN", text: "Connect DC Only" },
                    { key: "EMP", text: "Employability DC Only" },
                    { key: "HEA", text: "Health DC Only" },
                  ],
                }),
                PropertyPaneHorizontalRule(),                
                PropertyPaneDropdown('division',{
                  label:"Please Choose Your Division",
                  options:[
                    { key : 'Assessments', text : 'Assessments'},
                    { key : 'Central', text : 'Central'},
                    { key : 'Connect', text : 'Connect'},
                    { key : 'Employability', text : 'Employability'},
                    { key : 'Health', text : 'Health'},
                  ]
                }),
                PropertyPaneHorizontalRule(),                
                PropertyPaneDropdown('teamName',{
                  label:"Please Choose Your Team",
                  options:[
                    { key : 'Process Design', text : 'Process Design'},
                    { key : 'IPES Wales', text : 'IPES Wales'},
                    { key : 'Clinical', text : 'Clinical'},
                    { key : 'Central Operations Support', text : 'Central Operations Support'},
                    { key : '', text : 'Health'},
                  ]
                })                                
              ]
            }
          ]
        }
      ]
    };
  }
}

/*
  private async _renderlibraryTabsAsync(category:string): Promise<void> {
    //if(this.properties.libraries !== undefined){
      this.properties.libraries.sort()
    //}
      for(let x=0; x<this.properties.libraries.length;x++){     
        console.log("libraryTabsAsync",this.properties.libraries[x],x);
        //await this._renderLibraryTabs(this.properties.libraries[x],category).then( async ()=> {
          //this._setLibraryListeners();
          // *** get custom tabs from termstore and add library column
          //await this.renderCustomTabsAsync();              
        //});  
      }
    return;
  }

  private async _renderLibraryTabs(library:string,labelName:string): Promise<void> {
    console.log('renderLibraryTabs');

    //const dataTarget:string=library.toLowerCase();
    let html: string = '';

    //console.log("renderlist",library);
    html = `<button class="btn btn-primary text-center mb-1" id="${library}_btn" type="button"><h6 class="libraryText">${library}</h6></button>`;

    if(this.domElement.querySelector('#libraryName') !== null){
      this.domElement.querySelector('#libraryName')!.innerHTML += html;
    }
    return;  
  }
*/