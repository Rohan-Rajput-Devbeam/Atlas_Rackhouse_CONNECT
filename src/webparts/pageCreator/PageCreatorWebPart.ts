//import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";
import { ClientsideText, CreateClientsidePage, PromotedState } from "@pnp/sp/clientside-pages";
import { IClientsidePage } from "@pnp/sp/clientside-pages";
import "@pnp/sp/comments/clientside-page";
import "@pnp/sp/webs";

import { sp, ClientsidePageFromFile, Web } from "@pnp/sp/presets/all";

import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PageCreatorWebPart.module.scss';
import * as strings from 'PageCreatorWebPartStrings';


import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/lists";

import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";


// import PnPTelemetry from "@pnp/telemetry-js";

// const telemetry = PnPTelemetry.getInstance();
// telemetry.optOut();
export interface IPropertyControlsTestWebPartProps {
  filePickerResult: IFilePickerResult;
}

export interface IPageCreatorWebPartProps {
  people: IPropertyFieldGroupOrPerson[];

  filePickerResult: any;
  hyperlinkProperty: any;
  name: boolean;
  description: string;
  Template: any;
  Language: any;
  ImageProperty: string;
  PageTitle: any;
  ParentPageURL: string;
  button: string;
  flag: any;
  PageTitleTemp: any;
  CreateFolder: any;
  Saved: any;
  TempSaved: any;
  TimeStamp: any;
  PermanentTimeStamp: any;
  Executed: any;
  FolderUrl: any;
  RackID: any;
  delete: boolean;

  LangEnglish: any;
  LangChinese: any;
  LangGerman: any;
  LangSpanish: any;
  LangFrench: any;
  LangPolish: any;
  LangJapanese: any;
  LangPortuguese: any;
  LangRussian: any;

  EnglishText: any;
  ChineseText: any;
  GermanText: any;
  SpanishText: any;
  FrenchText: any;
  PolishText: any;
  JapaneseText: any;
  PortugueseText: any;
  RussianText: any;

  NewTabCheckBox: any;
  CurrentTabCheckBox: any;

}
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';

export interface IPropertyControlsTestWebPartProps {
  people: IPropertyFieldGroupOrPerson[];
}

export interface IPropertyControlsTestWebPartProps {
  filePickerResult: IFilePickerResult;
}


export default class PageCreatorWebPart extends BaseClientSideWebPart<IPageCreatorWebPartProps> {
  checkboxProperty1: string;


  public SaveButton(SaveButtonValue: any): any {
    this.properties.Saved = 1;
    // console.log("SAVE BUTTON CLICKED...");
    //Save TimeStamp value permanently:

    return this.properties.Saved;
  }

  public async DeleteButton(DeleteButtonValue: any): Promise<any> {
    var siteUrl = this.context.pageContext.web.absoluteUrl ///Get Site Url
    // console.log(siteUrl)

    const myArray = siteUrl.split("/");
    let siteName = myArray[myArray.length - 1].split(".")[0]; ///Get Site Name
    // console.log(siteName)
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, ":Please wait while your Rackhouse is being deleted. Do not close the window or tab");

    const page = await ClientsidePageFromFile(sp.web.getFileByServerRelativePath(`/sites/${siteName}/SitePages/` + this.properties.RackID + ".aspx"));
    await page.delete().then(function (response) {
      console.log("File Deleted !");
    }).catch(function (exception) {
      alert("Something went wrong :" + exception);
    });
    await sp.web.getFolderByServerRelativePath("Rackhouse Documents/" + this.properties.RackID).delete().then(function (response) {
      console.log("Rackhouse Documents- Folder Deleted !");
    }).catch(function (exception) {
      alert("Something went wrong :" + exception);
    });

    await sp.web.getFolderByServerRelativePath("Rackhouse Archive/" + this.properties.RackID).delete().then(function (response) {
      console.log("Rackhouse Archive - Folder Deleted !");
    }).catch(function (exception) {
      alert("Something went wrong :" + exception);
    });

    this.properties.delete = true;
    this.properties.ParentPageURL = null;
    this.properties.ImageProperty = null;
    this.properties.PageTitle = null;

    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.render();

    return
  }



  public simpleTextBoxValidationMethod(value: string): string {
    if (value.length == 0 || value.length < 7) {
      return "Please provide a hyperlink!";
    } else {
      return "";
    }
  }

  public async render(): Promise<void> {
    var siteUrl = this.context.pageContext.web.absoluteUrl ///Get Site Url
    // console.log(siteUrl)

    const myArray = siteUrl.split("/");
    let siteName = myArray[myArray.length - 1].split(".")[0]; ///Get Site Name
    // console.log(siteName)


    //If link is selected as a templete then the hyperlink will be the one that user enters in the hyperlink box....
    if (this.properties.Template != "Link" && (this.properties.ParentPageURL != null || this.properties.ParentPageURL != "")) {

      this.properties.ParentPageURL = `${this.context.pageContext.web.absoluteUrl}/SitePages/` + this.properties.RackID + ".aspx";

    }


    //---->>Set User Language based on user preference .....
    var userEmail = this.context.pageContext.user.email;
    this.context.spHttpClient.get(`${siteUrl}/_api/web/lists/getbytitle('Preference')/Items?&$filter=Title eq '${userEmail}'`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          // console.log(responseJSON.value);
          var prefLanguage = responseJSON.value.map(function (obj: { Language: any; }) {
            return obj.Language;
          });
          // console.log(prefLanguage)




          this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/Web/CurrentUser/Groups`,
            SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
              response.json().then(async (responseJSON: any) => {
                // console.log(responseJSON.value);
                var finalArray = responseJSON.value.map(function (obj: { Title: any; }) {
                  return obj.Title;
                });
                ///console.log(finalArray);//Array Retrieved from Current users Groups.....

                if (this.properties.people && this.properties.people.length > 0) {
                  ///console.log(JSON.stringify(this.properties.people));

                  const GroupArray = this.properties.people.map((obj: { fullName: any; }) => {
                    return obj.fullName;
                  });
                  var usrFullname = this.context.pageContext.user.displayName;
                  var Groupintersections = finalArray.filter(e => GroupArray.indexOf(e) !== -1);
                  // console.log(Groupintersections)

                  ///console.log(GroupArray);//Array Of Group in property pane
                  if (GroupArray.includes(usrFullname) || Groupintersections.length > 0) {
                    // console.log("Current User Present In The Group");



                    if (this.properties.delete == true) {

                      this.domElement.innerHTML = `
      <head>
      <link rel="preconnect" href="https://fonts.googleapis.com">
      <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
      <link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap" rel="stylesheet">
      </head>
      <div id="LoaderId">
      <script src='https://kit.fontawesome.com/a076d05399.js' crossorigin='anonymous'></script>
        
          <div class="ms-rte-embedcode ms-rte-embedwp" >
          <div class="${styles.MainContainer}"
          style="background-image: url(${escape(this.properties.ImageProperty)});
          box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
          background-repeat: no-repeat;width:100%;height:160px;
          background-size:cover;text-align: center;
          background-position: center;">
          <h1>Rackhouse Deleted</h1>
          <h5 style="color: #cc0a0a;">Please delete this webpart!</h5>
          
        </div></div></div>
        `;

                    }
                    else {



                      //Test Run--->
                      //var currentPageUrl = this.context.pageContext.web.absoluteUrl + this.context.pageContext.site.serverRequestPath;

                      // console.log(currentPageUrl);
                      // //Get the current timestamp.....
                      // this.properties.TimeStamp = new Date().getTime();

                      sp.setup({
                        spfxContext: this.context
                      });

                      try {
                        // Set Image URL received from the file picker component--->
                        const myObj = (this.properties.filePickerResult);
                        // console.log(myObj.fileAbsoluteUrl);
                        this.properties.ImageProperty = myObj.fileAbsoluteUrl;
                      }
                      catch (err) {

                      }

                      this.domElement.innerHTML = `
    <head>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap" rel="stylesheet">
</head>
<div id="LoaderId">
<script src='https://kit.fontawesome.com/a076d05399.js' crossorigin='anonymous'></script>
  
    <div class="ms-rte-embedcode ms-rte-embedwp" >
    <div class="${styles.MainContainer}"
    style="background-image: url(${escape(this.properties.ImageProperty)});
    box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
    background-repeat: no-repeat;width:100%;height:160px;
    background-size:cover;
    background-position: center;">

    <a  class="${styles.callToAction}" onMouseOver="this.style.color='#FFFFFF'; style.backgroundColor='#565656E6'" onMouseOut="this.style.color='#FFFFFF'; style.backgroundColor='#565656BF'" 
    
    style="
    display: block;
    font-family: 'Oswald' !important;
    float: left;
    overflow-wrap: break-word;
    width:80%;
    position: absolute;
    background: #565656BF;
    margin-top: 1.5em;
    //vertical-align: middle;
    text-align: right;
    text-decoration: none;
   font-size: 25px;
    padding: 0.25em 0.5em 0.25em calc(2% + 0em);
    color: #FFFFFF;
    text-transform: uppercase;" href="${escape(this.properties.ParentPageURL)}" target="_blank" unselectable="on" >
    
    ${prefLanguage[0].includes("English") && this.properties.LangEnglish == true ?
                          this.properties.EnglishText :
                          prefLanguage[0].includes("Chinese") && this.properties.LangChinese == true ?
                            this.properties.ChineseText :
                            prefLanguage[0].includes("German") && this.properties.LangGerman == true ?
                              this.properties.GermanText :
                              prefLanguage[0].includes("Spanish") && this.properties.LangSpanish == true ?
                                this.properties.SpanishText :
                                prefLanguage[0].includes("French") && this.properties.LangFrench == true ?
                                  this.properties.FrenchText :
                                  prefLanguage[0].includes("Polish") && this.properties.LangPolish == true ?
                                    this.properties.PolishText :
                                    prefLanguage[0].includes("Japanese") && this.properties.LangJapanese == true ?
                                      this.properties.JapaneseText :
                                      prefLanguage[0].includes("Portuguese") && this.properties.LangPortuguese == true ?
                                        this.properties.PortugueseText :
                                        prefLanguage[0].includes("Russian") && this.properties.LangRussian == true ?
                                          this.properties.RussianText :
                                          `${escape(this.properties.PageTitle)}`

                        }


    <i style="
    border: solid #FFFFFF;
    font-color: #FFFFFF;    
    border-width: 0 4px 4px 0;
    display: inline-block;
    padding: 7px;
    height:7px; width:7px;
    transform: rotate(-45deg);
    -webkit-transform: rotate(-45deg);">
    </i></a>
    
  </div></div></div>
  `;
                      //  console.log("Saved-"+this.properties.Saved);

                      if ((this.properties.Saved != 2) && (this.properties.Saved != 0)) {
                        //Get the current timestamp.....
                        this.properties.TimeStamp = new Date().getTime();

                        //Load the webpart on clicking save...
                        this.context.statusRenderer.displayLoadingIndicator(this.domElement, ":Please wait while your Rackhouse is being created. Do not close the window or tab");

                        console.log("Folder & Page Creation Started.....");

                        //Creation of A new Folder on saving the webpart...
                        const folderAddResult = await sp.web.rootFolder.folders.getByName("Rackhouse Documents").folders.add("Rack" + this.properties.TimeStamp);

                           //Creation of A new Folder in Rackhouse Archive Library saving the webpart...
                           const folderAddResult1 = await sp.web.rootFolder.folders.getByName("Rackhouse Archive").folders.add("Rack" + this.properties.TimeStamp);

                        //------------>
                        //console.log(folderAddResult.data.ServerRelativeUrl);
                        var folderurl = folderAddResult.data.ServerRelativeUrl;// Get created folders url...

                        console.log("Created page relative url-" + folderurl)



                        //Copy Template Layout and Create a site page...
                        const page = await ClientsidePageFromFile(sp.web.getFileByServerRelativePath(`/sites/${siteName}/SitePages/DND/` + this.properties.Template + ".aspx"));
                        console.log("Selected Page Templete-" + this.properties.Template);
                        // console.log(page);

                        // Create a published copy of the page
                        const pageCopy = await page.copy(sp.web, "Rack" + this.properties.TimeStamp, this.properties.PageTitle, false);
                        // pageCopy.addSection().addControl(new ClientsideText("https://devbeam.sharepoint.com" + folderurl)); //Add Text Section--->

                        // console.log(folderurl);

                        // turn off comments
                        await pageCopy.disableComments();
                        await pageCopy.save();


                        console.log("New Page Created By Title: " + pageCopy.title);
                        const myArray2 = await folderurl.split("/");
                        let rackID = myArray2[myArray2.length - 1].split(".")[0]; ///Get RackID
                        console.log(rackID)



                        // const pageCopy1 = await CreateClientsidePage(Web("/SitePages/Rackhouse Pages"),this.properties.PageTitle, this.properties.PageTitle);
                        // console.log("New Page Created By Title: " + pageCopy1.title);

                        this.properties.ParentPageURL = `${this.context.pageContext.web.absoluteUrl}/SitePages/` + "Rack" + this.properties.TimeStamp + ".aspx";
                        this.properties.delete = false;
                        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
                        this.properties.RackID = "Rack" + this.properties.TimeStamp;
                        console.log("RackID=" + this.properties.RackID)


                        this.domElement.innerHTML = `
    <head>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap" rel="stylesheet">
</head>
<div id="LoaderId">
<script src='https://kit.fontawesome.com/a076d05399.js' crossorigin='anonymous'></script>
  
    <div class="ms-rte-embedcode ms-rte-embedwp" >
    <div class="${styles.MainContainer}"
    style="background-image: url(${escape(this.properties.ImageProperty)});
    background-repeat: no-repeat;width:100%;height:160px;
    background-size:cover;
    background-position: center;">

    <a  class="${styles.callToAction}" onMouseOver="this.style.color='#FFFFFF'; style.backgroundColor='#565656E6'" onMouseOut="this.style.color='#FFFFFF'; style.backgroundColor='#565656BF'" 
    
    style="
    display: block;
    font-family: 'Oswald' !important;
    float: left;
    overflow-wrap: break-word;
    width:80%;
    position: absolute;
    background: #565656BF;
    margin-top: 1.5em;
    //vertical-align: middle;
    text-align: right;
    text-decoration: none;
   font-size: 25px;
    padding: 0.25em 0.5em 0.25em calc(2% + 0em);
    color: #FFFFFF;
    text-transform: uppercase;" href="${escape(this.properties.ParentPageURL)}" target="_blank" unselectable="on" >
   
    ${prefLanguage[0].includes("English") && this.properties.LangEnglish == true ?
                            this.properties.EnglishText :
                            prefLanguage[0].includes("Chinese") && this.properties.LangChinese == true ?
                              this.properties.ChineseText :
                              prefLanguage[0].includes("German") && this.properties.LangGerman == true ?
                                this.properties.GermanText :
                                prefLanguage[0].includes("Spanish") && this.properties.LangSpanish == true ?
                                  this.properties.SpanishText :
                                  prefLanguage[0].includes("French") && this.properties.LangFrench == true ?
                                    this.properties.FrenchText :
                                    prefLanguage[0].includes("Polish") && this.properties.LangPolish == true ?
                                      this.properties.PolishText :
                                      prefLanguage[0].includes("Japanese") && this.properties.LangJapanese == true ?
                                        this.properties.JapaneseText :
                                        prefLanguage[0].includes("Portuguese") && this.properties.LangPortuguese == true ?
                                          this.properties.PortugueseText :
                                          prefLanguage[0].includes("Russian") && this.properties.LangRussian == true ?
                                            this.properties.RussianText :
                                            `${escape(this.properties.PageTitle)}`

                          }


    <i style="
    border: solid #FFFFFF;
    font-color: #FFFFFF;    
    border-width: 0 4px 4px 0;
    display: inline-block;
    padding: 7px;
    height:7px; width:7px;
    transform: rotate(-45deg);
    -webkit-transform: rotate(-45deg);">
    </i></a>
    
  </div></div></div>
  `;
                        this.properties.Saved = 0;

                      }


                      else {
                        // var PageTitleTemp = this.properties.PageTitle;
                        //console.log("PageTitle-->" + PageTitleTemp);
                        console.log("Click on save to create a Folder and Page...");
                      }

                    }
                  }
                  else {
                    this.domElement.innerHTML = `
              <div><h5>Permission required to view this webpart!</h5></div>
            `;
                  }
                }
                else {

                  if (this.properties.delete == true) {

                    this.domElement.innerHTML = `
    <head>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap" rel="stylesheet">
    </head>
    <div id="LoaderId">
    <script src='https://kit.fontawesome.com/a076d05399.js' crossorigin='anonymous'></script>
      
        <div class="ms-rte-embedcode ms-rte-embedwp" >
        <div class="${styles.MainContainer}"
        style="background-image: url(${escape(this.properties.ImageProperty)});
        box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
        background-repeat: no-repeat;width:100%;height:160px;
        background-size:cover;text-align: center;
        background-position: center;">
        <h1>Rackhouse Deleted</h1>
        <h5 style="color: #cc0a0a;">Please delete this webpart!</h5>
        
      </div></div></div>
      `;

                  }
                  else {



                    //Test Run--->
                    //var currentPageUrl = this.context.pageContext.web.absoluteUrl + this.context.pageContext.site.serverRequestPath;

                    // console.log(currentPageUrl);
                    //Get the current timestamp.....
                    // this.properties.TimeStamp = new Date().getTime();

                    sp.setup({
                      spfxContext: this.context
                    });

                    try {
                      // Set Image URL received from the file picker component--->
                      const myObj = (this.properties.filePickerResult);
                      // console.log(myObj.fileAbsoluteUrl);
                      this.properties.ImageProperty = myObj.fileAbsoluteUrl;
                    }
                    catch (err) {

                    }

                    this.domElement.innerHTML = `
  <head>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap" rel="stylesheet">
</head>
<div id="LoaderId">
<script src='https://kit.fontawesome.com/a076d05399.js' crossorigin='anonymous'></script>

  <div class="ms-rte-embedcode ms-rte-embedwp" >
  <div class="${styles.MainContainer}"
  style="background-image: url(${escape(this.properties.ImageProperty)});
  box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
  background-repeat: no-repeat;width:100%;height:160px;
  background-size:cover;
  background-position: center;">

  <a  class="${styles.callToAction}" onMouseOver="this.style.color='#FFFFFF'; style.backgroundColor='#565656E6'" onMouseOut="this.style.color='#FFFFFF'; style.backgroundColor='#565656BF'" 
  
  style="
  display: block;
  font-family: 'Oswald' !important;
  float: left;
  overflow-wrap: break-word;
  width:80%;
  position: absolute;
  background: #565656BF;
  margin-top: 1.5em;
  //vertical-align: middle;
  text-align: right;
  text-decoration: none;
 font-size: 25px;
  padding: 0.25em 0.5em 0.25em calc(2% + 0em);
  color: #FFFFFF;
  text-transform: uppercase;" href="${escape(this.properties.ParentPageURL)}" target="_blank" unselectable="on" >

  ${prefLanguage[0].includes("English") && this.properties.LangEnglish == true ?
                        this.properties.EnglishText :
                        prefLanguage[0].includes("Chinese") && this.properties.LangChinese == true ?
                          this.properties.ChineseText :
                          prefLanguage[0].includes("German") && this.properties.LangGerman == true ?
                            this.properties.GermanText :
                            prefLanguage[0].includes("Spanish") && this.properties.LangSpanish == true ?
                              this.properties.SpanishText :
                              prefLanguage[0].includes("French") && this.properties.LangFrench == true ?
                                this.properties.FrenchText :
                                prefLanguage[0].includes("Polish") && this.properties.LangPolish == true ?
                                  this.properties.PolishText :
                                  prefLanguage[0].includes("Japanese") && this.properties.LangJapanese == true ?
                                    this.properties.JapaneseText :
                                    prefLanguage[0].includes("Portuguese") && this.properties.LangPortuguese == true ?
                                      this.properties.PortugueseText :
                                      prefLanguage[0].includes("Russian") && this.properties.LangRussian == true ?
                                        this.properties.RussianText :
                                        `${escape(this.properties.PageTitle)}`

                      }


  <i style="
  border: solid #FFFFFF;
  font-color: #FFFFFF;    
  border-width: 0 4px 4px 0;
  display: inline-block;
  padding: 7px;
  height:7px; width:7px;
  transform: rotate(-45deg);
  -webkit-transform: rotate(-45deg);">
  </i></a>
  
</div></div></div>
`;
                    //  console.log("Saved-"+this.properties.Saved);

                    if ((this.properties.Saved != 2) && (this.properties.Saved != 0)) {
                      this.properties.TimeStamp = new Date().getTime();

                      //Load the webpart on clicking save...
                      this.context.statusRenderer.displayLoadingIndicator(this.domElement, ":Please wait while your Rackhouse is being created. Do not close the window or tab");

                      console.log("Folder & Page Creation Started.....");

                      //Creation of A new Folder on saving the webpart...
                      const folderAddResult = await sp.web.rootFolder.folders.getByName("Rackhouse Documents").folders.add("Rack" + this.properties.TimeStamp);

                       //Creation of A new Folder in Rackhouse Archive Library saving the webpart...
                       const folderAddResult1 = await sp.web.rootFolder.folders.getByName("Rackhouse Archive").folders.add("Rack" + this.properties.TimeStamp);

                      //------------>
                      //console.log(folderAddResult.data.ServerRelativeUrl);
                      var folderurl = folderAddResult.data.ServerRelativeUrl;// Get created folders url...

                      console.log("Created page relative url-" + folderurl)


                      //Copy Template Layout and Create a site page...
                      const page = await ClientsidePageFromFile(sp.web.getFileByServerRelativePath(`/sites/${siteName}/SitePages/DND/` + this.properties.Template + ".aspx"));
                      console.log("Selected Page Templete-" + this.properties.Template);
                      // console.log(page);

                      // Create a published copy of the page
                      const pageCopy = await page.copy(sp.web, "Rack" + this.properties.TimeStamp, this.properties.PageTitle, false);
                      // pageCopy.addSection().addControl(new ClientsideText("https://devbeam.sharepoint.com" + folderurl)); //Add Text Section--->

                      // console.log(folderurl);

                      // turn off comments
                      await pageCopy.disableComments();
                      await pageCopy.save();


                      console.log("New Page Created By Title: " + pageCopy.title);


                      // const pageCopy1 = await CreateClientsidePage(Web("/SitePages/Rackhouse Pages"),this.properties.PageTitle, this.properties.PageTitle);
                      // console.log("New Page Created By Title: " + pageCopy1.title);

                      this.properties.ParentPageURL = `${this.context.pageContext.web.absoluteUrl}/SitePages/` + "Rack" + this.properties.TimeStamp + ".aspx";
                      this.properties.delete = false;
                      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
                      this.properties.RackID = "Rack" + this.properties.TimeStamp;
                      console.log("RackID=" + this.properties.RackID)

                      this.domElement.innerHTML = `
  <head>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap" rel="stylesheet">
</head>
<div id="LoaderId">
<script src='https://kit.fontawesome.com/a076d05399.js' crossorigin='anonymous'></script>

  <div class="ms-rte-embedcode ms-rte-embedwp" >
  <div class="${styles.MainContainer}"
  style="background-image: url(${escape(this.properties.ImageProperty)});
  background-repeat: no-repeat;width:100%;height:160px;
  background-size:cover;
  background-position: center;">

  <a  class="${styles.callToAction}" onMouseOver="this.style.color='#FFFFFF'; style.backgroundColor='#565656E6'" onMouseOut="this.style.color='#FFFFFF'; style.backgroundColor='#565656BF'" 
  
  style="
  display: block;
  font-family: 'Oswald' !important;
  float: left;
  overflow-wrap: break-word;
  width:80%;
  position: absolute;
  background: #565656BF;
  margin-top: 1.5em;
  //vertical-align: middle;
  text-align: right;
  text-decoration: none;
 font-size: 25px;
  padding: 0.25em 0.5em 0.25em calc(2% + 0em);
  color: #FFFFFF;
  text-transform: uppercase;" href="${escape(this.properties.ParentPageURL)}" target="_blank" unselectable="on" >
  ${prefLanguage[0].includes("English") && this.properties.LangEnglish == true ?
                          this.properties.EnglishText :
                          prefLanguage[0].includes("Chinese") && this.properties.LangChinese == true ?
                            this.properties.ChineseText :
                            prefLanguage[0].includes("German") && this.properties.LangGerman == true ?
                              this.properties.GermanText :
                              prefLanguage[0].includes("Spanish") && this.properties.LangSpanish == true ?
                                this.properties.SpanishText :
                                prefLanguage[0].includes("French") && this.properties.LangFrench == true ?
                                  this.properties.FrenchText :
                                  prefLanguage[0].includes("Polish") && this.properties.LangPolish == true ?
                                    this.properties.PolishText :
                                    prefLanguage[0].includes("Japanese") && this.properties.LangJapanese == true ?
                                      this.properties.JapaneseText :
                                      prefLanguage[0].includes("Portuguese") && this.properties.LangPortuguese == true ?
                                        this.properties.PortugueseText :
                                        prefLanguage[0].includes("Russian") && this.properties.LangRussian == true ?
                                          this.properties.RussianText :
                                          `${escape(this.properties.PageTitle)}`

                        }
  


  <i style="
  border: solid #FFFFFF;
  font-color: #FFFFFF;    
  border-width: 0 4px 4px 0;
  display: inline-block;
  padding: 7px;
  height:7px; width:7px;
  transform: rotate(-45deg);
  -webkit-transform: rotate(-45deg);">
  </i></a>
  
</div></div></div>
`;
                      this.properties.Saved = 0;

                    }


                    else {
                      // var PageTitleTemp = this.properties.PageTitle;
                      //console.log("PageTitle-->" + PageTitleTemp);
                      console.log("Click on save to create a Folder and Page...");
                    }

                  }
                }
              });
            });

        });
      });
  }



  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let MySaveButton: any;
    let deleteButton: any;
    let hyperlinkProperty: any;
    let EnglishProperty: any;
    let ChineseProperty: any;
    let GermanProperty: any;
    let SpanishProperty: any;
    let FrenchProperty: any;
    let PolishProperty: any;
    let JapaneseProperty: any;
    let PortugueseProperty: any;
    let RussianProperty: any;

    let NewTabCheckBox: any;
    let CurrentTabCheckBox: any;
    let PropertyPageID: any;


    // console.log(this.properties.NewTabCheckBox)

    if (this.properties.Template == "Link") {
      MySaveButton = PropertyPaneLabel('', {
        text: "",
      });

      PropertyPageID = ""

      hyperlinkProperty = PropertyPaneTextField('hyperlinkProperty', {
        label: 'Hyperlink',
        value: this.properties.hyperlinkProperty,
        // errorMessage: this.validateDescription.bind(this),
        // description: 'Link should start with "https://".',
        onGetErrorMessage: this.simpleTextBoxValidationMethod,
        deferredValidationTime: 1000,
        placeholder: "Link should start with https://"

      });

      this.properties.ParentPageURL = this.properties.hyperlinkProperty


    }
    else {

      hyperlinkProperty = ""
      this.properties.hyperlinkProperty = "" // Reset Hyperlink field.....

      if ((this.properties.Saved == 0) || (this.properties.Saved == 1) || (this.properties.Template == null)) {

        MySaveButton = PropertyPaneButton('Button', {
          disabled: true,
          text: "Create Page",
          buttonType: PropertyPaneButtonType.Normal,
          onClick: this.SaveButton.bind(this)
        });

        PropertyPageID = PropertyPaneTextField('RackID', {
          label: 'Property Page ID',
          value: this.properties.RackID
        });



      }
      else {
        MySaveButton = PropertyPaneButton('Button', {
          disabled: false,
          text: "Create Page",
          buttonType: PropertyPaneButtonType.Normal,
          onClick: this.SaveButton.bind(this)
        });

        PropertyPageID = ""


      }
    }

    if (this.properties.delete == true) {
      deleteButton = PropertyPaneButton('Button', {
        disabled: true,
        text: "Delete",
        buttonType: PropertyPaneButtonType.Primary,
        onClick: this.DeleteButton.bind(this)
      });
    }
    else {
      deleteButton = PropertyPaneButton('Button', {
        disabled: false,
        text: "Delete",
        buttonType: PropertyPaneButtonType.Primary,
        onClick: this.DeleteButton.bind(this)
      });

    }

    //////////////////////////

    if (this.properties.LangEnglish == true) {
      EnglishProperty = PropertyPaneTextField('EnglishText', {
        label: "",
        value: this.properties.EnglishText
      })
    }
    else {
      EnglishProperty = ""
    };
    ////////////////////////////////////////////////////////////
    if (this.properties.LangChinese == true) {
      ChineseProperty = PropertyPaneTextField('ChineseText', {
        label: "",
        value: this.properties.ChineseText
      })
    }
    else {
      ChineseProperty = ""
    };
    /////////////////////////////////////////////////////////////
    if (this.properties.LangGerman == true) {
      GermanProperty = PropertyPaneTextField('GermanText', {
        label: "",
        value: this.properties.GermanText
      })
    }
    else {
      GermanProperty = ""
    };
    ////////////////////////////////////////////////////////////
    if (this.properties.LangSpanish == true) {
      SpanishProperty = PropertyPaneTextField('SpanishText', {
        label: "",
        value: this.properties.SpanishText
      })
    }
    else {
      SpanishProperty = ""
    };
    ////////////////////////////////////////////////////////////
    if (this.properties.LangFrench == true) {
      FrenchProperty = PropertyPaneTextField('FrenchText', {
        label: "",
        value: this.properties.FrenchText
      })
    }
    else {
      FrenchProperty = ""
    };
    ///////////////////////////////////////////////////////////////
    if (this.properties.LangPolish == true) {
      PolishProperty = PropertyPaneTextField('PolishText', {
        label: "",
        value: this.properties.PolishText
      })
    }
    else {
      PolishProperty = ""
    };
    //////////////////////////////////////////////////////////////
    if (this.properties.LangJapanese == true) {
      JapaneseProperty = PropertyPaneTextField('JapaneseText', {
        label: "",
        value: this.properties.JapaneseText
      })
    }
    else {
      JapaneseProperty = ""
    };
    /////////////////////////////////////////////////////////////
    if (this.properties.LangPortuguese == true) {
      PortugueseProperty = PropertyPaneTextField('PortugueseText', {
        label: "",
        value: this.properties.PortugueseText
      })
    }
    else {
      PortugueseProperty = ""
    };
    //////////////////////////////////////////////////////////
    if (this.properties.LangRussian == true) {
      RussianProperty = PropertyPaneTextField('RussianText', {
        label: "",
        value: this.properties.RussianText
      })
    }
    else {
      RussianProperty = ""
    };
    ///////////////////////////////////////////////////////////




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
                PropertyPaneTextField('PageTitle', {
                  label: 'Page Title',
                  value: this.properties.PageTitle,
                  maxLength: 50,
                  description: 'Max Char length is 50.'
                }),
                PropertyPaneCheckbox('LangEnglish', {
                  text: "English",
                  checked: false,
                  disabled: false
                }),
                EnglishProperty,
                PropertyPaneCheckbox('LangChinese', {
                  text: "Chinese",
                  checked: false,
                  disabled: false
                }),
                ChineseProperty,
                PropertyPaneCheckbox('LangGerman', {
                  text: "German",
                  checked: false,
                  disabled: false
                }),
                GermanProperty,
                PropertyPaneCheckbox('LangSpanish', {
                  text: "Spanish",
                  checked: false,
                  disabled: false
                }),
                SpanishProperty,
                PropertyPaneCheckbox('LangFrench', {
                  text: "French",
                  checked: false,
                  disabled: false
                }),
                FrenchProperty,
                PropertyPaneCheckbox('LangPolish', {
                  text: "Polish",
                  checked: false,
                  disabled: false
                }),
                PolishProperty,
                PropertyPaneCheckbox('LangJapanese', {
                  text: "Japanese",
                  checked: false,
                  disabled: false
                }),
                JapaneseProperty,
                PropertyPaneCheckbox('LangPortuguese', {
                  text: "Portuguese",
                  checked: false,
                  disabled: false
                }),
                PortugueseProperty,
                PropertyPaneCheckbox('LangRussian', {
                  text: "Russian",
                  checked: false,
                  disabled: false
                }),
                RussianProperty,

                PropertyFieldFilePicker('filePicker', {
                  context: this.context,
                  filePickerResult: this.properties.filePickerResult,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: async (e: IFilePickerResult) => {
                    console.log(e);
                    console.log(e.downloadFileContent());
                    //for uploaded images
                    if (e.fileAbsoluteUrl == null) {
                      await e.downloadFileContent()
                        .then(async r => {
                          console.log(r, e)
                          let fileresult = await sp.web.getFolderByServerRelativeUrl("/sites/CONNECTII/SiteAssets/RackhouseImages/").files.addUsingPath(e.fileName.replace(/ /g, "_").replace(/\(|\)|\[|\]/g, "_"), r, { Overwrite: true });
                          e = { ...e, fileAbsoluteUrl: this.context.pageContext.web.absoluteUrl + fileresult.data.ServerRelativeUrl.substring(16) } //Will need to chane substring if Site name changes---->
                          this.properties.filePickerResult = e;

                        });
                    }
                    //for stock images/url/link images
                    else {
                      this.properties.filePickerResult = e;
                    }

                    console.log(this.properties.filePickerResult, e);

                  },
                  onChanged: (e: IFilePickerResult) => { this.properties.filePickerResult = e; },
                  key: "filePickerId",
                  buttonLabel: "Image Picker",
                  label: "Select Image",

                }),

                // PropertyPaneTextField('ImageProperty', {
                //   label: 'Select Image'
                // }),
                // PropertyPaneTextField('ParentPageURL', {
                //   label: 'Parent Page URL',
                //   value: this.properties.ParentPageURL
                // }),
                PropertyPaneDropdown('Template', {
                  label: 'Template',
                  options: [
                    { key: 'Library', text: 'Library' },
                    { key: 'People', text: 'People' },
                    { key: 'Location', text: 'Location' },
                    { key: 'Link', text: 'Link' }

                  ],
                  //  selectedKey: 'Library',
                }),
                hyperlinkProperty,

                PropertyFieldPeoplePicker('people', {
                  label: 'People Picker',
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'

                }),

                PropertyPageID,
                MySaveButton,
                deleteButton



              ]
            }
          ]
        }
      ]
    };
  }
}

