//import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";
import { ClientsideText, CreateClientsidePage, PromotedState } from "@pnp/sp/clientside-pages";

import { sp, ClientsidePageFromFile, Web } from "@pnp/sp/presets/all";



import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule,
  PropertyPaneTextField
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


  public SaveButton(SaveButtonValue: any): any {
    this.properties.Saved = 1;
    // console.log("SAVE BUTTON CLICKED...");
    //Save TimeStamp value permanently:

    return this.properties.Saved;
  }

  public async DeleteButton(DeleteButtonValue: any): Promise<any> {
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, ":Please wait while your Rackhouse is being deleted. Do not close the window or tab");

    const page = await ClientsidePageFromFile(sp.web.getFileByServerRelativePath("/sites/ModernConnect/SitePages/" + this.properties.RackID + ".aspx"));
    await page.delete().then(function (response) {
      console.log("File Deleted !");
    }).catch(function (exception) {
      alert("Something went wrong :" + exception);
    });
    await sp.web.getFolderByServerRelativePath("Rackhouse Documents/" + this.properties.RackID).delete().then(function (response) {
      console.log("Folder Deleted !");
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

  public async render(): Promise<void> {

    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/Web/CurrentUser/Groups`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then(async (responseJSON: any) => {
          // console.log(responseJSON.value);
          var finalArray = responseJSON.value.map(function (obj: { Title: any; }) {
            return obj.Title;
          });
          ///console.log(finalArray);//Array Retrieved from Current users Groups.....

          if (this.properties.people ) {
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
                //Get the current timestamp.....
                this.properties.TimeStamp = new Date().getTime();

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
    ${escape(this.properties.PageTitle)} 


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
                  //Load the webpart on clicking save...
                  this.context.statusRenderer.displayLoadingIndicator(this.domElement, ":Please wait while your Rackhouse is being created. Do not close the window or tab");

                  console.log("Folder & Page Creation Started.....");

                  //Creation of A new Folder on saving the webpart...
                  const folderAddResult = await sp.web.rootFolder.folders.getByName("Rackhouse Documents").folders.add("Rack" + this.properties.TimeStamp);
                  //------------>
                  //console.log(folderAddResult.data.ServerRelativeUrl);
                  var folderurl = folderAddResult.data.ServerRelativeUrl;// Get created folders url...

                  console.log("Created page relative url-" + folderurl)


                  //Copy Template Layout and Create a site page...
                  const page = await ClientsidePageFromFile(sp.web.getFileByServerRelativePath("/sites/ModernConnect/SitePages/DND/" + this.properties.Template + ".aspx"));
                  console.log("Selected Page Templete-" + this.properties.Template);
                  // console.log(page);

                  // Create a published copy of the page
                  const pageCopy = await page.copy(sp.web, "Rack" + this.properties.TimeStamp, this.properties.PageTitle, false);
                  // pageCopy.addSection().addControl(new ClientsideText("https://devbeam.sharepoint.com" + folderurl)); //Add Text Section--->

                  // console.log(folderurl);

                  await pageCopy.save();

                  console.log("New Page Created By Title: " + pageCopy.title);


                  // const pageCopy1 = await CreateClientsidePage(Web("/SitePages/Rackhouse Pages"),this.properties.PageTitle, this.properties.PageTitle);
                  // console.log("New Page Created By Title: " + pageCopy1.title);

                  this.properties.ParentPageURL = "https://devbeam.sharepoint.com/sites/ModernConnect/SitePages/" + "Rack" + this.properties.TimeStamp + ".aspx";
                  this.properties.delete = false;
                  this.context.statusRenderer.clearLoadingIndicator(this.domElement);
                  this.properties.RackID = "Rack" + this.properties.TimeStamp;

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
    ${escape(this.properties.PageTitle)} 


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
        });
      });
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let MySaveButton: any;
    let deleteButton: any;
    if ((this.properties.Saved == 0) || (this.properties.Saved == 1) || (this.properties.Template == null)) {
      MySaveButton = PropertyPaneButton('Button', {
        disabled: true,
        text: "Save",
        buttonType: PropertyPaneButtonType.Normal,
        onClick: this.SaveButton.bind(this)
      });
    }
    else {
      MySaveButton = PropertyPaneButton('Button', {
        disabled: false,
        text: "Save",
        buttonType: PropertyPaneButtonType.Normal,
        onClick: this.SaveButton.bind(this)
      });
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
                PropertyFieldFilePicker('filePicker', {
                  context: this.context,
                  filePickerResult: this.properties.filePickerResult,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => { this.properties.filePickerResult = e; },
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
                    { key: 'Event', text: 'Event' },
                    { key: 'People', text: 'People' },
                    { key: 'Location', text: 'Location' }
                  ],
                  //  selectedKey: 'Library',
                }),

                PropertyPaneDropdown('Language', {
                  label: 'Language',
                  options: [

                    { key: 'Arabic (Saudi Arabia)', text: 'Arabic (Saudi Arabia)' },
                    { key: 'Bulgarian (Bulgaria)', text: 'Bulgarian (Bulgaria)' },
                    { key: 'Chinese (Hong Kong S.A.R.)', text: 'Chinese (Hong Kong S.A.R.)' },
                    { key: 'Chinese (People s Republic of China)', text: 'Chinese (People s Republic of China)' },
                    { key: 'Chinese (Taiwan)', text: 'Chinese (Taiwan)' },
                    { key: 'Croatian (Croatia)', text: 'Croatian (Croatia)' },
                    { key: 'Czech (Czech Republic)', text: 'Czech (Czech Republic)' },
                    { key: 'Danish (Denmark)', text: 'Danish (Denmark)' },
                    { key: 'Dutch (Netherlands)', text: 'Dutch (Netherlands)' },
                    { key: 'English', text: 'English' },
                    { key: 'Estonian (Estonia)', text: 'Estonian (Estonia)' },
                    { key: 'Finnish (Finland)', text: 'Finnish (Finland)' },
                    { key: 'French (France)', text: 'French (France)' },
                    { key: 'German (Germany)', text: 'German (Germany)' },
                    { key: 'Greek (Greece)', text: 'Greek (Greece)' },
                    { key: 'Hebrew (Israel)', text: 'Hebrew (Israel)' },
                    { key: 'Hindi (India)', text: 'Hindi (India)' },
                    { key: 'Hungarian (Hungary)', text: 'Hungarian (Hungary)' },
                    { key: 'Indonesian (Indonesia)', text: 'Indonesian (Indonesia)' },
                    { key: 'Italian (Italy)', text: 'Italian (Italy)' },
                    { key: 'Japanese (Japan)', text: 'Japanese (Japan)' },
                    { key: 'Korean (Korea)', text: 'Korean (Korea)' },
                    { key: 'Latvian (Latvia)', text: 'Latvian (Latvia)' },
                    { key: 'Lithuanian (Lithuania)', text: 'Lithuanian (Lithuania)' },
                    { key: 'Malay (Malaysia)', text: 'Malay (Malaysia)' },
                    { key: 'Norwegian (Bokmal) (Norway)', text: 'Norwegian (Bokmal) (Norway)' },
                    { key: 'Polish (Poland)', text: 'Polish (Poland)' },
                    { key: 'Portuguese (Brazil)', text: 'Portuguese (Brazil)' },
                    { key: 'Portuguese (Portugal)', text: 'Portuguese (Portugal)' },
                    { key: 'Romanian (Romania)', text: 'Romanian (Romania)' },
                    { key: 'Russian (Russia)', text: 'Russian (Russia)' },
                    { key: 'Serbian (Latin) (Serbia)', text: 'Serbian (Latin) (Serbia)' },
                    { key: 'Slovak (Slovakia)', text: 'Slovak (Slovakia)' },
                    { key: 'Slovenian (Slovenia)', text: 'Slovenian (Slovenia)' },
                    { key: 'Spanish (Spain)', text: 'Spanish (Spain)' },
                    { key: 'Swedish (Sweden)', text: 'Swedish (Sweden)' },
                    { key: 'Thai (Thailand)', text: 'Thai (Thailand)' },
                    { key: 'Turkish (Turkey)', text: 'Turkish (Turkey)' },
                    { key: 'Ukrainian (Ukraine)', text: 'Ukrainian (Ukraine)' },
                    { key: 'Urdu (Islamic Republic of Pakistan)', text: 'Urdu (Islamic Republic of Pakistan)' },
                    { key: 'Vietnamese (Vietnam)', text: 'Vietnamese (Vietnam)' }
                  ],
                  selectedKey: 'English',
                }),
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

