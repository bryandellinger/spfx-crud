import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloListitemsWebPart.module.scss';
import * as strings from 'HelloListitemsWebPartStrings';
import { ISPListItem } from "./ISPListitems";

export interface IHelloListitemsWebPartProps {
  description: string;
}

export default class HelloListitemsWebPart extends BaseClientSideWebPart<IHelloListitemsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloListitems }">
        <div class="${ styles.container }">
        <h3>List Items</h3>
        <ul></ul>
        <span class="${styles.label}">Select Operation</span>
        <select>
          <option value="Create">Create</option>
          <option value="Read">Read</option>
          <option value="Update">Update</option>
          <option value="Delete">Delete</option>
        </select>
        <button type='button' class='ms-Button'>
          <span class='ms-Button-label'>Run Operation</span>
        </button>
        <p>Select operation and click the button.</p>
        </div> 
      </div>`;

      this._itemsList = this.domElement.getElementsByTagName("UL")[0] as HTMLUListElement;
      this._operationSelect = this.domElement.getElementsByTagName("SELECT")[0] as HTMLSelectElement;
      this._operationResults = this.domElement.getElementsByTagName("P")[0] as HTMLParagraphElement;
      this._runOperation = this._runOperation.bind(this);
      this._readAllItems = this._readAllItems.bind(this);
      const button: HTMLButtonElement = this.domElement.getElementsByTagName("BUTTON")[0] as HTMLButtonElement;
      button.onclick = this._runOperation;
      this._readAllItems();

  }

  private _itemsList: HTMLUListElement = null;
  private _operationSelect: HTMLSelectElement = null;
  private _operationResults: HTMLParagraphElement = null;
  private _runOperation(): void{
     const operation: HTMLOptionElement =
      this._operationSelect[this._operationSelect.selectedIndex] as
      HTMLOptionElement;
     this._createListItem = this._createListItem.bind(this);
     this._readListItem = this._readListItem.bind(this);
     this._updateListItem = this._updateListItem.bind(this);
     this._deleteListItem = this._deleteListItem.bind(this);
     switch(operation.value){
       case "Create":
       this._createListItem();
       break;
       case "Read":
       this._readListItem();
       break;
       case "Update":
       this._updateListItem();
       break;
       case "Delete":
       this._deleteListItem();
       break;
     }
    };
  
  private _createListItem(): void {
    const url: string = this.context.pageContext.site.absoluteUrl+
    "/_api/web/lists/getbytitle('MyList')/items";
    const itemDefinition: any ={
      "Title": "SPFX created item",
      "Info": "Info Column Value"
    };
    const spHttpClientOptions : ISPHttpClientOptions = {
      "body": JSON.stringify(itemDefinition)
     };
     this.context.spHttpClient.post(url, SPHttpClient.configurations.v1,
      spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 201){
          this._operationResults.innerHTML ="Create: List item created successfully.";
          this._readAllItems();
        }else{
          this._operationResults.innerHTML ="Create: List item created failed." +
          response.status + " - " + response.statusText;
        }
      });
}

  private _readListItem(): void {
    const id: number = 1;
    this._getListItem(id).then(listItem => {
      this._operationResults.innerHTML = `
      <div>
        Read list item <br/>
        Title: ${listItem.Title} <br/>
        Info: ${listItem.Info}
      </div>
      `;
    })
    .catch(error =>{
      this._operationResults.innerHTML = "Read: Operation failed. " + error.message;
    });
   }

   private _getListItem(id: number) : Promise<ISPListItem>{
    const url: string = this.context.pageContext.site.absoluteUrl+
    "/_api/web/lists/getbytitle('MyList')/items?$select=Title,Id,Info&$filter=Id eq " + id;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    })
     .then((listItems: any ) => {
       const untypedItem: any = listItems.value[0];
       const listItem: ISPListItem = untypedItem as ISPListItem;
       return listItem;
     }) as Promise <ISPListItem>;
   }

   private _updateListItem(): void {
    const url: string = this.context.pageContext.site.absoluteUrl+
    "/_api/web/lists/getbytitle('MyList')/items(1)";
    const itemDefinition : any = {
      "Title": "Modified Title Field Value!"
    };
    const headers : any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers,
      "body": JSON.stringify(itemDefinition)
    };
    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if(response.status === 204){
        this._operationResults.innerHTML = "Update: List item updated successfully.";
      } else {
        this._operationResults.innerHTML = "Update: List Item update failed. "
        + response.status + " - " + response.statusText;
      }
    });
   }

   private _deleteListItem(): void {this._operationResults.innerHTML =
    "Delete: Not Implemented";
   }

  private _readAllItems(): void{
    this._getListItems().then(listItems => {
      let itemsStr: string = "";
      listItems.forEach(listItem =>{
        itemsStr += `<li>${listItem.Title}</li>`;
      })
      this._itemsList.innerHTML = itemsStr;
    });
  }

  private _getListItems(): Promise<ISPListItem[]>{
    const url: string = this.context.pageContext.site.absoluteUrl+
    "/_api/web/lists/getbytitle('MyList')/items";
    return this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    }).then(json =>{
      return json.value;
    }) as Promise<ISPListItem[]>
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
