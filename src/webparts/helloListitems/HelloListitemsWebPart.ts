import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';
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
  private _runOperation(): void{ alert("Not Implemented")};

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
    debugger;
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
