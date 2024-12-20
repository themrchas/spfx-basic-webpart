import { Version } from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneButton

} from '@microsoft/sp-webpart-base';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DynamicPropsWebPartWebPart.module.scss';
import * as strings from 'DynamicPropsWebPartWebPartStrings';

import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http'

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  ItemCount: number;
}

export interface IDynamicPropsWebPartWebPartProps {
  description: string;
  propertyPaneText2: string;
  propertyPaneChkBox: boolean;
  propertyPaneDropDown: string;
  propertyPaneToggle: boolean;
  propertyPaneButtonText: string;

}

export default class DynamicPropsWebPartWebPart extends BaseClientSideWebPart<IDynamicPropsWebPartWebPartProps> {

  public render(): void {

    console.log('this.context is',this.context);
    console.log("this.contex.pageContext is",this.context.pageContext);


    //this.properties.description = "Yes - money";

    console.log("The interface desciption variable is",this.properties.description);


    this.domElement.innerHTML = `
      <div class="${ styles.dynamicPropsWebPart }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint! ${this.context.pageContext.user.displayName}</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p>${escape(this.properties.propertyPaneText2)}</p>
              <p>${this.properties.propertyPaneChkBox}</p>
              <p>${escape(this.properties.propertyPaneDropDown)}</p>
              <p>${this.properties.propertyPaneToggle}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
              <p><button type='button' id='btn1'>Press</button></p>
              <p><button type='button' id='btn2'>Display lists</button></p>
              <p id='lists'></p>

            </div>
          </div>
        </div>
      </div>`;

      this.domElement.querySelector('#btn1').addEventListener('click',this.btnModifyPropertyPaneTextField.bind(this));
      this.domElement.querySelector('#btn2').addEventListener('click',this.btnDisplayListNames.bind(this));
      //this.getSiteLists();
  }

  private btnDisplayListNames(targetId: string) : void {

    console.log('btnDisplayListNames this',this);

    this.getSiteLists().then(
      
        //Used a non arrow function in order to be able to use 'this' referring to the 
        function (rawListData: ISPLists) {

        console.log('then: that is',this);

        console.log('rawListData is',rawListData);

        const htmlLists  = document.createElement('ul');

        for(let item of rawListData.value)  {

          let li = document.createElement('li');
          li.innerHTML = `<li><div style="display:flex;justify-content:space-between;width:200px"><div>${item.Title}</div><div>${item.ItemCount}</div>`
          htmlLists.appendChild(li);
          //htmlLists += `<li><div style="display:flex;justify-content:space-between;width:80px"><div>${item.Title}</div><div>${item.ItemCount}</div>`
      }

        const listContainer: Element = this.domElement.querySelector("#lists").appendChild(htmlLists);

       // this.render(); Not needed as appendChild is used

      }.bind(this)

    )


  } //

  private btnModifyPropertyPaneTextField() : void {
    console.log("btnModifyPropertyPaneText: 'this' ",this);
    this.properties.propertyPaneText2 = "Text created by button click";
    this.render();
  }

  
  private getSiteLists(): Promise<ISPLists> {

    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + '/_api/web/lists', SPHttpClient.configurations.v1)
    .then(async (response: SPHttpClientResponse) => {
      let result =  await response.json();
      console.log('lists are',result);

      let listInfo:ISPLists = { value: [] };

      for (let item of result.value) 
        listInfo.value.push({Title:item.Title,ItemCount:item.ItemCount })


      return listInfo;
    });

  } //getSiteLists 

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
                PropertyPaneTextField('propertyPaneText2',{
                  label: "Multi-line Text Field",
                  multiline: true
                }),
                PropertyPaneCheckbox('propertyPaneChkBox',{
                  text: 'Checkbox'
                }),
                PropertyPaneDropdown('propertyPaneDropDown',{
                  label: 'Dropdown',
                  options: [
                    { key: '1', text: 'One'},
                    { key: '2', text: 'Two'},
                  ]
                }),
                PropertyPaneToggle('propertyPaneToggle',{
                  label: 'Toggle',
                  onText: 'On',
                  offText: 'Off'
                }),

                /* Note that 'propertyPaneButton'is simply an internal identifier for the control used by SharePoint.
                * There is no inherent functionality to this parameter.
                */
                PropertyPaneButton('Property PaneButton Identifier', {
                  text:this.properties.propertyPaneButtonText,
                  onClick: () => { 
                    alert('PropertyPaneButton click with value');
                    this.properties.propertyPaneText2 = "Modified mulit-line text field using PropertyPaneButton"
                    this.render();
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
