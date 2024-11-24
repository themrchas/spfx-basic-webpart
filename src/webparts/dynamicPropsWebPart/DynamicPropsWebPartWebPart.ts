import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
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
  Id: string;
}

export interface IDynamicPropsWebPartWebPartProps {
  description: string;
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
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;

      this.getSiteLists();
  }

  
  private getSiteLists(): Promise<ISPLists> {

    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + '/_api/web/lists', SPHttpClient.configurations.v1)
    .then(async (response: SPHttpClientResponse) => {
      let result =  await response.json();
      console.log('lists are',result);
      return result;
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
