import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField, 
  PropertyPaneCheckbox,  
  PropertyPaneToggle,  
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

export interface IHelloWorldWebPartProps {
  description: string;
  isChecked:boolean;
  dropdownfield1:string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
 
  public render(): void {
    let ddValue ="";
    if(this.properties.dropdownfield1 =="1")
    {
      ddValue = "item one";      
    }
    if(this.properties.dropdownfield1 =="2")
    {
      ddValue = "item two";      
    }
    if(this.properties.dropdownfield1 =="3")
    {
      ddValue = "item three";      
    }
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <p>Checkbox value from property pane: ${escape(this.properties.isChecked? "checked":"not checked")}</p>
              <p>drop down value from property pane: ${escape(ddValue)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'page 1' + strings.PropertyPaneDescription
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
        },
        {
          header: {
          description: "Page2"
          },
          groups: 
          [
            {
              groupName: "Group 2",
              groupFields: 
              [
                PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
                }),
                PropertyPaneCheckbox('isChecked',
                {
                text:"Checkbox",
                checked:true 
                })
              ]
            }
          ]
        },
        {
          header:{
          description:"Page 3"
          },
          groups:
          [
            {
              groupName : "Group 3",
              groupFields:
              [
                PropertyPaneDropdown('dropdownfield1',{
                label:"Dropdown",
                options:
                [
                  {key:1,text:"test 1"},
                  {key:2,text:"test 2"},
                  {key:3,text:"test 3"}
                ]
                })
              ]
            },
            {
              groupName:"Group 4  ",
              groupFields:
              [
                PropertyPaneToggle('toggle1',{
                label:"Toggle property"
                })
              ]
            }
          ]
        }


      ]
    };
  }
}
