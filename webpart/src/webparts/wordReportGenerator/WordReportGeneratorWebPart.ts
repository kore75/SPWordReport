import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WordReportGeneratorWebPartStrings';
import WordReportGenerator from './components/WordReportGenerator';
import { IWordReportGeneratorProps } from './components/IWordReportGeneratorProps';

import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { update, get } from '@microsoft/sp-lodash-subset';
import { SPDataService } from '../../service/SPDataService';
import { ISPDataService } from '../../service/ISPDataService';
import { SPFx, spfi } from '@pnp/sp';
import { IDocumentLibraryInformation } from '@pnp/sp/sites';
import { MockSPDataService } from '../../service/MockSPDataService';
import { ISpListInfo } from './ISpListInfo';


export interface IWordReportGeneratorWebPartProps {
  externalApiUrl: string;
  reportDocLib?: ISpListInfo;
  reportDocItem?: ISpListInfo;
  reportDocList?: ISpListInfo;
}

export default class WordReportGeneratorWebPart extends BaseClientSideWebPart<IWordReportGeneratorWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _itemsDropDown: PropertyPaneAsyncDropdown;
  private _dataService: ISPDataService;

  public render(): void {
   
    const element: React.ReactElement<IWordReportGeneratorProps> = React.createElement(
      WordReportGenerator,
      {
        externalApiUrl: this.properties.externalApiUrl,
        reportDocLib: this.properties.reportDocLib ,
        reportDocItem:this.properties.reportDocItem,
        reportDocList : this.properties.reportDocList,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        dataService:this._dataService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._dataService=new SPDataService(this.context);

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

  private onReportLibChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return {Id:newValue.key,Title:newValue.text} as ISpListInfo; });
    // reset selected item
    this.properties.reportDocItem = undefined;
    // store new value in web part properties
    update(this.properties, 'reportDocItem', (): any => { return this.properties.reportDocItem; });
    // refresh web part
    this.render();
    
      // reset selected values in item dropdown
    this._itemsDropDown.properties.selectedKey = this.properties.reportDocLib?.Id ?? "";
      // allow to load items         
    
    this._itemsDropDown.properties.disabled = false;
    this._itemsDropDown.render(); 
      
    // load items and re-render items dropdown
    
  }

  private onReportListChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return {Id:newValue.key,Title:newValue.text} as ISpListInfo; });
    // reset selected item    
    // store new value in web part properties
   
    this.render(); 
    // load items and re-render items dropdown
    
  }

  private async loadDocLists(): Promise<IDropdownOption[]> {
      
    

    var items=await this._dataService.loadSiteCollectionDocLibs();
    return items;

  }

  private async loadLists(): Promise<IDropdownOption[]> {
      
    var items=await  this._dataService.loadSiteCollectionLists();
    return items;

  }

  private onReportListItemChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return {Id:newValue.key,Title:newValue.text} as ISpListInfo; });
    // refresh web part
    this.render();
  }

  private async loadItems(): Promise<IDropdownOption[]> {
    if (!this.properties.reportDocLib) {
      // resolve to empty options since no list has been selected
      return [];
    }
    const wp: WordReportGeneratorWebPart = this;
   
    if(wp.properties.reportDocLib!=null)
    {
      return await this._dataService.loadItems(wp.properties.reportDocLib.Id);
    }
    return [];
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
     // reference to item dropdown needed later after selecting a list
     this._itemsDropDown = new PropertyPaneAsyncDropdown('reportDocItem', {
      label: strings.ReportDocLabel,
      loadOptions: this.loadItems.bind(this),
      onPropertyChange: this.onReportListItemChange.bind(this),
      selectedKey: this.properties.reportDocItem?.Id ?? "",
      // should be disabled if no list has been selected
      disabled: !this.properties.reportDocLib
    });
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
                PropertyPaneTextField('externalApiUrl', {
                  label: strings.ExternalApiUrl                  
                }),
                new PropertyPaneAsyncDropdown('reportDocLib', {
                  label: strings.ReportDocLibLabel,
                  loadOptions: this.loadDocLists.bind(this),                 
                  onPropertyChange: this.onReportLibChange.bind(this),
                  selectedKey: this.properties.reportDocLib?.Id ?? ""
                }),
                this._itemsDropDown                         
              ]
            },
            {
              groupName: strings.ReportList,
              groupFields: [                               
                new PropertyPaneAsyncDropdown('reportDocList', {
                  label: strings.ReportDocLibLabel,
                  loadOptions: this.loadLists.bind(this),                 
                  onPropertyChange: this.onReportListChange.bind(this),
                  selectedKey: this.properties.reportDocList?.Id ?? ""
                })                                  
              ]
            }
          ]
        }
      ]
    };
  }
}
