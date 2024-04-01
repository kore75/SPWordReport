import { IDropdownOption } from "office-ui-fabric-react";
import { ISPDataService, IWeatherData } from "./ISPDataService";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/sites";
import { IReportFileRequest } from "./IReportFileRequest";
import { IReportFileResult } from "./IReportFileResult";

const awaitTimeout = (delay:number) =>
  new Promise(resolve => setTimeout(resolve, delay));


export class MockSPDataService implements ISPDataService
{
    CreateReport(request: IReportFileRequest): Promise<IReportFileResult> {
      throw new Error("Method not implemented.");
    }
    GetWheatherData(): Promise<IWeatherData[]> {
      throw new Error("Method not implemented.");
    }
    async loadSiteCollectionDocLibs(): Promise<IDropdownOption[]> {

        await awaitTimeout(2000);
        const res:IDropdownOption[]=[{
            key: 'Dokumente',
            text: 'Dokumente'
         },
         {
            key: 'myDocuments',
            text: 'My Documents'
        }];
        return res;
                               
    }
    loadSiteCollectionLists(): Promise<IDropdownOption[]> {
        throw new Error("Method not implemented.");
    }
    async loadItems(splist: string): Promise<IDropdownOption[]> {
        await awaitTimeout(2000);
        const items :{[key:string]: IDropdownOption[]} ={
            "Dokumente": [
              {
                key: 'spfx_presentation.pptx',
                text: 'SPFx for the masses'
              },
              {
                key: 'hello-world.spapp',
                text: 'hello-world.spapp'
              }
            ],
            "myDocuments": [
              {
                key: 'isaiah_cv.docx',
                text: 'Isaiah CV'
              },
              {
                key: 'isaiah_expenses.xlsx',
                text: 'Isaiah Expenses'
              }
            ]
          };

        return items[splist];
    }

}