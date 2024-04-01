
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { IReportFileRequest } from './IReportFileRequest';
import { IReportFileResult } from './IReportFileResult';

export interface IWeatherData{
    date:Date,
    temperaturec:number,
    summary:string,
    temperaturef:number
}

export interface  ISPDataService
{
    loadSiteCollectionDocLibs(): Promise<IDropdownOption[]>;
    loadSiteCollectionLists(): Promise<IDropdownOption[]>;
    loadItems(splist:string): Promise<IDropdownOption[]>;
    GetWheatherData(): Promise<IWeatherData[]>;
    CreateReport(request:IReportFileRequest):Promise<IReportFileResult>;
}