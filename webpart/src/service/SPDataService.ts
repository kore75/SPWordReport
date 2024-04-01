import { IDropdownOption } from "office-ui-fabric-react";
import { ISPDataService, IWeatherData } from "./ISPDataService";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/sites";
import "@pnp/sp/items/get-all";
import { IDocumentLibraryInformation } from "@pnp/sp/sites";
import { getSP } from './pnpjsConfig';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AadHttpClient, HttpClientResponse,IHttpClientOptions } from '@microsoft/sp-http';
import { IReportFileRequest } from "./IReportFileRequest";
import { IReportFileResult } from "./IReportFileResult";


 

export class SPDataService implements ISPDataService
{
    private _context : WebPartContext;
    private readonly _apiId = 'api://13e67028-6a75-4c76-98cb-84118d194216';

    public constructor(context : WebPartContext) {
        this._context= context;
       
    }
    async CreateReport(request: IReportFileRequest): Promise<IReportFileResult> {
        const bodyString:string=JSON.stringify(request);
        const spOpts: IHttpClientOptions = { 
            body: bodyString
        };

        const client=await this._context.aadHttpClientFactory.getClient(this._apiId);
        const response= await client.post('https://localhost:7282/ReportGenerator',AadHttpClient.configurations.v1,spOpts);
        let data:IReportFileResult = await response.json();
        return data;        
    }
    async loadSiteCollectionDocLibs(): Promise<IDropdownOption[]> {


        // within a webpart, application customizer, or adaptive card extension where the context object is available
        const sp = getSP(this._context);
        //const info = await sp.web.getContextInfo();
        const web = this._context.pageContext.web.absoluteUrl;
        const docLibs: IDocumentLibraryInformation[] = await sp.site.getDocumentLibraries(web);
        return docLibs.map<IDropdownOption>((item)=>{return {key: item.Id, text: item.Title }});
               
    }
    async loadSiteCollectionLists(): Promise<IDropdownOption[]> {
         // within a webpart, application customizer, or adaptive card extension where the context object is available
         const sp = getSP(this._context);
         const lists=await sp.web.lists.select("BaseTemplate","Id","Title")();
         const items=lists.filter((item)=>item.BaseTemplate==100).map<IDropdownOption>((item)=>{
            {return {key: item.Id, text: item.Title }}
         });
         return items;

    }
    async loadItems(splist: string): Promise<IDropdownOption[]> {
         // within a webpart, application customizer, or adaptive card extension where the context object is available
         const sp = getSP(this._context);
         const spList=await sp.web.lists.getById(splist);
         let result:IDropdownOption[]=[];
        
         const queryResult=await spList.select("BaseTemplate")();

         if (queryResult.BaseTemplate == 101) { // Document Library template ID is 101
            const items=await spList.items.select('File','Id').expand('File')();
            result=items.map<IDropdownOption>((item=>{ console.log(item);return {key: item.Id, text: item.File.Name }}));
          } else {
             const items=await spList.items.getAll();
             console.log("The list is not a document library.");
             result=items.map<IDropdownOption>((item=>{return {key: item.Id, text: item.Title }}));
          }
        
         
         return result;
    }
    async GetWheatherData(): Promise<IWeatherData[]>{
        let client=await this._context.aadHttpClientFactory.getClient(this._apiId);
        let response= await client.get('https://localhost:7282/WeatherForecast',AadHttpClient.configurations.v1);
        let data:IWeatherData[] = await response.json();
        return data;
        
    }

}