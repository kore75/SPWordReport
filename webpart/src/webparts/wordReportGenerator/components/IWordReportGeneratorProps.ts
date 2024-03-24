
import { ISPDataService } from "../../../service/ISPDataService";
import {ISpListInfo} from"../ISpListInfo"
export interface IWordReportGeneratorProps {
  externalApiUrl: string;
  reportDocLib?: ISpListInfo;
  reportDocItem?: ISpListInfo;
  reportDocList?: ISpListInfo;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  dataService:ISPDataService;
  
}
