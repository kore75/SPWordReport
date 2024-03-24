
import {ISpListInfo} from"../ISpListInfo"
export interface IWordReportGeneratorProps {
  description: string;
  reportDocLib?: ISpListInfo;
  reportDocItem?: ISpListInfo;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
}
