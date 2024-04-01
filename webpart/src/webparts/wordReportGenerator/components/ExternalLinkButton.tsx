import * as React from 'react';
import styles from './WordReportGenerator.module.scss';
import { IReportFileRequest } from '../../../service/IReportFileRequest';
import { ISPDataService } from '../../../service/ISPDataService';
import { IReportFileResult } from '../../../service/IReportFileResult';

interface LinkItem{
  Name:string;
  Url:string,
  Request:IReportFileRequest,
  SPDataService:ISPDataService
}

  const ExternalLinkButton: React.FC<LinkItem> = (props: LinkItem)=>{
    const {
      Name,
      Url,
      Request,
      SPDataService
    } = props;

    const [creating, setCreating] = React.useState<boolean>(false);
    const [error, setError] = React.useState<string>("");
    const [createdDoc, setCreatedDoc] = React.useState<IReportFileResult | undefined>(undefined);

    const createReport = async (url:string):Promise<void> => {
        try{
          setError("");
          setCreating(true);
          //window.open(url, "_blank", "noreferrer");
          setCreatedDoc(await SPDataService.CreateReport(Request));
          setCreating(false);
        }
        catch(err){
          setError(err);
        }        
      };

      const openInNewTab = async (url:string):Promise<void> => {              
          window.open(url, "_blank", "noreferrer");        
      };


    return (     
      <div>
        {
         error ==="" ?
         (<button disabled={creating} className={styles.links}  onClick={()=>createReport(Url)}>{Name}</button>)
        :(<strong>{error}</strong>)
        }
        { 
             createdDoc!==undefined ? 
             (<><br/><button className={styles.links} onClick={()=>openInNewTab(createdDoc?.filePath)}>{createdDoc?.createdFileName}</button></>):(<></>)
        }           
      </div>     
      
    );
  }


  export default ExternalLinkButton;