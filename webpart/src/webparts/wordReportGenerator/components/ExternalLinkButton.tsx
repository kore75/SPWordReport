import * as React from 'react';
import styles from './WordReportGenerator.module.scss';
import { IReportFileRequest } from '../../../service/IReportFileRequest';
import { ISPDataService } from '../../../service/ISPDataService';

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

    const openInNewTab = async (url:string):Promise<void> => {
        setCreating(true);
        //window.open(url, "_blank", "noreferrer");
        await SPDataService.CreateReport(Request);
        setCreating(false);

      };


    return (       
       <button disabled={creating} className={styles.links}  onClick={()=>openInNewTab(Url)}>{Name}</button>       
    );
  }


  export default ExternalLinkButton;