import * as React from 'react';
import styles from './WordReportGenerator.module.scss';
import type { IWordReportGeneratorProps } from './IWordReportGeneratorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { useBoolean } from '@uifabric/react-hooks';
import { ISpListInfo } from '../ISpListInfo';
import ExternalLinkButton from './ExternalLinkButton';
import { IWeatherData } from '../../../service/ISPDataService';

interface reportListItem{
  Id:string;
  Title:string
}

const WordReportGenerator: React.FC<IWordReportGeneratorProps> = (props: IWordReportGeneratorProps)=>{
  const {
    externalApiUrl,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName,
    reportDocLib,
    reportDocItem,
    reportDocList,
    dataService
  } = props;

  const [listetems, setListItems] = React.useState<reportListItem[]>([]);
  const [wdata, setWData] = React.useState<IWeatherData[]>([]);
  const [loading, setloading] = React.useState<boolean>(true);

 

  const viewFields: IViewField[] = [
    {
        name: 'Id',
        minWidth: 50,
        maxWidth: 100,
    },
    {
        name: 'Title',
        minWidth: 100,
        maxWidth: 150,             
    },
    {
      name:'CreateReport',
      displayName:'Create Report',      
      minWidth: 100,
      maxWidth: 150,             
      render:(item)=>{
        return(
          <ExternalLinkButton Name='Create Report' Url={item.CreateReport} />          
        )
      }
  }
  ];
  const viewFieldsWData: IViewField[] = [
    {
        name: 'date',
        minWidth: 50,
        maxWidth: 100,
    },
    {
        name: 'temperatureC',      
        minWidth: 100,
        maxWidth: 150,             
    },
    {
      name: 'summary',
      minWidth: 100,
      maxWidth: 150,             
    },
    {
      name: 'temperatureF',
      minWidth: 100,
      maxWidth: 150,             
    }   
  ];

  const showWeatherData:boolean=false;

  const loadItems=async ()=>{

    if(showWeatherData){
      let twdata=await dataService.GetWheatherData();   
      setWData(twdata);    
    }

    if(reportDocList!=null ){
      
      let allItems= await dataService.loadItems(reportDocList.Id);
      let reportItems=allItems.map<reportListItem>((item:any)=>{return {Id:item.key,Title:item.text,CreateReport:externalApiUrl+item.key}});
      setloading(false);
      setListItems(reportItems);                  
    }
    else setloading(true);
  }

  React.useEffect(() => { loadItems() }, [props.reportDocList,props.externalApiUrl]);

  return (
    <section className={`${styles.wordReportGenerator} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web API Url: <strong>{escape(externalApiUrl)}</strong></div>
        <div>Report Vorlagenliste: <strong>{escape(reportDocLib?.Title ?? "")}</strong></div>
        <div>Report Vorlagen: <strong>{escape(reportDocItem?.Title ?? "")}</strong></div>
        <div>Report Liste: <strong>{escape(reportDocList?.Title ?? "")}</strong></div>
      </div>
      <div>
        
        { showWeatherData ? (
          <div>
            <h3> Weather data from sample</h3>
         <ListView
                        items={wdata}
                        viewFields={viewFieldsWData}
                        compact={true}
                        selectionMode={SelectionMode.single}                                                                    
                        stickyHeader={true}
         />
         </div>
        ):(<></> )}
        <h3>Welcome to SharePoint Framework!</h3>
        { loading ? (<div>Loading, Please wait...</div>):
        (
        <ListView
                        items={listetems}
                        viewFields={viewFields}
                        compact={true}
                        selectionMode={SelectionMode.single}                                                                    
                        stickyHeader={true}
                        />
         )}                        
               
      </div>
    </section>
  );
}
export default WordReportGenerator;
