import * as React from 'react';
import styles from './WordReportGenerator.module.scss';
import type { IWordReportGeneratorProps } from './IWordReportGeneratorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { useBoolean } from '@uifabric/react-hooks';
import { ISpListInfo } from '../ISpListInfo';

interface reportListItem{
  Id:string;
  Title:string
}

const WordReportGenerator: React.FC<IWordReportGeneratorProps> = (props: IWordReportGeneratorProps)=>{
  const {
    description,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName,
    reportDocLib,
    reportDocItem,
    reportDocList,
    dataService
  } = props;

  const [loading, setLoading] = React.useState<Boolean>(true);

  let listItems:reportListItem[]=[];

  const viewFields: IViewField[] = [
    {
        name: 'Id',
        displayName: 'Id',
        minWidth: 100,
        maxWidth: 150,
        sorting: true
    },
    {
        name: 'Title',
        displayName: 'Title',
        minWidth: 100,
        maxWidth: 150
    }    
  ];

  const loadItems=()=>{

    if(reportDocList!=null){
      dataService.loadItems(reportDocList.Id).then((items)=>{
          listItems=items.map<reportListItem>((item:any)=>{return {Id:item.key,Title:item.text}});
          setLoading(false);
      }
      );
    }
    else if(loading===false){
      setLoading(true);
    }

  };

  loadItems();
  return (
    <section className={`${styles.wordReportGenerator} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
        <div>Report Vorlagenliste: <strong>{escape(reportDocLib?.Title ?? "")}</strong></div>
        <div>Report Vorlagen: <strong>{escape(reportDocItem?.Title ?? "")}</strong></div>
        <div>Report Liste: <strong>{escape(reportDocList?.Title ?? "")}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        { loading ? (<div>Loading, Please wait...</div>):
        (
        <ListView
                        items={listItems}
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
