interface LinkItem{
    Name:string;
    Url:string
  }

import * as React from 'react';
import styles from './WordReportGenerator.module.scss';

  const ExternalLinkButton: React.FC<LinkItem> = (props: LinkItem)=>{
    const {
      Name,
      Url      
    } = props;

    const openInNewTab = (url:string) => {
        window.open(url, "_blank", "noreferrer");
      };


    return (       
       <button className={styles.links} onClick={()=>openInNewTab(Url)}>{Name}</button>       
    );
  }


  export default ExternalLinkButton;