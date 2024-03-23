
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

export interface  ISPDataService
{
    loadSiteCollectionDocLibs(): Promise<IDropdownOption[]>;
    loadSiteCollectionLists(): Promise<IDropdownOption[]>;
    loadItems(splist:string): Promise<IDropdownOption[]>;
}