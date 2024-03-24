declare interface IWordReportGeneratorWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ReportList:string;
  DescriptionFieldLabel: string;
  ReportDocLibLabel: string;
  ReportDocLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'WordReportGeneratorWebPartStrings' {
  const strings: IWordReportGeneratorWebPartStrings;
  export = strings;
}
