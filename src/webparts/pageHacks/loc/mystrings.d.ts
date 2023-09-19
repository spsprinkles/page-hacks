declare interface IPageHacksWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PageWidthFieldLabel: string;
  PageTypeFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
}

declare module 'PageHacksWebPartStrings' {
  const strings: IPageHacksWebPartStrings;
  export = strings;
}
