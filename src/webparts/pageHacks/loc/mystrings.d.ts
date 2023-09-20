declare interface IPageHacksWebPartStrings {
  PropertyPaneDescription: string;
  PageNavigationFieldDescription: string;
  PageNavigationFieldLabel: string;
  PageTypeFieldDescription: string;
  PageTypeFieldLabel: string;
  PageWidthFieldDescription: string;
  PageWidthFieldLabel: string;
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
