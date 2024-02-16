declare interface IPageHacksWebPartStrings {
  PropertyPaneDescription: string;
  PageHeaderFieldLabel: string;
  PageNavigationFieldLabel: string;
  PagePaddingFieldLabel: string;
  PageTypeFieldDescription: string;
  PageTypeFieldLabel: string;
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
