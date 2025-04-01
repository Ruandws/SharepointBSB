declare interface IDiretorioWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  title: string;  // Já adicionada
  SearchPlaceholder: string;  // Nova propriedade
  WelcomeTitle: string;  // Nova propriedade
  WelcomeSubtitle: string;  // Nova propriedade
}

declare module 'DiretorioWebPartStrings' {
  const strings: IDiretorioWebPartStrings;
  export = strings;
}
declare module "*.png" {
  const value: string;
  export default value;
}