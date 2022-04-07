declare interface INewWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'NewWebPartStrings' {
  const strings: INewWebPartStrings;
  export = strings;
}
