declare interface ICepWinsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;

  // Header
  WebPartTitle: string;
  PrevMonth: string;
  NextMonth: string;

  // Tabs
  TabMyWins: string;
  TabCommunity: string;

  // Filter
  FilterAll: string;

  // Actions
  Retry: string;
  CopyText: string;
  Copied: string;
  ShowMore: string;

  // Empty / error states
  NotConfigured: string;
  LoadError: string;
  EmptyFilteredWins: string;
  EmptyMyWins: string;
  EmptyCommunityWins: string;
}

declare module 'CepWinsWebPartStrings' {
  const strings: ICepWinsWebPartStrings;
  export = strings;
}
