declare interface ICepLeaderboardWebPartStrings {
  // Property Pane
  PropertyPaneDescription: string;
  BasicGroupName: string;

  // Header
  WebPartTitle: string;
  CurrentMonth: string;

  // Loading / errors
  Loading: string;
  LoadError: string;
  Retry: string;

  // Filter tabs
  FilterGlobal: string;
  FilterDepartment: string;
  FilterTeam: string;

  // Podium
  PodiumTitle: string;
  PodiumFirstPlace: string;
  PodiumSecondPlace: string;
  PodiumThirdPlace: string;
  NoDataForPodium: string;

  // Leaderboard table
  ColumnRank: string;
  ColumnUser: string;
  ColumnDepartment: string;
  ColumnPoints: string;
  ColumnLevel: string;
  YouLabel: string;
  SearchPlaceholder: string;
  NoResults: string;
  LoadMore: string;

  // Aggregate stats
  StatsAvgPoints: string;
  StatsTotalUsers: string;
  StatsActiveUsers: string;
}

declare module 'CepLeaderboardWebPartStrings' {
  const strings: ICepLeaderboardWebPartStrings;
  export = strings;
}
