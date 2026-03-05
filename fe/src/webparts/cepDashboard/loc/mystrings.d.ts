declare interface ICepDashboardWebPartStrings {
  // Property Pane
  PropertyPaneDescription: string;
  BasicGroupName: string;

  // Header
  WebPartTitle: string;
  ViewingProfileOf: string;
  BackToMyProfile: string;

  // Loading / errors
  Loading: string;
  LoadError: string;
  NotEnrolledTitle: string;
  NotEnrolledMessage: string;
  JoinProgram: string;
  Retry: string;

  // Stats
  PointsThisMonth: string;
  TotalPoints: string;
  GlobalRank: string;
  TeamRank: string;
  LastActivity: string;
  Level: string;
  NextLevel: string;
  PointsToNextLevel: string;

  // App usage breakdown
  UsageBreakdownTitle: string;
  PromptCount: string;
  NoUsage: string;
  FilterWeek: string;
  FilterMonth: string;
  GetInspired: string;

  // Badges
  BadgesTitle: string;
  NoBadgesYet: string;
  BadgeEarnedOn: string;
  LockedBadge: string;

  // Other user view (aggregated)
  OtherUserPointsThisMonth: string;
  OtherUserTotalPoints: string;
  OtherUserLevel: string;
  OtherUserBadges: string;
}

declare module 'CepDashboardWebPartStrings' {
  const strings: ICepDashboardWebPartStrings;
  export = strings;
}
