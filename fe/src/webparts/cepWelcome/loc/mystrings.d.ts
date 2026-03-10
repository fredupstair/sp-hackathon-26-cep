declare interface ICepWelcomeWebPartStrings {
  // Property Pane
  PropertyPaneDescription: string;
  FunctionAppGroupName: string;
  FunctionAppBaseUrlLabel: string;
  FunctionAppBaseUrlDescription: string;

  // General
  Loading: string;
  Retry: string;
  BackButton: string;
  ContinueButton: string;

  // Not-configured banner
  NotConfiguredTitle: string;
  NotConfiguredMessage: string;

  // Setup prompt (edit mode, empty welcome text)
  SetupTitle: string;
  SetupDescription: string;
  SetupButton: string;

  // WelcomeStep
  ProgramTitle: string;
  HelloGreeting: string;
  FeatureTrack: string;
  FeaturePoints: string;
  FeatureLevels: string;
  FeatureLeaderboard: string;
  GetStarted: string;
  LetsStart: string;

  // RulesStep
  HowItWorksTitle: string;
  RuleUseTitle: string;
  RuleUseBody: string;
  RuleEarnTitle: string;
  RuleEarnBody: string;
  RuleLevelTitle: string;
  RuleLevelBody: string;
  PrivacyTitle: string;
  PrivacyItem1: string;
  PrivacyItem2: string;
  PrivacyItem3: string;
  PrivacyItem4: string;
  PrivacyItem5: string;
  PrivacyNote: string;

  // PreferencesStep
  ProfileTitle: string;
  ProfileSubtitle: string;
  DepartmentLabel: string;
  DepartmentPlaceholder: string;
  TeamLabel: string;
  TeamPlaceholder: string;
  NotificationsTitle: string;
  NudgesLabel: string;
  NudgesOn: string;
  NudgesOff: string;
  NudgesNote: string;

  // ConsentStep
  AlmostInTitle: string;
  AlmostInSubtitle: string;
  EnrollmentSummaryTitle: string;
  NotificationsEnabled: string;
  NotificationsDisabled: string;
  DataCollectionTitle: string;
  DataItem1: string;
  DataItem2: string;
  DataItem3: string;
  DataNote: string;
  ConsentLabel: string;
  JoinButton: string;
  JoinButtonLoading: string;

  // Actions / outcomes
  JoinSuccess: string;
  LeaveButton: string;
  LeaveButtonLoading: string;
  LeaveSuccess: string;

  // Errors
  JoinError: string;
  LeaveError: string;

  // Enrolled view (from OptIn)
  PointsThisMonth: string;
  TotalPoints: string;
  GlobalRank: string;
  TeamRank: string;
  LastActivity: string;
  PreferencesTitle: string;
  NudgesToggleLabel: string;
  SettingsPanelTitle: string;

  // Leave dialog
  LeaveDialogTitle: string;
  LeaveDialogMessage: string;
  LeaveDialogConfirm: string;
  LeaveDialogCancel: string;

  // WelcomeTextEditor (property pane)
  GeneratorTitle: string;
  OrgNameLabel: string;
  OrgNameDescription: string;
  OrgNamePlaceholder: string;
  OrgNameRequired: string;
  ToneLabel: string;
  GenerateButton: string;
  GeneratingButton: string;
  GenerationFailed: string;
  FallbackWarning: string;
  WelcomeTextLabel: string;
  WelcomeTextDescription: string;
  WelcomeTextPlaceholder: string;
  ClearWelcomeText: string;

  // AI personalised welcome (end-user)
  AiWelcomeStaticIntro: string;
  AiWelcomeStaticRoleHint: string;
  AiWelcomeClosing: string;
  AiWelcomeGenerating: string;
  AiWelcomeBadge: string;
  CopilotBubbleLabel: string;
  CopilotBubbleLoading: string;
  CopilotLoadingHint: string;
  CopilotLoadingHint2: string;
  CopilotLoadingHint3: string;
  CopilotLoadingHint4: string;
  CopilotLoadingHint5: string;
  CopilotChatUserMessage: string;

  // Inline welcome editor (edit-mode, in-webpart)
  InlineEditorTitle: string;
  InlineEditorSave: string;
  InlineEditorDiscard: string;
  InlineEditorInputPlaceholder: string;
  InlineEditorEditButton: string;
  EditModeNotice: string;

  // ── Dashboard strings ──────────────────────────────────────────────────

  // Header
  WebPartTitle: string;
  ViewingProfileOf: string;
  BackToMyProfile: string;

  // Loading / errors (dashboard-specific)
  LoadError: string;
  NotEnrolledTitle: string;
  NotEnrolledMessage: string;
  JoinProgram: string;

  // Stats
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
  TabUniverse: string;
  TabStats: string;
  UniversePromptsLabel: string;
  UniverseOfTotal: string;
  UniversePoints: string;
  UniverseLockedWorlds: string;
  UniverseLockedHint: string;

  // App tooltips
  AppTooltipWord: string;
  AppTooltipExcel: string;
  AppTooltipPowerPoint: string;
  AppTooltipOutlook: string;
  AppTooltipTeams: string;
  AppTooltipOneNote: string;
  AppTooltipLoop: string;
  AppTooltipBizChat: string;
  AppTooltipWebChat: string;
  AppTooltipM365App: string;
  AppTooltipForms: string;
  AppTooltipSharePoint: string;
  AppTooltipWhiteboard: string;

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

  // Leaderboard
  LeaderboardLastUpdated: string;

  // Copilot Win card
  WinCardTitle: string;
  WinStatPoints: string;
}

declare module 'CepWelcomeWebPartStrings' {
  const strings: ICepWelcomeWebPartStrings;
  export = strings;
}
