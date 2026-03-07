declare interface ICepOptinWebPartStrings {
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

  // Enrolled view
  PointsThisMonth: string;
  TotalPoints: string;
  GlobalRank: string;
  TeamRank: string;
  LastActivity: string;
  PreferencesTitle: string;
  NudgesToggleLabel: string;

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
  AiWelcomeGenerating: string;
  AiWelcomeBadge: string;

  // Inline welcome editor (edit-mode, in-webpart)
  InlineEditorTitle: string;
  InlineEditorSave: string;
  InlineEditorDiscard: string;
  InlineEditorInputPlaceholder: string;
  InlineEditorEditButton: string;
  EditModeNotice: string;
}

declare module 'CepOptinWebPartStrings' {
  const strings: ICepOptinWebPartStrings;
  export = strings;
}
