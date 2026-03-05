declare interface ICepOptinWebPartStrings {
  // Property Pane
  PropertyPaneDescription: string;
  FunctionAppGroupName: string;
  FunctionAppBaseUrlLabel: string;
  FunctionAppBaseUrlDescription: string;
  
  // Component UI
  WebPartTitle: string;
  WelcomeMessage: string;
  NotConfiguredTitle: string;
  NotConfiguredMessage: string;
  Loading: string;
  Retry: string;
  
  // Transparency section
  TransparencyTitle: string;
  TransparencyItem1: string;
  TransparencyItem2: string;
  TransparencyItem3: string;
  TransparencyNote: string;
  
  // Form fields
  DepartmentLabel: string;
  DepartmentPlaceholder: string;
  TeamLabel: string;
  TeamPlaceholder: string;
  NudgesLabel: string;
  NudgesOn: string;
  NudgesOff: string;
  NudgesDescription: string;
  
  // Consent
  ConsentLabel: string;
  
  // Actions
  JoinButton: string;
  JoinButtonLoading: string;
  JoinSuccess: string;
  LeaveButton: string;
  LeaveButtonLoading: string;
  LeaveSuccess: string;
  
  // Errors
  LoadError: string;
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
}

declare module 'CepOptinWebPartStrings' {
  const strings: ICepOptinWebPartStrings;
  export = strings;
}
