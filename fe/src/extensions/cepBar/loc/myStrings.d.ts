declare interface ICepBarApplicationCustomizerStrings {
  Title: string;
  // Bar – general
  BarLoading: string;
  BarPoints: string;
  BarLevel: string;
  BarRank: string;
  BarNotEnrolled: string;
  BarNotEnrolledDesc: string;
  JoinOverlayCta: string;
  BarViewDashboard: string;
  // Bar – streak / progress
  StreakLabel: string;
  StreakTooltip: string;
  BarProgressLabel: string;
  // Bar – win
  BarMarkWin: string;
  WinCalloutTitle: string;
  WinCalloutApp: string;
  WinCalloutNote: string;
  WinCalloutNotePlaceholder: string;
  WinCalloutShare: string;
  WinCalloutSubmit: string;
  WinLimitReached: string;
  WinAnimText: string;
  // Bar – notification preferences
  NudgesTooltipOn: string;
  NudgesTooltipOff: string;
  // Bar – suggestion
  DismissSuggestion: string;
  // Common
  Cancel: string;
}

declare module 'CepBarApplicationCustomizerStrings' {
  const strings: ICepBarApplicationCustomizerStrings;
  export = strings;
}
