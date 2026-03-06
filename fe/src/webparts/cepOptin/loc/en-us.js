define([], function() {
  return {
    // Property Pane
    "PropertyPaneDescription": "Configure the CEP backend connection",
    "FunctionAppGroupName": "Azure Functions Backend",
    "FunctionAppBaseUrlLabel": "Function App Base URL",
    "FunctionAppBaseUrlDescription": "Base URL of the Function App without trailing slash. E.g.: https://cep-functions.azurewebsites.net",

    // General
    "Loading": "Loading…",
    "Retry": "Retry",
    "BackButton": "Back",
    "ContinueButton": "Continue",

    // Not-configured banner
    "NotConfiguredTitle": "Web part not configured.",
    "NotConfiguredMessage": "SharePoint Tenant Properties have not been set. Run deploy/set-tenant-properties.ps1 to configure the backend URL.",

    // Setup prompt (edit mode, empty welcome text)
    "SetupTitle": "Configure your Copilot Engagement Program",
    "SetupDescription": "Set a welcome text for the enrollment wizard. Use the AI generator in the settings panel to create a personalised message in seconds.",
    "SetupButton": "Configure Welcome Text",

    // WelcomeStep
    "ProgramTitle": "Copilot Engagement Program",
    "HelloGreeting": "Hello, {0}! \uD83D\uDC4B",
    "FeatureTrack": "Track your Copilot usage across all Microsoft 365 apps",
    "FeaturePoints": "Earn 1 point for every prompt you send",
    "FeatureLevels": "Level up from Bronze \u2192 Silver \u2192 Gold",
    "FeatureLeaderboard": "Compete on global and team leaderboards",
    "GetStarted": "Get Started",

    // RulesStep
    "HowItWorksTitle": "How it works",
    "RuleUseTitle": "Use Copilot",
    "RuleUseBody": "Use Microsoft 365 Copilot in any app \u2014 Word, Excel, Outlook, Teams, and more \u2014 as part of your everyday work.",
    "RuleEarnTitle": "Earn Points",
    "RuleEarnBody": "1 point for every prompt you send. Points accumulate over the month and reset at month-end for a fresh competition.",
    "RuleLevelTitle": "Level Up",
    "RuleLevelBody": "Reach Bronze, Silver, and Gold tiers as your usage grows. Badges are permanent \u2014 earn them once, keep them forever.",
    "PrivacyTitle": "Privacy & Transparency",
    "PrivacyItem1": "Number of prompts sent to Copilot (content is never read)",
    "PrivacyItem2": "App used (Word, Excel, Outlook, Teams\u2026)",
    "PrivacyItem3": "Daily activity date (aggregated, not per-interaction)",
    "PrivacyItem4": "Prompt text, responses, and attachments \u2014 never stored",
    "PrivacyItem5": "Personal content or sensitive data \u2014 never collected",
    "PrivacyNote": "Tracking runs once per day in aggregate. You can withdraw at any time.",

    // PreferencesStep
    "ProfileTitle": "Your Profile",
    "ProfileSubtitle": "These optional details personalise your leaderboard view. You can update them at any time.",
    "DepartmentLabel": "Department",
    "DepartmentPlaceholder": "e.g. IT, Marketing, Finance",
    "TeamLabel": "Team",
    "TeamPlaceholder": "e.g. Cloud Team, North Sales",
    "NotificationsTitle": "Notifications",
    "NudgesLabel": "Teams engagement nudges",
    "NudgesOn": "Enabled",
    "NudgesOff": "Disabled",
    "NudgesNote": "Receive a Teams message when you have not used Copilot for 3 or more consecutive days \u2014 a friendly nudge to stay in the game!",

    // ConsentStep
    "AlmostInTitle": "You're almost in!",
    "AlmostInSubtitle": "Review your details before joining.",
    "EnrollmentSummaryTitle": "Enrollment summary",
    "NotificationsEnabled": "\uD83D\uDD14 Enabled",
    "NotificationsDisabled": "\uD83D\uDD15 Disabled",
    "DataCollectionTitle": "By joining, we may collect:",
    "DataItem1": "Number of prompts sent to Copilot (not their content)",
    "DataItem2": "Which Microsoft 365 apps you used Copilot in",
    "DataItem3": "Daily activity timestamps (aggregated, not per-interaction)",
    "DataNote": "Your prompt content and responses are never stored or read. You can leave the programme at any time.",
    "ConsentLabel": "I consent to the data collection described above and confirm that I want to join the Copilot Engagement Program.",
    "JoinButton": "Join the Program",
    "JoinButtonLoading": "Joining\u2026",

    // Actions / outcomes
    "JoinSuccess": "Enrollment complete! Welcome to the Copilot Engagement Program. \uD83C\uDF89",
    "LeaveButton": "Cancel enrollment",
    "LeaveButtonLoading": "Processing\u2026",
    "LeaveSuccess": "Enrollment cancelled. Your data will be retained for 90 days.",

    // Errors
    "JoinError": "Enrollment error: {0}",
    "LeaveError": "Error cancelling enrollment: {0}",

    // Enrolled view
    "PointsThisMonth": "Points this month",
    "TotalPoints": "Total points",
    "GlobalRank": "Global rank",
    "TeamRank": "Team rank",
    "LastActivity": "Last activity: {0}",
    "PreferencesTitle": "Notification preferences",
    "NudgesToggleLabel": "Teams engagement nudges after 3+ days of inactivity",

    // Leave dialog
    "LeaveDialogTitle": "Cancel enrollment",
    "LeaveDialogMessage": "Are you sure you want to leave the Copilot Engagement Program? Tracking will stop and you will be removed from the leaderboard. Your data will be retained for 90 days.",
    "LeaveDialogConfirm": "Yes, cancel enrollment",
    "LeaveDialogCancel": "Cancel",

    // WelcomeTextEditor (property pane)
    "GeneratorTitle": "\u2728 AI Welcome Text Generator",
    "OrgNameLabel": "Organisation name",
    "OrgNameDescription": "Used to personalise the generated text",
    "OrgNamePlaceholder": "e.g. Contoso, Fabrikam",
    "OrgNameRequired": "Please enter the organisation name first.",
    "ToneLabel": "Tone (optional)",
    "GenerateButton": "Generate with Copilot",
    "GeneratingButton": "Generating\u2026",
    "GenerationFailed": "Generation failed: {0}",
    "FallbackWarning": "M365 Copilot Chat API is not available in this tenant \u2014 a built-in template was used instead. You can edit the text below freely.",
    "WelcomeTextLabel": "Welcome text",
    "WelcomeTextDescription": "Shown to users on the first screen of the enrollment wizard. Edit freely. Use **word** for bold.",
    "WelcomeTextPlaceholder": "Describe the programme\u2026 e.g. \u2018Join our Copilot journey and start earning points today!\u2019",
    "ClearWelcomeText": "Clear welcome text"
  };
});
