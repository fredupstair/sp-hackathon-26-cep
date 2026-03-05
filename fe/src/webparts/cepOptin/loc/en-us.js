define([], function() {
  return {
    // Property Pane
    "PropertyPaneDescription": "Configure the CEP backend connection",
    "FunctionAppGroupName": "Azure Functions Backend",
    "FunctionAppBaseUrlLabel": "Function App Base URL",
    "FunctionAppBaseUrlDescription": "Base URL of the Function App without trailing slash. E.g.: https://cep-functions.azurewebsites.net",
    
    // Component UI
    "WebPartTitle": "Copilot Engagement Program",
    "WelcomeMessage": "Welcome, {0}! Enroll in the program to track your Copilot usage, earn points and climb the leaderboard.",
    "NotConfiguredTitle": "Web part not configured.",
    "NotConfiguredMessage": "SharePoint Tenant Properties have not been set.",
    "Loading": "Loading...",
    "Retry": "Retry",
    
    // Transparency section
    "TransparencyTitle": "What we collect",
    "TransparencyItem1": "Number of prompts sent to Copilot (no content)",
    "TransparencyItem2": "App used (Word, Excel, Outlook, Teams, etc.)",
    "TransparencyItem3": "Daily activity date (aggregated, not per interaction)",
    "TransparencyNote": "We do NOT store: prompt text, responses, attachments, mentions or any personal content. Tracking occurs once per day. You can unsubscribe at any time.",
    
    // Form fields
    "DepartmentLabel": "Department",
    "DepartmentPlaceholder": "e.g. IT, Marketing, Finance",
    "TeamLabel": "Team",
    "TeamPlaceholder": "e.g. Cloud Team, North Sales Team",
    "NudgesLabel": "Teams engagement notifications",
    "NudgesOn": "Enabled",
    "NudgesOff": "Disabled",
    "NudgesDescription": "You will receive a Teams message if you don't use Copilot for 3+ consecutive days.",
    
    // Consent
    "ConsentLabel": "I consent to the collection of the data described above and confirm that I want to participate in the Copilot Engagement Program.",
    
    // Actions
    "JoinButton": "Enroll in the program",
    "JoinButtonLoading": "Enrolling...",
    "JoinSuccess": "Enrollment complete! Welcome to the Copilot Engagement Program.",
    "LeaveButton": "Cancel enrollment",
    "LeaveButtonLoading": "Processing...",
    "LeaveSuccess": "Enrollment cancelled. Your data will be retained for 90 days.",
    
    // Errors
    "LoadError": "Error loading: {0}",
    "JoinError": "Error during enrollment: {0}",
    "LeaveError": "Error during cancellation: {0}",
    
    // Enrolled view
    "PointsThisMonth": "Points this month",
    "TotalPoints": "Total points",
    "GlobalRank": "Global rank",
    "TeamRank": "Team rank",
    "LastActivity": "Last activity: {0}",
    "PreferencesTitle": "Notification preferences",
    "NudgesToggleLabel": "Teams engagement notifications after 3+ days of inactivity",
    
    // Leave dialog
    "LeaveDialogTitle": "Cancel enrollment",
    "LeaveDialogMessage": "Are you sure you want to leave the Copilot Engagement Program? Tracking will stop and you will be removed from the leaderboard. Your data will be retained for 90 days.",
    "LeaveDialogConfirm": "Yes, cancel enrollment",
    "LeaveDialogCancel": "Cancel"
  }
});