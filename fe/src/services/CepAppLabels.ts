/**
 * Human-readable labels and icons for canonical Copilot app keys.
 * Shared across all CEP web parts.
 */

export const APP_LABELS: Record<string, string> = {
  Word: 'Word',
  Excel: 'Excel',
  PowerPoint: 'PowerPoint',
  Outlook: 'Outlook',
  Teams: 'Teams',
  OneNote: 'OneNote',
  Loop: 'Loop',
  BizChat: 'Copilot Chat',
  WebChat: 'Web Chat',
  M365App: 'M365 App',
  Forms: 'Forms',
  SharePoint: 'SharePoint',
  Whiteboard: 'Whiteboard',
};

export function getAppLabel(appKey: string): string {
  return APP_LABELS[appKey] ?? appKey;
}

/** Fluent UI icon names and brand colours for each app key (same as AppUsageChart). */
export const APP_FLUENT_ICONS: Record<string, { iconName: string; color: string }> = {
  Word:       { iconName: 'WordLogo',       color: '#2b579a' },
  Excel:      { iconName: 'ExcelLogo',      color: '#217346' },
  PowerPoint: { iconName: 'PowerPointLogo', color: '#b7472a' },
  Outlook:    { iconName: 'OutlookLogo',    color: '#0078d4' },
  Teams:      { iconName: 'TeamsLogo',      color: '#6264a7' },
  OneNote:    { iconName: 'OneNoteLogo',    color: '#7719aa' },
  Loop:       { iconName: 'Sync',           color: '#008272' },
  BizChat:    { iconName: 'Chat',           color: '#0f6cbd' },
  WebChat:    { iconName: 'Globe',          color: '#3a96dd' },
  M365App:    { iconName: 'Waffle',         color: '#5c2d91' },
  Forms:      { iconName: 'ClipboardList',  color: '#038387' },
  SharePoint: { iconName: 'SharePointLogo', color: '#036c70' },
  Whiteboard: { iconName: 'EditCreate',     color: '#4a8fb8' },
};

export function getAppFluentIcon(appKey: string): { iconName: string; color: string } {
  return APP_FLUENT_ICONS[appKey] ?? { iconName: 'Waffle', color: '#5c2d91' };
}

/** @deprecated use getAppFluentIcon instead */
export const APP_ICONS: Record<string, string> = {};
/** @deprecated use getAppFluentIcon instead */
export function getAppIcon(appKey: string): string {
  return '';
}
