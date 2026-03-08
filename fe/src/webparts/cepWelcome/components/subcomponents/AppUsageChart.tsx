import * as React from 'react';
import { Icon } from '@fluentui/react';
import type { IUserUsage } from '../../../../services/CepApiModels';
import styles from '../CepWelcome.module.scss';
import * as strings from 'CepWelcomeWebPartStrings';

interface IAppUsageChartProps {
  usage: IUserUsage | undefined;
  /** Show skeleton placeholders */
  loading?: boolean;
}

const APP_LABEL: Record<string, string> = {
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

const APP_TOOLTIP: Record<string, string> = {
  Word: strings.AppTooltipWord,
  Excel: strings.AppTooltipExcel,
  PowerPoint: strings.AppTooltipPowerPoint,
  Outlook: strings.AppTooltipOutlook,
  Teams: strings.AppTooltipTeams,
  OneNote: strings.AppTooltipOneNote,
  Loop: strings.AppTooltipLoop,
  BizChat: strings.AppTooltipBizChat,
  WebChat: strings.AppTooltipWebChat,
  M365App: strings.AppTooltipM365App,
  Forms: strings.AppTooltipForms,
  SharePoint: strings.AppTooltipSharePoint,
  Whiteboard: strings.AppTooltipWhiteboard,
};

const APP_CLASS: Record<string, string> = {
  Word: styles.appWord,
  Excel: styles.appExcel,
  PowerPoint: styles.appPowerPoint,
  Outlook: styles.appOutlook,
  Teams: styles.appTeams,
  OneNote: styles.appOneNote,
  Loop: styles.appLoop,
  BizChat: styles.appBizChat,
  WebChat: styles.appWebChat,
  M365App: styles.appM365App,
  Forms: styles.appForms,
  SharePoint: styles.appSharePoint,
  Whiteboard: styles.appWhiteboard,
};

const APP_ICON: Record<string, { iconName: string; color: string }> = {
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

const fmt = (tpl: string, ...args: (string | number)[]): string =>
  tpl.replace(/{(\d+)}/g, (m, i) => (args[i] !== undefined ? String(args[i]) : m));

const ALL_APP_KEYS = Object.keys(APP_LABEL);

export const AppUsageChart: React.FC<IAppUsageChartProps> = ({ usage, loading }) => {
  if (loading || !usage) {
    return (
      <div className={styles.card}>
        <div className={`${styles.skeletonLine} ${styles.skeletonLineMedium}`} style={{ height: 16, marginBottom: 16 }} />
        <div className={styles.appBarList}>
          {[100, 85, 60, 40, 20, 10].map((w, i) => (
            <div key={i} className={styles.appBarRow}>
              <div className={`${styles.skeletonLine} ${styles.skeletonLineShort}`} style={{ marginBottom: 0 }} />
              <div className={styles.skeletonBar} style={{ width: `${w}%` }} />
              <div className={`${styles.skeletonLine}`} style={{ width: 50, marginBottom: 0 }} />
            </div>
          ))}
        </div>
      </div>
    );
  }

  const breakdownMap: Record<string, number> = {};
  (usage.breakdown ?? []).forEach((e) => { breakdownMap[e.appKey] = e.promptCount; });

  const allEntries = ALL_APP_KEYS
    .map((appKey) => ({ appKey, promptCount: breakdownMap[appKey] ?? 0 }))
    .sort((a, b) => b.promptCount - a.promptCount);

  const totalCount = allEntries.reduce((sum, e) => sum + e.promptCount, 0);
  const scale = totalCount > 0 ? totalCount : 1;

  const usedEntries = allEntries.filter((e) => e.promptCount > 0);
  const unusedEntries = allEntries.filter((e) => e.promptCount === 0);

  const renderRow = (entry: { appKey: string; promptCount: number }, dimmed: boolean): JSX.Element => {
    const pct = Math.round((entry.promptCount / scale) * 100);
    return (
      <div
        key={entry.appKey}
        className={`${styles.appBarRow} ${dimmed ? styles.appBarRowDimmed : ''}`}
        title={APP_TOOLTIP[entry.appKey] ?? ''}
      >
        <div className={styles.appBarLabel}>
          {APP_ICON[entry.appKey] && (
            <Icon
              iconName={APP_ICON[entry.appKey].iconName}
              className={styles.appBarIcon}
              style={{ color: dimmed ? undefined : APP_ICON[entry.appKey].color }}
            />
          )}
          {APP_LABEL[entry.appKey] ?? entry.appKey}
        </div>
        <div className={`${styles.appBarTrack} ${dimmed ? styles.appBarTrackDimmed : ''}`}>
          {!dimmed && (
            <div
              className={`${styles.appBarFill} ${APP_CLASS[entry.appKey] ?? styles.appM365App}`}
              style={{ width: `${pct}%` }}
            />
          )}
        </div>
        <div className={`${styles.appBarCount} ${dimmed ? styles.appBarCountDimmed : ''}`}>
          {dimmed ? '—' : `${fmt(strings.PromptCount, entry.promptCount)} · ${pct}%`}
        </div>
      </div>
    );
  };

  return (
    <div className={styles.card}>
      <div className={styles.sectionTitle}>{strings.UsageBreakdownTitle}</div>
      <div className={styles.appBarList}>
        {usedEntries.map((entry) => renderRow(entry, false))}
        {unusedEntries.length > 0 && usedEntries.length > 0 && (
          <div className={styles.appBarDivider} />
        )}
        {unusedEntries.map((entry) => renderRow(entry, true))}
      </div>
    </div>
  );
};
