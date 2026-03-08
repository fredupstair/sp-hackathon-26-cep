import * as React from 'react';
import type { IUserUsage } from '../../../../services/CepApiModels';
import styles from '../CepWelcome.module.scss';
import * as strings from 'CepWelcomeWebPartStrings';

interface IAppUsageChartProps {
  usage: IUserUsage;
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

const fmt = (tpl: string, ...args: (string | number)[]): string =>
  tpl.replace(/{(\d+)}/g, (m, i) => (args[i] !== undefined ? String(args[i]) : m));

const ALL_APP_KEYS = Object.keys(APP_LABEL);

export const AppUsageChart: React.FC<IAppUsageChartProps> = ({ usage }) => {
  const breakdownMap: Record<string, number> = {};
  (usage.breakdown ?? []).forEach((e) => { breakdownMap[e.appKey] = e.promptCount; });

  const allEntries = ALL_APP_KEYS
    .map((appKey) => ({ appKey, promptCount: breakdownMap[appKey] ?? 0 }))
    .sort((a, b) => b.promptCount - a.promptCount);

  const maxCount = allEntries[0].promptCount > 0 ? allEntries[0].promptCount : 1;

  return (
    <div className={styles.card}>
      <div className={styles.sectionTitle}>{strings.UsageBreakdownTitle}</div>
      <div className={styles.appBarList}>
        {allEntries.map((entry) => (
          <div key={entry.appKey} className={styles.appBarRow} title={APP_TOOLTIP[entry.appKey] ?? ''}>
            <div className={styles.appBarLabel}>
              {APP_LABEL[entry.appKey] ?? entry.appKey}
            </div>
            <div className={styles.appBarTrack}>
              <div
                className={`${styles.appBarFill} ${APP_CLASS[entry.appKey] ?? styles.appM365App}`}
                style={{ width: `${(entry.promptCount / maxCount) * 100}%` }}
              />
            </div>
            <div className={styles.appBarCount}>
              {fmt(strings.PromptCount, entry.promptCount)}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};
