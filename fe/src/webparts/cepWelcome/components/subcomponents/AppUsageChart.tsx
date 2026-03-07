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
  M365Chat: 'Copilot Chat',
  Other: 'Other',
};

const APP_CLASS: Record<string, string> = {
  Word: styles.appWord,
  Excel: styles.appExcel,
  PowerPoint: styles.appPowerPoint,
  Outlook: styles.appOutlook,
  Teams: styles.appTeams,
  OneNote: styles.appOneNote,
  Loop: styles.appLoop,
  M365Chat: styles.appM365Chat,
  Other: styles.appOther,
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
          <div key={entry.appKey} className={styles.appBarRow}>
            <div className={styles.appBarLabel}>
              {APP_LABEL[entry.appKey] ?? entry.appKey}
            </div>
            <div className={styles.appBarTrack}>
              <div
                className={`${styles.appBarFill} ${APP_CLASS[entry.appKey] ?? styles.appOther}`}
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
