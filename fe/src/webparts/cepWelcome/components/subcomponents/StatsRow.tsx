import * as React from 'react';
import type { IUserSummary, IUserUsage } from '../../../../services/CepApiModels';
import { getLevelIcon, getLevelLabel, getNextLevelLabel, isTopLevel } from '../../../../services/CepLevelPresentation';
import styles from '../CepWelcome.module.scss';
import * as strings from 'CepWelcomeWebPartStrings';

interface IStatsRowProps {
  summary: IUserSummary;
  /** Usage data to extract win stats */
  usage?: IUserUsage;
  /** Monthly points needed to reach Silver (default 500) */
  silverThreshold?: number;
  /** Monthly points needed to reach Gold (default 2000) */
  goldThreshold?: number;
  /** Show skeleton placeholders */
  loading?: boolean;
}

const fmt = (tpl: string, ...args: (string | number)[]): string =>
  tpl.replace(/{(\d+)}/g, (m, i) => (args[i] !== undefined ? String(args[i]) : m));

export const StatsRow: React.FC<IStatsRowProps> = ({
  summary,
  usage,
  silverThreshold = 500,
  goldThreshold = 2000,
  loading,
}) => {
  if (loading) {
    return (
      <div className={styles.statsRow}>
        {[0, 1, 2, 3].map((i) => (
          <div key={i} className={styles.statCard}>
            <div className={`${styles.skeletonLine} ${styles.skeletonLineShort}`} />
            <div className={styles.skeletonBlock} />
            <div className={`${styles.skeletonLine} ${styles.skeletonLineMedium}`} style={{ marginTop: 8 }} />
          </div>
        ))}
      </div>
    );
  }

  const { currentLevel, monthlyPoints, totalPoints, globalRank, teamRank } = summary;

  // Extract win stats from usage breakdown
  const winEntry = usage?.breakdown?.find((e) => e.appKey === 'Win');
  const winCount = winEntry?.promptCount ?? 0;
  const winPoints = winEntry?.pointsEarned ?? 0;

  let progressPct = 0;
  let nextLevelName = '';
  let pointsToGo = 0;
  let progressClass = styles.progressBronze;
  const currentLabel = getLevelLabel(currentLevel);

  if (currentLabel === 'Explorer') {
    nextLevelName = getNextLevelLabel(currentLevel);
    pointsToGo = Math.max(0, silverThreshold - monthlyPoints);
    progressPct = Math.min(monthlyPoints / silverThreshold, 1) * 100;
    progressClass = styles.progressBronze;
  } else if (currentLabel === 'Practitioner') {
    nextLevelName = getNextLevelLabel(currentLevel);
    pointsToGo = Math.max(0, goldThreshold - monthlyPoints);
    progressPct = Math.min(monthlyPoints / goldThreshold, 1) * 100;
    progressClass = styles.progressSilver;
  } else {
    progressPct = 100;
    progressClass = styles.progressGold;
  }

  return (
    <div className={styles.statsRow}>
      {/* ── Level progress ── */}
      <div className={styles.statCard}>
        <div className={styles.statLabel}>{strings.Level}</div>
        <div className={styles.statValue} style={{ fontSize: 22 }}>
          {getLevelIcon(currentLevel)} {getLevelLabel(currentLevel)}
        </div>
        <div className={styles.levelProgress}>
          <div className={styles.progressTrack}>
            <div
              className={`${styles.progressFill} ${progressClass}`}
              style={{ width: `${progressPct}%` }}
            />
          </div>
          {!isTopLevel(currentLevel) ? (
            <div className={styles.statSubValue}>
              {fmt(strings.PointsToNextLevel, pointsToGo, nextLevelName)}
            </div>
          ) : (
            <div className={styles.statSubValue}>🏆 Master level reached!</div>
          )}
        </div>
      </div>

      {/* ── Points ── */}
      <div className={styles.statCard}>
        <div className={styles.statLabel}>{strings.PointsThisMonth}</div>
        <div className={styles.statValue}>{monthlyPoints.toLocaleString()}</div>
        <div className={styles.statSubValue}>
          {strings.TotalPoints}: {totalPoints.toLocaleString()}
        </div>
      </div>

      {/* ── Ranks ── */}
      <div className={styles.statCard}>
        <div className={styles.statLabel}>{strings.GlobalRank}</div>
        <div className={styles.statValue}>
          {globalRank ? `#${globalRank}` : '–'}
        </div>
        {!!teamRank && (
          <div className={styles.statSubValue}>
            {strings.TeamRank}: #{teamRank}
          </div>
        )}
      </div>

      {/* ── Copilot Wins ── */}
      <div className={styles.statCard}>
        <div className={styles.statLabel}>{strings.WinCardTitle}</div>
        <div className={styles.statValue}>{winCount}</div>
        <div className={styles.statSubValue}>
          {strings.WinStatPoints.replace('{0}', String(winPoints))}
        </div>
      </div>
    </div>
  );
};
