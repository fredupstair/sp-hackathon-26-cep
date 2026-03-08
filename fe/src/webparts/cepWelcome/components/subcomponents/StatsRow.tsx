import * as React from 'react';
import type { IUserSummary } from '../../../../services/CepApiModels';
import { getLevelIcon, getLevelLabel, getNextLevelLabel, isTopLevel } from '../../../../services/CepLevelPresentation';
import styles from '../CepWelcome.module.scss';
import * as strings from 'CepWelcomeWebPartStrings';

interface IStatsRowProps {
  summary: IUserSummary;
  /** Monthly points needed to reach Silver (default 500) */
  silverThreshold?: number;
  /** Monthly points needed to reach Gold (default 2000) */
  goldThreshold?: number;
}

const fmt = (tpl: string, ...args: (string | number)[]): string =>
  tpl.replace(/{(\d+)}/g, (m, i) => (args[i] !== undefined ? String(args[i]) : m));

export const StatsRow: React.FC<IStatsRowProps> = ({
  summary,
  silverThreshold = 500,
  goldThreshold = 2000,
}) => {
  const { currentLevel, monthlyPoints, totalPoints, globalRank, teamRank } = summary;

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
          {globalRank !== undefined ? `#${globalRank}` : '–'}
        </div>
        {teamRank !== undefined && (
          <div className={styles.statSubValue}>
            {strings.TeamRank}: #{teamRank}
          </div>
        )}
      </div>
    </div>
  );
};
