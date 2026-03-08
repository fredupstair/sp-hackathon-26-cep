import * as React from 'react';
import type { ILeaderboardPage, ILeaderboardEntry } from '../../../../services/CepApiModels';
import { getLevelLabel } from '../../../../services/CepLevelPresentation';
import styles from '../CepWelcome.module.scss';
import * as strings from 'CepWelcomeWebPartStrings';

interface IMiniLeaderboardProps {
  leaderboard: ILeaderboardPage;
}

const MEDALS: Record<number, string> = { 1: '🥇', 2: '🥈', 3: '🥉' };

const LEVEL_CLASS: Record<string, string> = {
  Bronze: styles.bronze,
  Silver: styles.silver,
  Gold: styles.gold,
  Explorer: styles.bronze,
  Practitioner: styles.silver,
  Master: styles.gold,
};

const LeaderboardRow: React.FC<{ entry: ILeaderboardEntry }> = ({ entry }) => {
  const rowClass = entry.isCurrentUser ? styles.lbRowCurrentUser : styles.lbRow;
  return (
    <tr className={rowClass}>
      <td className={styles.lbRankCell}>
        {MEDALS[entry.rank] !== undefined ? (
          <span>{MEDALS[entry.rank]}</span>
        ) : (
          `#${entry.rank}`
        )}
      </td>
      <td>
        {entry.displayName}
        {entry.isCurrentUser && <span style={{ marginLeft: 6, opacity: 0.6 }}>← you</span>}
      </td>
      <td className={styles.lbLevelCell}>
        <span className={`${styles.levelPill} ${LEVEL_CLASS[entry.currentLevel] ?? ''}`}>
          {getLevelLabel(entry.currentLevel)}
        </span>
      </td>
      <td className={styles.lbPointsCell}>{entry.monthlyPoints.toLocaleString()} pts</td>
    </tr>
  );
};

export const MiniLeaderboard: React.FC<IMiniLeaderboardProps> = ({ leaderboard }) => {
  const topEntries = leaderboard.entries.slice(0, 5);
  const currentUserInTop = topEntries.some((e) => e.isCurrentUser);
  const showCurrentUserSeparately =
    !currentUserInTop &&
    !!leaderboard.currentUserEntry &&
    (leaderboard.currentUserEntry.rank > 5);

  return (
    <div className={styles.card}>
      <div className={styles.leaderboardHeader}>
        <div className={styles.leaderboardTitle}>
          🏆 Leaderboard – {leaderboard.month}
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 2 }}>
          <a className={styles.leaderboardSeeAll} href="#">
            See full leaderboard →
          </a>
          {leaderboard.lastUpdated && (
            <span style={{ fontSize: 11, opacity: 0.55 }}>
              {strings.LeaderboardLastUpdated.replace('{0}',
                new Date(leaderboard.lastUpdated).toLocaleString(undefined, {
                  day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'
                })
              )}
            </span>
          )}
        </div>
      </div>

      <table className={styles.leaderboardTable}>
        <tbody>
          {topEntries.map((entry) => (
            <LeaderboardRow key={`${entry.rank}-${entry.displayName}`} entry={entry} />
          ))}
          {showCurrentUserSeparately && (
            <>
              <tr className={styles.lbSeparator}>
                <td colSpan={4}>· · ·</td>
              </tr>
              <LeaderboardRow entry={leaderboard.currentUserEntry!} />
            </>
          )}
        </tbody>
      </table>
    </div>
  );
};
