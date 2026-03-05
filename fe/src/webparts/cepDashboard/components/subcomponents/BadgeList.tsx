import * as React from 'react';
import type { ICepBadge } from '../../../../services/CepApiModels';
import styles from '../CepDashboard.module.scss';
import * as strings from 'CepDashboardWebPartStrings';

interface IBadgeListProps {
  badges: ICepBadge[];
}

const BADGE_EMOJI: Record<string, string> = {
  FirstSteps: '🚀',
  CrossAppExplorer: '🌐',
  WeeklyWarrior: '⚔️',
  MonthlyMaster: '👑',
  ConsistencyKing: '🔥',
};

const fmt = (tpl: string, ...args: (string | number)[]): string =>
  tpl.replace(/{(\d+)}/g, (m, i) => (args[i] !== undefined ? String(args[i]) : m));

export const BadgeList: React.FC<IBadgeListProps> = ({ badges }) => (
  <div className={styles.card}>
    <div className={styles.sectionTitle}>{strings.BadgesTitle}</div>
    {badges.length === 0 ? (
      <div className={styles.emptyState}>{strings.NoBadgesYet}</div>
    ) : (
      <div className={styles.badgesGrid}>
        {badges.map((badge) => (
          <div
            key={badge.badgeKey}
            className={styles.badgeItem}
            title={badge.description}
          >
            <div className={styles.badgeIcon}>
              {badge.iconUrl ? (
                <img src={badge.iconUrl} alt={badge.badgeName} />
              ) : (
                <span>{BADGE_EMOJI[badge.badgeKey] ?? '🏅'}</span>
              )}
            </div>
            <div className={styles.badgeName}>{badge.badgeName}</div>
            <div className={styles.badgeDate}>
              {fmt(strings.BadgeEarnedOn, new Date(badge.earnedDate).toLocaleDateString())}
            </div>
          </div>
        ))}
      </div>
    )}
  </div>
);
