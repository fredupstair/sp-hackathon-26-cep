import * as React from 'react';
import type { ICepBadge } from '../../../../services/CepApiModels';
import styles from '../CepWelcome.module.scss';
import * as strings from 'CepWelcomeWebPartStrings';

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

export const BadgeList: React.FC<IBadgeListProps> = ({ badges }) => {
  const sorted = [...badges].sort(
    (a, b) => new Date(b.earnedDate).getTime() - new Date(a.earnedDate).getTime()
  );

  return (
    <div className={styles.card}>
      <div className={styles.sectionTitle}>{strings.BadgesTitle}</div>
      {sorted.length === 0 ? (
        <div className={styles.emptyState}>{strings.NoBadgesYet}</div>
      ) : (
        <div className={styles.badgeTimeline}>
          {sorted.map((badge, index) => (
            <div key={badge.badgeKey} className={styles.badgeTimelineItem}>
              <div className={styles.badgeTimelineDate}>
                {new Date(badge.earnedDate).toLocaleDateString(undefined, {
                  day: '2-digit', month: 'short', year: 'numeric',
                })}
              </div>
              <div className={styles.badgeTimelineDotCol}>
                <div className={styles.badgeTimelineDot} />
                {index < sorted.length - 1 && (
                  <div className={styles.badgeTimelineLine} />
                )}
              </div>
              <div className={styles.badgeTimelineContent}>
                <div className={styles.badgeTimelineIcon}>
                  {badge.iconUrl ? (
                    <img src={badge.iconUrl} alt={badge.badgeName} />
                  ) : (
                    <span>{BADGE_EMOJI[badge.badgeKey] ?? '🏅'}</span>
                  )}
                </div>
                <div>
                  <div className={styles.badgeName}>{badge.badgeName}</div>
                  <div className={styles.badgeDate}>{badge.description}</div>
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
