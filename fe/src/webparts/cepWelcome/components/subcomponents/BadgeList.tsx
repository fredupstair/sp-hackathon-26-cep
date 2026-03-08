import * as React from 'react';
import type { ICepBadge } from '../../../../services/CepApiModels';
import styles from '../CepWelcome.module.scss';
import * as strings from 'CepWelcomeWebPartStrings';

interface IBadgeListProps {
  badges: ICepBadge[];
  /** Show skeleton placeholders */
  loading?: boolean;
}

const BADGE_EMOJI: Record<string, string> = {
  FirstSteps: '🚀',
  CrossAppExplorer: '🌐',
  WeeklyWarrior: '⚔️',
  MonthlyMaster: '👑',
  ConsistencyKing: '🔥',
};

export const BadgeList: React.FC<IBadgeListProps> = ({ badges, loading }) => {
  if (loading) {
    return (
      <div className={styles.card}>
        <div className={`${styles.skeletonLine} ${styles.skeletonLineMedium}`} style={{ height: 16, marginBottom: 16 }} />
        {[0, 1].map((i) => (
          <div key={i} className={styles.badgeTimelineItem}>
            <div className={`${styles.skeletonLine}`} style={{ width: 70, marginBottom: 0 }} />
            <div className={styles.badgeTimelineDotCol}>
              <div className={styles.skeletonCircle} style={{ width: 12, height: 12 }} />
            </div>
            <div className={styles.badgeTimelineContent}>
              <div className={styles.skeletonCircle} style={{ width: 40, height: 40 }} />
              <div style={{ flex: 1 }}>
                <div className={`${styles.skeletonLine} ${styles.skeletonLineMedium}`} />
                <div className={`${styles.skeletonLine} ${styles.skeletonLineLong}`} />
              </div>
            </div>
          </div>
        ))}
      </div>
    );
  }

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
