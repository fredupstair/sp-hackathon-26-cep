import * as React from 'react';
import { Persona, PersonaSize } from '@fluentui/react';
import type { IUserSummary, ICepBadge } from '../../../../services/CepApiModels';
import styles from '../CepWelcome.module.scss';
import * as strings from 'CepWelcomeWebPartStrings';
import { BadgeList } from './BadgeList';

interface IOtherUserViewProps {
  summary: IUserSummary;
  badges: ICepBadge[];
}

const LEVEL_ICON: Record<string, string> = {
  Bronze: '🥉',
  Silver: '🥈',
  Gold: '🥇',
};

const LEVEL_CLASS: Record<string, string> = {
  Bronze: styles.bronze,
  Silver: styles.silver,
  Gold: styles.gold,
};

export const OtherUserView: React.FC<IOtherUserViewProps> = ({ summary, badges }) => (
  <>
    {/* Header card */}
    <div className={styles.otherUserHeader}>
      <Persona
        text={summary.displayName}
        size={PersonaSize.size56}
        secondaryText={`${summary.department}${summary.team ? ' · ' + summary.team : ''}`}
      />
      <span className={`${styles.levelPill} ${LEVEL_CLASS[summary.currentLevel] ?? ''}`}>
        {LEVEL_ICON[summary.currentLevel]} {summary.currentLevel}
      </span>
    </div>

    {/* Aggregated stats (no app breakdown – privacy) */}
    <div className={styles.statsRow}>
      <div className={styles.statCard}>
        <div className={styles.statLabel}>{strings.OtherUserPointsThisMonth}</div>
        <div className={styles.statValue}>{summary.monthlyPoints.toLocaleString()}</div>
      </div>
      <div className={styles.statCard}>
        <div className={styles.statLabel}>{strings.OtherUserTotalPoints}</div>
        <div className={styles.statValue}>{summary.totalPoints.toLocaleString()}</div>
      </div>
      <div className={styles.statCard}>
        <div className={styles.statLabel}>{strings.OtherUserLevel}</div>
        <div className={styles.statValue} style={{ fontSize: 22 }}>
          {summary.currentLevel}
        </div>
      </div>
    </div>

    {/* Badges are public */}
    <BadgeList badges={badges} />
  </>
);
