import * as React from 'react';
import { Persona, PersonaSize } from '@fluentui/react';
import type { IUserSummary } from '../../../../services/CepApiModels';
import styles from '../CepWelcome.module.scss';

interface IDashboardHeaderProps {
  summary: IUserSummary;
  month: string;
  isCurrentMonth: boolean;
  onPrevMonth: () => void;
  onNextMonth: () => void;
}

function formatMonthLabel(month: string): string {
  const [year, m] = month.split('-').map(Number);
  return new Date(year, m - 1, 1).toLocaleDateString(undefined, {
    month: 'long',
    year: 'numeric',
  });
}

export const DashboardHeader: React.FC<IDashboardHeaderProps> = ({
  summary,
  month,
  isCurrentMonth,
  onPrevMonth,
  onNextMonth,
}) => (
  <div className={styles.headerCard}>
    <Persona
      text={summary.displayName}
      size={PersonaSize.size72}
      secondaryText={summary.email}
      hidePersonaDetails
      imageUrl={`/_layouts/15/userphoto.aspx?AccountName=${encodeURIComponent(summary.email)}&Size=L`}
    />
    <div className={styles.headerInfo}>
      <div className={styles.displayName}>{summary.displayName}</div>
      <div className={styles.headerMeta}>
        {summary.department}
        {summary.team ? ` · ${summary.team}` : ''}
      </div>
    </div>

    <div className={styles.monthPicker}>
      <button className={styles.monthBtn} onClick={onPrevMonth} title="Previous month">
        ‹
      </button>
      <span className={styles.monthLabel}>{formatMonthLabel(month)}</span>
      <button
        className={styles.monthBtn}
        onClick={onNextMonth}
        disabled={isCurrentMonth}
        title="Next month"
      >
        ›
      </button>
    </div>
  </div>
);
