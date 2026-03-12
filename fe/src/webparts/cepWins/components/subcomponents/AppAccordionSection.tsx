import * as React from 'react';
import { useState } from 'react';
import { Icon } from '@fluentui/react';
import styles from '../CepWins.module.scss';
import { WinCard } from './WinCard';
import type { IWinItem, IWinItemShared } from '../../../../services/CepApiModels';
import { getAppLabel, getAppFluentIcon } from '../../../../services/CepAppLabels';

const INITIAL_VISIBLE = 5;

interface IAppAccordionSectionProps {
  appKey: string;
  wins: (IWinItem | IWinItemShared)[];
  showAuthor?: boolean;
  defaultExpanded?: boolean;
}

export const AppAccordionSection: React.FC<IAppAccordionSectionProps> = ({
  appKey,
  wins,
  showAuthor,
  defaultExpanded = true,
}) => {
  const [expanded, setExpanded] = useState(defaultExpanded);
  const [showAll, setShowAll] = useState(false);

  const visibleWins = showAll ? wins : wins.slice(0, INITIAL_VISIBLE);
  const hiddenCount = wins.length - INITIAL_VISIBLE;

  return (
    <div className={styles.accordionSection}>
      <button
        className={styles.accordionHeader}
        onClick={() => setExpanded(!expanded)}
        aria-expanded={expanded}
      >
        <span className={`${styles.accordionChevron}${expanded ? ' ' + styles.expanded : ''}`}>▶</span>
        <span className={styles.accordionIcon}>
          <Icon
            iconName={getAppFluentIcon(appKey).iconName}
            style={{ color: getAppFluentIcon(appKey).color, fontSize: 18 }}
          />
        </span>
        <span className={styles.accordionTitle}>{getAppLabel(appKey)}</span>
        <span className={styles.accordionCount}>{wins.length}</span>
      </button>

      {expanded && (
        <div className={styles.accordionBody}>
          {visibleWins.map(win => (
            <WinCard key={win.id} win={win} showAuthor={showAuthor} />
          ))}
          {!showAll && hiddenCount > 0 && (
            <button
              className={styles.showMoreBtn}
              onClick={() => setShowAll(true)}
            >
              + Mostra altri {hiddenCount}
            </button>
          )}
        </div>
      )}
    </div>
  );
};
