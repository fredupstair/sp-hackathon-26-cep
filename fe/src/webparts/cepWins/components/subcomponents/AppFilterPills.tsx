import * as React from 'react';
import { Icon } from '@fluentui/react';
import { getAppLabel, getAppFluentIcon } from '../../../../services/CepAppLabels';
import styles from '../CepWins.module.scss';
import * as strings from 'CepWinsWebPartStrings';

interface IAppFilterPillsProps {
  availableAppKeys: string[];
  selectedAppKey: string | undefined;
  onSelect: (appKey: string | undefined) => void;
}

export const AppFilterPills: React.FC<IAppFilterPillsProps> = ({
  availableAppKeys,
  selectedAppKey,
  onSelect,
}) => {
  return (
    <div className={styles.filterPills}>
      <button
        className={`${styles.pill}${!selectedAppKey ? ' ' + styles.pillActive : ''}`}
        onClick={() => onSelect(undefined)}
      >
        {strings.FilterAll}
      </button>
      {availableAppKeys.map(key => (
        <button
          key={key}
          className={`${styles.pill}${selectedAppKey === key ? ' ' + styles.pillActive : ''}`}
          onClick={() => onSelect(key)}
        >
          <Icon
            iconName={getAppFluentIcon(key).iconName}
            style={{ color: getAppFluentIcon(key).color, fontSize: 14, marginRight: 4, verticalAlign: 'middle' }}
          />
          {getAppLabel(key)}
        </button>
      ))}
    </div>
  );
};
