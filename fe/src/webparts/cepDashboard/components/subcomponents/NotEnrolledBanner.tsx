import * as React from 'react';
import { PrimaryButton } from '@fluentui/react';
import styles from '../CepDashboard.module.scss';
import * as strings from 'CepDashboardWebPartStrings';

interface INotEnrolledBannerProps {
  /** Optional callback – if omitted the CTA button is not rendered */
  onJoinClick?: () => void;
}

export const NotEnrolledBanner: React.FC<INotEnrolledBannerProps> = ({ onJoinClick }) => (
  <div className={styles.notEnrolledBanner}>
    <div className={styles.notEnrolledIcon}>🤖</div>
    <div className={styles.notEnrolledTitle}>{strings.NotEnrolledTitle}</div>
    <div className={styles.notEnrolledMessage}>{strings.NotEnrolledMessage}</div>
    {onJoinClick && (
      <PrimaryButton text={strings.JoinProgram} onClick={onJoinClick} />
    )}
  </div>
);
