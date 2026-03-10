import * as React from 'react';
import * as strings from 'CepBarApplicationCustomizerStrings';
import styles from './CepBar.module.scss';

interface IJoinOverlayProps {
  programUrl: string;
  onDismiss: () => void;
}

export class JoinOverlay extends React.Component<IJoinOverlayProps> {
  public render(): React.ReactElement {
    const { programUrl, onDismiss } = this.props;

    return (
      <div
        className={styles.joinOverlayBackdrop}
        onClick={onDismiss}
        role="dialog"
        aria-modal="true"
        aria-label="Join Copilot Engagement Program"
      >
        <div
          className={styles.joinOverlayCard}
          onClick={(e) => e.stopPropagation()}
        >
          <button
            className={styles.joinOverlayClose}
            onClick={onDismiss}
            aria-label="Close"
          >
            ×
          </button>

          <div className={styles.joinOverlayLogo}>✦</div>
          <h2 className={styles.joinOverlayTitle}>Copilot Engagement Program</h2>

          <div className={styles.joinOverlayBody}>
            <p className={styles.joinOverlayTextEN}>
              Join the <strong>Copilot Engagement Program</strong> and start earning
              points for every Copilot interaction! Track your progress, climb the
              leaderboard, and unlock exclusive badges.
            </p>
            <p className={styles.joinOverlayTextIT}>
              Iscriviti al <strong>Copilot Engagement Program</strong> e inizia a
              guadagnare punti per ogni interazione con Copilot! Monitora i tuoi
              progressi, scala la classifica e sblocca badge esclusivi.
            </p>
          </div>

          <a
            href={programUrl}
            className={styles.joinOverlayCta}
            target="_blank"
            rel="noreferrer"
          >
            {strings.JoinOverlayCta}
          </a>
        </div>
      </div>
    );
  }
}
