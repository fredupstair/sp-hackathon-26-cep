import * as React from 'react';
import { useState } from 'react';
import { IconButton, TooltipHost } from '@fluentui/react';
import type { IWinItem, IWinItemShared } from '../../../../services/CepApiModels';
import styles from '../CepWins.module.scss';
import * as strings from 'CepWinsWebPartStrings';

interface IWinCardProps {
  win: IWinItem | IWinItemShared;
  showAuthor?: boolean;
}

function isSharedWin(w: IWinItem | IWinItemShared): w is IWinItemShared {
  return 'displayName' in w;
}

export const WinCard: React.FC<IWinCardProps> = ({ win, showAuthor }) => {
  const [copied, setCopied] = useState(false);

  const handleCopy = async (): Promise<void> => {
    if (!win.note) return;
    await navigator.clipboard.writeText(win.note);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const dateLabel = new Date(win.date + 'T00:00:00').toLocaleDateString(undefined, {
    day: 'numeric',
    month: 'short',
  });

  return (
    <div className={styles.winCard}>
      <div className={styles.winCardMeta}>
        <span className={styles.winDate}>{dateLabel}</span>
        {showAuthor && isSharedWin(win) && (
          <span className={styles.winAuthor}>{win.displayName}</span>
        )}
      </div>
      <p className={styles.winNote}>{win.note || <em>—</em>}</p>
      <div className={styles.winCardActions}>
        <TooltipHost content={copied ? strings.Copied : strings.CopyText}>
          <IconButton
            iconProps={{ iconName: copied ? 'CheckMark' : 'Copy' }}
            ariaLabel={strings.CopyText}
            className={copied ? styles.copiedBtn : ''}
            onClick={handleCopy}
            disabled={!win.note}
          />
        </TooltipHost>
      </div>
    </div>
  );
};
