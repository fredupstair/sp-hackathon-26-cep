import * as React from 'react';
import { Callout, DirectionalHint } from '@fluentui/react';
import * as strings from 'CepBarApplicationCustomizerStrings';
import styles from './CepBar.module.scss';

const APP_OPTIONS: Array<{ key: string; text: string }> = [
  { key: 'Word',        text: 'Word' },
  { key: 'Excel',       text: 'Excel' },
  { key: 'PowerPoint',  text: 'PowerPoint' },
  { key: 'Outlook',     text: 'Outlook' },
  { key: 'Teams',       text: 'Teams' },
  { key: 'OneNote',     text: 'OneNote' },
  { key: 'Loop',        text: 'Loop' },
  { key: 'BizChat',     text: 'Copilot Chat' },
  { key: 'WebChat',     text: 'Web Chat' },
  { key: 'M365App',     text: 'M365 App' },
  { key: 'Forms',       text: 'Forms' },
  { key: 'SharePoint',  text: 'SharePoint' },
  { key: 'Whiteboard',  text: 'Whiteboard' },
];

interface IWinCalloutProps {
  target: HTMLButtonElement;
  onDismiss: () => void;
  onSubmit: (appKey: string, note: string, isShared: boolean) => Promise<void>;
}

interface IWinCalloutState {
  appKey: string;
  note: string;
  isShared: boolean;
  submitting: boolean;
  limitReached: boolean;
}

export class WinCallout extends React.Component<IWinCalloutProps, IWinCalloutState> {
  constructor(props: IWinCalloutProps) {
    super(props);
    this.state = {
      appKey: 'Word',
      note: '',
      isShared: false,
      submitting: false,
      limitReached: false,
    };
  }

  private _handleSubmit = async (): Promise<void> => {
    const { onSubmit } = this.props;
    const { appKey, note, isShared, submitting } = this.state;
    if (submitting) return;
    this.setState({ submitting: true });
    try {
      await onSubmit(appKey, note, isShared);
      // parent dismisses callout on success
    } catch (e: unknown) {
      const status = (e as { statusCode?: number })?.statusCode;
      if (status === 429) {
        this.setState({ limitReached: true, submitting: false });
      } else {
        this.setState({ submitting: false });
      }
    }
  };

  public render(): React.ReactElement {
    const { target, onDismiss } = this.props;
    const { appKey, note, isShared, submitting, limitReached } = this.state;

    return (
      <Callout
        target={target}
        onDismiss={onDismiss}
        directionalHint={DirectionalHint.topCenter}
        isBeakVisible
        gapSpace={6}
        setInitialFocus
      >
        <div className={styles.winCallout}>
          <div className={styles.winCalloutTitle}>🏆 {strings.WinCalloutTitle}</div>

          <div className={styles.winCalloutField}>
            <label className={styles.winCalloutLabel}>{strings.WinCalloutApp}</label>
            <select
              className={styles.winCalloutSelect}
              value={appKey}
              onChange={(e) => this.setState({ appKey: e.target.value })}
              disabled={submitting}
            >
              {APP_OPTIONS.map((o) => (
                <option key={o.key} value={o.key}>{o.text}</option>
              ))}
            </select>
          </div>

          <div className={styles.winCalloutField}>
            <label className={styles.winCalloutLabel}>{strings.WinCalloutNote}</label>
            <textarea
              className={styles.winCalloutTextarea}
              maxLength={140}
              value={note}
              placeholder={strings.WinCalloutNotePlaceholder}
              onChange={(e) => this.setState({ note: e.target.value })}
              disabled={submitting}
            />
          </div>

          <div className={styles.winCalloutCheckboxRow}>
            <input
              type="checkbox"
              id="cep-win-share"
              checked={isShared}
              onChange={(e) => this.setState({ isShared: e.target.checked })}
              disabled={submitting}
            />
            <label htmlFor="cep-win-share">{strings.WinCalloutShare}</label>
          </div>

          {limitReached && (
            <div className={styles.winLimitMsg}>{strings.WinLimitReached}</div>
          )}

          <div className={styles.winCalloutActions}>
            <button className={styles.winCalloutCancel} onClick={onDismiss} disabled={submitting}>
              {strings.Cancel}
            </button>
            <button
              className={styles.winCalloutSubmit}
              onClick={this._handleSubmit}
              disabled={submitting || limitReached}
            >
              ✨ {strings.WinCalloutSubmit}
            </button>
          </div>
        </div>
      </Callout>
    );
  }
}
