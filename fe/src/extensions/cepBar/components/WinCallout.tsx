import * as React from 'react';
import { Callout, DirectionalHint } from '@fluentui/react';
import * as strings from 'CepBarApplicationCustomizerStrings';
import styles from './CepBar.module.scss';

interface IAppOption {
  key: string;
  text: string;
  initial: string;
  bg: string;
}

const APP_OPTIONS: IAppOption[] = [
  { key: 'Word',        text: 'Word',         initial: 'W',  bg: '#185ABD' },
  { key: 'Excel',       text: 'Excel',        initial: 'X',  bg: '#107C41' },
  { key: 'PowerPoint',  text: 'PPT',          initial: 'P',  bg: '#C43E1C' },
  { key: 'Outlook',     text: 'Outlook',      initial: 'O',  bg: '#0078D4' },
  { key: 'Teams',       text: 'Teams',        initial: 'T',  bg: '#6264A7' },
  { key: 'OneNote',     text: 'OneNote',      initial: 'N',  bg: '#7719AA' },
  { key: 'Loop',        text: 'Loop',         initial: 'L',  bg: '#0F6CBD' },
  { key: 'BizChat',     text: 'Copilot Chat', initial: '✨', bg: '#0F6CBD' },
  { key: 'WebChat',     text: 'Web Chat',     initial: '🌐', bg: '#4A4A6A' },
  { key: 'M365App',     text: 'M365 App',     initial: 'M',  bg: '#D83B01' },
  { key: 'Forms',       text: 'Forms',        initial: 'F',  bg: '#007744' },
  { key: 'SharePoint',  text: 'SharePoint',   initial: 'SP', bg: '#0078D4' },
  { key: 'Whiteboard',  text: 'Whiteboard',   initial: 'WB', bg: '#0091DA' },
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
            <div className={styles.appGrid}>
              {APP_OPTIONS.map((o) => (
                <button
                  key={o.key}
                  type="button"
                  className={`${styles.appBtn} ${appKey === o.key ? styles.appBtnSelected : ''}`}
                  onClick={() => this.setState({ appKey: o.key })}
                  disabled={submitting}
                  title={o.text}
                >
                  <div
                    className={styles.appIcon}
                    style={{ background: o.bg }}
                  >
                    {o.initial}
                  </div>
                  <span className={styles.appLabel}>{o.text}</span>
                </button>
              ))}
            </div>
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
