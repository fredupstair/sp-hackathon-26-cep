import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import * as strings from 'CepBarApplicationCustomizerStrings';
import type { CepApiClient } from '../../../services/CepApiClient';
import type { IUserSummary, ICepSuggestion } from '../../../services/CepApiModels';
import { getLevelIcon, getLevelLabel, getNextLevelLabel, isTopLevel } from '../../../services/CepLevelPresentation';
import { WinCallout } from './WinCallout';
import { JoinOverlay } from './JoinOverlay';
import styles from './CepBar.module.scss';

// ─── Props ────────────────────────────────────────────────────────────────────

export interface ICepBarProps {
  apiClient: CepApiClient | undefined;
  dashboardPageUrl: string;
  optinPageUrl: string;
  practitionerThreshold: number;
  masterThreshold: number;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function p2(n: number): string { return n < 10 ? '0' + n : '' + n; }

function formatDate(d: Date): string {
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`;
}

function computeStreak(recentActiveDays: string[]): number {
  if (!recentActiveDays.length) return 0;
  const daysSet = new Set(recentActiveDays);
  const today     = formatDate(new Date());
  const yesterday = formatDate(new Date(Date.now() - 86400000));

  const startStr = daysSet.has(today) ? today : daysSet.has(yesterday) ? yesterday : null;
  if (!startStr) return 0;

  let streak = 0;
  const pointer = new Date(startStr + 'T00:00:00');
  while (daysSet.has(formatDate(pointer))) {
    streak++;
    pointer.setDate(pointer.getDate() - 1);
  }
  return streak;
}

function levelColorClass(level: string): string {
  if (level === 'Master') return styles.gold;
  if (level === 'Practitioner') return styles.silver;
  return styles.bronze;
}

function progressClass(level: string): string {
  if (level === 'Practitioner') return styles.progressSilver;
  if (level === 'Master') return styles.progressGold;
  return styles.progressBronze;
}

function fmt(tpl: string, ...args: (string | number)[]): string {
  return tpl.replace(/{(\d+)}/g, (m, i) => (args[+i] !== undefined ? String(args[+i]) : m));
}

// ─── State ────────────────────────────────────────────────────────────────────

type LoadState = 'loading' | 'not_configured' | 'not_enrolled' | 'ready' | 'error';

interface ICepBarState {
  loadState: LoadState;
  summary: IUserSummary | undefined;
  suggestion: ICepSuggestion | undefined;
  streak: number;
  open: boolean;
  showWin: boolean;
  showJoin: boolean;
  winAnim: boolean;
  winAnimText: string;
  suggestionDismissed: boolean;
  nudgesEnabled: boolean;
  nudgeSaving: boolean;
  nudgeChecking: boolean;
}

// ─── Component ────────────────────────────────────────────────────────────────

export class CepBar extends React.Component<ICepBarProps, ICepBarState> {
  private readonly _winBtnRef = React.createRef<HTMLButtonElement>();
  private readonly _hostRef = React.createRef<HTMLDivElement>();

  constructor(props: ICepBarProps) {
    super(props);
    this.state = {
      loadState: 'loading',
      summary: undefined,
      suggestion: undefined,
      streak: 0,
      open: false,
      showWin: false,
      showJoin: false,
      winAnim: false,
      winAnimText: '',
      suggestionDismissed: false,
      nudgesEnabled: true,
      nudgeSaving: false,
      nudgeChecking: false,
    };
  }

  public componentDidMount(): void {
    this._load().catch(console.error);
    document.addEventListener('mousedown', this._handleClickOutside);
  }

  public componentWillUnmount(): void {
    document.removeEventListener('mousedown', this._handleClickOutside);
  }

  private _handleClickOutside = (e: MouseEvent): void => {
    if (this._hostRef.current && !this._hostRef.current.contains(e.target as Node)) {
      this.setState({ open: false, showWin: false });
    }
  };

  private async _load(): Promise<void> {
    const { apiClient } = this.props;
    if (!apiClient) {
      this.setState({ loadState: 'not_configured' });
      return;
    }

    try {
      const summary = await apiClient.getMeSummary();
      if (!summary) {
        this.setState({ loadState: 'not_enrolled' });
        return;
      }

      const streak = computeStreak(summary.recentActiveDays ?? []);

      let suggestion: ICepSuggestion | undefined;
      try {
        suggestion = await apiClient.getSuggestion();
      } catch { /* ignore */ }

      this.setState({ loadState: 'ready', summary, streak, suggestion, nudgesEnabled: summary.isEngagementNudgesEnabled });
    } catch {
      this.setState({ loadState: 'error' });
    }
  }

  private _handleWinSubmit = async (appKey: string, note: string, isShared: boolean): Promise<void> => {
    const { apiClient } = this.props;
    if (!apiClient) return;

    const resp = await apiClient.postWin({ appKey, note: note || undefined, isShared });
    const animText = fmt(strings.WinAnimText, resp.pointsAdded);

    this.setState((s) => ({
      showWin: false,
      winAnim: true,
      winAnimText: animText,
      summary: s.summary
        ? { ...s.summary, monthlyPoints: resp.totalMonthlyPoints }
        : s.summary,
    }));

    setTimeout(() => this.setState({ winAnim: false }), 1600);
  };

  private _toggleNudges = async (): Promise<void> => {
    const { apiClient } = this.props;
    if (!apiClient || this.state.nudgeSaving) return;
    const next = !this.state.nudgesEnabled;
    this.setState({ nudgesEnabled: next, nudgeSaving: true });
    try {
      await apiClient.updateMePreferences({ isEngagementNudgesEnabled: next });
    } catch {
      this.setState({ nudgesEnabled: !next });  // rollback on error
    } finally {
      this.setState({ nudgeSaving: false });
    }
  };

  private _toggleOpen = (): void => {
    const opening = !this.state.open;
    this.setState((s) => ({
      open: !s.open,
      showWin: false,
      nudgeChecking: opening,
    }));
    if (opening) {
      this._fetchNudgePreference()
        .catch(() => { /* keep cached value */ })
        .finally(() => this.setState({ nudgeChecking: false }));
    }
  };

  private _fetchNudgePreference = async (): Promise<void> => {
    const { apiClient } = this.props;
    if (!apiClient) return;
    const prefs = await apiClient.getMePreferences();
    this.setState({ nudgesEnabled: prefs.isEngagementNudgesEnabled });
  };

  public render(): React.ReactElement {
    const { loadState, summary, suggestion, streak, open, showWin, showJoin, winAnim, winAnimText, suggestionDismissed, nudgesEnabled, nudgeSaving, nudgeChecking } = this.state;
    const { dashboardPageUrl, optinPageUrl, practitionerThreshold, masterThreshold } = this.props;

    // Silently hide when not configured or error
    if (loadState === 'not_configured' || loadState === 'error') return <></>;

    // ── Loading chip ────────────────────────────────────────────────────────
    if (loadState === 'loading') {
      return (
        <div className={styles.cepHost} ref={this._hostRef}>
          <div className={styles.chipLoading}>
            <Spinner size={SpinnerSize.xSmall} />
            <span className={styles.barLoading}>{strings.BarLoading}</span>
          </div>
        </div>
      );
    }

    // ── Not enrolled chip ───────────────────────────────────────────────────
    if (loadState === 'not_enrolled') {
      return (
        <div className={styles.cepHost} ref={this._hostRef}>
          <div
            className={styles.chipNotEnrolled}
            onClick={() => this.setState({ showJoin: true })}
            role="button"
            tabIndex={0}
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') this.setState({ showJoin: true }); }}
          >
            <span className={styles.programName}>✦ Copilot Engagement Program</span>
            <span className={styles.joinBtn}>{strings.BarNotEnrolled}</span>
          </div>
          {showJoin && (
            <JoinOverlay
              programUrl={optinPageUrl || '#'}
              onDismiss={() => this.setState({ showJoin: false })}
            />
          )}
        </div>
      );
    }

    // ── Enrolled ────────────────────────────────────────────────────────────
    const s = summary!;
    const { currentLevel: level, monthlyPoints: monthly } = s;
    const levelLabel = getLevelLabel(level);

    let progressPct: number;
    let nextLevelLabel: string;
    let ptToGo: number;

    if (levelLabel === 'Explorer') {
      progressPct   = Math.min(monthly / practitionerThreshold, 1) * 100;
      ptToGo        = Math.max(0, practitionerThreshold - monthly);
      nextLevelLabel = getNextLevelLabel(level);
    } else if (levelLabel === 'Practitioner') {
      progressPct   = Math.min(monthly / masterThreshold, 1) * 100;
      ptToGo        = Math.max(0, masterThreshold - monthly);
      nextLevelLabel = getNextLevelLabel(level);
    } else {
      progressPct   = 100;
      ptToGo        = 0;
      nextLevelLabel = '';
    }

    const showSuggestion = !suggestionDismissed && !!suggestion;

    return (
      <div className={styles.cepHost} ref={this._hostRef}>
        {/* Win animation */}
        {winAnim && (
          <span className={styles.winAnim} aria-live="polite">{winAnimText}</span>
        )}

        {/* Flyout */}
        {open && (
          <div className={styles.flyout}>
            {/* Header */}
            <div className={styles.flyoutHeader}>
              <div className={styles.flyoutNameRow}>
                <span className={styles.flyoutDisplayName}>{s.displayName}</span>
                <button
                  className={styles.nudgeBtn}
                  onClick={this._toggleNudges}
                  title={nudgesEnabled ? strings.NudgesTooltipOn : strings.NudgesTooltipOff}
                  aria-label={nudgesEnabled ? strings.NudgesTooltipOn : strings.NudgesTooltipOff}
                  disabled={nudgeSaving || nudgeChecking}
                >
                  {nudgeChecking ? '🔔' : nudgesEnabled ? '🔔' : '🔕'}
                </button>
              </div>
              <div className={styles.flyoutLevelRow}>
                <span className={`${styles.levelPill} ${levelColorClass(level)}`}>
                  {getLevelIcon(level)} {levelLabel}
                </span>
                <span className={styles.flyoutPoints}>{fmt(strings.BarPoints, monthly)}</span>
              </div>
            </div>

            {/* Progress */}
            {!isTopLevel(level) && (
              <div className={styles.progressGroup}>
                <div className={styles.progressTrack}>
                  <div
                    className={`${styles.progressFill} ${progressClass(level)}`}
                    style={{ width: `${progressPct}%` }}
                  />
                </div>
                <div className={styles.progressLabel}>
                  {fmt(strings.BarProgressLabel, ptToGo, nextLevelLabel)}
                </div>
              </div>
            )}

            {/* Suggestion */}
            {showSuggestion && (
              <div className={styles.suggestion}>
                <span className={styles.suggestionIcon}>💡</span>
                <span className={styles.suggestionText}>{suggestion!.text}</span>
                <button
                  className={styles.dismissBtn}
                  onClick={() => this.setState({ suggestionDismissed: true })}
                  title={strings.DismissSuggestion}
                  aria-label={strings.DismissSuggestion}
                >×</button>
              </div>
            )}

            {/* Actions */}
            <div className={styles.flyoutActions}>
              <button
                ref={this._winBtnRef}
                className={styles.winBtn}
                onClick={() => this.setState((prev) => ({ showWin: !prev.showWin }))}
                aria-expanded={showWin}
              >
                ✨ {strings.BarMarkWin}
              </button>

              {dashboardPageUrl && (
                <a href={dashboardPageUrl} className={styles.dashBtn}>
                  {strings.BarViewDashboard} →
                </a>
              )}
            </div>

            {/* WinCallout */}
            {showWin && this._winBtnRef.current && (
              <WinCallout
                target={this._winBtnRef.current}
                onDismiss={() => this.setState({ showWin: false })}
                onSubmit={this._handleWinSubmit}
              />
            )}
          </div>
        )}

        {/* Chip */}
        <div className={styles.chip} onClick={this._toggleOpen} role="button" tabIndex={0} aria-expanded={open}>
          {streak > 0 && (
            <>
              <span className={styles.streak} title={fmt(strings.StreakTooltip, streak)}>
                🔥{fmt(strings.StreakLabel, streak)}
              </span>
              <span className={styles.separator} />
            </>
          )}

          <span className={`${styles.levelPill} ${levelColorClass(level)}`}>
            {getLevelIcon(level)}
          </span>

          <span className={styles.points}>{fmt(strings.BarPoints, monthly)}</span>

          <span className={`${styles.chevron} ${open ? styles.chevronOpen : ''}`}>▲</span>
        </div>
      </div>
    );
  }
}
