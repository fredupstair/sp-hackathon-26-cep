import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import * as strings from 'CepBarApplicationCustomizerStrings';
import type { CepApiClient } from '../../../services/CepApiClient';
import type { IUserSummary, ICepSuggestion } from '../../../services/CepApiModels';
import { WinCallout } from './WinCallout';
import styles from './CepBar.module.scss';

// ─── Props ────────────────────────────────────────────────────────────────────

export interface ICepBarProps {
  apiClient: CepApiClient | undefined;
  dashboardPageUrl: string;
  optinPageUrl: string;
  silverThreshold: number;
  goldThreshold: number;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function p2(n: number): string { return n < 10 ? '0' + n : '' + n; }

function formatDate(d: Date): string {
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`;
}

/**
 * Computes the current streak from a set of recent active day strings (YYYY-MM-DD).
 * A streak is active if today or yesterday is in the set.
 */
function computeStreak(recentActiveDays: string[]): number {
  if (!recentActiveDays.length) return 0;
  const daysSet = new Set(recentActiveDays);
  const today     = formatDate(new Date());
  const yesterday = formatDate(new Date(Date.now() - 86400000));

  // Streak must be anchored to today or yesterday
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
  if (level === 'Gold')   return styles.gold;
  if (level === 'Silver') return styles.silver;
  return styles.bronze;
}

function levelEmoji(level: string): string {
  if (level === 'Gold')   return '🥇';
  if (level === 'Silver') return '🥈';
  return '🥉';
}

function progressClass(level: string): string {
  if (level === 'Silver') return styles.progressSilver;
  if (level === 'Gold')   return styles.progressGold;
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
  showWin: boolean;
  winAnim: boolean;
  winAnimText: string;
  suggestionDismissed: boolean;
}

// ─── Component ────────────────────────────────────────────────────────────────

export class CepBar extends React.Component<ICepBarProps, ICepBarState> {
  private readonly _winBtnRef = React.createRef<HTMLButtonElement>();

  constructor(props: ICepBarProps) {
    super(props);
    this.state = {
      loadState: 'loading',
      summary: undefined,
      suggestion: undefined,
      streak: 0,
      showWin: false,
      winAnim: false,
      winAnimText: '',
      suggestionDismissed: false,
    };
  }

  public componentDidMount(): void {
    this._load().catch(console.error);
  }

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

      // fetch suggestion in background; non-critical
      let suggestion: ICepSuggestion | undefined;
      try {
        suggestion = await apiClient.getSuggestion();
      } catch { /* ignore */ }

      this.setState({ loadState: 'ready', summary, streak, suggestion });
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

  public render(): React.ReactElement {
    const { loadState, summary, suggestion, streak, showWin, winAnim, winAnimText, suggestionDismissed } = this.state;
    const { dashboardPageUrl, optinPageUrl, silverThreshold, goldThreshold } = this.props;

    // Silently hide bar when not configured or in error
    if (loadState === 'not_configured' || loadState === 'error') return <></>;

    if (loadState === 'loading') {
      return (
        <div className={styles.cepBar}>
          <div className={styles.barLoading}>
            <Spinner size={SpinnerSize.xSmall} />
            <span>{strings.BarLoading}</span>
          </div>
        </div>
      );
    }

    if (loadState === 'not_enrolled') {
      return (
        <div className={styles.notEnrolledBar}>
          <span className={styles.programName}>✦ Copilot Engagement Program</span>
          <span className={styles.suggestionText}>{strings.BarNotEnrolledDesc}</span>
          {optinPageUrl && (
            <a href={optinPageUrl} className={styles.joinBtn}>{strings.BarNotEnrolled}</a>
          )}
        </div>
      );
    }

    // ── Enrolled / ready ────────────────────────────────────────────────────
    const s = summary!;
    const { currentLevel: level, monthlyPoints: monthly } = s;

    let progressPct: number;
    let nextLevelLabel: string;
    let ptToGo: number;

    if (level === 'Bronze') {
      progressPct   = Math.min(monthly / silverThreshold, 1) * 100;
      ptToGo        = Math.max(0, silverThreshold - monthly);
      nextLevelLabel = 'Silver';
    } else if (level === 'Silver') {
      progressPct   = Math.min(monthly / goldThreshold, 1) * 100;
      ptToGo        = Math.max(0, goldThreshold - monthly);
      nextLevelLabel = 'Gold';
    } else {
      progressPct   = 100;
      ptToGo        = 0;
      nextLevelLabel = '';
    }

    const showSuggestion = !suggestionDismissed && !!suggestion;

    return (
      <div className={styles.cepBar}>

        {/* Streak */}
        {streak > 0 && (
          <>
            <span
              className={styles.streak}
              title={fmt(strings.StreakTooltip, streak)}
            >
              <span className={styles.streakFire}>🔥</span>
              <span className={styles.streakCount}>{fmt(strings.StreakLabel, streak)}</span>
            </span>
            <span className={styles.separator} />
          </>
        )}

        {/* Level pill */}
        <span className={`${styles.levelPill} ${levelColorClass(level)}`}>
          {levelEmoji(level)} {level}
        </span>

        {/* Points */}
        <span className={styles.points}>{fmt(strings.BarPoints, monthly)}</span>

        {/* Progress bar */}
        {level !== 'Gold' && (
          <span className={styles.progressGroup}>
            <span className={styles.progressTrack}>
              <span
                className={`${styles.progressFill} ${progressClass(level)}`}
                style={{ width: `${progressPct}%` }}
              />
            </span>
            <span className={styles.progressLabel}>
              {fmt(strings.BarProgressLabel, ptToGo, nextLevelLabel)}
            </span>
          </span>
        )}

        <span className={styles.separator} />

        {/* Suggestion */}
        {showSuggestion && (
          <>
            <span className={styles.suggestion}>
              <span className={styles.suggestionIcon}>💡</span>
              <span className={styles.suggestionText}>{suggestion!.text}</span>
              <button
                className={styles.dismissBtn}
                onClick={() => this.setState({ suggestionDismissed: true })}
                title={strings.DismissSuggestion}
                aria-label={strings.DismissSuggestion}
              >×</button>
            </span>
            <span className={styles.separator} />
          </>
        )}

        {/* Win button + callout */}
        <button
          ref={this._winBtnRef}
          className={styles.winBtn}
          onClick={() => this.setState((prev) => ({ showWin: !prev.showWin }))}
          aria-expanded={showWin}
        >
          ✨ {strings.BarMarkWin}
        </button>

        {showWin && this._winBtnRef.current && (
          <WinCallout
            target={this._winBtnRef.current}
            onDismiss={() => this.setState({ showWin: false })}
            onSubmit={this._handleWinSubmit}
          />
        )}

        {/* Win animation */}
        {winAnim && (
          <span className={styles.winAnim} aria-live="polite">{winAnimText}</span>
        )}

        {/* Dashboard link */}
        {dashboardPageUrl && (
          <a href={dashboardPageUrl} className={styles.dashBtn}>
            {strings.BarViewDashboard} →
          </a>
        )}
      </div>
    );
  }
}
