import * as React from 'react';
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DefaultButton,
  Link,
} from '@fluentui/react';
import styles from './CepDashboard.module.scss';
import type { ICepDashboardProps } from './ICepDashboardProps';
import type {
  IUserSummary,
  IUserUsage,
  ICepBadge,
  ILeaderboardPage,
} from '../../../services/CepApiModels';
import * as strings from 'CepDashboardWebPartStrings';
import { DashboardHeader } from './subcomponents/DashboardHeader';
import { StatsRow } from './subcomponents/StatsRow';
import { AppUsageChart } from './subcomponents/AppUsageChart';
import { BadgeList } from './subcomponents/BadgeList';
import { MiniLeaderboard } from './subcomponents/MiniLeaderboard';
import { NotEnrolledBanner } from './subcomponents/NotEnrolledBanner';
import { OtherUserView } from './subcomponents/OtherUserView';

// ─── Month helpers ────────────────────────────────────────────────────────────

// ES5-compatible zero-pad (avoids padStart which requires es2017 lib)
function p2(n: number): string { return n < 10 ? '0' + n : '' + n; }

function currentMonth(): string {
  const d = new Date();
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}`;
}

function getMonthRange(month: string): { from: string; to: string } {
  const [year, m] = month.split('-').map(Number);
  const lastDay = new Date(year, m, 0).getDate();
  return { from: `${month}-01`, to: `${month}-${p2(lastDay)}` };
}

function shiftMonth(month: string, delta: number): string {
  const [year, m] = month.split('-').map(Number);
  const d = new Date(year, m - 1 + delta, 1);
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}`;
}

const fmt = (tpl: string, ...args: (string | number)[]): string =>
  tpl.replace(/{(\d+)}/g, (hit, i) => (args[i] !== undefined ? String(args[i]) : hit));

// ─── State ────────────────────────────────────────────────────────────────────

type LoadState = 'loading' | 'not_configured' | 'not_enrolled' | 'ready' | 'error';

interface ICepDashboardState {
  loadState: LoadState;
  errorMessage: string;
  month: string;
  summary: IUserSummary | undefined;
  usage: IUserUsage | undefined;
  badges: ICepBadge[];
  leaderboard: ILeaderboardPage | undefined;
}

// ─── Component ────────────────────────────────────────────────────────────────

export default class CepDashboard extends React.Component<ICepDashboardProps, ICepDashboardState> {
  constructor(props: ICepDashboardProps) {
    super(props);
    this.state = {
      loadState: 'loading',
      errorMessage: '',
      month: currentMonth(),
      summary: undefined,
      usage: undefined,
      badges: [],
      leaderboard: undefined,
    };
  }

  public componentDidMount(): void {
    this._loadData().catch(console.error);
  }

  public componentDidUpdate(
    prevProps: ICepDashboardProps,
    prevState: ICepDashboardState
  ): void {
    if (
      prevProps.apiClient !== this.props.apiClient ||
      prevProps.viewedUserEmail !== this.props.viewedUserEmail ||
      prevState.month !== this.state.month
    ) {
      this._loadData().catch(console.error);
    }
  }

  // ─── Data loading ────────────────────────────────────────────────────────

  private async _loadData(): Promise<void> {
    const { apiClient, functionAppBaseUrl } = this.props;
    if (!apiClient) {
      if (!functionAppBaseUrl) {
        this.setState({ loadState: 'not_configured' });
      } else {
        // URL is set but AAD client failed to initialise
        this.setState({
          loadState: 'error',
          errorMessage: 'Failed to initialise API client. Check the browser console for details.',
        });
      }
      return;
    }

    this.setState({ loadState: 'loading', errorMessage: '' });
    const { month } = this.state;
    const { from, to } = getMonthRange(month);

    try {
      const [summary, usage, badges, leaderboard] = await Promise.all([
        apiClient.getMeSummary(month),
        apiClient.getMeUsage(from, to).catch((): IUserUsage => ({
          from, to, breakdown: [], totalPrompts: 0, totalPoints: 0,
        })),
        apiClient.getMeBadges().catch((): ICepBadge[] => []),
        apiClient.getLeaderboard('global', month, 1, 5).catch((): undefined => undefined),
      ]);

      if (!summary) {
        this.setState({ loadState: 'not_enrolled' });
        return;
      }

      this.setState({ loadState: 'ready', summary, usage, badges, leaderboard });
    } catch (e) {
      this.setState({
        loadState: 'error',
        errorMessage: fmt(strings.LoadError, (e as Error).message),
      });
    }
  }

  // ─── Month navigation ───────────────────────────────────────────────────

  private _handlePrevMonth = (): void => {
    this.setState((s) => ({ month: shiftMonth(s.month, -1) }));
  };

  private _handleNextMonth = (): void => {
    this.setState((s) => {
      const next = shiftMonth(s.month, 1);
      return { month: next > currentMonth() ? currentMonth() : next };
    });
  };

  // ─── Render ──────────────────────────────────────────────────────────────

  public render(): React.ReactElement<ICepDashboardProps> {
    const { hasTeamsContext, viewedUserEmail } = this.props;
    const { loadState, errorMessage, month, summary, usage, badges, leaderboard } = this.state;
    const isCurrentMonth = month === currentMonth();

    return (
      <section
        className={`${styles.cepDashboard} ${hasTeamsContext ? styles.teams : ''}`}
      >
        {/* ── Loading ── */}
        {loadState === 'loading' && (
          <div style={{ textAlign: 'center', padding: 48 }}>
            <Spinner size={SpinnerSize.large} label={strings.Loading} />
          </div>
        )}

        {/* ── Error ── */}
        {loadState === 'error' && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline
            actions={
              <DefaultButton
                text={strings.Retry}
                onClick={() => this._loadData()}
              />
            }
          >
            {errorMessage}
          </MessageBar>
        )}

        {/* ── Not configured ── */}
        {loadState === 'not_configured' && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            <strong>{strings.NotConfiguredTitle}</strong>{' '}
            {strings.NotConfiguredMessage}
          </MessageBar>
        )}

        {/* ── Not enrolled ── */}
        {loadState === 'not_enrolled' && <NotEnrolledBanner />}

        {/* ── My Dashboard ── */}
        {loadState === 'ready' && summary && !viewedUserEmail && (
          <>
            <DashboardHeader
              summary={summary}
              month={month}
              isCurrentMonth={isCurrentMonth}
              onPrevMonth={this._handlePrevMonth}
              onNextMonth={this._handleNextMonth}
            />
            <StatsRow summary={summary} />
            {usage && <AppUsageChart usage={usage} />}
            <BadgeList badges={badges} />
            {leaderboard && <MiniLeaderboard leaderboard={leaderboard} />}
          </>
        )}

        {/* ── Other user (aggregated) ── */}
        {loadState === 'ready' && summary && viewedUserEmail && (
          <>
            <div style={{ marginBottom: 12 }}>
              <Link onClick={() => { /* caller clears viewedUserEmail */ }}>
                ← {strings.BackToMyProfile}
              </Link>
            </div>
            <OtherUserView summary={summary} badges={badges} />
          </>
        )}
      </section>
    );
  }
}

