import * as React from 'react';
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DefaultButton,
  PrimaryButton,
  Pivot,
  PivotItem,
  SearchBox,
} from '@fluentui/react';
import styles from './CepLeaderboard.module.scss';
import type { ICepLeaderboardProps } from './ICepLeaderboardProps';
import type { ILeaderboardEntry, ILeaderboardPage, LeaderboardScope } from '../../../services/CepApiModels';
import * as strings from 'CepLeaderboardWebPartStrings';

// ─── Helpers ─────────────────────────────────────────────────────────────────

function p2(n: number): string { return n < 10 ? '0' + n : '' + n; }

function currentMonth(): string {
  const d = new Date();
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}`;
}

function shiftMonth(month: string, delta: number): string {
  const [year, m] = month.split('-').map(Number);
  const d = new Date(year, m - 1 + delta, 1);
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}`;
}

function fmtMonth(month: string): string {
  const [year, m] = month.split('-').map(Number);
  const d = new Date(year, m - 1, 1);
  return d.toLocaleString('default', { month: 'long', year: 'numeric' });
}

const fmt = (tpl: string, ...args: (string | number)[]): string =>
  tpl.replace(/{(\d+)}/g, (hit, i) => (args[i] !== undefined ? String(args[i]) : hit));

const PAGE_SIZE = 20;

const MEDALS: Record<number, string> = { 1: '🥇', 2: '🥈', 3: '🥉' };

const LEVEL_CLASS: Record<string, string> = {
  Bronze: styles.bronze,
  Silver: styles.silver,
  Gold: styles.gold,
};

// ─── State ────────────────────────────────────────────────────────────────────

type LoadState = 'loading' | 'not_configured' | 'ready' | 'error';

interface ICepLeaderboardState {
  loadState: LoadState;
  errorMessage: string;
  month: string;
  scope: LeaderboardScope;
  buffer: ILeaderboardEntry[];
  totalEntries: number;
  currentUserEntry: ILeaderboardEntry | undefined;
  page: number;
  hasMore: boolean;
  isLoadingMore: boolean;
  searchQuery: string;
}

// ─── Podium ───────────────────────────────────────────────────────────────────

interface IPodiumCardProps {
  entry: ILeaderboardEntry;
  placement: 'first' | 'second' | 'third';
}

const PodiumCard: React.FC<IPodiumCardProps> = ({ entry, placement }) => {
  const colClass = [
    styles.podiumCol,
    placement === 'first'  ? styles.podiumFirst  :
    placement === 'second' ? styles.podiumSecond :
    styles.podiumThird,
    entry.isCurrentUser ? styles.podiumCurrentUser : '',
  ].join(' ');

  return (
    <div className={colClass}>
      <div className={styles.podiumMedal}>{MEDALS[entry.rank]}</div>
      <div className={styles.podiumName} title={entry.displayName}>
        {entry.displayName}
        {entry.isCurrentUser && <span className={styles.youTag}>{strings.YouLabel}</span>}
      </div>
      <div className={styles.podiumMeta}>
        <span className={`${styles.levelPill} ${LEVEL_CLASS[entry.currentLevel] ?? ''}`}>
          {entry.currentLevel}
        </span>
      </div>
      <div className={styles.podiumPoints}>{entry.monthlyPoints.toLocaleString()} pts</div>
      <div className={styles.podiumBase} />
    </div>
  );
};

// ─── Leaderboard Row ──────────────────────────────────────────────────────────

const LeaderboardRow: React.FC<{ entry: ILeaderboardEntry }> = ({ entry }) => {
  const rowClass = entry.isCurrentUser ? styles.lbRowCurrentUser : styles.lbRow;
  const rankDisplay = MEDALS[entry.rank] !== undefined
    ? <span title={`#${entry.rank}`}>{MEDALS[entry.rank]}</span>
    : `#${entry.rank}`;

  return (
    <tr className={rowClass}>
      <td className={styles.lbRankCell}>{rankDisplay}</td>
      <td className={styles.lbNameCell}>
        {entry.displayName}
        {entry.isCurrentUser && <span className={styles.youTag}>{strings.YouLabel}</span>}
      </td>
      <td className={styles.lbDeptCell}>{entry.department ?? '—'}</td>
      <td className={styles.lbLevelCell}>
        <span className={`${styles.levelPill} ${LEVEL_CLASS[entry.currentLevel] ?? ''}`}>
          {entry.currentLevel}
        </span>
      </td>
      <td className={styles.lbPointsCell}>{entry.monthlyPoints.toLocaleString()}</td>
    </tr>
  );
};

// ─── Main Component ───────────────────────────────────────────────────────────

export default class CepLeaderboard extends React.Component<ICepLeaderboardProps, ICepLeaderboardState> {

  constructor(props: ICepLeaderboardProps) {
    super(props);
    this.state = {
      loadState: 'loading',
      errorMessage: '',
      month: currentMonth(),
      scope: 'global',
      buffer: [],
      totalEntries: 0,
      currentUserEntry: undefined,
      page: 1,
      hasMore: false,
      isLoadingMore: false,
      searchQuery: '',
    };
  }

  public componentDidMount(): void {
    this._loadFirstPage().catch(console.error);
  }

  public componentDidUpdate(
    prevProps: ICepLeaderboardProps,
    prevState: ICepLeaderboardState
  ): void {
    if (
      prevProps.apiClient !== this.props.apiClient ||
      prevState.scope !== this.state.scope ||
      prevState.month !== this.state.month
    ) {
      this._loadFirstPage().catch(console.error);
    }
  }

  // ─── Data loading ───────────────────────────────────────────────────────

  private async _loadFirstPage(): Promise<void> {
    const { apiClient, functionAppBaseUrl } = this.props;
    if (!apiClient) {
      this.setState({ loadState: functionAppBaseUrl ? 'error' : 'not_configured', errorMessage: 'Failed to initialise API client.' });
      return;
    }
    this.setState({ loadState: 'loading', errorMessage: '', buffer: [], page: 1, hasMore: false, searchQuery: '' });
    const { scope, month } = this.state;
    try {
      const result: ILeaderboardPage = await apiClient.getLeaderboard(scope, month, 1, PAGE_SIZE);
      this.setState({
        loadState: 'ready',
        buffer: result.entries,
        totalEntries: result.totalEntries,
        currentUserEntry: result.currentUserEntry,
        page: 1,
        hasMore: result.entries.length < result.totalEntries,
      });
    } catch (e) {
      this.setState({ loadState: 'error', errorMessage: fmt(strings.LoadError, (e as Error).message) });
    }
  }

  private _handleLoadMore = async (): Promise<void> => {
    const { apiClient } = this.props;
    if (!apiClient) return;
    this.setState({ isLoadingMore: true });
    const { scope, month, page, buffer } = this.state;
    const nextPage = page + 1;
    try {
      const result: ILeaderboardPage = await apiClient.getLeaderboard(scope, month, nextPage, PAGE_SIZE);
      const combined = [...buffer, ...result.entries];
      this.setState({
        buffer: combined,
        page: nextPage,
        hasMore: combined.length < result.totalEntries,
        isLoadingMore: false,
      });
    } catch {
      this.setState({ isLoadingMore: false });
    }
  };

  // ─── Event handlers ─────────────────────────────────────────────────────

  private _handlePrevMonth = (): void => {
    this.setState((s) => ({ month: shiftMonth(s.month, -1) }));
  };

  private _handleNextMonth = (): void => {
    this.setState((s) => {
      const next = shiftMonth(s.month, 1);
      return { month: next > currentMonth() ? currentMonth() : next };
    });
  };

  private _handleScopeChange = (item?: PivotItem): void => {
    if (!item) return;
    const map: Record<string, LeaderboardScope> = {
      global: 'global',
      department: 'department',
      team: 'team',
    };
    const newScope = map[item.props.itemKey ?? ''] ?? 'global';
    if (newScope !== this.state.scope) {
      this.setState({ scope: newScope });
    }
  };

  private _handleSearch = (newValue?: string): void => {
    this.setState({ searchQuery: newValue ?? '' });
  };

  // ─── Render ─────────────────────────────────────────────────────────────

  public render(): React.ReactElement<ICepLeaderboardProps> {
    const { hasTeamsContext } = this.props;
    const {
      loadState, errorMessage, month, scope,
      buffer, totalEntries, currentUserEntry,
      hasMore, isLoadingMore, searchQuery,
    } = this.state;

    const isCurrentMonth = month === currentMonth();

    // client-side search filter
    const filtered = searchQuery.trim()
      ? buffer.filter((e) => e.displayName.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1)
      : buffer;

    // podium: first 3 entries from the full (unfiltered) buffer
    const top3 = buffer.slice(0, 3);
    const [first, second, third] = [top3[0], top3[1], top3[2]];
    const hasPodium = top3.length >= 3;

    // current user: show separator row only when not in the visible filtered list
    const currentUserInFiltered = currentUserEntry && filtered.some((e) => e.isCurrentUser);
    const showCurrentUserSeparator = !currentUserInFiltered && !!currentUserEntry && !searchQuery.trim();

    return (
      <section className={`${styles.cepLeaderboard} ${hasTeamsContext ? styles.teams : ''}`}>

        {/* ── Loading ── */}
        {loadState === 'loading' && (
          <div className={styles.spinnerWrap}>
            <Spinner size={SpinnerSize.large} label={strings.Loading} />
          </div>
        )}

        {/* ── Error ── */}
        {loadState === 'error' && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline
            actions={<DefaultButton text={strings.Retry} onClick={() => this._loadFirstPage()} />}
          >
            {errorMessage}
          </MessageBar>
        )}

        {/* ── Not configured ── */}
        {loadState === 'not_configured' && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            Leaderboard not configured. Run <code>deploy/set-tenant-properties.ps1</code> to set up the Function App URL.
          </MessageBar>
        )}

        {/* ── Ready ── */}
        {loadState === 'ready' && (
          <>
            {/* Header */}
            <div className={styles.headerRow}>
              <h2 className={styles.title}>🏆 {strings.WebPartTitle}</h2>
              <div className={styles.monthPicker}>
                <button
                  className={styles.monthBtn}
                  onClick={this._handlePrevMonth}
                  aria-label="Previous month"
                >‹</button>
                <span className={styles.monthLabel}>{fmtMonth(month)}</span>
                <button
                  className={styles.monthBtn}
                  onClick={this._handleNextMonth}
                  disabled={isCurrentMonth}
                  aria-label="Next month"
                >›</button>
              </div>
            </div>

            {/* Scope tabs */}
            <Pivot
              selectedKey={scope}
              onLinkClick={this._handleScopeChange}
              styles={{ root: { marginBottom: 16 } }}
            >
              <PivotItem headerText={strings.FilterGlobal} itemKey="global" />
              <PivotItem headerText={strings.FilterDepartment} itemKey="department" />
              <PivotItem headerText={strings.FilterTeam} itemKey="team" />
            </Pivot>

            {/* Stats */}
            <div className={styles.statsBar}>
              <span className={styles.statChip}>
                <strong>{totalEntries}</strong> {strings.StatsTotalUsers}
              </span>
            </div>

            {/* Podium */}
            {hasPodium && (
              <div className={styles.podiumSection}>
                <div className={styles.podiumRow}>
                  {second && <PodiumCard entry={second} placement="second" />}
                  {first  && <PodiumCard entry={first}  placement="first"  />}
                  {third  && <PodiumCard entry={third}  placement="third"  />}
                </div>
              </div>
            )}
            {!hasPodium && buffer.length === 0 && (
              <MessageBar messageBarType={MessageBarType.info}>{strings.NoDataForPodium}</MessageBar>
            )}

            {/* Search */}
            {buffer.length > 0 && (
              <div className={styles.searchWrap}>
                <SearchBox
                  placeholder={strings.SearchPlaceholder}
                  value={searchQuery}
                  onChange={(_, v) => this._handleSearch(v)}
                  onClear={() => this._handleSearch('')}
                  styles={{ root: { maxWidth: 320 } }}
                />
              </div>
            )}

            {/* Table */}
            {filtered.length > 0 && (
              <div className={styles.tableWrap}>
                <table className={styles.lbTable}>
                  <thead>
                    <tr>
                      <th className={styles.lbRankCell}>{strings.ColumnRank}</th>
                      <th className={styles.lbNameCell}>{strings.ColumnUser}</th>
                      <th className={styles.lbDeptCell}>{strings.ColumnDepartment}</th>
                      <th className={styles.lbLevelCell}>{strings.ColumnLevel}</th>
                      <th className={styles.lbPointsCell}>{strings.ColumnPoints}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((entry) => (
                      <LeaderboardRow key={`${entry.rank}-${entry.displayName}`} entry={entry} />
                    ))}
                    {showCurrentUserSeparator && currentUserEntry && (
                      <>
                        <tr className={styles.lbSeparator}>
                          <td colSpan={5}>· · ·</td>
                        </tr>
                        <LeaderboardRow entry={currentUserEntry} />
                      </>
                    )}
                  </tbody>
                </table>
              </div>
            )}

            {filtered.length === 0 && searchQuery.trim() && (
              <p className={styles.noResults}>{strings.NoResults}</p>
            )}

            {/* Load more */}
            {hasMore && !searchQuery.trim() && (
              <div className={styles.loadMoreWrap}>
                <PrimaryButton
                  text={isLoadingMore ? '...' : strings.LoadMore}
                  disabled={isLoadingMore}
                  onClick={this._handleLoadMore}
                />
              </div>
            )}
          </>
        )}
      </section>
    );
  }
}
