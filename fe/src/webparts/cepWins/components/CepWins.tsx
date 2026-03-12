import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import {
  Pivot,
  PivotItem,
  Shimmer,
  ShimmerElementType,
  MessageBar,
  MessageBarType,
  DefaultButton,
  IconButton,
} from '@fluentui/react';
import styles from './CepWins.module.scss';
import type { ICepWinsProps } from './ICepWinsProps';
import type { IWinItem, IWinItemShared } from '../../../services/CepApiModels';
import { AppAccordionSection } from './subcomponents/AppAccordionSection';
import { AppFilterPills } from './subcomponents/AppFilterPills';
import * as strings from 'CepWinsWebPartStrings';

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
  return d.toLocaleString(undefined, { month: 'long', year: 'numeric' });
}

function groupByApp<T extends { appKey: string }>(items: T[]): Map<string, T[]> {
  const map = new Map<string, T[]>();
  for (const item of items) {
    const key = item.appKey || 'Unknown';
    if (!map.has(key)) map.set(key, []);
    map.get(key)!.push(item);
  }
  return map;
}

// ─── Loading Skeleton ────────────────────────────────────────────────────────

const LoadingSkeleton: React.FC = () => (
  <div className={styles.loadingContainer}>
    {[180, 220, 160, 200].map((w, i) => (
      <div key={i} className={styles.shimmerRow}>
        <Shimmer
          shimmerElements={[
            { type: ShimmerElementType.line, width: w, height: 48 },
            { type: ShimmerElementType.gap, width: '100%' },
          ]}
        />
      </div>
    ))}
  </div>
);

// ─── Component ───────────────────────────────────────────────────────────────

const CepWins: React.FC<ICepWinsProps> = (props) => {
  const { apiClient } = props;

  const [selectedMonth, setSelectedMonth] = useState(currentMonth());
  const [activeTab, setActiveTab] = useState<'mine' | 'community'>('mine');
  const [filterAppKey, setFilterAppKey] = useState<string | undefined>(undefined);
  const [myWins, setMyWins] = useState<IWinItem[] | undefined>(undefined);
  const [sharedWins, setSharedWins] = useState<IWinItemShared[] | undefined>(undefined);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);

  const loadData = useCallback(async (): Promise<void> => {
    if (!apiClient) return;
    setLoading(true);
    setError(undefined);
    try {
      if (activeTab === 'mine') {
        const resp = await apiClient.getMeWins(selectedMonth);
        setMyWins(resp.items);
      } else {
        const resp = await apiClient.getSharedWins(selectedMonth);
        setSharedWins(resp.items);
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : strings.LoadError.replace('{0}', ''));
    } finally {
      setLoading(false);
    }
  }, [apiClient, activeTab, selectedMonth]);

  useEffect(() => {
    setFilterAppKey(undefined);
    loadData().catch(() => { /* handled inside loadData */ });
  }, [selectedMonth, activeTab]); // eslint-disable-line react-hooks/exhaustive-deps

  const isNextMonthDisabled = selectedMonth >= currentMonth();

  const activeWins: (IWinItem | IWinItemShared)[] =
    activeTab === 'mine' ? (myWins ?? []) : (sharedWins ?? []);

  const filteredWins = filterAppKey
    ? activeWins.filter(w => w.appKey === filterAppKey)
    : activeWins;

  // All app keys present in the full (unfiltered) dataset → for filter pills
  const allAppKeys = Array.from(groupByApp(activeWins).keys()).sort();

  // Group filtered wins by app, sort sections by count descending
  const sortedAppEntries = Array.from(groupByApp(filteredWins).entries())
    .sort((a, b) => b[1].length - a[1].length);

  return (
    <div className={styles.cepWins}>

      {/* ── Header ── */}
      <div className={styles.header}>
        <h2 className={styles.title}>🌟 Copilot WIN</h2>
        <div className={styles.monthNav}>
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            ariaLabel={strings.PrevMonth}
            onClick={() => setSelectedMonth(m => shiftMonth(m, -1))}
          />
          <span className={styles.monthLabel}>{fmtMonth(selectedMonth)}</span>
          <IconButton
            iconProps={{ iconName: 'ChevronRight' }}
            ariaLabel={strings.NextMonth}
            disabled={isNextMonthDisabled}
            onClick={() => setSelectedMonth(m => shiftMonth(m, +1))}
          />
        </div>
      </div>

      {/* ── Tabs ── */}
      <Pivot
        selectedKey={activeTab}
        onLinkClick={item => {
          if (item?.props.itemKey) {
            setActiveTab(item.props.itemKey as 'mine' | 'community');
          }
        }}
      >
        <PivotItem headerText={strings.TabMyWins} itemKey="mine" />
        <PivotItem headerText={strings.TabCommunity} itemKey="community" />
      </Pivot>

      {/* ── Body ── */}
      {!apiClient ? (
        <MessageBar messageBarType={MessageBarType.warning}>
          {strings.NotConfigured}
        </MessageBar>
      ) : loading ? (
        <LoadingSkeleton />
      ) : error ? (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {error}
          <DefaultButton
            text={strings.Retry}
            onClick={() => loadData().catch(() => { /* noop */ })}
            style={{ marginLeft: 8 }}
          />
        </MessageBar>
      ) : (
        <>
          {/* App filter pills — only shown when >1 app present */}
          {allAppKeys.length > 1 && (
            <AppFilterPills
              availableAppKeys={allAppKeys}
              selectedAppKey={filterAppKey}
              onSelect={setFilterAppKey}
            />
          )}

          {/* Accordion sections, one per app */}
          {sortedAppEntries.length === 0 ? (
            <div className={styles.emptyState}>
              {filterAppKey
                ? strings.EmptyFilteredWins
                : activeTab === 'mine'
                  ? strings.EmptyMyWins
                  : strings.EmptyCommunityWins}
            </div>
          ) : (
            sortedAppEntries.map(([appKey, wins], index) => (
              <AppAccordionSection
                key={appKey}
                appKey={appKey}
                wins={wins}
                showAuthor={activeTab === 'community'}
                defaultExpanded={index < 3}
              />
            ))
          )}
        </>
      )}
    </div>
  );
};

export default CepWins;
