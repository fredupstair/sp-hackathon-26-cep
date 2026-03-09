import * as React from 'react';
import { Icon } from '@fluentui/react';
import type { IUserUsage } from '../../../../services/CepApiModels';
import styles from '../CepWelcome.module.scss';
import * as strings from 'CepWelcomeWebPartStrings';
import { AppUsageChart } from './AppUsageChart';

interface ICopilotUniverseProps {
  usage: IUserUsage | undefined;
  loading?: boolean;
}

// ─── App metadata ─────────────────────────────────────────────────────────────

const APP_LABEL: Record<string, string> = {
  Word:        'Word',
  Excel:       'Excel',
  PowerPoint:  'PowerPoint',
  Outlook:     'Outlook',
  Teams:       'Teams',
  OneNote:     'OneNote',
  Loop:        'Loop',
  BizChat:     'Copilot Chat',
  WebChat:     'Web Chat',
  M365App:     'M365 App',
  Forms:       'Forms',
  SharePoint:  'SharePoint',
  Whiteboard:  'Whiteboard',
};

const APP_ICON: Record<string, string> = {
  Word:        'WordLogo',
  Excel:       'ExcelLogo',
  PowerPoint:  'PowerPointLogo',
  Outlook:     'OutlookLogo',
  Teams:       'TeamsLogo',
  OneNote:     'OneNoteLogo',
  Loop:        'Sync',
  BizChat:     'Chat',
  WebChat:     'Globe',
  M365App:     'Waffle',
  Forms:       'ClipboardList',
  SharePoint:  'SharePointLogo',
  Whiteboard:  'EditCreate',
};

const APP_COLOR: Record<string, string> = {
  Word:        '#2b579a',
  Excel:       '#217346',
  PowerPoint:  '#b7472a',
  Outlook:     '#0078d4',
  Teams:       '#6264a7',
  OneNote:     '#7719aa',
  Loop:        '#008272',
  BizChat:     '#0f6cbd',
  WebChat:     '#3a96dd',
  M365App:     '#5c2d91',
  Forms:       '#038387',
  SharePoint:  '#036c70',
  Whiteboard:  '#4a8fb8',
};

const APP_TOOLTIP: Record<string, string> = {
  Word:        strings.AppTooltipWord,
  Excel:       strings.AppTooltipExcel,
  PowerPoint:  strings.AppTooltipPowerPoint,
  Outlook:     strings.AppTooltipOutlook,
  Teams:       strings.AppTooltipTeams,
  OneNote:     strings.AppTooltipOneNote,
  Loop:        strings.AppTooltipLoop,
  BizChat:     strings.AppTooltipBizChat,
  WebChat:     strings.AppTooltipWebChat,
  M365App:     strings.AppTooltipM365App,
  Forms:       strings.AppTooltipForms,
  SharePoint:  strings.AppTooltipSharePoint,
  Whiteboard:  strings.AppTooltipWhiteboard,
};

const ALL_APP_KEYS = Object.keys(APP_LABEL);

// ─── Layout constants ─────────────────────────────────────────────────────────

const ORBIT_R    = 110;   // px – distance from hub centre to planet centre
const CANVAS_SZ  = 300;   // px – canvas width & height
const CENTER     = CANVAS_SZ / 2;
const FLOAT_BASE = 3;     // s  – base float period
const MIN_PLANET = 40;    // px – smallest planet diameter
const MAX_PLANET = 72;    // px – largest planet diameter

const calcSize = (count: number, max: number): number =>
  Math.round(MIN_PLANET + (max > 0 ? count / max : 0) * (MAX_PLANET - MIN_PLANET));

// ─── Component ────────────────────────────────────────────────────────────────

export const CopilotUniverse: React.FC<ICopilotUniverseProps> = ({ usage, loading }) => {
  const [hoveredApp,  setHoveredApp]  = React.useState<string | null>(null);
  const [selectedApp, setSelectedApp] = React.useState<string | null>(null);
  const [tab,         setTab]         = React.useState<'universe' | 'stats'>('universe');

  // ── Skeleton ──────────────────────────────────────────────────────────────
  if (loading || !usage) {
    return (
      <div className={styles.card}>
        <div className={`${styles.skeletonLine} ${styles.skeletonLineMedium}`} style={{ marginBottom: 20 }} />
        <div
          className={styles.skeleton}
          style={{ width: 240, height: 240, borderRadius: '50%', margin: '0 auto 16px' }}
        />
      </div>
    );
  }

  // ── Data prep ─────────────────────────────────────────────────────────────
  const bMap: Record<string, { promptCount: number; pointsEarned: number }> = {};
  (usage.breakdown ?? []).forEach(e => {
    bMap[e.appKey] = { promptCount: e.promptCount, pointsEarned: e.pointsEarned };
  });

  const allEntries = ALL_APP_KEYS.map(k => ({
    appKey:       k,
    promptCount:  bMap[k]?.promptCount  ?? 0,
    pointsEarned: bMap[k]?.pointsEarned ?? 0,
  }));

  const activeApps = allEntries
    .filter(e => e.promptCount > 0)
    .sort((a, b) => b.promptCount - a.promptCount);

  const lockedApps = allEntries.filter(e => e.promptCount === 0);
  const maxCount   = activeApps[0]?.promptCount ?? 1;
  const total      = Math.max(usage.totalPrompts, 1);

  // Precompute hovered planet geometry for canvas-level tooltip
  const hoveredIdx = hoveredApp ? activeApps.findIndex(e => e.appKey === hoveredApp) : -1;
  const hoveredGeo = hoveredIdx >= 0 ? (() => {
    const hEntry = activeApps[hoveredIdx];
    const hsz    = calcSize(hEntry.promptCount, maxCount);
    const hAngle = (hoveredIdx / activeApps.length) * 2 * Math.PI - Math.PI / 2;
    const hcx    = CENTER + ORBIT_R * Math.cos(hAngle);
    const hcy    = CENTER + ORBIT_R * Math.sin(hAngle);
    return { entry: hEntry, sz: hsz, cx: hcx, cy: hcy, above: hcy > CENTER };
  })() : null;

  const selEntry   = selectedApp
    ? {
        appKey:       selectedApp,
        promptCount:  bMap[selectedApp]?.promptCount  ?? 0,
        pointsEarned: bMap[selectedApp]?.pointsEarned ?? 0,
      }
    : null;

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div className={styles.card}>

      {/* Title + tab switcher */}
      <div className={styles.universeHeader}>
        <span className={styles.universeTitleText}>
          ✦ {strings.UsageBreakdownTitle}
        </span>
        <div className={styles.usageTabs}>
          <button
            className={`${styles.usageTab} ${tab === 'universe' ? styles.usageTabActive : ''}`}
            onClick={() => setTab('universe')}
          >
            🌌 {strings.TabUniverse}
          </button>
          <button
            className={`${styles.usageTab} ${tab === 'stats' ? styles.usageTabActive : ''}`}
            onClick={() => setTab('stats')}
          >
            📊 {strings.TabStats}
          </button>
        </div>
      </div>

      {/* Stats tab */}
      {tab === 'stats' && (
        <AppUsageChart usage={usage} loading={loading} embedded />
      )}

      {/* Universe tab */}
      {tab === 'universe' && activeApps.length === 0 && (
        <div className={styles.emptyState}>{strings.NoUsage}</div>
      )}
      {tab === 'universe' && activeApps.length > 0 && (
        <>
          {/* ── Orbit canvas (desktop) ── */}
          <div
            className={styles.universeCanvas}
            style={{ width: CANVAS_SZ, height: CANVAS_SZ }}
          >
            {/* Decorative orbit ring */}
            <div
              className={styles.orbitRing}
              style={{ width: ORBIT_R * 2, height: ORBIT_R * 2 }}
            />

            {/* Central hub */}
            <div className={styles.universeHub}>
              <span className={styles.hubStar}>🌟</span>
              <span className={styles.hubCount}>{usage.totalPrompts}</span>
              <span className={styles.hubLabel}>{strings.UniversePromptsLabel}</span>
            </div>

            {/* SVG connector lines from hub to each planet */}
            <svg
              className={styles.universeLines}
              width={CANVAS_SZ}
              height={CANVAS_SZ}
            >
              {activeApps.map((entry, i) => {
                const angle = (i / activeApps.length) * 2 * Math.PI - Math.PI / 2;
                return (
                  <line
                    key={entry.appKey}
                    x1={CENTER} y1={CENTER}
                    x2={CENTER + ORBIT_R * Math.cos(angle)}
                    y2={CENTER + ORBIT_R * Math.sin(angle)}
                    stroke="rgba(92,45,145,0.13)"
                    strokeWidth="1"
                    strokeDasharray="4 5"
                  />
                );
              })}
            </svg>

            {/* Fixed planets, gently floating */}
            {activeApps.map((entry, i) => {
              const sz    = calcSize(entry.promptCount, maxCount);
              const angle = (i / activeApps.length) * 2 * Math.PI - Math.PI / 2;
              const px    = CENTER + ORBIT_R * Math.cos(angle) - sz / 2;
              const py    = CENTER + ORBIT_R * Math.sin(angle) - sz / 2;
              const dur   = `${FLOAT_BASE + (i % 3) * 0.7}s`;
              const delay = `${(i * 0.45) % 2.5}s`;
              const paused = hoveredApp  === entry.appKey;
              const sel    = selectedApp === entry.appKey;

              return (
                <div
                  key={entry.appKey}
                  className={[
                    styles.orbitPlanet,
                    paused ? styles.orbitPlanetPaused   : '',
                    sel    ? styles.orbitPlanetSelected : '',
                  ].filter(Boolean).join(' ')}
                  style={{
                    width:             sz,
                    height:            sz,
                    left:              px,
                    top:               py,
                    background:        APP_COLOR[entry.appKey] ?? '#0078d4',
                    animationDuration: dur,
                    animationDelay:    delay,
                  }}
                  onMouseEnter={() => setHoveredApp(entry.appKey)}
                  onMouseLeave={() => setHoveredApp(null)}
                  onClick={() => setSelectedApp(sel ? null : entry.appKey)}
                  role="button"
                  tabIndex={0}
                  aria-label={`${APP_LABEL[entry.appKey]} – ${entry.promptCount} ${strings.UniversePromptsLabel}`}
                  onKeyDown={e => e.key === 'Enter' && setSelectedApp(sel ? null : entry.appKey)}
                >
                  <div className={styles.planetInner}>
                    <Icon
                      iconName={APP_ICON[entry.appKey] ?? 'Waffle'}
                      className={styles.planetIcon}
                    />
                    {sz >= 54 && (
                      <span className={styles.planetLabel}>
                        {APP_LABEL[entry.appKey]}
                      </span>
                    )}
                  </div>
                </div>
              );
            })}

            {/* Canvas-level tooltip – pixel coords, immune to planet transforms */}
            {hoveredGeo && (
              <div
                className={[
                  styles.universeTooltip,
                  hoveredGeo.above ? styles.universeTooltipAbove : styles.universeTooltipBelow,
                ].join(' ')}
                style={{
                  left:      hoveredGeo.cx,
                  top:       hoveredGeo.above
                               ? hoveredGeo.cy - hoveredGeo.sz / 2 - 8
                               : hoveredGeo.cy + hoveredGeo.sz / 2 + 8,
                  transform: hoveredGeo.above
                               ? 'translate(-50%, -100%)'
                               : 'translateX(-50%)',
                }}
              >
                <div className={styles.tooltipApp}>
                  <Icon
                    iconName={APP_ICON[hoveredGeo.entry.appKey] ?? 'Waffle'}
                    style={{ color: APP_COLOR[hoveredGeo.entry.appKey], fontSize: 16 }}
                  />
                  <strong>{APP_LABEL[hoveredGeo.entry.appKey]}</strong>
                </div>
                <div className={styles.tooltipStats}>
                  {hoveredGeo.entry.promptCount} {strings.UniversePromptsLabel}
                  &nbsp;·&nbsp;
                  {Math.round((hoveredGeo.entry.promptCount / total) * 100)}%
                </div>
                <div className={styles.tooltipDesc}>
                  {APP_TOOLTIP[hoveredGeo.entry.appKey]}
                </div>
              </div>
            )}
          </div>

          {/* ── Mobile list (replaces canvas on small screens) ── */}
          <div className={styles.universeMobileList}>
            {activeApps.map(entry => (
              <div
                key={entry.appKey}
                className={`${styles.universeMobileRow} ${selectedApp === entry.appKey ? styles.universeMobileRowSelected : ''}`}
                onClick={() => setSelectedApp(selectedApp === entry.appKey ? null : entry.appKey)}
                role="button"
                tabIndex={0}
                onKeyDown={e => e.key === 'Enter' && setSelectedApp(selectedApp === entry.appKey ? null : entry.appKey)}
              >
                <div
                  className={styles.universeMobilePlanet}
                  style={{ background: APP_COLOR[entry.appKey] ?? '#0078d4' }}
                >
                  <Icon
                    iconName={APP_ICON[entry.appKey] ?? 'Waffle'}
                    style={{ color: '#fff', fontSize: 16 }}
                  />
                </div>
                <div className={styles.universeMobileInfo}>
                  <div className={styles.universeMobileLabel}>
                    {APP_LABEL[entry.appKey]}
                  </div>
                  <div className={styles.universeMobileCount}>
                    {entry.promptCount} {strings.UniversePromptsLabel}
                    &nbsp;·&nbsp;
                    {Math.round((entry.promptCount / total) * 100)}%
                  </div>
                </div>
                <div className={styles.universeMobileArrow}>›</div>
              </div>
            ))}
          </div>
        </>
      )}

      {/* ── Selected planet detail card ── */}
      {tab === 'universe' && selEntry && (
        <div
          className={styles.planetDetailCard}
          style={{ borderLeftColor: APP_COLOR[selEntry.appKey] ?? '#0078d4' }}
        >
          <div className={styles.planetDetailHeader}>
            <Icon
              iconName={APP_ICON[selEntry.appKey] ?? 'Waffle'}
              style={{ color: APP_COLOR[selEntry.appKey], fontSize: 22 }}
            />
            <strong className={styles.planetDetailName}>
              {APP_LABEL[selEntry.appKey]}
            </strong>
            <button
              className={styles.planetDetailClose}
              onClick={() => setSelectedApp(null)}
              aria-label="Chiudi dettaglio"
            >
              ✕
            </button>
          </div>
          <div className={styles.planetDetailBody}>
            <div className={styles.planetDetailStat}>
              <span className={styles.planetDetailValue}>
                {selEntry.promptCount}
              </span>
              <span className={styles.planetDetailDimLabel}>{strings.UniversePromptsLabel}</span>
            </div>
            <div className={styles.planetDetailStat}>
              <span className={styles.planetDetailValue}>
                {Math.round((selEntry.promptCount / total) * 100)}%
              </span>
              <span className={styles.planetDetailDimLabel}>{strings.UniverseOfTotal}</span>
            </div>
            {selEntry.pointsEarned > 0 && (
              <div className={styles.planetDetailStat}>
                <span className={styles.planetDetailValue}>
                  {selEntry.pointsEarned}
                </span>
                <span className={styles.planetDetailDimLabel}>{strings.UniversePoints}</span>
              </div>
            )}
          </div>
          <p className={styles.planetDetailDesc}>
            {APP_TOOLTIP[selEntry.appKey]}
          </p>
        </div>
      )}

      {/* ── Locked worlds ── */}
      {tab === 'universe' && lockedApps.length > 0 && (
        <div className={styles.lockedSection}>
          <div className={styles.lockedDivider}>{strings.UniverseLockedWorlds}</div>
          <div className={styles.lockedGrid}>
            {lockedApps.map(app => (
              <div
                key={app.appKey}
                className={styles.lockedPlanet}
                title={APP_TOOLTIP[app.appKey]}
                data-tip={strings.UniverseLockedHint.replace('{0}', APP_LABEL[app.appKey])}
              >
                <div className={styles.lockedIconStack}>
                  {/* default: lock */}
                  <div className={styles.lockedIconWrap}>
                    <Icon iconName="Lock" />
                  </div>
                  {/* hover: real app icon with brand colour */}
                  <div
                    className={styles.lockedAppIconWrap}
                    style={{ color: APP_COLOR[app.appKey] }}
                  >
                    <Icon iconName={APP_ICON[app.appKey] ?? 'Waffle'} />
                  </div>
                </div>
                <span className={styles.lockedLabel}>{APP_LABEL[app.appKey]}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};
