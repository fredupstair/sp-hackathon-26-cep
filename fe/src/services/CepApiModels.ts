// ─── Enrollment ──────────────────────────────────────────────────────────────

export interface IEnrollmentJoinRequest {
  department: string;
  team: string;
  isEngagementNudgesEnabled: boolean;
}

// ─── User / Me ────────────────────────────────────────────────────────────────

export type CepLevel = 'Bronze' | 'Silver' | 'Gold';

export interface IUserSummary {
  userId: string;
  displayName: string;
  email: string;
  department: string;
  team: string;
  enrollmentDate: string;          // ISO date string
  currentLevel: CepLevel;
  totalPoints: number;
  monthlyPoints: number;
  isActive: boolean;
  lastActivityDate?: string;        // ISO date string, optional
  isEngagementNudgesEnabled: boolean;
  globalRank?: number;
  teamRank?: number;
  departmentRank?: number;
  month: string;                    // YYYY-MM the summary refers to
}

export interface IAppUsageBreakdown {
  appKey: string;                   // Word | Excel | PowerPoint | Outlook | Teams | OneNote | Loop | M365Chat | Other
  promptCount: number;
  pointsEarned: number;
}

export interface IUserUsage {
  from: string;
  to: string;
  breakdown: IAppUsageBreakdown[];
  totalPrompts: number;
  totalPoints: number;
}

// ─── Badges ──────────────────────────────────────────────────────────────────

export interface ICepBadge {
  badgeKey: string;
  badgeName: string;
  description: string;
  earnedDate: string;               // ISO date string
  iconUrl?: string;
}

// ─── Leaderboard ─────────────────────────────────────────────────────────────

export type LeaderboardScope = 'global' | 'department' | 'team';

export interface ILeaderboardEntry {
  rank: number;
  displayName: string;
  department?: string;
  team?: string;
  monthlyPoints: number;
  currentLevel: CepLevel;
  isCurrentUser: boolean;
}

export interface ILeaderboardPage {
  scope: LeaderboardScope;
  month: string;
  totalEntries: number;
  entries: ILeaderboardEntry[];
  currentUserEntry?: ILeaderboardEntry;
}

// ─── Admin ───────────────────────────────────────────────────────────────────

export interface ICepConfig {
  syncFrequency: string;
  pointsPerPrompt: number;
  levelThresholdSilver: number;
  levelThresholdGold: number;
  inactivityDaysForNudge: number;
  leaderboardRefreshNotificationEnabled: boolean;
  maxUsersPerIngestionBatch: number;
  timeZone: string;
}

export interface ICepSyncState {
  lastSuccessfulRunUtc?: string;
  lastRunStatus: 'Success' | 'Failure' | 'Running' | 'Unknown';
  lastRunCorrelationId?: string;
  lastRunSummary?: string;
}
