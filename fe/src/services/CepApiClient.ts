import { AadHttpClient } from '@microsoft/sp-http';
import {
  IEnrollmentJoinRequest,
  IUserSummary,
  IUserUsage,
  ICepBadge,
  ILeaderboardPage,
  LeaderboardScope,
  ICepConfig,
  ICepSyncState,
  ICepWinRequest,
  ICepWinResponse,
  ICepSuggestion,
} from './CepApiModels';

export class CepApiError extends Error {
  public readonly statusCode: number;
  constructor(message: string, statusCode: number) {
    super(message);
    this.name = 'CepApiError';
    this.statusCode = statusCode;
  }
}

export class CepApiClient {
  private readonly _client: AadHttpClient;
  private readonly _baseUrl: string;
  private readonly _userId: string;
  private readonly _userPrincipalName: string;
  private readonly _displayName: string;

  constructor(client: AadHttpClient, baseUrl: string, userId: string, userPrincipalName: string, displayName: string) {
    this._client = client;
    this._baseUrl = baseUrl.replace(/\/$/, '');
    this._userId = userId;
    this._userPrincipalName = userPrincipalName;
    this._displayName = displayName;
  }

  /** Default extra headers sent with every request */
  private get _defaultHeaders(): Record<string, string> {
    return this._userId ? { 'X-User-Id': this._userId } : {};
  }

  // ─── Enrollment ────────────────────────────────────────────────────────────

  public async join(data: Omit<IEnrollmentJoinRequest, 'aadUserId' | 'userPrincipalName' | 'displayName' | 'email'>): Promise<void> {
    const body: IEnrollmentJoinRequest = {
      ...data,
      aadUserId: this._userId,
      userPrincipalName: this._userPrincipalName,
      displayName: this._displayName,
      email: this._userPrincipalName,
    };
    const response = await this._client.post(
      `${this._baseUrl}/api/enrollment/join`,
      AadHttpClient.configurations.v1,
      {
        body: JSON.stringify(body),
        headers: { ...this._defaultHeaders, 'Content-Type': 'application/json' },
      }
    );
    await this._assertOk(response);
  }

  public async leave(): Promise<void> {
    const response = await this._client.post(
      `${this._baseUrl}/api/enrollment/leave`,
      AadHttpClient.configurations.v1,
      {
        body: JSON.stringify({ aadUserId: this._userId }),
        headers: { ...this._defaultHeaders, 'Content-Type': 'application/json' },
      }
    );
    await this._assertOk(response);
  }

  // ─── Me / Dashboard ────────────────────────────────────────────────────────

  /**
   * Returns the user summary for the given month (default: current month).
   * Returns undefined if the user is not enrolled (404).
   */
  public async getMeSummary(month?: string): Promise<IUserSummary | undefined> {
    const qs = month ? `?month=${encodeURIComponent(month)}` : '';
    const response = await this._client.get(
      `${this._baseUrl}/api/me/summary${qs}`,
      AadHttpClient.configurations.v1,
      { headers: this._defaultHeaders }
    );
    if (response.status === 404) return undefined;
    await this._assertOk(response);
    return response.json() as Promise<IUserSummary>;
  }

  public async getMeUsage(from: string, to: string): Promise<IUserUsage> {
    const qs = `?from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}`;
    const response = await this._client.get(
      `${this._baseUrl}/api/me/usage${qs}`,
      AadHttpClient.configurations.v1,
      { headers: this._defaultHeaders }
    );
    await this._assertOk(response);
    return response.json() as Promise<IUserUsage>;
  }

  public async getMeBadges(): Promise<ICepBadge[]> {
    const response = await this._client.get(
      `${this._baseUrl}/api/me/badges`,
      AadHttpClient.configurations.v1,
      { headers: this._defaultHeaders }
    );
    await this._assertOk(response);
    return response.json() as Promise<ICepBadge[]>;
  }

  // ─── Leaderboard ───────────────────────────────────────────────────────────

  public async getLeaderboard(
    scope: LeaderboardScope,
    month: string,
    page: number = 1,
    pageSize: number = 20
  ): Promise<ILeaderboardPage> {
    const qs = `?scope=${scope}&month=${encodeURIComponent(month)}&page=${page}&pageSize=${pageSize}`;
    const response = await this._client.get(
      `${this._baseUrl}/api/leaderboard${qs}`,
      AadHttpClient.configurations.v1,
      { headers: this._defaultHeaders }
    );
    await this._assertOk(response);
    return response.json() as Promise<ILeaderboardPage>;
  }

  // ─── Bar: Win ──────────────────────────────────────────────────────────────

  /**
   * Records a manual "Copilot Win" for the calling user.
   * Aggregated per day in CEP_ActivityLog (AppKey = "Win"); max 10 wins/day.
   * Returns 429 CepApiError when the daily limit is reached.
   */
  public async postWin(data: ICepWinRequest): Promise<ICepWinResponse> {
    const response = await this._client.post(
      `${this._baseUrl}/api/me/wins`,
      AadHttpClient.configurations.v1,
      {
        body: JSON.stringify(data),
        headers: { ...this._defaultHeaders, 'Content-Type': 'application/json' },
      }
    );
    await this._assertOk(response);
    return response.json() as Promise<ICepWinResponse>;
  }

  // ─── Me: Preferences ──────────────────────────────────────────────────────

  public async getMePreferences(): Promise<{ isEngagementNudgesEnabled: boolean }> {
    const response = await this._client.get(
      `${this._baseUrl}/api/me/preferences`,
      AadHttpClient.configurations.v1,
      { headers: this._defaultHeaders }
    );
    await this._assertOk(response);
    return response.json() as Promise<{ isEngagementNudgesEnabled: boolean }>;
  }

  public async updateMePreferences(prefs: { isEngagementNudgesEnabled?: boolean }): Promise<void> {
    const response = await this._client.post(
      `${this._baseUrl}/api/me/preferences`,
      AadHttpClient.configurations.v1,
      {
        body: JSON.stringify(prefs),
        headers: { ...this._defaultHeaders, 'Content-Type': 'application/json' },
      }
    );
    await this._assertOk(response);
  }

  // ─── Bar: Suggestion ───────────────────────────────────────────────────────

  /** Returns a contextual Copilot suggestion based on the user's least-used apps. */
  public async getSuggestion(): Promise<ICepSuggestion> {
    const response = await this._client.get(
      `${this._baseUrl}/api/me/suggestion`,
      AadHttpClient.configurations.v1,
      { headers: this._defaultHeaders }
    );
    await this._assertOk(response);
    return response.json() as Promise<ICepSuggestion>;
  }

  // ─── Admin ─────────────────────────────────────────────────────────────────

  public async getAdminConfig(): Promise<ICepConfig> {
    const response = await this._client.get(
      `${this._baseUrl}/api/admin/config`,
      AadHttpClient.configurations.v1,
      { headers: this._defaultHeaders }
    );
    await this._assertOk(response);
    return response.json() as Promise<ICepConfig>;
  }

  public async updateAdminConfig(config: Partial<ICepConfig>): Promise<void> {
    const response = await this._client.post(
      `${this._baseUrl}/api/admin/config`,
      AadHttpClient.configurations.v1,
      {
        body: JSON.stringify(config),
        headers: { ...this._defaultHeaders, 'Content-Type': 'application/json' },
      }
    );
    await this._assertOk(response);
  }

  public async getAdminStatus(): Promise<ICepSyncState> {
    const response = await this._client.get(
      `${this._baseUrl}/api/admin/status`,
      AadHttpClient.configurations.v1,
      { headers: this._defaultHeaders }
    );
    await this._assertOk(response);
    return response.json() as Promise<ICepSyncState>;
  }

  // ─── Helpers ───────────────────────────────────────────────────────────────

  private async _assertOk(response: { ok: boolean; status: number; text: () => Promise<string> }): Promise<void> {
    if (!response.ok) {
      let detail = '';
      try { detail = await response.text(); } catch { /* ignore */ }
      throw new CepApiError(
        `API error ${response.status}${detail ? ': ' + detail : ''}`,
        response.status
      );
    }
  }
}
