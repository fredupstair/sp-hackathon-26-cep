import type { CepApiClient } from '../../../services/CepApiClient';

export interface ICepDashboardProps {
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  userAadId: string;
  /** When set, the dashboard shows the profile of another user (aggregated view) */
  viewedUserEmail?: string;
  /** Optional pre-initialised API client; if not provided the web part is not yet configured */
  apiClient?: CepApiClient;
}
