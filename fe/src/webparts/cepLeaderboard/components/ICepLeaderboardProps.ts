import type { CepApiClient } from '../../../services/CepApiClient';

export interface ICepLeaderboardProps {
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  userAadId: string;
  /** Optional pre-initialised API client */
  apiClient?: CepApiClient;
}
