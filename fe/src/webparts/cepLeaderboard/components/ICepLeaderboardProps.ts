import type { CepApiClient } from '../../../services/CepApiClient';

export interface ICepLeaderboardProps {
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  userAadId: string;
  /** Optional pre-initialised API client */
  apiClient?: CepApiClient;
  /** Base URL of the Azure Function App – empty means tenant properties are not configured */
  functionAppBaseUrl?: string;
}
