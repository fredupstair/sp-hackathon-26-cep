import type { CepApiClient } from '../../../services/CepApiClient';

export interface ICepWinsProps {
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  userAadId: string;
  apiClient: CepApiClient | undefined;
  functionAppBaseUrl: string;
}
