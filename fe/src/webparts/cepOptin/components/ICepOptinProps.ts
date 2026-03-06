import { DisplayMode } from '@microsoft/sp-core-library';
import { CepApiClient } from '../../../services/CepApiClient';

export interface ICepOptinProps {
  /** Base URL of the Azure Function App (used for API calls and AAD resource) */
  functionAppBaseUrl: string;
  /** Fully initialised API client, undefined if not yet configured */
  apiClient: CepApiClient | undefined;
  userDisplayName: string;
  userEmail: string;
  /** AAD Object ID (OID) of the current user, sent as X-User-Id to the backend */
  userAadId: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  /** Welcome text shown on Step 1 of the enrollment wizard (configured by the page editor) */
  welcomeText: string;
  /** SPFx display mode — used to show the editor-only setup card when welcomeText is empty */
  displayMode: DisplayMode;
  /** Opens the SPFx property pane so editors can configure the webpart */
  onConfigureClick: () => void;
}
