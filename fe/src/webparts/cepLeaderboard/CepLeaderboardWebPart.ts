import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { AadHttpClient, SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'CepLeaderboardWebPartStrings';
import CepLeaderboard from './components/CepLeaderboard';
import { ICepLeaderboardProps } from './components/ICepLeaderboardProps';
import { CepApiClient } from '../../services/CepApiClient';

// Tenant Properties keys – same as other CEP web parts
const TENANT_KEY_BASE_URL  = 'CEP_FunctionAppBaseUrl';
const TENANT_KEY_CLIENT_ID = 'CEP_FunctionAppClientId';

// No user-editable properties – config comes from SharePoint Tenant Properties
export type ICepLeaderboardWebPartProps = Record<string, never>;

export default class CepLeaderboardWebPart extends BaseClientSideWebPart<ICepLeaderboardWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _apiClient: CepApiClient | undefined;
  private _functionAppBaseUrl: string = '';
  private _functionAppClientId: string = '';

  protected async onInit(): Promise<void> {
    await super.onInit();
    await this._loadTenantProperties();
    await this._initApiClient();
  }

  private async _loadTenantProperties(): Promise<void> {
    const siteAbsUrl = this.context.pageContext.site.absoluteUrl;
    try {
      const [baseUrlResp, clientIdResp] = await Promise.all([
        this.context.spHttpClient.get(
          `${siteAbsUrl}/_api/web/GetStorageEntity('${TENANT_KEY_BASE_URL}')`,
          SPHttpClient.configurations.v1
        ),
        this.context.spHttpClient.get(
          `${siteAbsUrl}/_api/web/GetStorageEntity('${TENANT_KEY_CLIENT_ID}')`,
          SPHttpClient.configurations.v1
        ),
      ]);
      const baseUrlData  = await baseUrlResp.json();
      const clientIdData = await clientIdResp.json();
      this._functionAppBaseUrl  = (baseUrlData.Value  as string | null)?.trim() ?? '';
      this._functionAppClientId = (clientIdData.Value as string | null)?.trim() ?? '';
      if (!this._functionAppBaseUrl || !this._functionAppClientId) {
        console.warn('[CepLeaderboard] Tenant Properties not found. Run deploy/set-tenant-properties.ps1.');
      }
    } catch (e) {
      console.error('[CepLeaderboard] Error retrieving Tenant Properties:', e);
    }
  }

  private async _initApiClient(): Promise<void> {
    const baseUrl  = this._functionAppBaseUrl;
    const clientId = this._functionAppClientId;
    if (!baseUrl || !clientId) return;
    try {
      const aadResourceUri = `api://${clientId}`;
      const aadClient: AadHttpClient = await this.context.aadHttpClientFactory.getClient(aadResourceUri);
      this._apiClient = new CepApiClient(
        aadClient,
        baseUrl,
        this.context.pageContext.aadInfo?.userId?.toString() ?? this.context.pageContext.user.email,
        this.context.pageContext.user.email,
        this.context.pageContext.user.displayName
      );
    } catch (e) {
      console.error('[CepLeaderboard] Failed to initialise AAD HTTP client:', e);
    }
  }

  public render(): void {
    const element: React.ReactElement<ICepLeaderboardProps> = React.createElement(
      CepLeaderboard,
      {
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        userAadId: this.context.pageContext.aadInfo?.userId?.toString() ?? '',
        apiClient: this._apiClient,
        functionAppBaseUrl: this._functionAppBaseUrl,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: []
        }
      ]
    };
  }
}
