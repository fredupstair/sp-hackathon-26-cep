import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { AadHttpClient, SPHttpClient } from '@microsoft/sp-http';

import CepOptin from './components/CepOptin';
import { ICepOptinProps } from './components/ICepOptinProps';
import { CepApiClient } from '../../services/CepApiClient';

// Tenant Properties keys (Storage Entities) – set via deploy/set-tenant-properties.ps1
const TENANT_KEY_BASE_URL   = 'CEP_FunctionAppBaseUrl';
const TENANT_KEY_CLIENT_ID  = 'CEP_FunctionAppClientId';

// No user-editable properties – config comes from SharePoint Tenant Properties
export type ICepOptinWebPartProps = Record<string, never>;

export default class CepOptinWebPart extends BaseClientSideWebPart<ICepOptinWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _aadClient: AadHttpClient | undefined;
  private _apiClient: CepApiClient | undefined;

  /** Resolved from SharePoint Tenant Properties on init */
  private _functionAppBaseUrl: string = '';
  private _functionAppClientId: string = '';

  protected async onInit(): Promise<void> {
    await super.onInit();
    await this._loadTenantProperties();
    await this._initApiClient();
  }

  /** Reads CEP configuration from SharePoint Tenant Properties (Storage Entities). */
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

      const baseUrlData   = await baseUrlResp.json();
      const clientIdData  = await clientIdResp.json();

      this._functionAppBaseUrl  = (baseUrlData.Value  as string | null)?.trim() ?? '';
      this._functionAppClientId = (clientIdData.Value as string | null)?.trim() ?? '';

      if (!this._functionAppBaseUrl || !this._functionAppClientId) {
        console.warn(
          '[CepOptin] Tenant Properties non trovate. ' +
          'Eseguire deploy/set-tenant-properties.ps1 per impostarle.'
        );
      }
    } catch (e) {
      console.error('[CepOptin] Errore nel recupero delle Tenant Properties:', e);
    }
  }

  private async _initApiClient(): Promise<void> {
    const baseUrl  = this._functionAppBaseUrl;
    const clientId = this._functionAppClientId;
    if (!baseUrl || !clientId) return;
    try {
      // The resource URI for AAD token acquisition MUST be the App ID URI (api://<clientId>)
      // NOT the Function App URL – the approved grant is scoped to the App ID URI.
      const aadResourceUri = `api://${clientId}`;
      this._aadClient = await this.context.aadHttpClientFactory.getClient(aadResourceUri);
      this._apiClient = new CepApiClient(
        this._aadClient,
        baseUrl,
        this.context.pageContext.aadInfo?.userId?.toString() ?? this.context.pageContext.user.email,
        this.context.pageContext.user.email,
        this.context.pageContext.user.displayName
      );
    } catch (e) {
      console.error('[CepOptin] Failed to initialise AAD HTTP client:', e);
    }
  }

  public render(): void {
    const element: React.ReactElement<ICepOptinProps> = React.createElement(
      CepOptin,
      {
        functionAppBaseUrl: this._functionAppBaseUrl,
        apiClient: this._apiClient,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        userAadId: this.context.pageContext.aadInfo?.userId?.toString() ?? '',
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
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
    // La configurazione proviene dalle Tenant Properties di SharePoint.
    // Usare deploy/set-tenant-properties.ps1 per impostare i valori.
    return { pages: [] };
  }
}

