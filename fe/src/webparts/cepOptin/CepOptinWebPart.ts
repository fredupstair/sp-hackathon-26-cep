import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneLabel,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { AadHttpClient, MSGraphClientV3, SPHttpClient } from '@microsoft/sp-http';

import CepOptin from './components/CepOptin';
import { ICepOptinProps } from './components/ICepOptinProps';
import { CepApiClient } from '../../services/CepApiClient';

// Tenant Properties keys (Storage Entities) – set via deploy/set-tenant-properties.ps1
const TENANT_KEY_BASE_URL  = 'CEP_FunctionAppBaseUrl';
const TENANT_KEY_CLIENT_ID = 'CEP_FunctionAppClientId';

export interface ICepOptinWebPartProps {
  /** Welcome text shown on the first step of the enrollment wizard */
  welcomeText: string;
  /** Organisation name used for AI text generation (not shown to end users) */
  organizationName: string;
}

export default class CepOptinWebPart extends BaseClientSideWebPart<ICepOptinWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _aadClient: AadHttpClient | undefined;
  private _apiClient: CepApiClient | undefined;
  private _graphClient: MSGraphClientV3 | undefined;

  /** Resolved from SharePoint Tenant Properties on init */
  private _functionAppBaseUrl: string = '';
  private _functionAppClientId: string = '';

  /** DOM node used by the custom property pane field */
  private _welcomeEditorContainer: HTMLElement | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();
    await this._loadTenantProperties();
    await this._initApiClient();
    try {
      this._graphClient = await this.context.msGraphClientFactory.getClient('3');
    } catch (e) {
      console.warn('[CepOptin] MSGraphClientV3 not available:', e);
    }
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

      const baseUrlData  = await baseUrlResp.json();
      const clientIdData = await clientIdResp.json();

      this._functionAppBaseUrl  = (baseUrlData.Value  as string | null)?.trim() ?? '';
      this._functionAppClientId = (clientIdData.Value as string | null)?.trim() ?? '';

      if (!this._functionAppBaseUrl || !this._functionAppClientId) {
        console.warn(
          '[CepOptin] Tenant Properties not found. ' +
          'Run deploy/set-tenant-properties.ps1 to set them.'
        );
      }
    } catch (e) {
      console.error('[CepOptin] Error retrieving Tenant Properties:', e);
    }
  }

  private async _initApiClient(): Promise<void> {
    const baseUrl  = this._functionAppBaseUrl;
    const clientId = this._functionAppClientId;
    if (!baseUrl || !clientId) return;
    try {
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
        apiClient:          this._apiClient,
        userDisplayName:    this.context.pageContext.user.displayName,
        userEmail:          this.context.pageContext.user.email,
        userAadId:          this.context.pageContext.aadInfo?.userId?.toString() ?? '',
        isDarkTheme:        this._isDarkTheme,
        hasTeamsContext:    !!this.context.sdks.microsoftTeams,
        welcomeText:        this.properties.welcomeText ?? '',
        displayMode:        this.displayMode,
        onConfigureClick:   () => this.context.propertyPane.open(),
        graphClient:        this._graphClient,
        organizationName:   this.properties.organizationName ?? '',
        onWelcomeTextSave:  (text: string, orgName: string) => {
          this.properties.welcomeText      = text;
          this.properties.organizationName = orgName;
          this.render();
        },
      }
    );
    // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText',    semanticColors.bodyText    || null);
      this.domElement.style.setProperty('--link',        semanticColors.link        || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Copilot Engagement Program — opt-in experience settings.' },
          groups: [
            {
              groupName: '✨ Welcome Text',
              groupFields: [
                PropertyPaneLabel('welcomeText', {
                  text: 'Welcome text and organisation name are managed directly on the page via the ✨ Edit welcome text button in SharePoint edit mode.',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}



