import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
} from '@microsoft/sp-application-base';
import { AadHttpClient, SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'CepBarApplicationCustomizerStrings';
import { CepApiClient } from '../../services/CepApiClient';
import { CepBar } from './components/CepBar';

const LOG_SOURCE: string = 'CepBarApplicationCustomizer';

// Keys must match what deploy/set-tenant-properties.ps1 sets
const TENANT_KEY_BASE_URL  = 'CEP_FunctionAppBaseUrl';
const TENANT_KEY_CLIENT_ID = 'CEP_FunctionAppClientId';

/**
 * Properties set via the Custom Action registration (clientSideComponentProperties JSON).
 * See deploy/deploy-azure.ps1 for the registration command.
 */
export interface ICepBarApplicationCustomizerProperties {
  /** Absolute URL of the page hosting the CEP Dashboard web part. */
  dashboardPageUrl: string;
  /** Absolute URL of the page hosting the CEP Opt-in web part. */
  optinPageUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CepBarApplicationCustomizer
  extends BaseApplicationCustomizer<ICepBarApplicationCustomizerProperties> {

  private _container: HTMLDivElement | undefined;
  private _apiClient: CepApiClient | undefined;
  private _functionAppBaseUrl: string = '';
  private _functionAppClientId: string = '';
  private _silverThreshold: number = 500;
  private _goldThreshold: number = 1500;

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    await this._loadTenantProperties();
    await this._initApiClient();
    this._renderChip();
    return Promise.resolve();
  }

  protected onDispose(): void {
    if (this._container) {
      // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
      ReactDom.unmountComponentAtNode(this._container);
      this._container.remove();
      this._container = undefined;
    }
  }

  // ── Tenant Properties ──────────────────────────────────────────────────────

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
        Log.warn(LOG_SOURCE,
          'Tenant Properties CEP_FunctionAppBaseUrl / CEP_FunctionAppClientId not found. ' +
          'Run deploy/set-tenant-properties.ps1.');
      }
    } catch (e) {
      Log.error(LOG_SOURCE, e as Error);
    }
  }

  // ── AAD client ─────────────────────────────────────────────────────────────

  private async _initApiClient(): Promise<void> {
    const baseUrl  = this._functionAppBaseUrl;
    const clientId = this._functionAppClientId;
    if (!baseUrl || !clientId) return;

    try {
      const aadResourceUri = `api://${clientId}`;
      const aadClient: AadHttpClient =
        await this.context.aadHttpClientFactory.getClient(aadResourceUri);

      this._apiClient = new CepApiClient(
        aadClient,
        baseUrl,
        this.context.pageContext.aadInfo?.userId?.toString() ??
          this.context.pageContext.user.email,
        this.context.pageContext.user.email,
        this.context.pageContext.user.displayName
      );
    } catch (e) {
      Log.error(LOG_SOURCE, e as Error);
    }
  }

  // ── Render ─────────────────────────────────────────────────────────────────

  private _renderChip(): void {
    if (!this._container) {
      this._container = document.createElement('div');
      this._container.id = 'cep-chip-root';
      document.body.appendChild(this._container);
    }

    const element = React.createElement(CepBar, {
      apiClient:         this._apiClient,
      dashboardPageUrl:  this.properties.dashboardPageUrl  ?? '',
      optinPageUrl:      this.properties.optinPageUrl      ?? '',
      silverThreshold:   this._silverThreshold,
      goldThreshold:     this._goldThreshold,
    });

    // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
    ReactDom.render(element, this._container);
  }
}
