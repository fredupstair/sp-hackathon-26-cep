import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { AadHttpClient } from '@microsoft/sp-http';

import CepOptin from './components/CepOptin';
import { ICepOptinProps } from './components/ICepOptinProps';
import { CepApiClient } from '../../services/CepApiClient';

export interface ICepOptinWebPartProps {
  functionAppBaseUrl: string;
  /** Client ID (appId) of the CEP-Backend app registration – used to build the AAD resource URI */
  functionAppClientId: string;
}

export default class CepOptinWebPart extends BaseClientSideWebPart<ICepOptinWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _aadClient: AadHttpClient | undefined;
  private _apiClient: CepApiClient | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();
    await this._initApiClient();
  }

  private async _initApiClient(): Promise<void> {
    const baseUrl = this.properties.functionAppBaseUrl?.trim();
    const clientId = this.properties.functionAppClientId?.trim();
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
        functionAppBaseUrl: this.properties.functionAppBaseUrl || '',
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

  protected async onPropertyPaneFieldChanged(propertyPath: string): Promise<void> {
    if (propertyPath === 'functionAppBaseUrl' || propertyPath === 'functionAppClientId') {
      this._aadClient = undefined;
      this._apiClient = undefined;
      await this._initApiClient();
      this.render();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Configura la connessione al backend CEP' },
          groups: [
            {
              groupName: 'Backend Azure Functions',
              groupFields: [
                PropertyPaneTextField('functionAppBaseUrl', {
                  label: 'Function App Base URL',
                  placeholder: 'https://<nome-app>.azurewebsites.net',
                  description: "URL base della Function App, senza trailing slash.",
                  onGetErrorMessage: (value: string) => {
                    if (!value) return 'Campo obbligatorio';
                    try { new URL(value); return ''; } catch { return 'URL non valido'; }
                  },
                }),
                PropertyPaneTextField('functionAppClientId', {
                  label: 'App Registration Client ID',
                  placeholder: '00000000-0000-0000-0000-000000000000',
                  description: "Client ID (appId) dell'app registration CEP-Backend in Entra ID.",
                  onGetErrorMessage: (value: string) => {
                    if (!value) return 'Campo obbligatorio';
                    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value)
                      ? '' : 'Deve essere un GUID valido';
                  },
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

