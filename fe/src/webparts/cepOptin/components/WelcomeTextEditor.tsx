import * as React from 'react';
import {
  Stack, TextField,
  PrimaryButton, DefaultButton, Spinner, SpinnerSize,
  MessageBar, MessageBarType, Label, Separator,
} from '@fluentui/react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as strings from 'CepOptinWebPartStrings';
import { CopilotChatService } from '../../../services/CopilotChatService';
import { QUICK_PROMPTS } from '../../../services/CopilotSystemPrompts';

export interface IWelcomeTextEditorProps {
  graphClient: MSGraphClientV3 | undefined;
  welcomeText: string;
  organizationName: string;
  onPropertiesChange: (changes: Partial<{
    welcomeText: string;
    organizationName: string;
  }>) => void;
}

interface IWelcomeTextEditorState {
  orgName: string;
  generating: boolean;
  errorMessage: string;
  fallbackWarning: boolean;
  generatedText: string;
  selectedQuickPromptKey: string;
}

export class WelcomeTextEditor extends React.Component<
  IWelcomeTextEditorProps,
  IWelcomeTextEditorState
> {
  constructor(props: IWelcomeTextEditorProps) {
    super(props);
    this.state = {
      orgName:                props.organizationName || '',
      generating:             false,
      errorMessage:           '',
      fallbackWarning:        false,
      generatedText:          props.welcomeText || '',
      selectedQuickPromptKey: '',
    };
  }

  // ─── Handlers ─────────────────────────────────────────────────────────────

  private _handleGenerate = async (): Promise<void> => {
    const { orgName, selectedQuickPromptKey } = this.state;

    if (!orgName.trim()) {
      this.setState({ errorMessage: strings.OrgNameRequired });
      return;
    }

    this.setState({ generating: true, errorMessage: '', fallbackWarning: false });
    try {
      const { graphClient } = this.props;
      if (!graphClient) {
        const fallback = CopilotChatService.generateFallbackText(orgName.trim());
        this.setState({ generating: false, fallbackWarning: true, generatedText: fallback });
        this.props.onPropertiesChange({ welcomeText: fallback, organizationName: orgName.trim() });
        return;
      }
      const service = new CopilotChatService(graphClient);
      const quickPrompt = QUICK_PROMPTS.find(p => p.key === selectedQuickPromptKey);
      const { text, fromFallback } = await service.generateWelcomeText(
        orgName.trim(),
        quickPrompt?.instruction
      );

      this.setState({ generating: false, generatedText: text, fallbackWarning: fromFallback });
      this.props.onPropertiesChange({
        welcomeText:      text,
        organizationName: orgName.trim(),
      });
    } catch (e) {
      this.setState({
        generating:   false,
        errorMessage: strings.GenerationFailed.replace('{0}', (e as Error).message),
      });
    }
  };

  private _handleTextChange = (_e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    const text = value ?? '';
    this.setState({ generatedText: text });
    this.props.onPropertiesChange({ welcomeText: text });
  };

  private _handleClear = (): void => {
    this.setState({ generatedText: '', fallbackWarning: false, errorMessage: '' });
    this.props.onPropertiesChange({ welcomeText: '' });
  };

  private _toggleQuickPrompt = (key: string): void => {
    this.setState(prev => ({
      selectedQuickPromptKey: prev.selectedQuickPromptKey === key ? '' : key,
    }));
  };

  // ─── Render ───────────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const {
      orgName, generating, errorMessage, fallbackWarning,
      generatedText, selectedQuickPromptKey,
    } = this.state;

    return (
      <Stack tokens={{ childrenGap: 12 }} style={{ padding: '8px 0' }}>
        <TextField
          label={strings.OrgNameLabel}
          description={strings.OrgNameDescription}
          placeholder={strings.OrgNamePlaceholder}
          value={orgName}
          onChange={(_e, v) => {
            this.setState({ orgName: v ?? '' });
            this.props.onPropertiesChange({ organizationName: v ?? '' });
          }}
        />

        {/* Quick prompt chips */}
        <Stack tokens={{ childrenGap: 6 }}>
          <Label>{strings.ToneLabel}</Label>
          <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
            {QUICK_PROMPTS.map(p => {
              const isSelected = selectedQuickPromptKey === p.key;
              return (
                <DefaultButton
                  key={p.key}
                  text={p.label}
                  onClick={() => this._toggleQuickPrompt(p.key)}
                  styles={{
                    root: {
                      height:      28,
                      fontSize:    12,
                      padding:     '0 10px',
                      minWidth:    'auto',
                      borderColor: isSelected ? 'var(--themePrimary, #0078d4)' : undefined,
                      background:  isSelected ? 'var(--themeLight, #deecf9)'   : undefined,
                      color:       isSelected ? 'var(--themeDark, #005a9e)'     : undefined,
                    },
                    rootHovered: {
                      borderColor: 'var(--themePrimary, #0078d4)',
                    },
                  }}
                />
              );
            })}
          </Stack>
        </Stack>

        <PrimaryButton
          text={generating ? strings.GeneratingButton : strings.GenerateButton}
          iconProps={generating ? undefined : { iconName: 'Glimmer' }}
          onClick={this._handleGenerate}
          disabled={generating || !orgName.trim()}
        />

        {generating && (
          <Spinner size={SpinnerSize.small} label={strings.GeneratingButton} labelPosition="right" />
        )}

        {errorMessage && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ errorMessage: '' })}
          >
            {errorMessage}
          </MessageBar>
        )}

        {fallbackWarning && !errorMessage && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            onDismiss={() => this.setState({ fallbackWarning: false })}
          >
            {strings.FallbackWarning}
          </MessageBar>
        )}

        <Separator />

        <TextField
          label={strings.WelcomeTextLabel}
          description={strings.WelcomeTextDescription}
          multiline
          rows={7}
          value={generatedText}
          onChange={this._handleTextChange}
          placeholder={strings.WelcomeTextPlaceholder}
        />

        {generatedText && (
          <DefaultButton
            text={strings.ClearWelcomeText}
            iconProps={{ iconName: 'Delete' }}
            onClick={this._handleClear}
            styles={{ root: { color: 'var(--red, #a4262c)', borderColor: 'var(--red, #a4262c)' } }}
          />
        )}
      </Stack>
    );
  }
}
