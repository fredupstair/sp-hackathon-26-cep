import * as React from 'react';
import {
  Stack, Text,
  PrimaryButton, DefaultButton, IconButton,
  MessageBar, MessageBarType, TextField, Label,
} from '@fluentui/react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as strings from 'CepOptinWebPartStrings';
import { CopilotChatService } from '../../../services/CopilotChatService';
import { QUICK_PROMPTS } from '../../../services/CopilotSystemPrompts';
import styles from './CepOptin.module.scss';

// ─── Props ────────────────────────────────────────────────────────────────────

export interface IInlineWelcomeEditorProps {
  graphClient:      MSGraphClientV3 | undefined;
  organizationName: string;
  welcomeText:      string;
  hasExistingText:  boolean;
  onSave:           (text: string, orgName: string) => void;
  onDiscard?:       () => void; // shown only when hasExistingText === true
}

// ─── Component ────────────────────────────────────────────────────────────────

export const InlineWelcomeEditor: React.FC<IInlineWelcomeEditorProps> = ({
  graphClient,
  organizationName,
  welcomeText,
  hasExistingText,
  onSave,
  onDiscard,
}) => {
  const [orgName, setOrgName]          = React.useState(organizationName || '');
  const [selectedToneKey, setToneKey]  = React.useState('');
  const [promptText, setPromptText]    = React.useState('');
  const [generatedText, setGenText]    = React.useState(welcomeText || '');
  const [generating, setGenerating]    = React.useState(false);
  const [errorMessage, setError]       = React.useState('');
  const [fallbackWarning, setFallback] = React.useState(false);
  const textareaRef                    = React.useRef<HTMLTextAreaElement>(null);

  const autoResize = (): void => {
    const el = textareaRef.current;
    if (el) { el.style.height = 'auto'; el.style.height = `${el.scrollHeight}px`; }
  };

  const handleGenerate = async (): Promise<void> => {
    if (!orgName.trim()) { setError(strings.OrgNameRequired); return; }
    setGenerating(true);
    setError('');
    setFallback(false);
    try {
      if (!graphClient) {
        setGenText(CopilotChatService.generateFallbackText(orgName.trim()));
        setFallback(true);
      } else {
        const service = new CopilotChatService(graphClient);
        const { text, fromFallback } = await service.generateWelcomeText(
          orgName.trim(),
          promptText.trim() || undefined
        );
        setGenText(text);
        setFallback(fromFallback);
      }
    } catch (e) {
      setError(strings.GenerationFailed.replace('{0}', (e as Error).message));
    } finally {
      setGenerating(false);
    }
  };

  const handleToneClick = (key: string): void => {
    if (selectedToneKey === key) {
      setToneKey('');
      setPromptText('');
    } else {
      const p = QUICK_PROMPTS.find(q => q.key === key);
      setToneKey(key);
      setPromptText(p?.instruction ?? '');
    }
    setTimeout(autoResize, 0);
  };

  return (
    <Stack className={styles.inlineEditor} tokens={{ childrenGap: 20 }}>

      {/* ── Header ─────────────────────────────────────────────────────── */}
      <Stack horizontal verticalAlign="start" horizontalAlign="space-between">
        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="large" className={styles.inlineEditorTitle}>
            {strings.InlineEditorTitle}
          </Text>
          <Text className={styles.inlineEditorSubtitle}>
            Enter your organisation name, choose a style and click ✨ to generate your welcome text. You can edit it freely before saving.
          </Text>
        </Stack>
        {onDiscard && (
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            onClick={onDiscard}
            title={strings.InlineEditorDiscard}
            ariaLabel={strings.InlineEditorDiscard}
          />
        )}
      </Stack>

      {/* ── Organisation name ──────────────────────────────────────────── */}
      <Stack tokens={{ childrenGap: 4 }}>
        <Label className={styles.fieldLabel}>{strings.OrgNameLabel}</Label>
        <input
          type="text"
          className={styles.orgInput}
          placeholder="e.g. Contoso"
          value={orgName}
          onChange={e => setOrgName(e.target.value)}
          aria-label={strings.OrgNameLabel}
        />
      </Stack>

      {/* ── Prompt textarea ────────────────────────────────────────────── */}
      <Stack tokens={{ childrenGap: 4 }}>
        <Label className={styles.fieldLabel}>Prompt</Label>
        <div className={styles.promptWrapper}>
          <textarea
            ref={textareaRef}
            className={styles.promptTextarea}
            placeholder="Select a style below…"
            value={promptText}
            rows={3}
            readOnly
            aria-label="Prompt"
            aria-readonly="true"
          />
        </div>
      </Stack>

      {/* ── Style cards (3 × 2 grid) ────────────────────────────────────── */}
      <Stack tokens={{ childrenGap: 8 }}>
        <Text className={styles.toneLabel}>Choose a style</Text>
        <div className={styles.toneGrid}>
          {QUICK_PROMPTS.map(p => (
            <button
              key={p.key}
              className={[
                styles.toneChip,
                selectedToneKey === p.key ? styles.toneChipSelected : '',
              ].filter(Boolean).join(' ')}
              onClick={() => handleToneClick(p.key)}
              aria-pressed={selectedToneKey === p.key}
            >
              <span className={styles.toneChipLabel}>{p.label}</span>
              <span className={styles.toneChipDesc}>{p.description}</span>
            </button>
          ))}
        </div>
      </Stack>

      {/* ── Generate button ─────────────────────────────────────────────── */}
      <PrimaryButton
        text={generating ? strings.GeneratingButton : `✨ ${strings.GenerateButton}`}
        onClick={() => handleGenerate().catch(console.error)}
        disabled={generating || !orgName.trim() || !promptText.trim()}
        styles={{ root: { width: '100%' } }}
      />

      {/* ── Error / fallback messages ──────────────────────────────────── */}
      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>
          {errorMessage}
        </MessageBar>
      )}
      {fallbackWarning && !errorMessage && (
        <MessageBar messageBarType={MessageBarType.warning} onDismiss={() => setFallback(false)}>
          {strings.FallbackWarning}
        </MessageBar>
      )}

      {/* ── Welcome text ────────────────────────────────────────────────── */}
      <div className={styles.inlineEditorDivider} />
      <TextField
        label={strings.WelcomeTextLabel}
        description={strings.WelcomeTextDescription}
        multiline
        rows={5}
        value={generatedText}
        onChange={(_e, v) => setGenText(v ?? '')}
        placeholder={strings.WelcomeTextPlaceholder}
        styles={{
          fieldGroup: { borderRadius: 8, overflow: 'hidden' },
          field: { lineHeight: '1.6', padding: '10px 12px' },
        }}
      />
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text className={styles.aiBadge}>✨ {strings.AiWelcomeBadge}</Text>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text={strings.ClearWelcomeText}
            iconProps={{ iconName: 'Delete' }}
            onClick={() => setGenText('')}
            styles={{ root: { color: 'var(--red, #a4262c)', borderColor: 'var(--red, #a4262c)' } }}
          />
          <PrimaryButton
            text={strings.InlineEditorSave}
            iconProps={{ iconName: 'Save' }}
            onClick={() => onSave(generatedText, orgName)}
          />
        </Stack>
      </Stack>

    </Stack>
  );
};
