import * as React from 'react';
import { Stack, Text, PrimaryButton, Icon } from '@fluentui/react';
import * as strings from 'CepWelcomeWebPartStrings';
import styles from '../CepWelcome.module.scss';

/** Converts **word** tokens to <strong> spans, returns a React.ReactNode array */
function renderBoldText(text: string): React.ReactNode[] {
  const parts = text.split(/(\*\*[^*]+\*\*)/g);
  return parts.map((part, i) => {
    if (part.startsWith('**') && part.endsWith('**')) {
      return <strong key={i}>{part.slice(2, -2)}</strong>;
    }
    return part;
  });
}

export interface IWelcomeStepProps {
  welcomeText: string;
  userName: string;
  /** AI-personalised text (streamed progressively) */
  aiWelcomeText?: string;
  /** True while the conversation is being created — shows bouncing dots */
  aiWelcomeLoading?: boolean;
  /** True while the Copilot response is being streamed — shows typing cursor */
  aiWelcomeStreaming?: boolean;
  onNext: () => void;
}

export const WelcomeStep: React.FC<IWelcomeStepProps> = ({ welcomeText, userName, aiWelcomeText, aiWelcomeLoading, aiWelcomeStreaming, onNext }) => {
  const displayText = (welcomeText ?? '').trim();
  const firstName = userName.split(' ')[0];

  // Rotating loading hints
  const loadingHints = React.useMemo(() => [
    strings.CopilotLoadingHint,
    strings.CopilotLoadingHint2,
    strings.CopilotLoadingHint3,
    strings.CopilotLoadingHint4,
    strings.CopilotLoadingHint5,
  ], []);
  const [hintIndex, setHintIndex] = React.useState(0);
  React.useEffect(() => {
    if (!aiWelcomeLoading || aiWelcomeText) return;
    const timer = setInterval(() => {
      setHintIndex(i => (i + 1) % loadingHints.length);
    }, 2500);
    return () => clearInterval(timer);
  }, [aiWelcomeLoading, aiWelcomeText, loadingHints.length]);

  const aiInProgress = !!(aiWelcomeLoading || aiWelcomeStreaming);
  // Track when AI finishes to trigger blink
  const [blink, setBlink] = React.useState(false);
  const prevInProgress = React.useRef(aiInProgress);
  React.useEffect(() => {
    if (prevInProgress.current && !aiInProgress) {
      setBlink(true);
      const t = setTimeout(() => setBlink(false), 1500);
      return () => clearTimeout(t);
    }
    prevInProgress.current = aiInProgress;
    return undefined;
  }, [aiInProgress]);

  return (
    <Stack className={styles.stepContainer} tokens={{ childrenGap: 24 }}>
      {/* Hero */}
      <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
        <div className={styles.heroIconWrapper}>
          <Icon iconName="Trophy2" className={styles.heroIcon} />
        </div>
        <Text variant="xxLarge" className={styles.heroTitle}>
          {strings.ProgramTitle}
        </Text>
      </Stack>

      {/* Static welcome text box — always visible, never overridden */}
      <div className={styles.welcomeTextBox}>
        <Text variant="medium" className={styles.welcomeText}>
          {renderBoldText(displayText)}
        </Text>
      </div>

      {/* Chat-style Copilot conversation */}
      {(aiWelcomeLoading || aiWelcomeText) && (
        <div className={styles.chatContainer}>
          {/* User message bubble */}
          <div className={styles.chatRow + ' ' + styles.chatRowUser}>
            <div className={styles.chatBubbleUser}>
              <div className={styles.chatBubbleUserName}>
                <Icon iconName="Contact" className={styles.chatAvatarIcon} />
                <Text variant="smallPlus" className={styles.chatNameLabel}>{firstName}</Text>
              </div>
              <Text variant="medium">{strings.CopilotChatUserMessage}</Text>
            </div>
          </div>

          {/* Copilot response bubble */}
          <div className={styles.chatRow + ' ' + styles.chatRowCopilot}>
            <div className={styles.chatBubbleCopilot}>
              <div className={styles.chatBubbleCopilotName}>
                <Icon iconName="ChatBot" className={styles.copilotBubbleIcon} />
                <Text variant="smallPlus" className={styles.copilotBubbleLabel}>Copilot</Text>
              </div>
              {aiWelcomeLoading && !aiWelcomeText ? (
                <div className={styles.copilotBubbleLoading}>
                  <Text variant="small" className={styles.copilotLoadingHint} key={hintIndex}>
                    {loadingHints[hintIndex]}
                  </Text>
                  <div className={styles.copilotDotsRow}>
                    <span className={styles.copilotDot} />
                    <span className={styles.copilotDot} />
                    <span className={styles.copilotDot} />
                  </div>
                </div>
              ) : (
                <div className={styles.copilotBubbleBody}>
                  <Text variant="medium">
                    {renderBoldText(aiWelcomeText!)}
                    {aiWelcomeStreaming && (
                      <span className={styles.streamingCursor} />
                    )}
                  </Text>
                  {!aiWelcomeStreaming && aiWelcomeText && (
                    <div className={styles.aiBadge}>
                      ✨ {strings.AiWelcomeBadge}
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      <PrimaryButton
        className={[styles.ctaButton, blink ? styles.ctaBlink : ''].filter(Boolean).join(' ')}
        text={strings.GetStarted}
        iconProps={{ iconName: 'ChevronRight' }}
        disabled={aiInProgress}
        onClick={onNext}
      />
    </Stack>
  );
};
