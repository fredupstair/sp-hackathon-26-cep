import * as React from 'react';
import { Stack, Text, PrimaryButton, Icon } from '@fluentui/react';
import * as strings from 'CepWelcomeWebPartStrings';
import styles from '../CepWelcome.module.scss';

/**
 * Renders text that may contain:
 *  - **bold** markers  → <strong>
 *  - \n line breaks   → <br /> + block layout
 *  - • bullet lines   → flex row with a styled dot
 */
function renderFormattedText(text: string, trailingNode?: React.ReactNode): React.ReactNode[] {
  // Normalize literal \u2022 escape sequences (Copilot sometimes sends them as text)
  const normalised = text.replace(/\\u2022\s*/g, '- ');
  const lines = normalised.split('\n').filter(l => l.length > 0);
  return lines.map((line, lineIdx) => {
    const isLast = lineIdx === lines.length - 1;
    const isBullet = line.startsWith('- ') || line.startsWith('\u2022 ');
    const content  = isBullet ? line.slice(2) : line;
    const segments = content.split(/(\*\*[^*]+\*\*)/g);
    const boldParts: React.ReactNode[] = segments.map((seg, si) =>
      seg.startsWith('**') && seg.endsWith('**')
        ? <strong key={`s-${lineIdx}-${si}`}>{seg.slice(2, -2)}</strong>
        : seg
    );
    if (isBullet) {
      return (
        <div key={`line-${lineIdx}`} className={styles.bulletLine}>
          <span className={styles.bulletDot}>•</span>
          <span>{boldParts}{isLast && trailingNode}</span>
        </div>
      );
    } else {
      return <div key={`line-${lineIdx}`} className={styles.textLine}>{boldParts}{isLast && trailingNode}</div>;
    }
  });
}

/** Copilot logo: shows a placeholder SVG. Pass `copilotLogoUrl` prop to use your own image. */
const CopilotLogoIcon: React.FC<{ src?: string; className?: string }> = ({ src, className }) =>
  src
    ? <img src={src} className={className} alt="Copilot" />
    : (
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" className={className} aria-hidden="true">
        <defs>
          <linearGradient id="cg" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stopColor="#6B3FA0" />
            <stop offset="60%" stopColor="#2563EB" />
            <stop offset="100%" stopColor="#00B4C8" />
          </linearGradient>
        </defs>
        <circle cx="16" cy="16" r="16" fill="url(#cg)" />
        <polygon points="16,8 17.8,14.2 24,14.2 19,18.1 21,24 16,20.2 11,24 13,18.1 8,14.2 14.2,14.2" fill="white" opacity="0.92" />
      </svg>
    );

export interface IWelcomeStepProps {
  welcomeText: string;
  userName: string;
  /** Circular profile photo URL (from Graph /me/photo/$value) */
  userPhotoUrl?: string;
  /** Optional URL for the Copilot logo (absolute URL or SharePoint-hosted path). Falls back to built-in placeholder SVG. */
  copilotLogoUrl?: string;
  /** AI-personalised text (streamed progressively) */
  aiWelcomeText?: string;
  /** True while the conversation is being created — shows bouncing dots */
  aiWelcomeLoading?: boolean;
  /** True while the Copilot response is being streamed — shows typing cursor */
  aiWelcomeStreaming?: boolean;
  /** Triggers the AI chat generation (called only once) */
  onStartChat: () => void;
  onNext: () => void;
}

export const WelcomeStep: React.FC<IWelcomeStepProps> = ({ welcomeText, userName, userPhotoUrl, copilotLogoUrl, aiWelcomeText, aiWelcomeLoading, aiWelcomeStreaming, onStartChat, onNext }) => {
  const displayText = (welcomeText ?? '').trim();
  const firstName = userName.split(' ')[0];

  // Whether the chat section has been revealed (persists across re-renders via parent state)
  const chatVisible = !!(aiWelcomeLoading || aiWelcomeText);

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
  const aiDone = !!(aiWelcomeText && !aiWelcomeStreaming);

  // Blink: 3 rapid pulses when AI finishes, then repeat every 5 seconds
  const [blink, setBlink] = React.useState(false);
  const prevInProgress = React.useRef(aiInProgress);
  const blinkTimer = React.useRef<ReturnType<typeof setInterval> | undefined>(undefined);

  React.useEffect(() => {
    // Detect transition from in-progress to done
    if (prevInProgress.current && !aiInProgress && aiDone) {
      // Start blink cycle
      setBlink(true);
      blinkTimer.current = setInterval(() => {
        setBlink(true);
        // Each blink burst lasts ~1.5s (3 × 0.5s animation), then turns off
        setTimeout(() => setBlink(false), 1500);
      }, 3000);
    }
    prevInProgress.current = aiInProgress;
    return undefined;
  }, [aiInProgress, aiDone]);

  // If user already has aiWelcomeText on mount (came back), start blink cycle
  React.useEffect(() => {
    if (aiDone && !blinkTimer.current) {
      setBlink(true);
      blinkTimer.current = setInterval(() => {
        setBlink(true);
        setTimeout(() => setBlink(false), 1500);
      }, 5000);
    }
    return () => {
      if (blinkTimer.current) {
        clearInterval(blinkTimer.current);
        blinkTimer.current = undefined;
      }
    };
  }, [aiDone]);

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
          {renderFormattedText(displayText)}
        </Text>
      </div>

      {/* "Let's start" button — shown only before chat is revealed */}
      {!chatVisible && (
        <PrimaryButton
          className={styles.ctaButton}
          text={strings.LetsStart}
          iconProps={{ iconName: 'Play' }}
          onClick={onStartChat}
        />
      )}

      {/* Chat-style Copilot conversation */}
      {chatVisible && (
        <div className={styles.chatContainer}>
          {/* User message bubble */}
          <div className={styles.chatRow + ' ' + styles.chatRowUser}>
            <div className={styles.chatBubbleUser}>
              <div className={styles.chatBubbleUserName}>
              {userPhotoUrl
                ? <img src={userPhotoUrl} className={styles.userAvatar} alt={firstName} />
                : <Icon iconName="Contact" className={styles.chatAvatarIcon} />}
                <Text variant="smallPlus" className={styles.chatNameLabel}>{firstName}</Text>
              </div>
              <Text variant="medium">{strings.CopilotChatUserMessage}</Text>
            </div>
          </div>

          {/* Copilot response bubble */}
          <div className={styles.chatRow + ' ' + styles.chatRowCopilot}>
            <div className={styles.chatBubbleCopilot}>
              <div className={styles.chatBubbleCopilotName}>
                <CopilotLogoIcon src={copilotLogoUrl} className={styles.copilotLogoImg} />
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
                  <div className={styles.copilotBubbleText}>
                    {renderFormattedText(
                      aiWelcomeText!,
                      aiWelcomeStreaming ? <span className={styles.streamingCursor} /> : undefined
                    )}
                  </div>
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

      {/* Get Started — only visible after chat is revealed */}
      {chatVisible && (
        <PrimaryButton
          className={[styles.ctaButton, blink ? styles.ctaBlink : ''].filter(Boolean).join(' ')}
          text={strings.GetStarted}
          iconProps={{ iconName: 'ChevronRight' }}
          disabled={aiInProgress}
          onClick={onNext}
        />
      )}
    </Stack>
  );
};
