import * as React from 'react';
import { Stack, Text, PrimaryButton, Icon } from '@fluentui/react';
import * as strings from 'CepOptinWebPartStrings';
import styles from '../CepOptin.module.scss';

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
  /** AI-personalised text (replaces static welcomeText once generated) */
  aiWelcomeText?: string;
  /** True while the AI request is in-flight — shows a blinking cursor */
  aiWelcomeLoading?: boolean;
  onNext: () => void;
}

interface IFeaturePillProps {
  icon: string;
  text: string;
}

const FeaturePill: React.FC<IFeaturePillProps> = ({ icon, text }) => (
  <Stack
    horizontal
    verticalAlign="center"
    tokens={{ childrenGap: 10 }}
    className={styles.featurePill}
  >
    <Icon iconName={icon} className={styles.featurePillIcon} />
    <Text variant="medium">{text}</Text>
  </Stack>
);

export const WelcomeStep: React.FC<IWelcomeStepProps> = ({ welcomeText, userName, aiWelcomeText, aiWelcomeLoading, onNext }) => {
  const displayText = (welcomeText ?? '').trim();
  // The text to render: prefer AI version once available, otherwise fall back to static
  const bodyText = (aiWelcomeText ?? '').trim() || displayText;

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
        <Text variant="large" className={styles.heroGreeting}>
          {strings.HelloGreeting.replace('{0}', userName)}
        </Text>
      </Stack>

      {/* Welcome text box */}
      <div className={styles.welcomeTextBox}>
        <Text
          variant="medium"
          className={[
            styles.welcomeText,
            aiWelcomeText ? styles.aiTextReveal : '',
          ].filter(Boolean).join(' ')}
        >
          {renderBoldText(bodyText)}
          {aiWelcomeLoading && !aiWelcomeText && (
            <span
              className={styles.typingCursor}
              aria-label={strings.AiWelcomeGenerating}
            />
          )}
        </Text>
        {aiWelcomeText && (
          <div className={styles.aiBadge}>
            ✨ {strings.AiWelcomeBadge}
          </div>
        )}
      </div>

      {/* Feature bullets */}
      <Stack tokens={{ childrenGap: 8 }}>
        <FeaturePill icon="LineChart"   text={strings.FeatureTrack} />
        <FeaturePill icon="StarFill"    text={strings.FeaturePoints} />
        <FeaturePill icon="RibbonSolid" text={strings.FeatureLevels} />
        <FeaturePill icon="People"      text={strings.FeatureLeaderboard} />
      </Stack>

      <PrimaryButton
        className={styles.ctaButton}
        text={strings.GetStarted}
        iconProps={{ iconName: 'ChevronRight' }}
        onClick={onNext}
      />
    </Stack>
  );
};
