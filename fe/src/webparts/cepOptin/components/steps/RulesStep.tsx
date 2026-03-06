import * as React from 'react';
import {
  Stack, Text, PrimaryButton, DefaultButton,
  Icon, Separator,
} from '@fluentui/react';
import * as strings from 'CepOptinWebPartStrings';
import styles from '../CepOptin.module.scss';

export interface IRulesStepProps {
  onNext: () => void;
  onBack: () => void;
}

interface IRuleCardProps {
  icon: string;
  color: string;
  title: string;
  body: string;
}

const RuleCard: React.FC<IRuleCardProps> = ({ icon, color, title, body }) => (
  <Stack className={styles.ruleCard} tokens={{ childrenGap: 8 }}>
    <Icon iconName={icon} style={{ color, fontSize: 32 }} />
    <Text variant="mediumPlus"><strong>{title}</strong></Text>
    <Text variant="small">{body}</Text>
  </Stack>
);

interface IPrivacyItemProps {
  icon: string;
  color: string;
  text: string;
}

const PrivacyItem: React.FC<IPrivacyItemProps> = ({ icon, color, text }) => (
  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
    <Icon iconName={icon} style={{ color, fontSize: 12, flexShrink: 0 }} />
    <Text variant="small">{text}</Text>
  </Stack>
);

export const RulesStep: React.FC<IRulesStepProps> = ({ onNext, onBack }) => (
  <Stack className={styles.stepContainer} tokens={{ childrenGap: 20 }}>
    <Text variant="xLarge" className={styles.stepTitle}>How it works</Text>

    {/* 3 rule cards */}
    <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
      <RuleCard
        icon="LineChart"
        color="#0078d4"
        title="Use Copilot"
        body="Use Microsoft 365 Copilot in any app — Word, Excel, Outlook, Teams, and more — as part of your everyday work."
      />
      <RuleCard
        icon="StarFill"
        color="#ffb900"
        title="Earn Points"
        body="1 point for every prompt you send. Points accumulate over the month and reset at month-end for a fresh competition."
      />
      <RuleCard
        icon="RibbonSolid"
        color="#107c10"
        title="Level Up"
        body="Reach Bronze, Silver, and Gold tiers as your usage grows. Badges are permanent — earn them once, keep them forever."
      />
    </Stack>

    <Separator />

    {/* Privacy & transparency */}
    <Stack className={styles.transparencyBox} tokens={{ childrenGap: 8 }}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
        <Icon iconName="Shield" />
        <Text variant="mediumPlus"><strong>Privacy &amp; Transparency</strong></Text>
      </Stack>
      <Stack tokens={{ childrenGap: 6 }}>
        <PrivacyItem icon="CheckMark" color="#107c10" text="Number of prompts sent to Copilot (content is never read)" />
        <PrivacyItem icon="CheckMark" color="#107c10" text="App used (Word, Excel, Outlook, Teams…)" />
        <PrivacyItem icon="CheckMark" color="#107c10" text="Daily activity date (aggregated, not per-interaction)" />
        <PrivacyItem icon="Cancel"    color="#a4262c" text="Prompt text, responses, and attachments — never stored" />
        <PrivacyItem icon="Cancel"    color="#a4262c" text="Personal content or sensitive data — never collected" />
      </Stack>
      <Text variant="small" className={styles.note}>
        Tracking runs once per day in aggregate. You can withdraw at any time.
      </Text>
    </Stack>

    {/* Navigation */}
    <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="space-between">
      <DefaultButton text="Back" iconProps={{ iconName: 'ChevronLeft' }} onClick={onBack} />
      <PrimaryButton text="Continue" iconProps={{ iconName: 'ChevronRight' }} onClick={onNext} />
    </Stack>
  </Stack>
);
