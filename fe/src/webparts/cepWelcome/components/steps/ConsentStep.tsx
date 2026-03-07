import * as React from 'react';
import {
  Stack, Text, Checkbox, PrimaryButton, DefaultButton,
  MessageBar, MessageBarType, Icon, Separator,
  Spinner, SpinnerSize,
} from '@fluentui/react';
import * as strings from 'CepWelcomeWebPartStrings';
import styles from '../CepWelcome.module.scss';

export interface IConsentStepProps {
  department: string;
  team: string;
  enableNudges: boolean;
  consentChecked: boolean;
  submitting: boolean;
  errorMessage: string;
  onConsentChange: (checked: boolean) => void;
  onJoin: () => void;
  onBack: () => void;
}

interface ISummaryRowProps {
  label: string;
  value: string;
}

const SummaryRow: React.FC<ISummaryRowProps> = ({ label, value }) => (
  <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
    <Text
      variant="small"
      style={{ minWidth: 120, color: 'var(--bodySubtext, #605e5c)', flexShrink: 0 }}
    >
      {label}
    </Text>
    <Text variant="small"><strong>{value}</strong></Text>
  </Stack>
);

export const ConsentStep: React.FC<IConsentStepProps> = ({
  department,
  team,
  enableNudges,
  consentChecked,
  submitting,
  errorMessage,
  onConsentChange,
  onJoin,
  onBack,
}) => (
  <Stack className={styles.stepContainer} tokens={{ childrenGap: 20 }}>
    {/* Hero */}
    <Stack horizontalAlign="center" tokens={{ childrenGap: 8 }}>
      <Icon iconName="RocketLaunch" style={{ fontSize: 40, color: 'var(--themePrimary, #0078d4)' }} />
      <Text variant="xLarge" className={styles.stepTitle}>{strings.AlmostInTitle}</Text>
      <Text variant="medium" className={styles.stepSubtitle}>
        {strings.AlmostInSubtitle}
      </Text>
    </Stack>

    {/* Enrollment summary */}
    <Stack className={styles.summaryBox} tokens={{ childrenGap: 10 }}>
      <Text variant="mediumPlus"><strong>{strings.EnrollmentSummaryTitle}</strong></Text>
      <Separator />
      <SummaryRow label={strings.DepartmentLabel}    value={department || '\u2014'} />
      <SummaryRow label={strings.TeamLabel}          value={team       || '\u2014'} />
      <SummaryRow label={strings.NotificationsTitle} value={enableNudges ? strings.NotificationsEnabled : strings.NotificationsDisabled} />
    </Stack>

    {/* Data collection disclosure */}
    <Stack className={styles.transparencyBox} tokens={{ childrenGap: 8 }}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
        <Icon iconName="Info" />
        <Text variant="medium"><strong>{strings.DataCollectionTitle}</strong></Text>
      </Stack>
      <ul className={styles.transparencyList}>
        <li>{strings.DataItem1}</li>
        <li>{strings.DataItem2}</li>
        <li>{strings.DataItem3}</li>
      </ul>
      <Text variant="small" className={styles.note}>
        {strings.DataNote}
      </Text>
    </Stack>

    {/* Consent checkbox */}
    <Checkbox
      className={styles.consentCheckbox}
      label={strings.ConsentLabel}
      checked={consentChecked}
      onChange={(_e, checked) => onConsentChange(!!checked)}
      disabled={submitting}
    />

    {errorMessage && (
      <MessageBar messageBarType={MessageBarType.error} isMultiline>
        {errorMessage}
      </MessageBar>
    )}

    {/* Navigation */}
    <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="space-between">
      <DefaultButton
        text={strings.BackButton}
        iconProps={{ iconName: 'ChevronLeft' }}
        onClick={onBack}
        disabled={submitting}
      />
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
        {submitting && <Spinner size={SpinnerSize.small} />}
        <PrimaryButton
          text={submitting ? strings.JoinButtonLoading : strings.JoinButton}
          iconProps={submitting ? undefined : { iconName: 'AddFriend' }}
          disabled={!consentChecked || submitting}
          onClick={onJoin}
        />
      </Stack>
    </Stack>
  </Stack>
);
