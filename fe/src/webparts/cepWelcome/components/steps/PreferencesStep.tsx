import * as React from 'react';
import {
  Stack, Text, TextField, Toggle,
  PrimaryButton, DefaultButton, Icon,
} from '@fluentui/react';
import * as strings from 'CepWelcomeWebPartStrings';
import styles from '../CepWelcome.module.scss';

export interface IPreferencesStepProps {
  department: string;
  /** True when department was auto-fetched from Azure AD / Microsoft Graph */
  departmentReadOnly?: boolean;
  team: string;
  enableNudges: boolean;
  onDepartmentChange: (value: string) => void;
  onTeamChange: (value: string) => void;
  onNudgesChange: (value: boolean) => void;
  onNext: () => void;
  onBack: () => void;
}

export const PreferencesStep: React.FC<IPreferencesStepProps> = ({
  department,
  departmentReadOnly,
  team,
  enableNudges,
  onDepartmentChange,
  onTeamChange,
  onNudgesChange,
  onNext,
  onBack,
}) => (
  <Stack className={styles.stepContainer} tokens={{ childrenGap: 20 }}>
    <Stack tokens={{ childrenGap: 4 }}>
      <Text variant="xLarge" className={styles.stepTitle}>{strings.ProfileTitle}</Text>
      <Text variant="medium" className={styles.stepSubtitle}>
        {strings.ProfileSubtitle}
      </Text>
    </Stack>

    {/* Profile fields */}
    <Stack tokens={{ childrenGap: 14 }}>
      <Stack tokens={{ childrenGap: 4 }}>
        <Text variant="medium" style={{ fontWeight: 600 }}>{strings.DepartmentLabel}</Text>
        {departmentReadOnly ? (
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="mediumPlus">{department}</Text>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
              <Icon iconName="Lock" style={{ fontSize: 12, color: '#605e5c' }} />
              <Text variant="small" style={{ color: '#605e5c', fontStyle: 'italic' }}>
                Managed by your organization
              </Text>
            </Stack>
          </Stack>
        ) : (
          <TextField
            placeholder={strings.DepartmentPlaceholder}
            value={department}
            onChange={(_e, v) => onDepartmentChange(v ?? '')}
          />
        )}
      </Stack>
      <TextField
        label={strings.TeamLabel}
        placeholder={strings.TeamPlaceholder}
        value={team}
        onChange={(_e, v) => onTeamChange(v ?? '')}
      />
    </Stack>

    {/* Notifications */}
    <Stack className={styles.notificationsBox} tokens={{ childrenGap: 10 }}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
        <Icon iconName="Ringer" className={styles.notifIcon} />
        <Text variant="mediumPlus"><strong>{strings.NotificationsTitle}</strong></Text>
      </Stack>
      <Toggle
        label={strings.NudgesLabel}
        onText={strings.NudgesOn}
        offText={strings.NudgesOff}
        checked={enableNudges}
        onChange={(_e, checked) => onNudgesChange(!!checked)}
        inlineLabel
      />
      <Text variant="small" className={styles.note}>
        {strings.NudgesNote}
      </Text>
    </Stack>

    {/* Navigation */}
    <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="space-between">
      <DefaultButton text={strings.BackButton}     iconProps={{ iconName: 'ChevronLeft'  }} onClick={onBack} />
      <PrimaryButton text={strings.ContinueButton} iconProps={{ iconName: 'ChevronRight' }} onClick={onNext} />
    </Stack>
  </Stack>
);
