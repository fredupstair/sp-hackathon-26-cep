import * as React from 'react';
import {
  Stack, Text, TextField, Toggle,
  PrimaryButton, DefaultButton, Icon,
} from '@fluentui/react';
import * as strings from 'CepOptinWebPartStrings';
import styles from '../CepOptin.module.scss';

export interface IPreferencesStepProps {
  department: string;
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
      <TextField
        label={strings.DepartmentLabel}
        placeholder={strings.DepartmentPlaceholder}
        value={department}
        onChange={(_e, v) => onDepartmentChange(v ?? '')}
      />
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
