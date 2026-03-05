import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Checkbox,
  TextField,
  Toggle,
  Stack,
  Text,
  Separator,
  Icon,
  Dialog,
  DialogType,
  DialogFooter,
} from '@fluentui/react';
import styles from './CepOptin.module.scss';
import type { ICepOptinProps } from './ICepOptinProps';
import type { IUserSummary } from '../../../services/CepApiModels';
import * as strings from 'CepOptinWebPartStrings';

// Helper to format strings with placeholders like {0}, {1}
const formatString = (template: string, ...args: (string | number)[]): string => {
  return template.replace(/{(\d+)}/g, (match, index) => {
    return typeof args[index] !== 'undefined' ? String(args[index]) : match;
  });
};

type LoadingState = 'idle' | 'loading' | 'loaded' | 'error';

interface ICepOptinState {
  loadingState: LoadingState;
  userSummary: IUserSummary | undefined;   // undefined = not enrolled
  errorMessage: string;
  successMessage: string;
  // form fields (enrollment)
  department: string;
  team: string;
  enableNudges: boolean;
  consentChecked: boolean;
  // actions
  submitting: boolean;
  // leave confirmation dialog
  showLeaveDialog: boolean;
}

export default class CepOptin extends React.Component<ICepOptinProps, ICepOptinState> {

  constructor(props: ICepOptinProps) {
    super(props);
    this.state = {
      loadingState: 'idle',
      userSummary: undefined,
      errorMessage: '',
      successMessage: '',
      department: '',
      team: '',
      enableNudges: true,
      consentChecked: false,
      submitting: false,
      showLeaveDialog: false,
    };
  }

  public componentDidMount(): void {
    this._loadEnrollmentStatus().catch(console.error);
  }

  public componentDidUpdate(prevProps: ICepOptinProps): void {
    if (prevProps.apiClient !== this.props.apiClient) {
      this._loadEnrollmentStatus().catch(console.error);
    }
  }

  // ─── Data loading ──────────────────────────────────────────────────────────

  private async _loadEnrollmentStatus(): Promise<void> {
    const { apiClient } = this.props;
    if (!apiClient) {
      this.setState({ loadingState: 'loaded', userSummary: undefined });
      return;
    }
    this.setState({ loadingState: 'loading', errorMessage: '', successMessage: '' });
    try {
      const summary = await apiClient.getMeSummary();
      this.setState({ loadingState: 'loaded', userSummary: summary });
    } catch (e) {
      this.setState({ loadingState: 'error', errorMessage: formatString(strings.LoadError, (e as Error).message) });
    }
  }

  // ─── Actions ───────────────────────────────────────────────────────────────

  private _handleJoin = async (): Promise<void> => {
    const { apiClient } = this.props;
    const { department, team, enableNudges } = this.state;
    if (!apiClient) return;
    this.setState({ submitting: true, errorMessage: '', successMessage: '' });
    try {
      await apiClient.join({ department, team, isEngagementNudgesEnabled: enableNudges });
      const summary = await apiClient.getMeSummary();
      this.setState({ submitting: false, userSummary: summary, successMessage: strings.JoinSuccess });
    } catch (e) {
      this.setState({ submitting: false, errorMessage: formatString(strings.JoinError, (e as Error).message) });
    }
  };

  private _handleLeave = async (): Promise<void> => {
    const { apiClient } = this.props;
    if (!apiClient) return;
    this.setState({ submitting: true, showLeaveDialog: false, errorMessage: '', successMessage: '' });
    try {
      await apiClient.leave();
      this.setState({ submitting: false, userSummary: undefined, successMessage: strings.LeaveSuccess });
    } catch (e) {
      this.setState({ submitting: false, errorMessage: formatString(strings.LeaveError, (e as Error).message) });
    }
  };

  private _handleNudgeToggle = async (_ev: React.MouseEvent<HTMLElement>, checked?: boolean): Promise<void> => {
    const { userSummary } = this.state;
    if (!userSummary) return;
    // Optimistic update – backend update not yet implemented, placeholder
    this.setState({ userSummary: { ...userSummary, isEngagementNudgesEnabled: !!checked } });
  };

  // ─── Rendering helpers ─────────────────────────────────────────────────────

  private _renderNotConfigured(): React.ReactElement {
    return (
      <MessageBar messageBarType={MessageBarType.warning} isMultiline>
        <strong>{strings.NotConfiguredTitle}</strong>{' '}
        {strings.NotConfiguredMessage}
      </MessageBar>
    );
  }

  private _renderLoading(): React.ReactElement {
    return (
      <Stack className={styles.container} horizontalAlign="center" tokens={{ padding: 24 }}>
        <Spinner size={SpinnerSize.large} label={strings.Loading} />
      </Stack>
    );
  }

  private _renderEnrollmentForm(): React.ReactElement {
    const { userDisplayName } = this.props;
    const { department, team, enableNudges, consentChecked, submitting, errorMessage, successMessage } = this.state;
    const canSubmit = consentChecked && !submitting;

    return (
      <Stack className={`${styles.container} ${styles.enrollForm}`} tokens={{ childrenGap: 16 }}>
        {/* Header */}
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
          <Icon iconName="Trophy2" className={styles.headerIcon} />
          <Text variant="xLarge" className={styles.title}>{strings.WebPartTitle}</Text>
        </Stack>
        <Text variant="medium" className={styles.subtitle}>
          {formatString(strings.WelcomeMessage, userDisplayName)}
        </Text>

        <Separator />

        {/* Transparency */}
        <Stack className={styles.transparencyBox} tokens={{ childrenGap: 8 }}>
          <Text variant="mediumPlus"><Icon iconName="Shield" /> {strings.TransparencyTitle}</Text>
          <ul className={styles.transparencyList}>
            <li>{strings.TransparencyItem1}</li>
            <li>{strings.TransparencyItem2}</li>
            <li>{strings.TransparencyItem3}</li>
          </ul>
          <Text variant="small" className={styles.note}>
            {strings.TransparencyNote}
          </Text>
        </Stack>

        <Separator />

        {/* Form fields */}
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField
            label={strings.DepartmentLabel}
            placeholder={strings.DepartmentPlaceholder}
            value={department}
            onChange={(_e, v) => this.setState({ department: v || '' })}
          />
          <TextField
            label={strings.TeamLabel}
            placeholder={strings.TeamPlaceholder}
            value={team}
            onChange={(_e, v) => this.setState({ team: v || '' })}
          />
          <Toggle
            label={strings.NudgesLabel}
            onText={strings.NudgesOn}
            offText={strings.NudgesOff}
            checked={enableNudges}
            onChange={(_e, checked) => this.setState({ enableNudges: !!checked })}
            inlineLabel
          />
          <Text variant="small" className={styles.note}>
            {strings.NudgesDescription}
          </Text>
        </Stack>

        <Separator />

        {/* Consent */}
        <Checkbox
          label={strings.ConsentLabel}
          checked={consentChecked}
          onChange={(_e, checked) => this.setState({ consentChecked: !!checked })}
          className={styles.consentCheckbox}
        />

        {/* Messages */}
        {errorMessage && <MessageBar messageBarType={MessageBarType.error}>{errorMessage}</MessageBar>}
        {successMessage && <MessageBar messageBarType={MessageBarType.success}>{successMessage}</MessageBar>}

        {/* CTA */}
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={submitting ? strings.JoinButtonLoading : strings.JoinButton}
            disabled={!canSubmit}
            onClick={this._handleJoin}
            iconProps={submitting ? undefined : { iconName: 'AddFriend' }}
          />
        </Stack>
      </Stack>
    );
  }

  private _renderEnrolledView(): React.ReactElement {
    const { userSummary, submitting, errorMessage, successMessage, showLeaveDialog } = this.state;
    if (!userSummary) return <></>;

    const levelColor: Record<string, string> = { Bronze: '#CD7F32', Silver: '#A8A9AD', Gold: '#FFD700' };
    const levelIcon: Record<string, string> = { Bronze: 'Medal', Silver: 'Medal', Gold: 'FavoriteStarFill' };
    const color = levelColor[userSummary.currentLevel] || '#CD7F32';
    const icon = levelIcon[userSummary.currentLevel] || 'Medal';

    return (
      <Stack className={styles.container} tokens={{ childrenGap: 16 }}>
        {/* Header */}
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
          <Icon iconName="Trophy2" className={styles.headerIcon} />
          <Text variant="xLarge" className={styles.title}>{strings.WebPartTitle}</Text>
        </Stack>

        {/* Status card */}
        <Stack className={styles.statusCard} tokens={{ childrenGap: 12 }}>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Icon iconName={icon} style={{ color, fontSize: 28 }} />
            <Stack>
              <Text variant="large"><strong>{userSummary.displayName}</strong></Text>
              <Text variant="medium" style={{ color }}>{userSummary.currentLevel}</Text>
            </Stack>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 24 }} wrap>
            <Stack className={styles.metricBox}>
              <Text variant="xxLarge" className={styles.metricValue}>{userSummary.monthlyPoints}</Text>
              <Text variant="small" className={styles.metricLabel}>{strings.PointsThisMonth}</Text>
            </Stack>
            <Stack className={styles.metricBox}>
              <Text variant="xxLarge" className={styles.metricValue}>{userSummary.totalPoints}</Text>
              <Text variant="small" className={styles.metricLabel}>{strings.TotalPoints}</Text>
            </Stack>
            {userSummary.globalRank !== undefined && (
              <Stack className={styles.metricBox}>
                <Text variant="xxLarge" className={styles.metricValue}>#{userSummary.globalRank}</Text>
                <Text variant="small" className={styles.metricLabel}>{strings.GlobalRank}</Text>
              </Stack>
            )}
            {userSummary.teamRank !== undefined && (
              <Stack className={styles.metricBox}>
                <Text variant="xxLarge" className={styles.metricValue}>#{userSummary.teamRank}</Text>
                <Text variant="small" className={styles.metricLabel}>{strings.TeamRank}</Text>
              </Stack>
            )}
          </Stack>

          {userSummary.lastActivityDate && (
            <Text variant="small" className={styles.note}>
              {formatString(strings.LastActivity, new Date(userSummary.lastActivityDate).toLocaleDateString())}
            </Text>
          )}
        </Stack>

        <Separator />

        {/* Preferences */}
        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant="mediumPlus"><Icon iconName="Settings" /> {strings.PreferencesTitle}</Text>
          <Toggle
            label={strings.NudgesToggleLabel}
            onText={strings.NudgesOn}
            offText={strings.NudgesOff}
            checked={userSummary.isEngagementNudgesEnabled}
            onChange={this._handleNudgeToggle}
            inlineLabel
          />
        </Stack>

        {/* Messages */}
        {errorMessage && <MessageBar messageBarType={MessageBarType.error}>{errorMessage}</MessageBar>}
        {successMessage && <MessageBar messageBarType={MessageBarType.success}>{successMessage}</MessageBar>}

        {/* Leave */}
        <Stack>
          <DefaultButton
            text={submitting ? strings.LeaveButtonLoading : strings.LeaveButton}
            iconProps={{ iconName: 'Leave' }}
            disabled={submitting}
            onClick={() => this.setState({ showLeaveDialog: true })}
            className={styles.leaveButton}
          />
        </Stack>

        {/* Leave confirmation dialog */}
        <Dialog
          hidden={!showLeaveDialog}
          onDismiss={() => this.setState({ showLeaveDialog: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: strings.LeaveDialogTitle,
            subText: strings.LeaveDialogMessage,
          }}
        >
          <DialogFooter>
            <PrimaryButton text={strings.LeaveDialogConfirm} onClick={this._handleLeave} />
            <DefaultButton text={strings.LeaveDialogCancel} onClick={() => this.setState({ showLeaveDialog: false })} />
          </DialogFooter>
        </Dialog>
      </Stack>
    );
  }

  // ─── Main render ───────────────────────────────────────────────────────────

  public render(): React.ReactElement<ICepOptinProps> {
    const { functionAppBaseUrl, hasTeamsContext } = this.props;
    const { loadingState, userSummary } = this.state;

    const rootClass = `${styles.cepOptin} ${hasTeamsContext ? styles.teams : ''}`;

    if (!functionAppBaseUrl) {
      return <div className={rootClass}>{this._renderNotConfigured()}</div>;
    }

    if (loadingState === 'idle' || loadingState === 'loading') {
      return <div className={rootClass}>{this._renderLoading()}</div>;
    }

    if (loadingState === 'error') {
      return (
        <div className={rootClass}>
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            {this.state.errorMessage}
            <DefaultButton text={strings.Retry} onClick={() => this._loadEnrollmentStatus()} style={{ marginLeft: 8 }} />
          </MessageBar>
        </div>
      );
    }

    return (
      <div className={rootClass}>
        {userSummary && userSummary.isActive ? this._renderEnrolledView() : this._renderEnrollmentForm()}
      </div>
    );
  }
}

