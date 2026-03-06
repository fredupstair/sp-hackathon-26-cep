import * as React from 'react';
import {
  Stack, Text, Icon,
  Spinner, SpinnerSize,
  MessageBar, MessageBarType,
  DefaultButton, PrimaryButton,
  Toggle, Separator,
  Dialog, DialogType, DialogFooter,
} from '@fluentui/react';
import { DisplayMode } from '@microsoft/sp-core-library';
import * as strings from 'CepOptinWebPartStrings';
import styles from './CepOptin.module.scss';
import type { ICepOptinProps } from './ICepOptinProps';
import type { IUserSummary } from '../../../services/CepApiModels';
import { WelcomeStep }     from './steps/WelcomeStep';
import { RulesStep }       from './steps/RulesStep';
import { PreferencesStep } from './steps/PreferencesStep';
import { ConsentStep }     from './steps/ConsentStep';

// ─── Types ────────────────────────────────────────────────────────────────────

type LoadingState = 'idle' | 'loading' | 'loaded' | 'error';
type WizardStep  = 'welcome' | 'rules' | 'preferences' | 'consent';
const WIZARD_STEPS: WizardStep[] = ['welcome', 'rules', 'preferences', 'consent'];

interface ICepOptinState {
  loadingState:    LoadingState;
  userSummary:     IUserSummary | undefined; // undefined = not enrolled
  errorMessage:    string;
  successMessage:  string;
  // wizard
  currentStep:     WizardStep;
  // form fields
  department:      string;
  team:            string;
  enableNudges:    boolean;
  consentChecked:  boolean;
  // action state
  submitting:      boolean;
  showLeaveDialog: boolean;
}

// ─── Helper ───────────────────────────────────────────────────────────────────

interface IMetricBoxProps { value: string | number; label: string }
const MetricBox: React.FC<IMetricBoxProps> = ({ value, label }) => (
  <Stack style={{ minWidth: 90, alignItems: 'center', textAlign: 'center' }}>
    <Text variant="xxLarge" className={styles.metricValue}>{value}</Text>
    <Text variant="small"   className={styles.metricLabel}>{label}</Text>
  </Stack>
);

// ─── Component ────────────────────────────────────────────────────────────────

export default class CepOptin extends React.Component<ICepOptinProps, ICepOptinState> {

  constructor(props: ICepOptinProps) {
    super(props);
    this.state = {
      loadingState:    'idle',
      userSummary:     undefined,
      errorMessage:    '',
      successMessage:  '',
      currentStep:     'welcome',
      department:      '',
      team:            '',
      enableNudges:    true,
      consentChecked:  false,
      submitting:      false,
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
      this.setState({
        loadingState: 'error',
        errorMessage: `Error loading enrollment status: ${(e as Error).message}`,
      });
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
      this.setState({
        submitting:     false,
        userSummary:    summary,
        successMessage: strings.JoinSuccess,
      });
    } catch (e) {
      this.setState({
        submitting:   false,
        errorMessage: strings.JoinError.replace('{0}', (e as Error).message),
      });
    }
  };

  private _handleLeave = async (): Promise<void> => {
    const { apiClient } = this.props;
    if (!apiClient) return;
    this.setState({ submitting: true, showLeaveDialog: false, errorMessage: '', successMessage: '' });
    try {
      await apiClient.leave();
      this.setState({
        submitting:     false,
        userSummary:    undefined,
        currentStep:    'welcome',
        successMessage: strings.LeaveSuccess,
        consentChecked: false,
      });
    } catch (e) {
      this.setState({
        submitting:   false,
        errorMessage: strings.LeaveError.replace('{0}', (e as Error).message),
      });
    }
  };

  private _handleNudgeToggle = (_ev: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    const { userSummary } = this.state;
    if (!userSummary) return;
    this.setState({ userSummary: { ...userSummary, isEngagementNudgesEnabled: !!checked } });
  };

  // ─── Wizard navigation ─────────────────────────────────────────────────────

  private _goToStep = (step: WizardStep): void => {
    this.setState({ currentStep: step, errorMessage: '' });
  };

  private _goNext = (): void => {
    const idx = WIZARD_STEPS.indexOf(this.state.currentStep);
    if (idx < WIZARD_STEPS.length - 1) this._goToStep(WIZARD_STEPS[idx + 1]);
  };

  private _goBack = (): void => {
    const idx = WIZARD_STEPS.indexOf(this.state.currentStep);
    if (idx > 0) this._goToStep(WIZARD_STEPS[idx - 1]);
  };

  // ─── Step indicator ────────────────────────────────────────────────────────

  private _renderStepIndicator(): React.ReactElement {
    const currentIdx = WIZARD_STEPS.indexOf(this.state.currentStep);
    return (
      <Stack
        horizontal
        horizontalAlign="center"
        tokens={{ childrenGap: 8 }}
        className={styles.stepIndicator}
      >
        {WIZARD_STEPS.map((step, idx) => (
          <div
            key={step}
            className={[
              styles.stepDot,
              idx === currentIdx ? styles.stepDotActive : '',
              idx < currentIdx   ? styles.stepDotDone   : '',
            ].filter(Boolean).join(' ')}
          />
        ))}
      </Stack>
    );
  }

  // ─── Wizard ────────────────────────────────────────────────────────────────

  private _renderWizard(): React.ReactElement {
    const {
      currentStep, department, team,
      enableNudges, consentChecked, submitting, errorMessage,
    } = this.state;
    const { userDisplayName, welcomeText } = this.props;

    switch (currentStep) {
      case 'welcome':
        return (
          <WelcomeStep
            welcomeText={welcomeText}
            userName={userDisplayName}
            onNext={this._goNext}
          />
        );
      case 'rules':
        return <RulesStep onNext={this._goNext} onBack={this._goBack} />;
      case 'preferences':
        return (
          <PreferencesStep
            department={department}
            team={team}
            enableNudges={enableNudges}
            onDepartmentChange={v => this.setState({ department: v })}
            onTeamChange={v => this.setState({ team: v })}
            onNudgesChange={v => this.setState({ enableNudges: v })}
            onNext={this._goNext}
            onBack={this._goBack}
          />
        );
      case 'consent':
        return (
          <ConsentStep
            department={department}
            team={team}
            enableNudges={enableNudges}
            consentChecked={consentChecked}
            submitting={submitting}
            errorMessage={errorMessage}
            onConsentChange={checked => this.setState({ consentChecked: checked })}
            onJoin={this._handleJoin}
            onBack={this._goBack}
          />
        );
      default:
        return <></>;
    }
  }

  // ─── Enrolled view ─────────────────────────────────────────────────────────

  private _renderEnrolledView(): React.ReactElement {
    const { userSummary, submitting, errorMessage, successMessage, showLeaveDialog } = this.state;
    if (!userSummary) return <></>;

    const levelColor: Record<string, string> = {
      Bronze: '#CD7F32', Silver: '#A8A9AD', Gold: '#FFD700',
    };
    const levelIcon: Record<string, string> = {
      Bronze: 'Medal', Silver: 'Medal', Gold: 'FavoriteStarFill',
    };
    const color = levelColor[userSummary.currentLevel] || '#CD7F32';
    const icon  = levelIcon[userSummary.currentLevel]  || 'Medal';

    return (
      <Stack className={styles.container} tokens={{ childrenGap: 16 }}>

        {/* Header */}
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
          <Icon iconName="Trophy2" className={styles.headerIcon} />
          <Text variant="xLarge" className={styles.title}>{strings.ProgramTitle}</Text>
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
            <MetricBox value={userSummary.monthlyPoints}  label={strings.PointsThisMonth} />
            <MetricBox value={userSummary.totalPoints}    label={strings.TotalPoints} />
            {userSummary.globalRank !== undefined && (
              <MetricBox value={`#${userSummary.globalRank}`} label={strings.GlobalRank} />
            )}
            {userSummary.teamRank !== undefined && (
              <MetricBox value={`#${userSummary.teamRank}`} label={strings.TeamRank} />
            )}
          </Stack>

          {userSummary.lastActivityDate && (
            <Text variant="small" className={styles.note}>
              {`${strings.LastActivity} ${new Date(userSummary.lastActivityDate).toLocaleDateString()}`}
            </Text>
          )}
        </Stack>

        <Separator />

        {/* Preferences */}
        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant="mediumPlus">
            <Icon iconName="Settings" /> {strings.PreferencesTitle}
          </Text>
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
        {errorMessage   && <MessageBar messageBarType={MessageBarType.error  }>{errorMessage  }</MessageBar>}
        {successMessage && <MessageBar messageBarType={MessageBarType.success}>{successMessage}</MessageBar>}

        {/* Leave */}
        <DefaultButton
          text={submitting ? strings.LeaveButtonLoading : strings.LeaveButton}
          iconProps={{ iconName: 'Leave' }}
          disabled={submitting}
          onClick={() => this.setState({ showLeaveDialog: true })}
          className={styles.leaveButton}
        />

        {/* Leave confirmation dialog */}
        <Dialog
          hidden={!showLeaveDialog}
          onDismiss={() => this.setState({ showLeaveDialog: false })}
          dialogContentProps={{
            type:    DialogType.normal,
            title:   strings.LeaveDialogTitle,
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
    const { functionAppBaseUrl, hasTeamsContext, welcomeText, displayMode } = this.props;
    const { loadingState, userSummary, successMessage } = this.state;

    const rootClass = `${styles.cepOptin} ${hasTeamsContext ? styles.teams : ''}`;

    // Editor setup card — shown when welcomeText has not been configured yet
    if (!welcomeText?.trim() && displayMode === DisplayMode.Edit) {
      return <div className={rootClass}>{this._renderSetupPrompt()}</div>;
    }

    if (!functionAppBaseUrl) {
      return (
        <div className={rootClass}>
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            <strong>{strings.NotConfiguredTitle}</strong>{' '}
            SharePoint Tenant Properties have not been set.
            Run <code>deploy/set-tenant-properties.ps1</code> to configure the backend URL.
          </MessageBar>
        </div>
      );
    }

    if (loadingState === 'idle' || loadingState === 'loading') {
      return (
        <div className={rootClass}>
          <Stack horizontalAlign="center" tokens={{ padding: 32 }}>
            <Spinner size={SpinnerSize.large} label="Loading…" />
          </Stack>
        </div>
      );
    }

    if (loadingState === 'error') {
      return (
        <div className={rootClass}>
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            {this.state.errorMessage}
            <DefaultButton
              text="Retry"
              onClick={() => this._loadEnrollmentStatus().catch(console.error)}
              style={{ marginLeft: 8 }}
            />
          </MessageBar>
        </div>
      );
    }

    // Enrolled and active → show profile / dashboard card
    if (userSummary?.isActive) {
      return <div className={rootClass}>{this._renderEnrolledView()}</div>;
    }

    // Not enrolled → multi-step enrollment wizard
    return (
      <div className={rootClass}>
        <Stack className={styles.wizardContainer}>
          {this._renderStepIndicator()}
          {successMessage && (
            <MessageBar
              messageBarType={MessageBarType.success}
              onDismiss={() => this.setState({ successMessage: '' })}
            >
              {successMessage}
            </MessageBar>
          )}
          {this._renderWizard()}
        </Stack>
      </div>
    );
  }

  // ─── Setup prompt (edit mode, no welcome text configured) ─────────────────

  private _renderSetupPrompt(): React.ReactElement {
    return (
      <Stack
        horizontalAlign="center"
        tokens={{ childrenGap: 16 }}
        style={{ padding: '40px 24px', textAlign: 'center' }}
      >
        <Icon iconName="Lightbulb" style={{ fontSize: 48, color: 'var(--themePrimary, #0078d4)' }} />
        <Text variant="xLarge" style={{ fontWeight: 600 }}>
          {strings.SetupTitle}
        </Text>
        <Text variant="medium" style={{ maxWidth: 480, color: 'var(--bodySubtext)' }}>
          {strings.SetupDescription}
        </Text>
        <PrimaryButton
          text={strings.SetupButton}
          iconProps={{ iconName: 'Settings' }}
          onClick={this.props.onConfigureClick}
        />
      </Stack>
    );
  }
}

