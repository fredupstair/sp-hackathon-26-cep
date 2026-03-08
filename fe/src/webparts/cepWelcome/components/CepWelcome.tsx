import * as React from 'react';
import {
  Stack, Text, Icon,
  Spinner, SpinnerSize,
  MessageBar, MessageBarType,
  DefaultButton, PrimaryButton,
  Toggle,
  Dialog, DialogType, DialogFooter,
} from '@fluentui/react';
import { DisplayMode } from '@microsoft/sp-core-library';
import * as strings from 'CepWelcomeWebPartStrings';
import styles from './CepWelcome.module.scss';
import type { ICepWelcomeProps } from './ICepWelcomeProps';
import type {
  IUserSummary,
  IUserUsage,
  ICepBadge,
  ILeaderboardPage,
} from '../../../services/CepApiModels';
import { CopilotChatService } from '../../../services/CopilotChatService';
import { InlineWelcomeEditor } from './InlineWelcomeEditor';
import { WelcomeStep }     from './steps/WelcomeStep';
import { RulesStep }       from './steps/RulesStep';
import { PreferencesStep } from './steps/PreferencesStep';
import { ConsentStep }     from './steps/ConsentStep';
import { DashboardHeader } from './subcomponents/DashboardHeader';
import { StatsRow }        from './subcomponents/StatsRow';
import { AppUsageChart }   from './subcomponents/AppUsageChart';
import { BadgeList }       from './subcomponents/BadgeList';
import { MiniLeaderboard } from './subcomponents/MiniLeaderboard';

// ─── Types ────────────────────────────────────────────────────────────────────

type LoadingState = 'idle' | 'loading' | 'loaded' | 'error';
type WizardStep   = 'welcome' | 'rules' | 'preferences' | 'consent';
const WIZARD_STEPS: WizardStep[] = ['welcome', 'rules', 'preferences', 'consent'];

// ─── Month helpers ────────────────────────────────────────────────────────────

function p2(n: number): string { return n < 10 ? '0' + n : '' + n; }

function currentMonth(): string {
  const d = new Date();
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}`;
}

function getMonthRange(month: string): { from: string; to: string } {
  const [year, m] = month.split('-').map(Number);
  const lastDay = new Date(year, m, 0).getDate();
  return { from: `${month}-01`, to: `${month}-${p2(lastDay)}` };
}

function shiftMonth(month: string, delta: number): string {
  const [year, m] = month.split('-').map(Number);
  const d = new Date(year, m - 1 + delta, 1);
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}`;
}

// ─── State ────────────────────────────────────────────────────────────────────

interface ICepWelcomeState {
  loadingState:     LoadingState;
  userSummary:      IUserSummary | undefined;
  errorMessage:     string;
  successMessage:   string;
  // wizard
  currentStep:      WizardStep;
  department:       string;
  team:             string;
  enableNudges:     boolean;
  consentChecked:   boolean;
  submitting:       boolean;
  showLeaveDialog:  boolean;
  // AI personalised welcome
  aiWelcomeText:    string;
  aiWelcomeLoading: boolean;
  aiWelcomeStreaming: boolean;
  editingWelcome:   boolean;
  // dashboard data
  month:            string;
  usage:            IUserUsage | undefined;
  badges:           ICepBadge[];
  leaderboard:      ILeaderboardPage | undefined;
  dashboardLoading: boolean;
}

// ─── Component ────────────────────────────────────────────────────────────────

export default class CepWelcome extends React.Component<ICepWelcomeProps, ICepWelcomeState> {

  constructor(props: ICepWelcomeProps) {
    super(props);
    this.state = {
      loadingState:     'idle',
      userSummary:      undefined,
      errorMessage:     '',
      successMessage:   '',
      currentStep:      'welcome',
      department:       '',
      team:             '',
      enableNudges:     true,
      consentChecked:   false,
      submitting:       false,
      showLeaveDialog:  false,
      aiWelcomeText:    '',
      aiWelcomeLoading: false,
      aiWelcomeStreaming: false,
      editingWelcome:   props.displayMode === DisplayMode.Edit && !props.welcomeText?.trim(),
      month:            currentMonth(),
      usage:            undefined,
      badges:           [],
      leaderboard:      undefined,
      dashboardLoading: false,
    };
  }

  public componentDidMount(): void {
    this._loadEnrollmentStatus().catch(console.error);
    this._generatePersonalizedWelcome().catch(console.error);
  }

  public componentDidUpdate(prevProps: ICepWelcomeProps, prevState: ICepWelcomeState): void {
    if (prevProps.apiClient !== this.props.apiClient) {
      this._loadEnrollmentStatus().catch(console.error);
    }
    if (prevProps.graphClient !== this.props.graphClient && this.props.graphClient) {
      this._generatePersonalizedWelcome().catch(console.error);
    }
    // Reload dashboard data when month changes (enrolled only)
    if (prevState.month !== this.state.month && this.state.userSummary?.isActive) {
      this._loadDashboardData().catch(console.error);
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
      if (summary?.isActive) {
        await this._loadDashboardData();
      }
    } catch (e) {
      this.setState({
        loadingState: 'error',
        errorMessage: `Error loading enrollment status: ${(e as Error).message}`,
      });
    }
  }

  private async _loadDashboardData(): Promise<void> {
    const { apiClient } = this.props;
    if (!apiClient) return;

    this.setState({ dashboardLoading: true });
    const { month } = this.state;
    const { from, to } = getMonthRange(month);

    try {
      const [summary, usage, badges, leaderboard] = await Promise.all([
        apiClient.getMeSummary(month),
        apiClient.getMeUsage(from, to).catch((): IUserUsage => ({
          from, to, breakdown: [], totalPrompts: 0, totalPoints: 0,
        })),
        apiClient.getMeBadges().catch((): ICepBadge[] => []),
        apiClient.getLeaderboard('global', month, 1, 5).catch((): undefined => undefined),
      ]);
      this.setState({
        userSummary: summary ?? this.state.userSummary,
        usage,
        badges,
        leaderboard,
        dashboardLoading: false,
      });
    } catch (e) {
      this.setState({
        dashboardLoading: false,
        errorMessage: `Error loading dashboard: ${(e as Error).message}`,
      });
    }
  }

  private async _generatePersonalizedWelcome(): Promise<void> {
    const { graphClient, organizationName, userDisplayName, displayMode } = this.props;
    if (displayMode === DisplayMode.Edit) return;
    if (!graphClient || !userDisplayName) return;
    this.setState({ aiWelcomeLoading: true, aiWelcomeStreaming: false });
    try {
      let jobTitle = '';
      let department = '';
      try {
        const me = await graphClient
          .api('/me')
          .version('v1.0')
          .select('jobTitle,department')
          .get() as { jobTitle?: string; department?: string };
        jobTitle   = me.jobTitle   ?? '';
        department = me.department ?? '';
      } catch { /* ignore */ }

      const service = new CopilotChatService(graphClient);

      // Switch from loading dots to streaming cursor on first chunk
      const onChunk = (text: string): void => {
        this.setState({ aiWelcomeText: text, aiWelcomeLoading: false, aiWelcomeStreaming: true });
      };

      await service.streamPersonalizedText(
        userDisplayName,
        jobTitle,
        department,
        organizationName,
        onChunk,
      );

      // Stream complete
      this.setState({ aiWelcomeStreaming: false });
    } catch {
      this.setState({ aiWelcomeLoading: false, aiWelcomeStreaming: false });
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
      // Load dashboard data now that user is enrolled
      await this._loadDashboardData();
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
        usage:          undefined,
        badges:         [],
        leaderboard:    undefined,
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

  // ─── Month navigation ─────────────────────────────────────────────────────

  private _handlePrevMonth = (): void => {
    this.setState((s) => ({ month: shiftMonth(s.month, -1) }));
  };

  private _handleNextMonth = (): void => {
    this.setState((s) => {
      const next = shiftMonth(s.month, 1);
      return { month: next > currentMonth() ? currentMonth() : next };
    });
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
      aiWelcomeText, aiWelcomeLoading, aiWelcomeStreaming,
    } = this.state;
    const { userDisplayName, welcomeText } = this.props;

    switch (currentStep) {
      case 'welcome':
        return (
          <WelcomeStep
            welcomeText={welcomeText}
            userName={userDisplayName}
            aiWelcomeText={aiWelcomeText}
            aiWelcomeLoading={aiWelcomeLoading}
            aiWelcomeStreaming={aiWelcomeStreaming}
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

  // ─── Enrolled view (Dashboard + Settings) ─────────────────────────────────

  private _renderEnrolledView(): React.ReactElement {
    const {
      userSummary, month, usage, badges, leaderboard,
      submitting, errorMessage, successMessage, showLeaveDialog,
      dashboardLoading,
    } = this.state;
    if (!userSummary) return <></>;

    const isCurrentMonth = month === currentMonth();

    return (
      <Stack tokens={{ childrenGap: 0 }}>

        {/* ── Dashboard section ── */}
        {dashboardLoading ? (
          <div style={{ textAlign: 'center', padding: 32 }}>
            <Spinner size={SpinnerSize.small} label={strings.Loading} />
          </div>
        ) : (
          <>
            <DashboardHeader
              summary={userSummary}
              month={month}
              isCurrentMonth={isCurrentMonth}
              onPrevMonth={this._handlePrevMonth}
              onNextMonth={this._handleNextMonth}
            />
            <StatsRow summary={userSummary} />
            {usage && <AppUsageChart usage={usage} />}
            <BadgeList badges={badges} />
            {leaderboard && <MiniLeaderboard leaderboard={leaderboard} />}
          </>
        )}

        {/* ── Settings section ── */}
        <div className={styles.settingsSection}>
          <Stack tokens={{ childrenGap: 12 }}>
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
        </div>
      </Stack>
    );
  }

  // ─── Main render ───────────────────────────────────────────────────────────

  public render(): React.ReactElement<ICepWelcomeProps> {
    const { functionAppBaseUrl, hasTeamsContext, displayMode } = this.props;
    const { loadingState, userSummary, successMessage, editingWelcome } = this.state;

    const rootClass = `${styles.cepWelcome} ${hasTeamsContext ? styles.teams : ''}`;

    // ── Edit mode: inline AI welcome text editor ─────────────────────────────
    if (displayMode === DisplayMode.Edit && editingWelcome) {
      return <div className={rootClass}>{this._renderInlineEditor()}</div>;
    }

    // ── Determine main content ───────────────────────────────────────────────
    let mainContent: React.ReactElement;

    if (!functionAppBaseUrl) {
      mainContent = (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          <strong>{strings.NotConfiguredTitle}</strong>{' '}
          SharePoint Tenant Properties have not been set.
          Run <code>deploy/set-tenant-properties.ps1</code> to configure the backend URL.
        </MessageBar>
      );
    } else if (loadingState === 'idle' || loadingState === 'loading') {
      mainContent = (
        <Stack horizontalAlign="center" tokens={{ padding: 32 }}>
          <Spinner size={SpinnerSize.large} label="Loading…" />
        </Stack>
      );
    } else if (loadingState === 'error') {
      mainContent = (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {this.state.errorMessage}
          <DefaultButton
            text="Retry"
            onClick={() => this._loadEnrollmentStatus().catch(console.error)}
            style={{ marginLeft: 8 }}
          />
        </MessageBar>
      );
    } else if (userSummary?.isActive) {
      mainContent = this._renderEnrolledView();
    } else {
      mainContent = (
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
      );
    }

    // ── In edit mode, wrap with admin bar ────────────────────────────────────
    if (displayMode === DisplayMode.Edit) {
      return (
        <div className={rootClass}>
          {this._renderAdminBar()}
          {mainContent}
        </div>
      );
    }

    return <div className={rootClass}>{mainContent}</div>;
  }

  // ─── Inline AI editor (edit mode, full webpart) ────────────────────────────

  private _renderInlineEditor(): React.ReactElement {
    const { graphClient, organizationName, welcomeText, onWelcomeTextSave } = this.props;
    return (
      <InlineWelcomeEditor
        graphClient={graphClient}
        organizationName={organizationName}
        welcomeText={welcomeText}
        hasExistingText={!!welcomeText?.trim()}
        onSave={(text, orgName) => {
          onWelcomeTextSave(text, orgName);
          this.setState({ editingWelcome: false });
        }}
        onDiscard={() => this.setState({ editingWelcome: false })}
      />
    );
  }

  // ─── Admin bar (edit mode, editor closed) ─────────────────────────────────

  private _renderAdminBar(): React.ReactElement {
    return (
      <Stack
        horizontal
        verticalAlign="center"
        horizontalAlign="space-between"
        className={styles.adminBar}
      >
        <Text className={styles.adminBarLabel}>
          <Icon iconName="Edit" style={{ marginRight: 4, fontSize: 11 }} />
          {strings.EditModeNotice}
        </Text>
        <DefaultButton
          text={`✨ ${strings.InlineEditorEditButton}`}
          onClick={() => this.setState({ editingWelcome: true })}
          styles={{ root: { fontSize: 12, height: 28, padding: '0 12px', minWidth: 'auto' } }}
        />
      </Stack>
    );
  }
}
