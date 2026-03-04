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
      this.setState({ loadingState: 'error', errorMessage: `Errore nel caricamento: ${(e as Error).message}` });
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
      this.setState({ submitting: false, userSummary: summary, successMessage: 'Iscrizione completata! Benvenuto nel Copilot Engagement Program.' });
    } catch (e) {
      this.setState({ submitting: false, errorMessage: `Errore durante l'iscrizione: ${(e as Error).message}` });
    }
  };

  private _handleLeave = async (): Promise<void> => {
    const { apiClient } = this.props;
    if (!apiClient) return;
    this.setState({ submitting: true, showLeaveDialog: false, errorMessage: '', successMessage: '' });
    try {
      await apiClient.leave();
      this.setState({ submitting: false, userSummary: undefined, successMessage: 'Iscrizione annullata. I tuoi dati rimarranno disponibili per 90 giorni.' });
    } catch (e) {
      this.setState({ submitting: false, errorMessage: `Errore durante l'annullamento: ${(e as Error).message}` });
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
        <strong>Web part non configurata.</strong> Apri il pannello proprietà (icona a matita) e inserisci l&apos;URL della Function App.
      </MessageBar>
    );
  }

  private _renderLoading(): React.ReactElement {
    return (
      <Stack className={styles.container} horizontalAlign="center" tokens={{ padding: 24 }}>
        <Spinner size={SpinnerSize.large} label="Caricamento..." />
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
          <Text variant="xLarge" className={styles.title}>Copilot Engagement Program</Text>
        </Stack>
        <Text variant="medium" className={styles.subtitle}>
          Benvenuto/a, <strong>{userDisplayName}</strong>! Iscriviti al programma per tracciare il tuo utilizzo di Copilot, guadagnare punti e scalare la classifica.
        </Text>

        <Separator />

        {/* Transparency */}
        <Stack className={styles.transparencyBox} tokens={{ childrenGap: 8 }}>
          <Text variant="mediumPlus"><Icon iconName="Shield" /> Cosa raccogliamo</Text>
          <ul className={styles.transparencyList}>
            <li>Numero di prompt inviati a Copilot (nessun contenuto)</li>
            <li>App utilizzata (Word, Excel, Outlook, Teams, ecc.)</li>
            <li>Data giornaliera dell&apos;attività (aggregata, non per singola interazione)</li>
          </ul>
          <Text variant="small" className={styles.note}>
            Non vengono salvati: testi dei prompt, risposte, allegati, menzioni o qualsiasi contenuto personale.
            Il tracciamento avviene una volta al giorno. Puoi annullare l&apos;iscrizione in qualsiasi momento.
          </Text>
        </Stack>

        <Separator />

        {/* Form fields */}
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField
            label="Dipartimento"
            placeholder="es. IT, Marketing, Finance"
            value={department}
            onChange={(_e, v) => this.setState({ department: v || '' })}
          />
          <TextField
            label="Team"
            placeholder="es. Team Cloud, Team Vendite Nord"
            value={team}
            onChange={(_e, v) => this.setState({ team: v || '' })}
          />
          <Toggle
            label="Notifiche di engagement su Teams"
            onText="Attive"
            offText="Disattivate"
            checked={enableNudges}
            onChange={(_e, checked) => this.setState({ enableNudges: !!checked })}
            inlineLabel
          />
          <Text variant="small" className={styles.note}>
            Riceverai un messaggio su Teams se non usi Copilot per 3+ giorni consecutivi.
          </Text>
        </Stack>

        <Separator />

        {/* Consent */}
        <Checkbox
          label="Acconsento alla raccolta dei dati descritti sopra e confermo di voler partecipare al Copilot Engagement Program."
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
            text={submitting ? 'Iscrizione in corso...' : 'Iscrivimi al programma'}
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
          <Text variant="xLarge" className={styles.title}>Copilot Engagement Program</Text>
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
              <Text variant="small" className={styles.metricLabel}>Punti questo mese</Text>
            </Stack>
            <Stack className={styles.metricBox}>
              <Text variant="xxLarge" className={styles.metricValue}>{userSummary.totalPoints}</Text>
              <Text variant="small" className={styles.metricLabel}>Punti totali</Text>
            </Stack>
            {userSummary.globalRank !== undefined && (
              <Stack className={styles.metricBox}>
                <Text variant="xxLarge" className={styles.metricValue}>#{userSummary.globalRank}</Text>
                <Text variant="small" className={styles.metricLabel}>Classifica globale</Text>
              </Stack>
            )}
            {userSummary.teamRank !== undefined && (
              <Stack className={styles.metricBox}>
                <Text variant="xxLarge" className={styles.metricValue}>#{userSummary.teamRank}</Text>
                <Text variant="small" className={styles.metricLabel}>Classifica team</Text>
              </Stack>
            )}
          </Stack>

          {userSummary.lastActivityDate && (
            <Text variant="small" className={styles.note}>
              Ultima attività: {new Date(userSummary.lastActivityDate).toLocaleDateString('it-IT')}
            </Text>
          )}
        </Stack>

        <Separator />

        {/* Preferences */}
        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant="mediumPlus"><Icon iconName="Settings" /> Preferenze notifiche</Text>
          <Toggle
            label="Notifiche di engagement su Teams dopo 3+ giorni di inattività"
            onText="Attive"
            offText="Disattivate"
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
            text={submitting ? 'Elaborazione...' : 'Annulla iscrizione'}
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
            title: 'Annulla iscrizione',
            subText: 'Sei sicuro/a di voler uscire dal Copilot Engagement Program? Il tracciamento verrà interrotto e verrai rimosso/a dalla classifica. I tuoi dati verranno conservati per 90 giorni.',
          }}
        >
          <DialogFooter>
            <PrimaryButton text="Sì, annulla iscrizione" onClick={this._handleLeave} />
            <DefaultButton text="Annulla" onClick={() => this.setState({ showLeaveDialog: false })} />
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
            <DefaultButton text="Riprova" onClick={() => this._loadEnrollmentStatus()} style={{ marginLeft: 8 }} />
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

