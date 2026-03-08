define([], function() {
  return {
    // Property Pane
    "PropertyPaneDescription": "Configura la connessione al backend CEP",
    "FunctionAppGroupName": "Backend Azure Functions",
    "FunctionAppBaseUrlLabel": "Function App Base URL",
    "FunctionAppBaseUrlDescription": "URL base della Function App senza trailing slash. Es: https://cep-functions.azurewebsites.net",

    // General
    "Loading": "Caricamento\u2026",
    "Retry": "Riprova",
    "BackButton": "Indietro",
    "ContinueButton": "Continua",

    // Not-configured banner
    "NotConfiguredTitle": "Web part non configurata.",
    "NotConfiguredMessage": "Le Tenant Properties di SharePoint non sono state impostate. Esegui deploy/set-tenant-properties.ps1 per configurare l'URL del backend.",

    // Setup prompt (edit mode, empty welcome text)
    "SetupTitle": "Configura il tuo Copilot Engagement Program",
    "SetupDescription": "Imposta il testo di benvenuto per il wizard di iscrizione. Usa il generatore AI nel pannello impostazioni per creare un messaggio personalizzato in pochi secondi.",
    "SetupButton": "Configura il testo di benvenuto",

    // WelcomeStep
    "ProgramTitle": "Copilot Engagement Program",
    "HelloGreeting": "Ciao, {0}! \uD83D\uDC4B",
    "FeatureTrack": "Traccia il tuo utilizzo di Copilot in tutte le app Microsoft 365",
    "FeaturePoints": "Guadagna 1 punto per ogni prompt inviato",
    "FeatureLevels": "Sali di livello: Explorer \u2192 Practitioner \u2192 Master",
    "FeatureLeaderboard": "Compete nelle classifiche globali e di team",
    "GetStarted": "Inizia",

    // RulesStep
    "HowItWorksTitle": "Come funziona",
    "RuleUseTitle": "Usa Copilot",
    "RuleUseBody": "Utilizza Microsoft 365 Copilot in qualsiasi app \u2014 Word, Excel, Outlook, Teams e altro \u2014 come parte del tuo lavoro quotidiano.",
    "RuleEarnTitle": "Guadagna punti",
    "RuleEarnBody": "1 punto per ogni prompt inviato. I punti si accumulano nel mese e si azzerano a fine mese per una nuova competizione.",
    "RuleLevelTitle": "Sali di livello",
    "RuleLevelBody": "Raggiungi i livelli Explorer, Practitioner e Master man mano che il tuo utilizzo cresce. I badge sono permanenti \u2014 li guadagni una volta, li tieni per sempre.",
    "PrivacyTitle": "Privacy e Trasparenza",
    "PrivacyItem1": "Numero di prompt inviati a Copilot (il contenuto non viene mai letto)",
    "PrivacyItem2": "App utilizzata (Word, Excel, Outlook, Teams\u2026)",
    "PrivacyItem3": "Data giornaliera dell\u2019attivit\u00e0 (aggregata, non per singola interazione)",
    "PrivacyItem4": "Testo dei prompt, risposte e allegati \u2014 mai conservati",
    "PrivacyItem5": "Contenuto personale o dati sensibili \u2014 mai raccolti",
    "PrivacyNote": "Il tracciamento avviene una volta al giorno in forma aggregata. Puoi ritirarsi in qualsiasi momento.",

    // PreferencesStep
    "ProfileTitle": "Il tuo profilo",
    "ProfileSubtitle": "Questi dettagli facoltativi personalizzano la tua vista in classifica. Puoi aggiornarli in qualsiasi momento.",
    "DepartmentLabel": "Dipartimento",
    "DepartmentPlaceholder": "es. IT, Marketing, Finance",
    "TeamLabel": "Team",
    "TeamPlaceholder": "es. Team Cloud, Team Vendite Nord",
    "NotificationsTitle": "Notifiche",
    "NudgesLabel": "Notifiche di engagement su Teams",
    "NudgesOn": "Attive",
    "NudgesOff": "Disattivate",
    "NudgesNote": "Ricevi un messaggio su Teams quando non usi Copilot per 3 o pi\u00f9 giorni consecutivi \u2014 un promemoria amichevole per restare in gioco!",

    // ConsentStep
    "AlmostInTitle": "Ci siamo quasi!",
    "AlmostInSubtitle": "Controlla i tuoi dati prima di iscriverti.",
    "EnrollmentSummaryTitle": "Riepilogo iscrizione",
    "NotificationsEnabled": "\uD83D\uDD14 Attive",
    "NotificationsDisabled": "\uD83D\uDD15 Disattivate",
    "DataCollectionTitle": "Iscrivendoti, potremmo raccogliere:",
    "DataItem1": "Numero di prompt inviati a Copilot (non il contenuto)",
    "DataItem2": "Le app Microsoft 365 in cui hai usato Copilot",
    "DataItem3": "Timestamp giornalieri dell\u2019attivit\u00e0 (aggregati, non per interazione)",
    "DataNote": "Il contenuto dei prompt e le risposte non vengono mai conservati o letti. Puoi uscire dal programma in qualsiasi momento.",
    "ConsentLabel": "Acconsento alla raccolta dei dati descritti sopra e confermo di voler partecipare al Copilot Engagement Program.",
    "JoinButton": "Iscriviti al programma",
    "JoinButtonLoading": "Iscrizione\u2026",

    // Actions / outcomes
    "JoinSuccess": "Iscrizione completata! Benvenuto nel Copilot Engagement Program. \uD83C\uDF89",
    "LeaveButton": "Annulla iscrizione",
    "LeaveButtonLoading": "Elaborazione\u2026",
    "LeaveSuccess": "Iscrizione annullata. I tuoi dati rimarranno disponibili per 90 giorni.",

    // Errors
    "JoinError": "Errore durante l\u2019iscrizione: {0}",
    "LeaveError": "Errore durante l\u2019annullamento: {0}",

    // Enrolled view
    "PointsThisMonth": "Punti questo mese",
    "TotalPoints": "Punti totali",
    "GlobalRank": "Classifica globale",
    "TeamRank": "Classifica team",
    "LastActivity": "Ultima attivit\u00e0: {0}",
    "PreferencesTitle": "Preferenze notifiche",
    "NudgesToggleLabel": "Notifiche di engagement su Teams dopo 3+ giorni di inattivit\u00e0",

    // Leave dialog
    "LeaveDialogTitle": "Annulla iscrizione",
    "LeaveDialogMessage": "Sei sicuro/a di voler uscire dal Copilot Engagement Program? Il tracciamento verr\u00e0 interrotto e verrai rimosso/a dalla classifica. I tuoi dati verranno conservati per 90 giorni.",
    "LeaveDialogConfirm": "S\u00ec, annulla iscrizione",
    "LeaveDialogCancel": "Annulla",

    // WelcomeTextEditor (property pane)
    "GeneratorTitle": "\u2728 Generatore AI testo di benvenuto",
    "OrgNameLabel": "Nome organizzazione",
    "OrgNameDescription": "Usato per personalizzare il testo generato",
    "OrgNamePlaceholder": "es. Contoso, Fabrikam",
    "OrgNameRequired": "Inserisci prima il nome dell\u2019organizzazione.",
    "ToneLabel": "Tono (facoltativo)",
    "GenerateButton": "Genera con Copilot",
    "GeneratingButton": "Generazione\u2026",
    "GenerationFailed": "Generazione fallita: {0}",
    "FallbackWarning": "L\u2019API Copilot Chat di M365 non \u00e8 disponibile in questo tenant \u2014 \u00e8 stato usato un template predefinito. Puoi modificare il testo liberamente.",
    "WelcomeTextLabel": "Testo di benvenuto",
    "WelcomeTextDescription": "Mostrato agli utenti nella prima schermata del wizard di iscrizione. Modifica liberamente. Usa **parola** per il grassetto.",
    "WelcomeTextPlaceholder": "Descrivi il programma\u2026 es. \u2018Unisciti alla nostra avventura Copilot e inizia a guadagnare punti oggi!\u2019",
    "ClearWelcomeText": "Cancella testo di benvenuto",

    // AI personalised welcome (end-user)
    "AiWelcomeGenerating": "Personalizzazione in corso\u2026",
    "AiWelcomeBadge": "Personalizzato da AI",

    // Inline welcome editor (edit-mode, in-webpart)
    "InlineEditorTitle": "\u2728 Configura il testo di benvenuto",
    "InlineEditorSave": "Salva",
    "InlineEditorDiscard": "Annulla",
    "InlineEditorInputPlaceholder": "Nome dell\u2019organizzazione\u2026",
    "InlineEditorEditButton": "Modifica testo di benvenuto",
    "EditModeNotice": "Modalit\u00e0 modifica \u2014 anteprima testo di benvenuto",

    // ── Dashboard strings ──────────────────────────────────────────────────

    // Header
    "WebPartTitle": "La mia Dashboard CEP",
    "ViewingProfileOf": "Profilo di {0}",
    "BackToMyProfile": "Torna al mio profilo",

    // Loading / errors (dashboard-specific)
    "LoadError": "Errore nel caricamento dei dati: {0}",
    "NotEnrolledTitle": "Non sei iscritto",
    "NotEnrolledMessage": "Iscriviti al Copilot Engagement Program per vedere la tua dashboard.",
    "JoinProgram": "Iscriviti ora",

    // Stats
    "Level": "Livello",
    "NextLevel": "Prossimo livello",
    "PointsToNextLevel": "{0} punti a {1}",

    // App usage breakdown
    "UsageBreakdownTitle": "Utilizzo Copilot per app",
    "PromptCount": "{0} prompt",
    "NoUsage": "Nessuna attivit\u00e0",
    "FilterWeek": "Questa settimana",
    "FilterMonth": "Questo mese",
    "GetInspired": "Trova ispirazione su Prompt Gallery",

    // Badges
    "BadgesTitle": "Badge",
    "NoBadgesYet": "Nessun badge ancora. Continua ad usare Copilot!",
    "BadgeEarnedOn": "Ottenuto il {0}",
    "LockedBadge": "Bloccato \u2013 {0}",

    // Other user view (aggregated)
    "OtherUserPointsThisMonth": "Punti mensili",
    "OtherUserTotalPoints": "Punti totali",
    "OtherUserLevel": "Livello",
    "OtherUserBadges": "Badge"
  };
});
