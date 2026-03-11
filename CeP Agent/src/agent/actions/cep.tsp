import "@typespec/http";
import "@microsoft/typespec-m365-copilot";
import "../env.tsp";

using TypeSpec.Http;
using TypeSpec.M365.Copilot.Actions;

/**
 * Request body for logging a Copilot WIN.
 */
model WinRequest {
  /** Key of the M365 app used (e.g. "BizChat", "Word", "Outlook"). */
  appKey: string;

  /** Optional free-text note describing the win. */
  note?: string;

  /** Whether the win should be visible to the rest of the team. */
  isShared: boolean;
}

/**
 * Response returned after a WIN is logged.
 */
model WinResponse {
  /** Points added by this win. */
  pointsAdded: int32;

  /** User's total monthly points after this win. */
  totalMonthlyPoints: int32;

  /** Number of wins logged by the user today (after this one). */
  todayWinCount: int32;
}

@service
@server(CePAPI.SERVER_URL)
@useAuth(OAuth2Auth<[{
  type: OAuth2FlowType.authorizationCode,
  authorizationUrl: "https://login.microsoftonline.com/69d000be-840d-4db5-be28-0965a4fd72fc/oauth2/v2.0/authorize",
  tokenUrl: "https://login.microsoftonline.com/69d000be-840d-4db5-be28-0965a4fd72fc/oauth2/v2.0/token",
  scopes: ["api://0cb84638-30db-4bbd-936f-54a599840aec/user_impersonation"]
}]>)
@actions(CePAPI.ACTIONS_METADATA)
namespace CePAPI {
  /**
   * Metadata for the Copilot Engagement Program API actions.
   */
  const ACTIONS_METADATA = #{
    nameForHuman: "Copilot Engagement Program",
    descriptionForHuman: "Log a Copilot WIN and track your M365 Copilot engagement.",
    descriptionForModel: "Use this action to record a Copilot WIN on behalf of the calling user. Call it when the user describes a successful or impactful use of Microsoft 365 Copilot in any app. The appKey must match one of the supported values: BizChat, Word, Excel, PowerPoint, Outlook, Teams, OneNote, Loop, WebChat, M365App, Forms, SharePoint, Whiteboard.",
    legalInfoUrl: "https://learn.microsoft.com/legal/microsoft-365/microsoft-365-terms-of-service",
    privacyPolicyUrl: "https://privacy.microsoft.com/en-us/privacystatement"
  };

  /**
   * The base URL for the CEP Azure Functions API.
   * Set CEP_API_URL in env/.env.dev (and env/.env.local for local dev).
   */
  const SERVER_URL = Environment.CEP_API_URL;

  /**
   * Log a Copilot WIN for the calling user.
   *
   * @param body The win details: which M365 app was used, an optional note, and whether to share.
   * @returns Points earned and updated totals.
   */
  @route("/me/wins")
  @post op postWin(@body body: WinRequest): WinResponse;
}
