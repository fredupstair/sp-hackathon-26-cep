import "@typespec/http";
import "@typespec/openapi3";
import "@microsoft/typespec-m365-copilot"; 
import "./actions/cep.tsp";
import "./prompts/instructions.tsp";
import "./env.tsp";

using TypeSpec.M365.Copilot.Agents;

@agent(
  "CePAgent",
  "Gamify your Copilot usage by logging your wins with Microsoft 365 Copilot! Just tell me about a successful or impactful experience you had using Copilot in any M365 app, and I'll take care of the rest. Let's celebrate your Copilot wins together!"
)
@instructions(Prompts.INSTRUCTIONS)
@conversationStarter(#{
  title: "Mark this as a win",
  text: "Mark the prompt I just used as a Copilot win. Use the text I used as the description, remove any sensitive info, and deduce the app from the context."
})
@conversationStarter(#{
  title: "Log a shared Copilot win",
  text: "Use the prompt I just sent as the description of my win and log it as a shared Copilot win for the app I used."
})
@conversationStarter(#{
  title: "Log a Word win",
  text: "Copilot in Word helped me draft a proposal much faster than usual. Log this as a shared win."
})
@conversationStarter(#{
  title: "Log a PowerPoint win",
  text: "Copilot in PowerPoint helped me build a presentation from scratch in minutes. Log this as a shared win."
})
@conversationStarter(#{
  title: "Log a BizChat win",
  text: "I used BizChat to quickly answer a technical question from a colleague. Log this as a shared win."
})
namespace CePAgent {  
  // Uncomment this part to include custom actions in the agent
  // op searchIssues is global.GitHubAPI.searchIssues;

  // Log a Copilot WIN via the CEP backend
  op postWin is global.CePAPI.postWin;
}