const express = require("express");
const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication
} = require("botbuilder");

const { ClientSecretCredential } = require("@azure/identity");
const { AgentsClient } = require("@azure/ai-agents");

const app = express();
app.use(express.json());

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppType: process.env.MicrosoftAppType || "SingleTenant",
  MicrosoftAppId: process.env.MicrosoftAppId || "",
  MicrosoftAppPassword: process.env.MicrosoftAppPassword || "",
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId || ""
});

const adapter = new CloudAdapter(botFrameworkAuthentication);

const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID,
  process.env.AZURE_CLIENT_ID,
  process.env.AZURE_CLIENT_SECRET
);

const foundryClient = new AgentsClient(
  process.env.FOUNDRY_PROJECT_ENDPOINT,
  credential
);

async function chamarAgenteFoundry(mensagemCliente) {
  const agentId = process.env.FOUNDRY_AGENT_ID;

  if (!agentId) {
    throw new Error("FOUNDRY_AGENT_ID não configurado.");
  }

  if (!process.env.FOUNDRY_PROJECT_ENDPOINT) {
    throw new Error("FOUNDRY_PROJECT_ENDPOINT não configurado.");
  }

  const thread = await foundryClient.threads.create();

  await foundryClient.messages.create(thread.id, "user", mensagemCliente);

  let run = await foundryClient.runs.create(thread.id, agentId);

  const inicio = Date.now();
  const timeoutMs = 30000;

  while (
    run.status === "queued" ||
    run.status === "in_progress" ||
    run.status === "requires_action"
  ) {
    if (Date.now() - inicio > timeoutMs) {
      throw new Error("Timeout aguardando resposta do Foundry Agent.");
    }

    await new Promise((resolve) => setTimeout(resolve, 1000));

    run = await foundryClient.runs.get(thread.id, run.id);
  }

  if (run.status !== "completed") {
    console.error("Run não completou:", run);
    throw new Error(`Run do Foundry terminou com status: ${run.status}`);
  }

  const mensagens = await foundryClient.messages.list(thread.id);

  const listaMensagens = [];

  for await (const item of mensagens) {
    listaMensagens.push(item);
  }

  const respostaAssistente = listaMensagens.find(
    (msg) => msg.role === "assistant"
  );

  if (!respostaAssistente || !respostaAssistente.content) {
    throw new Error("Nenhuma resposta do assistente foi encontrada.");
  }

  const parteTexto = respostaAssistente.content.find(
    (contentItem) => contentItem.type === "text"
  );

  const textoResposta =
    parteTexto?.text?.value ||
    parteTexto?.text ||
    "Não consegui gerar uma resposta agora. Vou direcionar para um atendente.";

  return textoResposta;
}

adapter.onTurnError = async (context, error) => {
  console.error("Erro no bot:", error);

  try {
    await context.sendActivity(
      "Tive um problema ao processar sua mensagem. Vou direcionar para um atendente."
    );
  } catch (sendError) {
    console.error("Erro ao enviar mensagem de erro ao usuário:", sendError);
  }
};

app.get("/", (req, res) => {
  res.status(200).send("Micks Bot API online.");
});

app.post("/api/messages", async (req, res) => {
  try {
    await adapter.process(req, res, async (context) => {
      if (context.activity.type === "message") {
        const textoCliente = context.activity.text || "";

        console.log("Mensagem recebida:", textoCliente);

        try {
          const respostaFoundry = await chamarAgenteFoundry(textoCliente);

          console.log("Resposta Foundry:", respostaFoundry);

          await context.sendActivity(respostaFoundry);
        } catch (foundryError) {
          console.error("Erro ao chamar Foundry:", foundryError);

          await context.sendActivity(
            "No momento não consegui consultar meu assistente interno. Vou direcionar seu atendimento para um atendente."
          );
        }
      } else {
        console.log("Activity recebida:", context.activity.type);
      }
    });
  } catch (error) {
    console.error("Erro ao processar activity:", error);

    if (!res.headersSent) {
      res.status(500).send("Erro ao processar mensagem.");
    }
  }
});

const port = process.env.PORT || 3978;

app.listen(port, () => {
  console.log(`Micks Bot API rodando na porta ${port}`);

  console.log("MicrosoftAppType:", process.env.MicrosoftAppType || "não definido");
  console.log("MicrosoftAppId configurado:", process.env.MicrosoftAppId ? "sim" : "não");
  console.log("MicrosoftAppPassword configurado:", process.env.MicrosoftAppPassword ? "sim" : "não");
  console.log("MicrosoftAppTenantId configurado:", process.env.MicrosoftAppTenantId ? "sim" : "não");

  console.log("FOUNDRY_PROJECT_ENDPOINT configurado:", process.env.FOUNDRY_PROJECT_ENDPOINT ? "sim" : "não");
  console.log("FOUNDRY_AGENT_ID configurado:", process.env.FOUNDRY_AGENT_ID ? "sim" : "não");
  console.log("AZURE_CLIENT_ID configurado:", process.env.AZURE_CLIENT_ID ? "sim" : "não");
  console.log("AZURE_CLIENT_SECRET configurado:", process.env.AZURE_CLIENT_SECRET ? "sim" : "não");
  console.log("AZURE_TENANT_ID configurado:", process.env.AZURE_TENANT_ID ? "sim" : "não");
});