const express = require("express");
const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication
} = require("botbuilder");

const { ClientSecretCredential } = require("@azure/identity");
const { AIProjectClient } = require("@azure/ai-projects");

const app = express();
app.use(express.json());

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppType: process.env.MicrosoftAppType || "SingleTenant",
  MicrosoftAppId: process.env.MicrosoftAppId || "",
  MicrosoftAppPassword: process.env.MicrosoftAppPassword || "",
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId || ""
});

const adapter = new CloudAdapter(botFrameworkAuthentication);

function limparEndpoint(endpoint) {
  return (endpoint || "").replace(/\/$/, "");
}

let openaiClientPromise = null;

async function obterOpenAIClientDoProjeto() {
  if (openaiClientPromise) {
    return openaiClientPromise;
  }

  const projectEndpoint = limparEndpoint(process.env.FOUNDRY_PROJECT_ENDPOINT);

  if (!projectEndpoint) {
    throw new Error("FOUNDRY_PROJECT_ENDPOINT não configurado.");
  }

  if (!process.env.AZURE_TENANT_ID) {
    throw new Error("AZURE_TENANT_ID não configurado.");
  }

  if (!process.env.AZURE_CLIENT_ID) {
    throw new Error("AZURE_CLIENT_ID não configurado.");
  }

  if (!process.env.AZURE_CLIENT_SECRET) {
    throw new Error("AZURE_CLIENT_SECRET não configurado.");
  }

  const credential = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID,
    process.env.AZURE_CLIENT_ID,
    process.env.AZURE_CLIENT_SECRET
  );

  const project = new AIProjectClient(projectEndpoint, credential);

  openaiClientPromise = project.getOpenAIClient();

  return openaiClientPromise;
}

async function chamarAgenteFoundry(mensagemCliente) {
  const agentName = process.env.FOUNDRY_AGENT_NAME || "bot-micks";
  const agentVersion = process.env.FOUNDRY_AGENT_VERSION || "2";

  const openai = await obterOpenAIClientDoProjeto();

  const agentReference = {
    name: agentName,
    type: "agent_reference"
  };

  if (agentVersion) {
    agentReference.version = agentVersion;
  }

  const response = await openai.responses.create({
    input: mensagemCliente,
    agent_reference: agentReference
  });

  if (response.output_text) {
    return response.output_text;
  }

  if (Array.isArray(response.output)) {
    const textos = [];

    for (const item of response.output) {
      if (Array.isArray(item.content)) {
        for (const content of item.content) {
          if (content.text) {
            textos.push(content.text);
          }

          if (content.type === "output_text" && content.text) {
            textos.push(content.text);
          }
        }
      }
    }

    if (textos.length > 0) {
      return textos.join("\n");
    }
  }

  console.log("Resposta completa do Foundry:", JSON.stringify(response, null, 2));

  return "Não consegui gerar uma resposta agora. Vou direcionar para um atendente.";
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
  console.log("FOUNDRY_AGENT_NAME:", process.env.FOUNDRY_AGENT_NAME || "não definido");
  console.log("FOUNDRY_AGENT_VERSION:", process.env.FOUNDRY_AGENT_VERSION || "não definido");

  console.log("AZURE_TENANT_ID configurado:", process.env.AZURE_TENANT_ID ? "sim" : "não");
  console.log("AZURE_CLIENT_ID configurado:", process.env.AZURE_CLIENT_ID ? "sim" : "não");
  console.log("AZURE_CLIENT_SECRET configurado:", process.env.AZURE_CLIENT_SECRET ? "sim" : "não");
});