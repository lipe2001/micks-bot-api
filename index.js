const express = require("express");
const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication
} = require("botbuilder");

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

async function chamarAgenteFoundry(mensagemCliente) {
  const endpoint = limparEndpoint(process.env.AZURE_OPENAI_ENDPOINT);
  const apiKey = process.env.AZURE_OPENAI_API_KEY;
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT || "gpt-4.1-mini";
  const agentName = process.env.FOUNDRY_AGENT_NAME || "bot-micks";
  const agentVersion = process.env.FOUNDRY_AGENT_VERSION || "2";

  if (!endpoint) {
    throw new Error("AZURE_OPENAI_ENDPOINT não configurado.");
  }

  if (!apiKey) {
    throw new Error("AZURE_OPENAI_API_KEY não configurada.");
  }

  if (!deployment) {
    throw new Error("AZURE_OPENAI_DEPLOYMENT não configurado.");
  }

  const url = `${endpoint}/openai/v1/responses`;

  const body = {
    model: deployment,
    input: [
      {
        role: "user",
        content: mensagemCliente
      }
    ],
    agent_reference: {
      name: agentName,
      version: agentVersion,
      type: "agent_reference"
    }
  };

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "api-key": apiKey,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });

  const data = await response.json().catch(() => null);

  if (!response.ok) {
    console.error("Erro bruto do Foundry:", JSON.stringify(data, null, 2));

    throw new Error(
      `Foundry retornou erro HTTP ${response.status}: ${JSON.stringify(data)}`
    );
  }

  if (data && data.output_text) {
    return data.output_text;
  }

  if (data && Array.isArray(data.output)) {
    const textos = [];

    for (const item of data.output) {
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

  console.log("Resposta completa do Foundry:", JSON.stringify(data, null, 2));

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

  console.log("AZURE_OPENAI_ENDPOINT configurado:", process.env.AZURE_OPENAI_ENDPOINT ? "sim" : "não");
  console.log("AZURE_OPENAI_API_KEY configurado:", process.env.AZURE_OPENAI_API_KEY ? "sim" : "não");
  console.log("AZURE_OPENAI_DEPLOYMENT:", process.env.AZURE_OPENAI_DEPLOYMENT || "não definido");
  console.log("FOUNDRY_AGENT_NAME:", process.env.FOUNDRY_AGENT_NAME || "não definido");
  console.log("FOUNDRY_AGENT_VERSION:", process.env.FOUNDRY_AGENT_VERSION || "não definido");
});