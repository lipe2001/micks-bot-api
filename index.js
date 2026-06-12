const express = require("express");
const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication
} = require("botbuilder");

const { ClientSecretCredential } = require("@azure/identity");
const { AIProjectClient } = require("@azure/ai-projects");

const app = express();
app.use(express.json());

/**
 * Autenticação do Azure Bot Service
 * Usada para o Azure Bot conversar com este middleware no Render.
 */
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppType: process.env.MicrosoftAppType || "SingleTenant",
  MicrosoftAppId: process.env.MicrosoftAppId || "",
  MicrosoftAppPassword: process.env.MicrosoftAppPassword || "",
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId || ""
});

const adapter = new CloudAdapter(botFrameworkAuthentication);

/**
 * Memória local por conversa.
 *
 * Estrutura:
 * conversationId => {
 *   mensagens: [
 *     { role: "user", content: "..." },
 *     { role: "assistant", content: "..." }
 *   ],
 *   atualizadoEm: timestamp
 * }
 */
const historicoConversas = new Map();

/**
 * Quantidade máxima de mensagens salvas por conversa.
 * 12 mensagens = aproximadamente 6 interações cliente/bot.
 */
const MAX_MENSAGENS_HISTORICO = Number(process.env.MAX_MENSAGENS_HISTORICO || 12);

/**
 * Tempo para limpar conversas antigas da memória.
 * Padrão: 120 minutos.
 */
const TEMPO_EXPIRACAO_MINUTOS = Number(process.env.TEMPO_EXPIRACAO_MINUTOS || 120);
const TEMPO_EXPIRACAO_MS = TEMPO_EXPIRACAO_MINUTOS * 60 * 1000;

function limparEndpoint(endpoint) {
  return (endpoint || "").replace(/\/$/, "");
}

function limparHistoricosExpirados() {
  const agora = Date.now();

  for (const [conversationId, dados] of historicoConversas.entries()) {
    if (agora - dados.atualizadoEm > TEMPO_EXPIRACAO_MS) {
      historicoConversas.delete(conversationId);
      console.log(`Histórico expirado removido: ${conversationId}`);
    }
  }
}

function obterHistorico(conversationId) {
  limparHistoricosExpirados();

  if (!historicoConversas.has(conversationId)) {
    historicoConversas.set(conversationId, {
      mensagens: [],
      atualizadoEm: Date.now()
    });
  }

  return historicoConversas.get(conversationId);
}

function salvarInteracaoNoHistorico(conversationId, mensagemCliente, respostaBot) {
  const historico = obterHistorico(conversationId);

  historico.mensagens.push({
    role: "user",
    content: mensagemCliente
  });

  historico.mensagens.push({
    role: "assistant",
    content: respostaBot
  });

  if (historico.mensagens.length > MAX_MENSAGENS_HISTORICO) {
    historico.mensagens = historico.mensagens.slice(
      historico.mensagens.length - MAX_MENSAGENS_HISTORICO
    );
  }

  historico.atualizadoEm = Date.now();

  historicoConversas.set(conversationId, historico);
}

function montarInputComHistorico(conversationId, mensagemCliente) {
  const historico = obterHistorico(conversationId);

  const mensagens = [
    ...historico.mensagens,
    {
      role: "user",
      content: mensagemCliente
    }
  ];

  return mensagens;
}

let openaiClientPromise = null;

/**
 * Cria o client do projeto Foundry.
 * Usa:
 * - FOUNDRY_PROJECT_ENDPOINT
 * - AZURE_TENANT_ID
 * - AZURE_CLIENT_ID
 * - AZURE_CLIENT_SECRET
 */
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

/**
 * Chama o agente do Microsoft Foundry usando agent_reference.
 * Agora envia histórico da conversa junto com a nova mensagem.
 *
 * Usa:
 * - FOUNDRY_AGENT_NAME
 * - FOUNDRY_AGENT_VERSION
 */
async function chamarAgenteFoundry(conversationId, mensagemCliente) {
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

  const inputComHistorico = montarInputComHistorico(conversationId, mensagemCliente);

  console.log(
    `Enviando para Foundry | conversationId=${conversationId} | mensagens_no_contexto=${inputComHistorico.length}`
  );

  const response = await openai.responses.create(
    {
      input: inputComHistorico
    },
    {
      body: {
        agent_reference: agentReference
      }
    }
  );

  let respostaFinal = null;

  if (response.output_text) {
    respostaFinal = response.output_text;
  }

  if (!respostaFinal && Array.isArray(response.output)) {
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
      respostaFinal = textos.join("\n");
    }
  }

  if (!respostaFinal) {
    console.log("Resposta completa do Foundry:", JSON.stringify(response, null, 2));
    respostaFinal =
      "Não consegui gerar uma resposta agora. Vou direcionar para um atendente.";
  }

  salvarInteracaoNoHistorico(conversationId, mensagemCliente, respostaFinal);

  return respostaFinal;
}

/**
 * Tratamento global de erro do Bot Framework.
 */
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

/**
 * Health check simples.
 */
app.get("/", (req, res) => {
  res.status(200).send("Micks Bot API online.");
});

/**
 * Endpoint para ver status simples da memória.
 * Não expõe conteúdo das conversas, apenas quantidade.
 */
app.get("/debug/memoria", (req, res) => {
  limparHistoricosExpirados();

  res.status(200).json({
    conversas_em_memoria: historicoConversas.size,
    max_mensagens_por_conversa: MAX_MENSAGENS_HISTORICO,
    tempo_expiracao_minutos: TEMPO_EXPIRACAO_MINUTOS
  });
});

/**
 * Endpoint principal usado pelo Azure Bot Service.
 */
app.post("/api/messages", async (req, res) => {
  try {
    await adapter.process(req, res, async (context) => {
      if (context.activity.type === "message") {
        const textoCliente = context.activity.text || "";
        const conversationId =
          context.activity.conversation?.id ||
          context.activity.from?.id ||
          "conversa-sem-id";

        console.log("Mensagem recebida:", textoCliente);
        console.log("Conversation ID:", conversationId);

        try {
          const respostaFoundry = await chamarAgenteFoundry(
            conversationId,
            textoCliente
          );

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

  console.log("MAX_MENSAGENS_HISTORICO:", MAX_MENSAGENS_HISTORICO);
  console.log("TEMPO_EXPIRACAO_MINUTOS:", TEMPO_EXPIRACAO_MINUTOS);
});