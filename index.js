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

        await context.sendActivity(
          "Olá! Sou o assistente virtual da Micks Fibra. Recebi sua mensagem e estou pronto para ajudar. 😊"
        );
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
});