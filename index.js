const express = require("express");
const { BotFrameworkAdapter } = require("botbuilder");

const app = express();
app.use(express.json());

const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId || "",
  appPassword: process.env.MicrosoftAppPassword || ""
});

adapter.onTurnError = async (context, error) => {
  console.error("Erro no bot:", error);
  await context.sendActivity(
    "Tive um problema ao processar sua mensagem. Vou direcionar para um atendente."
  );
};

app.get("/", (req, res) => {
  res.status(200).send("Micks Bot API online.");
});

app.post("/api/messages", async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
    if (context.activity.type === "message") {
      const textoCliente = context.activity.text || "";

      console.log("Mensagem recebida:", textoCliente);

      await context.sendActivity(
        "Olá! Sou o assistente virtual da Micks Fibra. Recebi sua mensagem e estou pronto para ajudar. 😊"
      );
    }
  });
});

const port = process.env.PORT || 3978;

app.listen(port, () => {
  console.log(`Micks Bot API rodando na porta ${port}`);
});