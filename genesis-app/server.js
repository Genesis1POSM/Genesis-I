import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import { criarServicoDeAlertas } from "./alerts.js";
import { acharUsuario, mesclarMateriais } from "./matmerge.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
app.use(express.json({ limit: "15mb" }));

const PORT = process.env.PORT || 3000;
const DATABASE_URL = process.env.DATABASE_URL;

/* ============================================================
   STORAGE LAYER
   - If DATABASE_URL is set (Render Postgres), state is persisted
     there, in a single JSONB row, shared by everyone.
   - If DATABASE_URL is NOT set (e.g. running locally without a
     database configured yet), falls back to an in-memory store
     so `npm start` still works for a quick local test — but this
     mode does NOT persist across server restarts.
   ============================================================ */
let getState, setState;

if (DATABASE_URL) {
  const { default: pg } = await import("pg");
  const pool = new pg.Pool({
    connectionString: DATABASE_URL,
    ssl: DATABASE_URL.includes("localhost") ? false : { rejectUnauthorized: false },
  });

  await pool.query(`
    CREATE TABLE IF NOT EXISTS app_state (
      id INT PRIMARY KEY DEFAULT 1,
      data JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);

  getState = async () => {
    const { rows } = await pool.query("SELECT data FROM app_state WHERE id = 1");
    return rows[0]?.data ?? null;
  };
  /* MESCLA em vez de substituir o objeto inteiro: se algum cliente salvar com uma versão do
     front-end mais antiga (ou uma aba desatualizada aberta no navegador) que não conhece um campo
     mais novo (ex.: tmDue, planningItems), esse campo é simplesmente OMITIDO do payload enviado —
     e se aqui fizéssemos um "replace" puro, esse campo desapareceria do banco pra sempre na
     próxima gravação, mesmo sem ninguém ter mexido nele. Fazendo merge por chave de topo, um campo
     ausente no payload preserva o valor já salvo; só é sobrescrito quando o cliente realmente o envia. */
  setState = async (partial) => {
    const { rows } = await pool.query("SELECT data FROM app_state WHERE id = 1");
    const current = rows[0]?.data || {};
    const merged = { ...current, ...partial };
    await pool.query(
      `INSERT INTO app_state (id, data, updated_at) VALUES (1, $1, now())
       ON CONFLICT (id) DO UPDATE SET data = $1, updated_at = now()`,
      [merged]
    );
    return merged;
  };
  console.log("[genesis] Persistência: PostgreSQL conectado.");
} else {
  let memory = null;
  getState = async () => memory;
  setState = async (partial) => { memory = { ...(memory || {}), ...partial }; return memory; };
  console.warn("[genesis] AVISO: DATABASE_URL não definida — usando memória local (não persiste entre reinícios do servidor).");
}

/* ============================================================
   API
   ============================================================ */
app.get("/api/state", async (req, res) => {
  try {
    const data = await getState();
    res.json({ data });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Falha ao carregar o estado." });
  }
});

app.get("/api/materiais", async (req, res) => {
  try {
    const d = (await getState()) || {};
    res.json({ matReqs: d.matReqs || [], matDeleted: d.matDeleted || [] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Falha ao carregar materiais." });
  }
});

app.put("/api/state", async (req, res) => {
  try {
    const data = { ...(req.body.data || {}) };
    let extra = {};
    if (Array.isArray(data.matReqs)) {
      /* Materiais: quem salva é identificado pelo usuário/senha enviados nos cabeçalhos e só consegue
         alterar as linhas do seu departamento (Almoxarifado e administrador: todas) */
      const atual = (await getState()) || {};
      const usuario = acharUsuario(atual.users, req.get("x-genesis-user"), req.get("x-genesis-pass"));
      const r = mesclarMateriais(atual.matReqs, atual.matDeleted, data.matReqs, data.matDeleted, usuario);
      data.matReqs = r.rows; data.matDeleted = r.deleted;
      extra = { matReqs: r.rows, matDeleted: r.deleted };
    }
    await setState(data);
    res.json({ ok: true, ...extra });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Falha ao salvar o estado." });
  }
});

/* ============================================================
   ALERTA POR E-MAIL (jobs a vencer) — ver alerts.js
   /api/alerts/run?token=SEU_ALERT_TOKEN          → roda a verificação do dia (serve para um "ping" externo)
   /api/alerts/run?token=SEU_ALERT_TOKEN&force=1  → envia agora (teste), ignorando "já enviado hoje" e fim de semana
   ============================================================ */
const alertas = criarServicoDeAlertas({ getState, setState });
alertas.iniciar();
app.get("/api/alerts/run", async (req, res) => {
  if (!process.env.ALERT_TOKEN || req.query.token !== process.env.ALERT_TOKEN) return res.status(403).json({ error: "Token inválido." });
  try {
    res.json(await alertas.executar({ forcar: req.query.force === "1" }));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

/* ============================================================
   STATIC FRONTEND (Vite build output)
   ============================================================ */
app.use(express.static(path.join(__dirname, "dist")));
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "dist", "index.html"));
});

app.listen(PORT, () => {
  console.log(`[genesis] Servidor rodando na porta ${PORT}`);
});
