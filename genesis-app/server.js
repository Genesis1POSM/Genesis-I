import express from "express";
import path from "path";
import { fileURLToPath } from "url";

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
  setState = async (data) => {
    await pool.query(
      `INSERT INTO app_state (id, data, updated_at) VALUES (1, $1, now())
       ON CONFLICT (id) DO UPDATE SET data = $1, updated_at = now()`,
      [data]
    );
  };
  console.log("[genesis] Persistência: PostgreSQL conectado.");
} else {
  let memory = null;
  getState = async () => memory;
  setState = async (data) => { memory = data; };
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

app.put("/api/state", async (req, res) => {
  try {
    await setState(req.body.data);
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Falha ao salvar o estado." });
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
