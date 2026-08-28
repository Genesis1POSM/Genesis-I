import React, { useState, useMemo, useRef } from "react";
import {
  LayoutGrid, Ship, Wrench, Package, Wallet, Calculator,
  Plus, Trash2, ChevronDown, ChevronUp, ChevronRight, AlertTriangle,
  Download, Upload, FileText, LogOut, Lock, User, X, Settings, DollarSign, Clock
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell
} from "recharts";
import * as XLSX from "xlsx";

/* ============================================================
   THEME — offshore ops console: deep marine dark, amber beacon
   accent, teal secondary, mono for IDs/data, sans for UI.
   ============================================================ */
const Theme = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

    .genesis {
      --bg: #0A1220;
      --panel: #101B2E;
      --panel-alt: #16223A;
      --panel-raised: #1C2A44;
      --border: #22314C;
      --border-soft: #1A2740;
      --text: #E7ECF3;
      --text-dim: #8D9BB5;
      --text-faint: #5D6E8C;
      --accent: #F2A93B;
      --accent-dim: #7A5A24;
      --teal: #3FC1C9;
      --ok: #35D399;
      --warn: #F2C94C;
      --crit: #F2685B;
      --sans: 'IBM Plex Sans', system-ui, sans-serif;
      --mono: 'IBM Plex Mono', ui-monospace, monospace;

      background: var(--bg);
      color: var(--text);
      font-family: var(--sans);
      min-height: 100%;
      display: flex;
      flex-direction: column;
      font-size: 13px;
      line-height: 1.4;
    }
    .genesis * { box-sizing: border-box; }
    .genesis ::selection { background: var(--accent-dim); }

    /* ---------- Login ---------- */
    .g-login-wrap {
      min-height: 100vh; width: 100%; display: flex; align-items: center; justify-content: center;
      background: radial-gradient(circle at 30% 20%, #142138 0%, #0A1220 60%);
    }
    .g-login-card {
      width: 340px; background: var(--panel); border: 1px solid var(--border);
      border-radius: 6px; padding: 30px 28px; box-shadow: 0 20px 60px rgba(0,0,0,0.45);
    }
    .g-login-brand { font-family: var(--mono); font-weight: 700; font-size: 20px; color: var(--accent); letter-spacing: .5px; margin-bottom: 2px; }
    .g-login-sub { font-family: var(--mono); font-size: 10px; color: var(--text-faint); letter-spacing: 1px; text-transform: uppercase; margin-bottom: 24px; }
    .g-login-field { margin-bottom: 14px; }
    .g-login-field label { display: flex; align-items: center; gap: 6px; font-size: 11px; color: var(--text-dim); margin-bottom: 6px; text-transform: uppercase; letter-spacing: .5px; }
    .g-login-field input {
      width: 100%; background: var(--panel-raised); border: 1px solid var(--border); color: var(--text);
      font-family: var(--sans); font-size: 13px; padding: 9px 11px; border-radius: 3px;
    }
    .g-login-field input:focus { outline: none; border-color: var(--accent); }
    .g-login-error { color: var(--crit); font-size: 11.5px; margin-bottom: 12px; }
    .g-login-hint { margin-top: 18px; font-size: 10.5px; color: var(--text-faint); font-family: var(--mono); border-top: 1px solid var(--border-soft); padding-top: 12px; }

    /* ---------- Horizontal top nav ---------- */
    .g-topnav {
      display: flex; align-items: center; gap: 22px;
      padding: 0 22px; height: 54px; flex-shrink: 0;
      background: var(--panel); border-bottom: 1px solid var(--border);
    }
    .g-brand-mark { font-family: var(--mono); font-weight: 700; font-size: 15px; color: var(--accent); letter-spacing: .5px; white-space: nowrap; }
    .g-nav-row { display: flex; align-items: center; gap: 2px; flex: 1; overflow-x: auto; }
    .g-nav-item {
      display: flex; align-items: center; gap: 7px;
      padding: 8px 13px; border-radius: 4px; color: var(--text-dim); cursor: pointer;
      font-size: 12.5px; font-weight: 500; white-space: nowrap;
      border-bottom: 2px solid transparent;
    }
    .g-nav-item:hover { background: var(--panel-alt); color: var(--text); }
    .g-nav-item.active { color: var(--accent); border-bottom: 2px solid var(--accent); background: var(--panel-alt); }
    .g-logout { display: flex; align-items: center; gap: 6px; color: var(--text-faint); cursor: pointer; font-size: 12px; white-space: nowrap; }
    .g-logout:hover { color: var(--crit); }

    /* ---------- Global period filter bar ---------- */
    .g-filterbar {
      display: flex; align-items: flex-end; gap: 16px; flex-wrap: wrap;
      padding: 12px 22px; background: var(--panel-alt); border-bottom: 1px solid var(--border);
    }
    .g-field { display: flex; flex-direction: column; gap: 4px; }
    .g-field label { font-size: 9.5px; text-transform: uppercase; letter-spacing: .6px; color: var(--text-faint); font-family: var(--mono); }
    .g-field input, .g-field select {
      background: var(--panel-raised); border: 1px solid var(--border); color: var(--text);
      font-family: var(--mono); font-size: 12px; padding: 6px 8px; border-radius: 3px;
    }
    .g-field input:focus, .g-field select:focus { outline: none; border-color: var(--accent); }
    .g-mode-toggle { display: flex; border: 1px solid var(--border); border-radius: 4px; overflow: hidden; }
    .g-mode-toggle button {
      background: var(--panel-raised); border: none; color: var(--text-dim); padding: 7px 12px;
      font-size: 11.5px; font-weight: 500; cursor: pointer; font-family: var(--sans);
    }
    .g-mode-toggle button.active { background: var(--accent); color: #1a1204; font-weight: 600; }
    .g-filter-spacer { flex: 1; }
    .g-filter-summary { font-family: var(--mono); font-size: 11px; color: var(--text-faint); align-self: center; }
    .g-period-bar {
      display: flex; align-items: flex-end; gap: 14px; flex-wrap: wrap;
      background: var(--panel-alt); border: 1px solid var(--border);
      border-radius: 4px; padding: 12px 14px;
    }

    /* ---------- Page action row ---------- */
    .g-pageactions {
      display: flex; align-items: center; justify-content: space-between; gap: 10px; flex-wrap: wrap;
      padding: 16px 22px 0 22px;
    }
    .g-title { font-size: 16px; font-weight: 600; letter-spacing: .2px; }
    .g-title-sub { color: var(--text-faint); font-size: 11.5px; margin-top: 2px; font-family: var(--mono); }
    .g-body { padding: 16px 22px 40px 22px; overflow-y: auto; }

    .g-btn {
      display: inline-flex; align-items: center; gap: 6px;
      background: var(--panel-raised);
      border: 1px solid var(--border);
      color: var(--text);
      padding: 7px 12px;
      border-radius: 3px;
      font-size: 12px;
      font-weight: 500;
      cursor: pointer;
      font-family: var(--sans);
    }
    .g-btn:hover { border-color: var(--accent); color: var(--accent); }
    .g-btn.primary {
      background: var(--accent);
      border-color: var(--accent);
      color: #1a1204;
      font-weight: 600;
    }
    .g-btn.primary:hover { filter: brightness(1.08); color: #1a1204; }
    .g-btn.ghost { background: transparent; border-color: transparent; padding: 4px 6px; }
    .g-btn.danger:hover { border-color: var(--crit); color: var(--crit); }

    /* ---------- KPI cards ---------- */
    .g-kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 14px; }
    .g-kpi {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 4px;
      padding: 13px 14px;
      position: relative;
      overflow: hidden;
    }
    .g-kpi::before {
      content: '';
      position: absolute; left: 0; top: 0; bottom: 0; width: 2px;
      background: var(--kpi-accent, var(--teal));
    }
    .g-kpi.clickable { cursor: pointer; transition: border-color .12s, background .12s; }
    .g-kpi.clickable:hover { background: var(--panel-alt); }
    .g-kpi.active { background: var(--panel-alt); border-color: var(--kpi-accent, var(--teal)); }
    .g-kpi-label {
      font-size: 10px; text-transform: uppercase; letter-spacing: .8px;
      color: var(--text-faint); font-family: var(--mono); margin-bottom: 6px;
    }
    .g-kpi-value { font-size: 19px; font-weight: 600; font-family: var(--mono); }
    .g-kpi-value.small { font-size: 16px; }
    .g-section-label {
      font-size: 10.5px; text-transform: uppercase; letter-spacing: 1px; color: var(--text-faint);
      margin: 18px 0 8px 2px; font-family: var(--mono);
    }
    .g-section-label:first-child { margin-top: 0; }

    /* ---------- Panels / sections ---------- */
    .g-panel {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 4px;
      padding: 16px;
      margin-bottom: 16px;
    }
    .g-panel-head {
      display: flex; align-items: center; justify-content: space-between;
      margin-bottom: 12px;
      flex-wrap: wrap;
      gap: 8px;
    }
    .g-panel-title {
      font-size: 11px; text-transform: uppercase; letter-spacing: 1px;
      color: var(--text-dim); font-weight: 600;
    }
    .g-grid-2 { display: grid; grid-template-columns: 1.1fr 1fr; gap: 16px; }

    /* ---------- WP tag (signature element) ---------- */
    .g-tag {
      display: inline-flex; align-items: center;
      font-family: var(--mono); font-size: 10.5px; font-weight: 600;
      color: var(--accent);
      background: rgba(242,169,59,0.08);
      border: 1px solid var(--accent-dim);
      padding: 3px 8px 3px 9px;
      clip-path: polygon(0 0, 100% 0, 100% 100%, 8px 100%, 0 calc(100% - 8px));
      letter-spacing: .3px;
      white-space: nowrap;
    }

    /* ---------- Status pill ---------- */
    .g-pill {
      display: inline-flex; align-items: center; gap: 5px;
      font-size: 10.5px; font-weight: 600; padding: 3px 9px;
      border-radius: 20px; white-space: nowrap; font-family: var(--sans);
    }
    .g-dot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }

    /* ---------- Tables ---------- */
    .g-table { width: 100%; border-collapse: collapse; }
    .g-table-wrap { overflow-x: auto; }
    .g-edit-wrap {
      background: transparent; border: 1px solid transparent; color: var(--text); font-family: inherit;
      font-size: 12px; padding: 5px 6px; border-radius: 3px; width: 100%; min-width: 0;
      white-space: normal; word-break: break-word; resize: vertical; line-height: 1.35;
    }
    .g-edit-wrap:hover { border-color: var(--border); }
    .g-edit-wrap:focus { outline: none; border-color: var(--accent); background: var(--panel-raised); }
    .g-table th {
      text-align: left; font-size: 10px; text-transform: uppercase;
      letter-spacing: .6px; color: var(--text-faint); font-weight: 600;
      padding: 7px 8px; border-bottom: 1px solid var(--border);
      font-family: var(--mono); white-space: nowrap;
    }
    .g-table td {
      padding: 6px 8px; border-bottom: 1px solid var(--border-soft);
      font-size: 12px; vertical-align: middle; white-space: nowrap;
    }
    .g-table tr.g-row:hover { background: var(--panel-alt); }
    .g-table tr.g-expand-row { background: var(--panel-alt); }

    .g-edit {
      background: transparent; border: 1px solid transparent;
      color: var(--text); font-family: inherit; font-size: 12px;
      padding: 3px 5px; border-radius: 3px; width: 100%; min-width: 60px;
    }
    .g-edit:hover { border-color: var(--border); }
    .g-edit:focus { outline: none; border-color: var(--accent); background: var(--panel-raised); }
    select.g-edit { cursor: pointer; }
    .g-edit.mono { font-family: var(--mono); }
    .g-edit.num { text-align: right; font-family: var(--mono); }

    .g-list-item {
      display: flex; align-items: center; justify-content: space-between;
      padding: 8px 0; border-bottom: 1px solid var(--border-soft);
      font-size: 12px; gap: 10px;
    }
    .g-list-item:last-child { border-bottom: none; }

    /* ---------- Gantt ---------- */
    .g-gantt-wrap { overflow-x: auto; }
    .g-gantt { min-width: 1100px; }
    .g-gantt-header { display: flex; margin-left: 400px; border-bottom: 1px solid var(--border); padding-bottom: 6px; margin-bottom: 4px; }
    .g-gantt-day { flex: 1; text-align: center; font-family: var(--mono); font-size: 9.5px; color: var(--text-faint); }
    .g-gantt-row { display: flex; align-items: center; min-height: 56px; padding: 8px 10px; margin-bottom: 4px; border-radius: 6px; background: var(--panel); border: 1px solid var(--border-soft); transition: border-color .12s, background .12s; }
    .g-gantt-row:hover { border-color: var(--border); background: var(--panel-alt); }
    .g-gantt-taskinfo { width: 860px; flex-shrink: 0; padding-right: 14px; }
    .g-gantt-mini-label { font-size: 8.5px; text-transform: uppercase; letter-spacing: .4px; color: var(--text-faint); font-family: var(--mono); margin-bottom: 2px; }
    .g-gantt-mini-dt {
      background: var(--panel-raised); border: 1px solid var(--border); color: var(--text);
      font-family: var(--mono); font-size: 10px; padding: 4px 5px; border-radius: 3px; width: 128px;
    }
    .g-gantt-mini-dt:focus { outline: none; border-color: var(--accent); }
    .g-gantt-track { flex: 1; position: relative; height: 20px; background:
      repeating-linear-gradient(90deg, var(--border-soft) 0, var(--border-soft) 1px, transparent 1px, transparent calc(100% / var(--gantt-days, 10))); }
    .g-gantt-bar { position: absolute; top: 3px; height: 20px; border-radius: 5px; display: flex; align-items: center; padding: 0 7px; font-size: 9.5px; font-weight: 700; font-family: var(--mono); color: #0A1220; overflow: hidden; white-space: nowrap; cursor: pointer; box-shadow: 0 1px 4px rgba(0,0,0,0.3); }
    .g-gantt-bar-fill { position: absolute; left: 0; top: 0; bottom: 0; background: rgba(255,255,255,0.35); }

    .g-gantt-group-row {
      display: flex; align-items: center; justify-content: space-between; gap: 8px;
      margin-left: 0; padding: 7px 10px; margin-top: 10px;
      background: var(--panel-raised); border-left: 2px solid var(--accent); border-radius: 3px;
    }
    .g-gantt-group-title {
      background: transparent; border: none; color: var(--accent); font-weight: 700;
      font-size: 11.5px; text-transform: uppercase; letter-spacing: .6px; font-family: var(--sans);
      padding: 2px 4px; flex: 1; min-width: 100px;
    }
    .g-gantt-group-title:focus { outline: none; background: var(--panel-alt); border-radius: 3px; }
    .g-gantt-empty { margin-left: 400px; padding: 8px 0; color: var(--text-faint); font-size: 11px; font-style: italic; }
    .g-gantt-name-edit {
      font-size: 12px; font-family: var(--sans); padding: 4px 6px; min-height: 40px;
      width: 100%; flex: 1;
    }
    .g-gantt-editrow {
      display: flex; gap: 14px; flex-wrap: wrap; align-items: flex-end;
      margin-left: 400px; padding: 8px 10px; margin-bottom: 4px;
      background: var(--panel-alt); border: 1px solid var(--border-soft); border-radius: 4px;
    }

    /* ---------- Misc ---------- */
    .g-muted { color: var(--text-faint); }
    .g-flex { display: flex; align-items: center; gap: 8px; }
    .g-bar-bg { background: var(--panel-alt); border-radius: 20px; height: 6px; width: 100%; overflow: hidden; }
    .g-bar-fg { height: 100%; border-radius: 20px; }
    .g-alert {
      display: flex; align-items: flex-start; gap: 8px;
      background: rgba(242,104,91,0.08); border: 1px solid rgba(242,104,91,0.35);
      color: #F6A79E; padding: 8px 10px; border-radius: 4px; font-size: 11.5px; margin-bottom: 8px;
    }
  `}</style>
);

/* ============================================================
   HELPERS
   ============================================================ */
const pad2 = (n) => String(n).padStart(2, "0");
const fmt = (n) => "R$ " + Number(n || 0).toLocaleString("pt-BR");

const fmtDate = (d) => {
  if (!d) return "—";
  const [, m, day] = d.split("-");
  return `${day}/${m}`;
};
const fmtDateTime = (dt) => {
  if (!dt) return "—";
  const d = new Date(dt);
  if (isNaN(d)) return "—";
  return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
};
const fmtPeriodDate = (d) => {
  if (!d) return "—";
  const dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt)) return "—";
  return `${pad2(dt.getDate())}/${pad2(dt.getMonth() + 1)}/${dt.getFullYear()}`;
};

const uid = (prefix) => `${prefix}-${Math.random().toString(36).slice(2, 7).toUpperCase()}`;
const todayISO = () => new Date().toISOString().slice(0, 10);

const MONTH_NAMES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

const WP_STATUS = ["Planejamento", "Não iniciado", "Em andamento", "Concluído", "Cancelado"];
const WP_STATUS_DEFAULT_PROGRESS = { "Planejamento": 0, "Não iniciado": 0, "Em andamento": 50, "Concluído": 100, "Cancelado": 0 };
const MAT_STATUS = ["Solicitado", "Em aprovação", "Cotação", "Cotação recebida", "Em aprovação comercial", "PO emitida", "Em fabricação", "Em trânsito", "Recebido", "Entregue a bordo"];
const PAY_STATUS = ["Orçamento", "Aprovado", "PO emitida", "Serviço executado", "Medição aprovada", "NF recebida", "NF validada", "Pagamento programado", "Pago"];
const PRIORITY = ["Baixa", "Média", "Alta", "Crítica"];
/* categories + Orçado (USD) exactly as in the uploaded drill-down report */
const CATEGORIES = ["Drilling", "Elétrica", "Hse", "Hull & Structure", "Integridade", "Lubrificantes", "Marine", "Mecânica", "R&R Elétrica", "R&R Mecânica"];
const CATEGORY_BUDGET_USD = {
  "Drilling": 0,
  "Elétrica": 27900,
  "Hse": 6200,
  "Hull & Structure": 12772,
  "Integridade": 12400,
  "Lubrificantes": 31000,
  "Marine": 21700,
  "Mecânica": 27900,
  "R&R Elétrica": 18600,
  "R&R Mecânica": 18600,
};
const DISCIPLINES = CATEGORIES;

const paymentSituation = (p) => {
  if (p.status === "Pago") return "Pago";
  if (p.due && p.due < todayISO()) return "Atrasado";
  return "Pendente";
};
const situationColor = { Pago: "var(--ok)", Pendente: "var(--warn)", Atrasado: "var(--crit)" };
const daysLate = (p) => {
  if (!p.due) return 0;
  const diff = (new Date(todayISO()) - new Date(p.due)) / 86400000;
  return Math.max(0, Math.round(diff));
};

/* ---------- column maps for Excel export/import ---------- */
const WP_COLS = [
  ["id", "ID"], ["name", "Manutenção"], ["portCall", "Port Call"], ["group", "Categoria"], ["ganttCategory", "Categoria Operacional"], ["empresa", "Empresa"], ["md", "MD"], ["rc", "RC"], ["obs", "Observação"],
  ["discipline", "Disciplina (custo)"],
  ["budget", "Budget"], ["committed", "Comprometido"], ["actual", "Realizado"], ["forecast", "Forecast"],
  ["start", "Início"], ["end", "Fim"], ["status", "Status"], ["progress", "Progresso (%)"],
  ["dataRealInicio", "Data Real de Início"], ["dataRealFim", "Data Real de Conclusão"], ["repeatOf", "Repetição de (ID)"],
];
const MAT_COLS = [
  ["tmMaster", "TM Master"], ["departamento", "Departamento"], ["sap", "SAP"], ["descricao", "Descrição"],
  ["quantidade", "Quantidade"], ["priority", "Prioridade"], ["dataSolicitacao", "Data da solicitação"], ["dataNecessidade", "Data da Necessidade"],
  ["reserva", "Reserva"], ["rc", "RC"], ["po", "PO"], ["linhaPo", "Linha da PO"], ["valor", "Valor"],
  ["eta", "ETA"], ["obs", "Observação"], ["dataRecebimento", "Data de Recebimento"], ["status", "Status"],
  ["id", "ID"], ["wp", "Work Package"],
];
const PAY_COLS = [
  ["id", "ID"], ["service", "Serviço"], ["po", "PO"], ["poValue", "Valor PO"],
  ["nf", "NF"], ["nfValue", "Valor NF"], ["issue", "Emissão"], ["due", "Vencimento"], ["status", "Status"],
];
const STATUS_PAGAMENTO_OPTIONS = [
  "Aguardando Orçamento", "Aguardando Suprimentos", "Aguardando Medição", "Aprovação Pendente",
  "Aguardando NF", "Pagamento Programado", "Pago", "Cancelado",
];
const STATUS_PAGAMENTO_COLOR = {
  "Aguardando Orçamento": "#8D9BB5",
  "Aguardando Suprimentos": "#F2C94C",
  "Aguardando Medição": "#3FC1C9",
  "Aprovação Pendente": "#F2685B",
  "Aguardando NF": "#F2A93B",
  "Pagamento Programado": "#9B8CF2",
  "Pago": "#35D399",
  "Cancelado": "#5D6E8C",
};
const INV_COLS = [
  ["id", "ID"], ["date", "Data"], ["assunto", "Manutenção"], ["empresa", "Empresa"], ["md", "MD"],
  ["mdSentDate", "Data de Envio da MD"], ["diffDays", "Diferença de Dias"], ["daysOpenTotal", "Dias em Aberto Total"],
  ["rc", "RC"], ["serviceStatus", "Status do Serviço"], ["poContrato", "PO / Contrato"], ["medicao", "Medição"],
  ["valorTotal", "Valor Total"], ["saldoPo", "Saldo PO"], ["obs", "Observações"],
  ["statusPagamento", "Status Pagamento"], ["dataPagamento", "Data de Pagamento"],
];
const NUMERIC_KEYS = new Set(["budget", "committed", "actual", "forecast", "progress", "quantidade", "valor", "poValue", "nfValue", "diffDays", "daysOpenTotal", "valorTotal", "saldoPo"]);
const DATE_KEYS = new Set(["dataSolicitacao", "dataNecessidade", "eta", "dataRecebimento", "issue", "due", "date", "mdSentDate", "dataPagamento", "dataRealInicio", "dataRealFim"]);
const DATETIME_KEYS = new Set(["start", "end"]);

const cellToDateStr = (v) => {
  if (!v && v !== 0) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v);
};
const cellToDateTimeStr = (v) => {
  if (!v && v !== 0) return "";
  if (v instanceof Date) return `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}T${pad2(v.getHours())}:${pad2(v.getMinutes())}`;
  return String(v);
};
const rowsToSheet = (rows, cols) => rows.map((r) => Object.fromEntries(cols.map(([key, label]) => [label, r[key]])));
const sheetToRows = (json, cols) =>
  json.map((row) => {
    const out = {};
    cols.forEach(([key, label]) => {
      let v = row[label];
      if (v === undefined) v = "";
      if (NUMERIC_KEYS.has(key)) v = v === "" ? 0 : Number(v);
      else if (DATETIME_KEYS.has(key)) v = cellToDateTimeStr(v);
      else if (DATE_KEYS.has(key)) v = cellToDateStr(v);
      else v = String(v);
      out[key] = v;
    });
    return out;
  });

const statusColor = (status) => {
  const map = {
    "Concluído": "var(--ok)", "Pago": "var(--ok)", "Recebido": "var(--ok)", "Entregue a bordo": "var(--ok)",
    "Em andamento": "var(--teal)", "Em trânsito": "var(--teal)", "Em fabricação": "var(--teal)",
    "Crítico": "var(--crit)", "Atrasado": "var(--crit)",
    "Planejamento": "var(--text-faint)", "Não iniciado": "var(--text-faint)", "Solicitado": "var(--text-faint)", "Orçamento": "var(--text-faint)",
    "Cancelado": "var(--text-faint)",
  };
  return map[status] || "var(--warn)";
};

/* ============================================================
   REUSABLE EDITABLE CELLS
   ============================================================ */
const EText = ({ value, onChange, mono, align }) => (
  <input className={`g-edit ${mono ? "mono" : ""}`} style={{ textAlign: align }} value={value}
    onChange={(e) => onChange(e.target.value)} />
);
const ETextArea = ({ value, onChange, rows = 2 }) => (
  <textarea className="g-edit-wrap" rows={rows} value={value || ""}
    onChange={(e) => onChange(e.target.value)} />
);
const ENum = ({ value, onChange }) => (
  <input type="number" className="g-edit num" value={value}
    onChange={(e) => onChange(Number(e.target.value))} />
);
const EDate = ({ value, onChange }) => (
  <input type="date" className="g-edit mono" value={value || ""}
    onChange={(e) => onChange(e.target.value)} />
);
const EDateTime = ({ value, onChange }) => (
  <input type="datetime-local" className="g-edit mono" value={value || ""}
    onChange={(e) => onChange(e.target.value)} />
);
const ESelect = ({ value, onChange, options, style }) => (
  <select className="g-edit" style={style} value={value} onChange={(e) => onChange(e.target.value)}>
    {options.map((o) => <option key={o} value={o}>{o}</option>)}
  </select>
);
const Pill = ({ status }) => (
  <span className="g-pill" style={{ background: "var(--panel-raised)", color: "var(--text)" }}>
    <span className="g-dot" style={{ background: statusColor(status) }} />{status}
  </span>
);
const SituationPill = ({ situation }) => (
  <span className="g-pill" style={{ background: "var(--panel-raised)", color: "var(--text)" }}>
    <span className="g-dot" style={{ background: situationColor[situation] }} />{situation}
  </span>
);

/* dropdown de Status de Pagamento com destaque de cor forte, para chamar atenção nas tabelas */
const StatusPagamentoSelect = ({ value, onChange }) => {
  const color = STATUS_PAGAMENTO_COLOR[value] || "var(--warn)";
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{
        width: "100%", minWidth: 190, fontWeight: 700, fontSize: 12.5, cursor: "pointer",
        color, background: `${color}20`, border: `1.5px solid ${color}`, borderRadius: 5,
        padding: "6px 8px", fontFamily: "var(--sans)",
      }}
    >
      {STATUS_PAGAMENTO_OPTIONS.map((o) => <option key={o} value={o} style={{ color: "#000" }}>{o}</option>)}
    </select>
  );
};

/* mesma ideia, para o vocabulário de status usado no Dashboard de Valores (PAY_STATUS) */
const PAY_STATUS_COLOR = {
  "Orçamento": "#8D9BB5",
  "Aprovado": "#3FC1C9",
  "PO emitida": "#F2C94C",
  "Serviço executado": "#F2A93B",
  "Medição aprovada": "#3FC1C9",
  "NF recebida": "#F2C94C",
  "NF validada": "#9B8CF2",
  "Pagamento programado": "#9B8CF2",
  "Pago": "#35D399",
};
const PayStatusSelect = ({ value, onChange }) => {
  const color = PAY_STATUS_COLOR[value] || "var(--warn)";
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{
        width: "100%", minWidth: 190, fontWeight: 700, fontSize: 12.5, cursor: "pointer",
        color, background: `${color}20`, border: `1.5px solid ${color}`, borderRadius: 5,
        padding: "6px 8px", fontFamily: "var(--sans)",
      }}
    >
      {PAY_STATUS.map((o) => <option key={o} value={o} style={{ color: "#000" }}>{o}</option>)}
    </select>
  );
};

/* ---------- ordenação de tabela ao clicar no cabeçalho ---------- */
const compareRows = (a, b, key, dir) => {
  let av = a[key], bv = b[key];
  const an = parseFloat(av), bn = parseFloat(bv);
  const numeric = av !== "" && av !== null && av !== undefined && bv !== "" && bv !== null && bv !== undefined && !isNaN(an) && !isNaN(bn);
  if (numeric) return (an - bn) * dir;
  av = (av ?? "").toString().toLowerCase();
  bv = (bv ?? "").toString().toLowerCase();
  if (av < bv) return -dir;
  if (av > bv) return dir;
  return 0;
};
const sortRows = (rows, sort) => (sort.key ? [...rows].sort((a, b) => compareRows(a, b, sort.key, sort.dir)) : rows);
const SortTh = ({ children, sortKey, sort, setSort, style }) => {
  const active = sort.key === sortKey;
  return (
    <th
      style={{ cursor: "pointer", userSelect: "none", ...style }}
      onClick={() => setSort((s) => (s.key === sortKey ? { key: sortKey, dir: -s.dir } : { key: sortKey, dir: 1 }))}
      title="Clique para ordenar"
    >
      <span className="g-flex" style={{ gap: 4 }}>
        {children}
        <span style={{ fontSize: 9, opacity: active ? 1 : 0.3, color: active ? "var(--accent)" : undefined }}>
          {active && sort.dir === -1 ? "▼" : "▲"}
        </span>
      </span>
    </th>
  );
};

/* cabeçalho de tabela clicável para ordenar — usado nas 3 tabelas da aba Pagamentos */
const SortableTh = ({ label, sortKey, sort, setSort, style }) => {
  const active = sort.key === sortKey;
  return (
    <th
      style={{ ...style, cursor: "pointer", userSelect: "none" }}
      onClick={() => setSort((s) => (s.key === sortKey ? { key: sortKey, dir: -s.dir } : { key: sortKey, dir: 1 }))}
    >
      <span className="g-flex" style={{ gap: 3 }}>
        {label}
        {active
          ? (sort.dir === 1 ? <ChevronUp size={11} /> : <ChevronDown size={11} />)
          : <ChevronDown size={11} style={{ opacity: 0.25 }} />}
      </span>
    </th>
  );
};

/* aplica a ordenação corrente a uma lista, usando getters opcionais para colunas calculadas */
function applySort(rows, sort, getters = {}) {
  if (!sort.key) return rows;
  const getVal = getters[sort.key] || ((r) => r[sort.key]);
  return [...rows].sort((a, b) => {
    let av = getVal(a), bv = getVal(b);
    if (av === null || av === undefined) av = "";
    if (bv === null || bv === undefined) bv = "";
    if (typeof av === "number" && typeof bv === "number") return (av - bv) * sort.dir;
    return String(av).localeCompare(String(bv), "pt-BR", { numeric: true }) * sort.dir;
  });
}

/* ============================================================
   LOGIN SCREEN (client-side demo gate — no real backend/security)
   ============================================================ */
const DEFAULT_USERS = [
  { id: uid("USR"), name: "Lethicia", username: "lethicia", password: "GenesisI" },
  { id: uid("USR"), name: "Davi", username: "davi", password: "GenesisI" },
  { id: uid("USR"), name: "Mariana", username: "mariana", password: "GenesisI" },
];

function LoginScreen({ users, onLogin }) {
  const [u, setU] = useState("");
  const [p, setP] = useState("");
  const [error, setError] = useState("");

  const submit = (e) => {
    e.preventDefault();
    const match = users.find((x) => x.username.toLowerCase() === u.trim().toLowerCase() && x.password === p);
    if (match) { setError(""); onLogin(match); }
    else setError("Usuário ou senha inválidos.");
  };

  return (
    <div className="genesis g-login-wrap">
      <Theme />
      <form className="g-login-card" onSubmit={submit}>
        <div className="g-login-brand">GENESIS I</div>
        <div className="g-login-sub">Maintenance &amp; Port Call</div>
        <div className="g-login-field">
          <label><User size={12} />Usuário</label>
          <input value={u} onChange={(e) => setU(e.target.value)} autoFocus />
        </div>
        <div className="g-login-field">
          <label><Lock size={12} />Senha</label>
          <input type="password" value={p} onChange={(e) => setP(e.target.value)} />
        </div>
        {error && <div className="g-login-error">{error}</div>}
        <button type="submit" className="g-btn primary" style={{ width: "100%", justifyContent: "center" }}>Entrar</button>
      </form>
    </div>
  );
}

/* ============================================================
   ROOT — auth gate with multiple accounts
   ============================================================ */
/* ============================================================
   INITIAL / SEED DATA — usados apenas na primeiríssima vez que o
   backend compartilhado ainda não tem nada salvo. A partir daí,
   tudo abaixo é carregado do servidor e salvo de volta nele (ver
   Root()), independente de qual usuário estiver logado.
   ============================================================ */
const INITIAL_WORK_PACKAGES = [
    { id: "MAN-2026-001", name: "Alinhamento do Eixo do Compressor", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/01", empresa: "Norpem", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-002", name: "Detectores de Gases para Manutenção", discipline: "Hse", group: "Segurança", portCall: "Port Call 23/01", empresa: "Casa Offshore", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-003", name: "Manutenção do Motor do Bote", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/01", empresa: "Sea Services", md: "Sim", rc: "10303580", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-004", name: "Manutenção dos Radar Banda S", discipline: "Marine", group: "Bridge", portCall: "Port Call 23/01", empresa: "Radiomar", md: "Sim", rc: "4600003659", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-005", name: "Desmontagem da bomba do Motor 3", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/01", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-006", name: "PTA para Manutenção do Limitador do Cabo de Aço do Hook Princpal", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/01", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-007", name: "Manutenção no Motor 3 do Guincho Principal", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/01", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-008", name: "Montagem de Andaime – Guincho Principal", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 23/01", empresa: "AASJ", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-009", name: "Manutenção Preditiva Crane TTS", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/01", empresa: "Norpem", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-010", name: "Balsas Salva-Vidas – Retorno", discipline: "Marine", group: "Bridge", portCall: "Port Call 23/01", empresa: "Viking", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-011", name: "Elaboração do PGR", discipline: "Integridade", group: "Documental", portCall: "Port Call 23/01", empresa: "Traume", md: "Sim", rc: "10300888", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-01-23T08:00", end: "2026-01-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-012", name: "Manutenção no Radar Banda X", discipline: "Marine", group: "Bridge", portCall: "Port Call 03/02", empresa: "Radiomar", md: "Sim", rc: "4600003659", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-02-03T08:00", end: "2026-02-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-013", name: "Manutenção do Motor do Bote", discipline: "Mecânica", group: "Engine", portCall: "Port Call 03/02", empresa: "Sea Services", md: "Sim", rc: "10303580", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-02-03T08:00", end: "2026-02-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-014", name: "Certificação Equipamentos da Enfermaria", discipline: "Hse", group: "Segurança", portCall: "Port Call 03/02", empresa: "Measure", md: "Sim", rc: "Contrato", obs: "A empresa não compareceu", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-02-03T08:00", end: "2026-02-03T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-015", name: "Programação Serviços MI Electric - Testes Proteção Maior Confiabilidade dos Sistemas - SWTBs HV & LV - Protection Relay - Transformers HV & LV", discipline: "Elétrica", group: "Electrical", portCall: "Port Call 03/02", empresa: "M&I", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-02-03T08:00", end: "2026-02-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-016", name: "Manutenção de Antena de TV", discipline: "Marine", group: "Bridge", portCall: "Port Call 16/02", empresa: "Salestech", md: "Sim", rc: "10311549", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-02-16T08:00", end: "2026-02-16T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-017", name: "Retorno dos Extintores", discipline: "Hse", group: "Segurança", portCall: "Port Call 16/02", empresa: "Sollax", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-02-16T08:00", end: "2026-02-16T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-018", name: "Certificação dos equipamentos da enfermaria", discipline: "Hse", group: "Segurança", portCall: "Port Call 16/02", empresa: "Measure", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-02-16T08:00", end: "2026-02-16T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-019", name: "Resgatista para Manutenção no HIPAP", discipline: "Marine", group: "Bridge", portCall: "Port Call 16/02", empresa: "Setec", md: "Sim", rc: "10309583", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-02-16T08:00", end: "2026-02-16T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-020", name: "Programação Serviços MI Electric - Testes Proteção Maior Confiabilidade dos Sistemas - SWTBs HV & LV - Protection Relay - Transformers HV & LV", discipline: "Elétrica", group: "Electrical", portCall: "Port Call 03/03", empresa: "M&I", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-021", name: "Certificação de Luvas de Borracha", discipline: "Hse", group: "Segurança", portCall: "Port Call 03/03", empresa: "Measure", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-022", name: "Calibração Equipamentos Enfermaria", discipline: "Hse", group: "Segurança", portCall: "Port Call 03/03", empresa: "Ih Care", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-023", name: "Certificação Anual dos Guindastes e Teste de Carga", discipline: "Marine", group: "Bridge", portCall: "Port Call 03/03", empresa: "Oil States", md: "Sim", rc: "N/A", obs: "A empresa cancelou 1 dia antes do portcall", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-024", name: "Montagem de Andaime para Manutenção do Tugger Winch do Guindaste Offshore", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 03/03", empresa: "Priner", md: "Sim", rc: "10324367", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-025", name: "Resgatista para Radar Banda X", discipline: "Marine", group: "Bridge", portCall: "Port Call 03/03", empresa: "Setec", md: "Sim", rc: "10309582", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-026", name: "Reparo Radar Banda X", discipline: "Marine", group: "Bridge", portCall: "Port Call 03/03", empresa: "Radiomar", md: "Sim", rc: "4600003659", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-027", name: "Reparo do Duto de Ventilação", discipline: "Mecânica", group: "Engine", portCall: "Port Call 03/03", empresa: "Evetec", md: "Sim", rc: "10305941", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-028", name: "Substituição do EPIRB", discipline: "Marine", group: "Bridge", portCall: "Port Call 03/03", empresa: "Ocean Wave", md: "Sim", rc: "N/A", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-029", name: "Reparo das Defensas", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 03/03", empresa: "Evetec", md: "Sim", rc: "10313395", obs: "Só foi possivel realizar de 1 bordo", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-03T08:00", end: "2026-03-03T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-030", name: "Extintor de CO2", discipline: "Marine", group: "Bridge", portCall: "Port Call 17/03", empresa: "Sollax", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-17T08:00", end: "2026-03-17T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-031", name: "Análise de vibração nos BTT, redutora, DGs", discipline: "Mecânica", group: "Engine", portCall: "Port Call 17/03", empresa: "Norpem", md: "Sim", rc: "4263315", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-17T08:00", end: "2026-03-17T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-032", name: "Pintura das Marcações do Calado", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 17/03", empresa: "Evetec", md: "Sim", rc: "10313395", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-17T08:00", end: "2026-03-17T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-033", name: "Reparo das Defensas", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 17/03", empresa: "Evetec", md: "Sim", rc: "10313395", obs: "Só foi possivel realizar de 1 bordo", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-17T08:00", end: "2026-03-17T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-034", name: "Certificação lifting points para ovh DG#1", discipline: "Mecânica", group: "Engine", portCall: "Port Call 31/03", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "Não foi possivel terminar todos os olhais devido a janela operacional", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-31T08:00", end: "2026-03-31T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-035", name: "Instalação de cabo de proteção célula de carga", discipline: "Mecânica", group: "Engine", portCall: "Port Call 31/03", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-31T08:00", end: "2026-03-31T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-036", name: "Inspeção em luminárias", discipline: "Mecânica", group: "Engine", portCall: "Port Call 31/03", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-31T08:00", end: "2026-03-31T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-037", name: "Equalização das pressões das bbs de giro", discipline: "Mecânica", group: "Engine", portCall: "Port Call 31/03", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-31T08:00", end: "2026-03-31T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-038", name: "Troca de motor exaustor de bombordo", discipline: "Mecânica", group: "Engine", portCall: "Port Call 31/03", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-31T08:00", end: "2026-03-31T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-039", name: "Inspeção anual e teste quinquenal - Guindaste BB", discipline: "Marine", group: "Bridge", portCall: "Port Call 31/03", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "Sem autorização do porto, precisa ser realizado na proxima janela", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-31T08:00", end: "2026-03-31T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-040", name: "Teste de carga do sistema auxiliar 10ton", discipline: "Marine", group: "Bridge", portCall: "Port Call 31/03", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "Sem janela para realizar devido a manutenção no guindaste", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-03-31T08:00", end: "2026-03-31T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-041", name: "Calderaria DG#2", discipline: "Mecânica", group: "Engine", portCall: "Port Call 07/04", empresa: "Attech", md: "Sim", rc: "10368488", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-042", name: "Overhaul 20.000 cooling compressor", discipline: "Mecânica", group: "Engine", portCall: "Port Call 07/04", empresa: "Macnor", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-043", name: "Certificação olhais escotilha moon pool", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 07/04", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-044", name: "Inspeção anual e teste quinquenal - Guindaste BB", discipline: "Marine", group: "Bridge", portCall: "Port Call 07/04", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-045", name: "Teste de carga do sistema auxiliar 10 ton", discipline: "Marine", group: "Bridge", portCall: "Port Call 07/04", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "Sem janela para realizar devido a manutenção no guindaste", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-046", name: "Análise qualidade do ar", discipline: "Mecânica", group: "Engine", portCall: "Port Call 07/04", empresa: "Cimartec", md: "Sim", rc: "10327287", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-047", name: "Certificação lifting points para OVH DG#1", discipline: "Mecânica", group: "Engine", portCall: "Port Call 07/04", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-048", name: "Teste Hidrostático PLT", discipline: "Marine", group: "Bridge", portCall: "Port Call 07/04", empresa: "Survitec", md: "Sim", rc: "10324355", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-049", name: "Reparo em suporte dos pinos trava do AHC", discipline: "Mecânica", group: "Engine", portCall: "Port Call 07/04", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-050", name: "Manutenção Prev Serpentina heat exchanger BE", discipline: "Mecânica", group: "Engine", portCall: "Port Call 07/04", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-051", name: "Troca de Lub-oil cilindros do AHC", discipline: "Mecânica", group: "Engine", portCall: "Port Call 07/04", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-052", name: "Retirada de folga (backlash) do guindaste TTS", discipline: "Mecânica", group: "Engine", portCall: "Port Call 07/04", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-07T08:00", end: "2026-04-07T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-053", name: "Limpeza de dutos anual", discipline: "Mecânica", group: "Engine", portCall: "Port Call 29/04", empresa: "Cimartec", md: "Sim", rc: "10327287", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-29T08:00", end: "2026-04-29T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-054", name: "Limpeza tanques de água", discipline: "Mecânica", group: "Engine", portCall: "Port Call 29/04", empresa: "Tankclean", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-29T08:00", end: "2026-04-29T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-055", name: "Estudo trocador de calor", discipline: "Mecânica", group: "Engine", portCall: "Port Call 29/04", empresa: "Prismar", md: "Sim", rc: "N/A", obs: "Empresa não compareceu", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-29T08:00", end: "2026-04-29T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-056", name: "Pintura das Marcações do Calado", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 29/04", empresa: "Evetec", md: "Sim", rc: "10313395", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-29T08:00", end: "2026-04-29T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-057", name: "Troca de olhais", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 29/04", empresa: "Evetec", md: "Sim", rc: "N/A", obs: "Empresa não fabricou o olhal devido a pagamento", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-29T08:00", end: "2026-04-29T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-058", name: "NRs12/13 e 35", discipline: "Integridade", group: "Documental", portCall: "Port Call 29/04", empresa: "Tekee", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-29T08:00", end: "2026-04-29T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-059", name: "Teste de carga / Recertificação", discipline: "Marine", group: "Bridge", portCall: "Port Call 29/04", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "Sem janela para realizar devido a manutenção no guindaste", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-29T08:00", end: "2026-04-29T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-060", name: "Inspeção Compressor do Chiller", discipline: "Mecânica", group: "Engine", portCall: "Port Call 29/04", empresa: "Macnor", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-04-29T08:00", end: "2026-04-29T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-061", name: "Troca do Acoplamento Guindaste BE", discipline: "Marine", group: "Bridge", portCall: "Port Call 13/05", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "Acomplamento não chegou a tempo", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-062", name: "Limpeza tanques 62S / 13P / 14S", discipline: "Mecânica", group: "Engine", portCall: "Port Call 13/05", empresa: "Tankclean", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-063", name: "DP Annual Trial", discipline: "Marine", group: "Bridge", portCall: "Port Call 13/05", empresa: "All marine", md: "Sim", rc: "10338906", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-064", name: "Teste operacional do guindaste", discipline: "Mecânica", group: "Engine", portCall: "Port Call 13/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-065", name: "Preventiva em limit switch do guincho princ", discipline: "Mecânica", group: "Engine", portCall: "Port Call 13/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-066", name: "Instalar cabo de aço proteção célula de carga", discipline: "Mecânica", group: "Engine", portCall: "Port Call 13/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-067", name: "Reparo da chapa de desgaste do cabo AHC", discipline: "Mecânica", group: "Engine", portCall: "Port Call 13/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-068", name: "Troca de cabo piloto Servo válvula", discipline: "Mecânica", group: "Engine", portCall: "Port Call 13/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-069", name: "Reparo em olhais da trava do trolley do AHC", discipline: "Mecânica", group: "Engine", portCall: "Port Call 13/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-070", name: "Instalação do Motor - Trocador calor boreste", discipline: "Mecânica", group: "Engine", portCall: "Port Call 13/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-071", name: "Troca de rede - Trocador de calor", discipline: "Mecânica", group: "Engine", portCall: "Port Call 13/05", empresa: "Attech", md: "Sim", rc: "10378377", obs: "Rede não chegou a tempo para troca definitiva", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-072", name: "Teste de carga / Recertificação", discipline: "Marine", group: "Bridge", portCall: "Port Call 13/05", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-13T08:00", end: "2026-05-13T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-073", name: "Troca do Acoplamento Guindaste BE", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-074", name: "Calibração da Célula de 15 ppm", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "Engeprime", md: "Sim", rc: "10352768", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-075", name: "Adequação de Desenhos", discipline: "Integridade", group: "Documental", portCall: "Port Call 21/05", empresa: "Gran", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-076", name: "Instalação DGPS / Network", discipline: "Marine", group: "Bridge", portCall: "Port Call 21/05", empresa: "KM", md: "Não", rc: "Contrato", obs: "A KM não tinha disponibilidade", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-077", name: "Reaperto do tubo do guincho auxiliar", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-078", name: "Inspeção do sistema hidráulico", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-079", name: "Purga no sistema hidráulico do AHC", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-080", name: "Drenagem de óleo do PTO 2", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-081", name: "Engraxamento da cremalheira de giro", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-082", name: "Survey instalação sistema monitoramento", discipline: "Marine", group: "Bridge", portCall: "Port Call 21/05", empresa: "TWS", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-083", name: "Troca de rede - Trocador de calor", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "Attech", md: "Sim", rc: "10378377", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-084", name: "Estudo trocador de calor máquina", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "Autocomp", md: "Sim", rc: "10352766", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-085", name: "Termografia", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "Autocomp", md: "Sim", rc: "10352762", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-086", name: "Overhaul dos Compressores de Ar de Partida", discipline: "Mecânica", group: "Engine", portCall: "Port Call 21/05", empresa: "Autocomp", md: "Sim", rc: "10349772", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-087", name: "NRs 12 / 13 e 35", discipline: "Integridade", group: "Documental", portCall: "Port Call 21/05", empresa: "Tekee", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-088", name: "Teste de carga SB provision crane", discipline: "Marine", group: "Bridge", portCall: "Port Call 21/05", empresa: "Highbras", md: "Sim", rc: "10309413", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-05-21T08:00", end: "2026-05-21T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-089", name: "Tratamento e Pintura de corrosão do Guindaste TTS", discipline: "Marine", group: "Bridge", portCall: "Port Call 23/06", empresa: "Attech", md: "Sim", rc: "10368494", obs: "Devido a indisponibilidade do guindaste", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-090", name: "Reparo no DG#3", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "Wartsila", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-091", name: "Troca/Reparo das Válvulas Reguladores do Sistema de Ar", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "Autocomp", md: "Sim", rc: "10386081", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-092", name: "Calibração e Certificação dos Manômetros", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "Autocomp", md: "Sim", rc: "10395954", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-093", name: "Reaperto em Parafusos da Estrutura do AHC", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-094", name: "Reaperto em Parafusos do Sistema de Giro", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-095", name: "Medição de Folga em Rolamento de Giro", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-096", name: "Certificação de PFOS nos LGS e Inspeção semestral do CO2", discipline: "Hse", group: "Segurança", portCall: "Port Call 23/06", empresa: "Sollax", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-097", name: "Verificação do Sistema Chiller", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "Macnor", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-098", name: "Redundância da bomba do ROV", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "Attech", md: "Sim", rc: "10368489", obs: "Aguardando a chegada da rede", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-099", name: "Troca de Rolamentos dos Motores Elétricos dos Compressores", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "Autocomp", md: "Sim", rc: "10368491", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-100", name: "Instalação DGPS", discipline: "Marine", group: "Bridge", portCall: "Port Call 23/06", empresa: "KM", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-101", name: "NRs 12/13 e 35 - Realizado NR-12", discipline: "Integridade", group: "Documental", portCall: "Port Call 23/06", empresa: "Tekee", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-102", name: "Adequações de drops", discipline: "Marine", group: "Bridge", portCall: "Port Call 23/06", empresa: "Attech", md: "Sim", rc: "10406914", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-103", name: "Teste Geral", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-104", name: "Preparar Base do data Logger", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-105", name: "Substituir Filtros de Retorno de Tanque", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-106", name: "Verificação em Painel Elétrico da Cabine do Operador", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-107", name: "Instalação da Parte Elétrica do monitoramento do Bloco AHC", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-108", name: "Aperto de Parafusos dos Cilindros do AHC e Fechamento do Carro", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-109", name: "Instalação de Mangueiras dos Latches e Reparo em Olhal", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-110", name: "Implementação de Suporte e Proteções AHC", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-111", name: "Delineamento de Mangueiras", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-112", name: "Substituição do Radiador do Trocador de Calor Boreste", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "FAC", md: "Não", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-113", name: "Instalação do Compressor de Provisões", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/06", empresa: "Macnor", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-06-23T08:00", end: "2026-06-23T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-114", name: "Fire Fighting Equipment Annual", discipline: "Hse", group: "Segurança", portCall: "Port Call 06/07", empresa: "Sollax", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-115", name: "Substituição do Cabo de Aço do Tugger Winch no TTS", discipline: "Marine", group: "Bridge", portCall: "Port Call 06/07", empresa: "Attech", md: "Sim", rc: "10406915", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-116", name: "Fabricação e Instalação do Guarda-Corpo do Guindaste", discipline: "Marine", group: "Bridge", portCall: "Port Call 06/07", empresa: "Attech", md: "Sim", rc: "10406916", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-117", name: "Pintura e Tratamento do Guindaste TTS", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 06/07", empresa: "Attech", md: "Sim", rc: "10368494", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-118", name: "Reparo das Defensas", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 06/07", empresa: "Attech", md: "Sim", rc: "10406912", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-119", name: "Balanço de Carga", discipline: "Elétrica", group: "Electrical", portCall: "Port Call 06/07", empresa: "M&I", md: "Sim", rc: "10313396", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-120", name: "Calibração e Certificação dos Flowmeters", discipline: "Mecânica", group: "Engine", portCall: "Port Call 06/07", empresa: "Autocomp", md: "Sim", rc: "10369462", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-121", name: "Redundância da Bomba do ROV", discipline: "Mecânica", group: "Engine", portCall: "Port Call 06/07", empresa: "Attech", md: "Sim", rc: "10368489", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-122", name: "NRs 12/13 e 35", discipline: "Integridade", group: "Documental", portCall: "Port Call 06/07", empresa: "Tekee", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-123", name: "Adequação de Drops", discipline: "Marine", group: "Bridge", portCall: "Port Call 06/07", empresa: "Attech", md: "Sim", rc: "10406914", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-124", name: "Inspeção na Propulsão", discipline: "Mecânica", group: "Engine", portCall: "Port Call 06/07", empresa: "Wartsila", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-07-06T08:00", end: "2026-07-06T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-125", name: "NRs 12/13 e 35", discipline: "Integridade", group: "Documental", portCall: "Port Call 04/08", empresa: "Tekee", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-126", name: "Fire Fighting Equipment Annual", discipline: "Hse", group: "Segurança", portCall: "Port Call 04/08", empresa: "Sollax", md: "Sim", rc: "Contrato", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-127", name: "Troca e Calibração das Válvulas Reguladoras", discipline: "Mecânica", group: "Engine", portCall: "Port Call 04/08", empresa: "Autocomp", md: "Sim", rc: "10386081", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-128", name: "Teste de Carga dos Olhais do Berço do Guindaste", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 04/08", empresa: "Highbras", md: "Sim", rc: "10432166", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-129", name: "Fabricação e Instalação dos Olhais das Defensas", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 04/08", empresa: "Attech", md: "Sim", rc: "10432124", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-130", name: "Estudo dos Quadros Elétricos e AVR", discipline: "Mecânica", group: "Engine", portCall: "Port Call 04/08", empresa: "United Power", md: "Sim", rc: "10432173", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-131", name: "Inspeção dos CJCs", discipline: "Mecânica", group: "Engine", portCall: "Port Call 04/08", empresa: "United Power", md: "Sim", rc: "10432174", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-132", name: "Equipe de Resgate e Manutenção no HIPAP", discipline: "Mecânica", group: "Engine", portCall: "Port Call 04/08", empresa: "Attech", md: "Sim", rc: "10432150", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-133", name: "Equipe de Irata para Teste de Carga dos olhais", discipline: "Marine", group: "Bridge", portCall: "Port Call 04/08", empresa: "Attech", md: "Não", rc: "Regularização", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-134", name: "Overhaul Motores elétricos dos Compressores", discipline: "Mecânica", group: "Engine", portCall: "Port Call 04/08", empresa: "United Power", md: "Não", rc: "Regularização", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-135", name: "Fabricação e Instalação dos Olhais", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 04/08", empresa: "Attech", md: "Sim", rc: "10386083", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-136", name: "Anual de AIS EPIRB SART GMDSS", discipline: "Marine", group: "Bridge", portCall: "Port Call 04/08", empresa: "CWDMC", md: "Sim", rc: "10431860", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-137", name: "Anual do VDR", discipline: "Marine", group: "Bridge", portCall: "Port Call 04/08", empresa: "Radio Holland", md: "Sim", rc: "10431861", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-138", name: "Calibração da Bússola Magnética", discipline: "Marine", group: "Bridge", portCall: "Port Call 04/08", empresa: "Gyromarsat", md: "Sim", rc: "10431862", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-139", name: "Anual do Bote Resgate e Davit", discipline: "Marine", group: "Bridge", portCall: "Port Call 04/08", empresa: "Mapamar", md: "Sim", rc: "10432177", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-140", name: "Limpeza e Pintura das Marcas de Calado", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 04/08", empresa: "Attech", md: "Sim", rc: "10432142", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-04T08:00", end: "2026-08-04T17:00", status: "Concluído", progress: 100 },
    { id: "MAN-2026-141", name: "Adequação de Drops", discipline: "Marine", group: "Bridge", portCall: "Port Call 23/08", empresa: "Attech", md: "Sim", rc: "10406914", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-142", name: "Pintura e Tratamento do Guindaste TTS", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 23/08", empresa: "Attech", md: "Sim", rc: "10368494", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-143", name: "Tratamento e Reforma da Tampa Superior do Moonpool", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 23/08", empresa: "Attech", md: "Sim", rc: "10432115", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-144", name: "Fabricação e Instalação do Guarda-Corpo na lança do Guindaste", discipline: "Mecânica", group: "Guindaste", portCall: "Port Call 23/08", empresa: "Attech", md: "Sim", rc: "10432145", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-145", name: "Instalação do CJC nos Thrusters 2 e 5", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/08", empresa: "United Power", md: "Não", rc: "", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-146", name: "Mergulho para inspeção dos Thrusters", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/08", empresa: "Northsub", md: "Não", rc: "", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-147", name: "Inspeção dos Olhais da Lança do Guindaste", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 23/08", empresa: "Highbras", md: "Não", rc: "", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-148", name: "Solda do olhal do Convés", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call 23/08", empresa: "Attech", md: "Sim", rc: "10386083", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Em andamento", progress: 50 },
    { id: "MAN-2026-149", name: "Instalação da Bomba de Lastro", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/08", empresa: "Attech", md: "Não", rc: "4324509", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-150", name: "Overhaul dos motores elétricos da bomba FO Feed pump 1 e 2", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/08", empresa: "United Power", md: "Não", rc: "", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-151", name: "Diagrama Unifilares", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/08", empresa: "United Power", md: "Não", rc: "", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-152", name: "Balanço de Potência", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/08", empresa: "United Power", md: "Não", rc: "", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-153", name: "Estudo de Seletividade", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/08", empresa: "United Power", md: "Não", rc: "", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-154", name: "Desnhos atualizados dos Quadros Elétricos", discipline: "Mecânica", group: "Engine", portCall: "Port Call 23/08", empresa: "United Power", md: "Não", rc: "", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-23T08:00", end: "2026-08-23T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-155", name: "Calibração e Certificação dos Multimetros", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "Measure", md: "Sim", rc: "10410632", obs: "Sem retorno da equipe de Suprimentos. Cotação encerrada", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-156", name: "Instalação da escada do guindaste", discipline: "Hull & Structure", group: "Deck", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10413186", obs: "Sem retorno da equipe de Suprimentos. Cotação encerrada", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-157", name: "Estudo de Instalação de Refrigeração Quadros Eletricos", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10317742", obs: "Sem retorno da equipe de Suprimentos. Cotação encerrada", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-158", name: "Projeto de Construção de Almoxarifado dentro de um silo", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10413164", obs: "Sem retorno da equipe de Suprimentos. Cotação encerrada", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-159", name: "Fabricação e instalação das Redes dos Resfriadores nº 1 e nº 2 CuNiFe", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10413165", obs: "Sem retorno da equipe de Suprimentos. Cotação encerrada", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-160", name: "Comissionamento e Teste Operacional do Sistema BULK", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10413167", obs: "Sem retorno da equipe de Suprimentos. Cotação encerrada", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-161", name: "Certificação dos Olhais da Praça de Máquinas e Ponte Rolante", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10413179", obs: "Sem retorno da equipe de Suprimentos. Cotação encerrada", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-162", name: "Substituição da Rede de Bilge da Sala do Azimutal", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10413183", obs: "Sem retorno da equipe de Suprimentos. Cotação encerrada", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-163", name: "Recuperação e Comissionamento das Bombas do Sistema MUD", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10413185", obs: "Sem retorno da equipe de Suprimentos. Cotação encerrada", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-164", name: "Delineamento de Mangueiras", discipline: "Mecânica", group: "Outros", portCall: "Port Call — sem data definida", empresa: "", md: "Não", rc: "", obs: "", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-165", name: "Certificação dos Olhais da Praça de Máquinas", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10432162", obs: "Aguardando Suprimentos", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-166", name: "Substituição da Gate Valve do Sistema HiPAP", discipline: "Mecânica", group: "Engine", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10432169", obs: "Aguardando Suprimentos", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
    { id: "MAN-2026-167", name: "Reparo do Detector Multigás Modelo 4X", discipline: "Marine", group: "Bridge", portCall: "Port Call — sem data definida", empresa: "", md: "Sim", rc: "10432170", obs: "Aguardando Suprimentos", budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-30T08:00", end: "2026-08-30T17:00", status: "Não iniciado", progress: 0 },
];
const INITIAL_MATERIALS = [
    { id: "MAT-0084", wp: "", tmMaster: "TM-04521", departamento: "Manutenção", sap: "10098231", descricao: "Seal Kit - Bow Thruster", quantidade: 2, priority: "Crítica", dataSolicitacao: "2026-08-18", dataNecessidade: "2026-09-05", reserva: "RES-3321", rc: "10410632", po: "4600012345", linhaPo: "10", valor: 18500, eta: "2026-09-03", obs: "", dataRecebimento: "", status: "Em trânsito" },
    { id: "MAT-0085", wp: "", tmMaster: "TM-04521", departamento: "Manutenção", sap: "10098232", descricao: "O-Ring Set", quantidade: 6, priority: "Alta", dataSolicitacao: "2026-08-19", dataNecessidade: "2026-09-05", reserva: "RES-3322", rc: "10410633", po: "", linhaPo: "", valor: 2400, eta: "", obs: "Aguardando cotação", dataRecebimento: "", status: "Cotação" },
    { id: "MAT-0090", wp: "", tmMaster: "TM-05011", departamento: "Elétrica", sap: "10098240", descricao: "Bobina Estator", quantidade: 1, priority: "Crítica", dataSolicitacao: "2026-08-15", dataNecessidade: "2026-09-08", reserva: "RES-3340", rc: "10410650", po: "4600012390", linhaPo: "20", valor: 41200, eta: "2026-09-06", obs: "", dataRecebimento: "", status: "Em fabricação" },
    { id: "MAT-0091", wp: "", tmMaster: "TM-05011", departamento: "Elétrica", sap: "10098241", descricao: "Placa AVR", quantidade: 3, priority: "Média", dataSolicitacao: "2026-08-10", dataNecessidade: "2026-08-30", reserva: "RES-3341", rc: "10410651", po: "4600012391", linhaPo: "10", valor: 6800, eta: "2026-08-28", obs: "", dataRecebimento: "2026-08-27", status: "Recebido" },
    { id: "MAT-0092", wp: "", tmMaster: "TM-06120", departamento: "Mecânica", sap: "10098255", descricao: "Rotor Kit", quantidade: 1, priority: "Alta", dataSolicitacao: "2026-08-21", dataNecessidade: "2026-09-01", reserva: "RES-3355", rc: "10410670", po: "4600012410", linhaPo: "10", valor: 27300, eta: "", obs: "Fornecedor confirmou pedido", dataRecebimento: "", status: "PO emitida" },
];
const INITIAL_PAYMENTS = [
    { id: "PAY-001", service: "Manutenção de Defensas", po: "450001245", poValue: 132000, nf: "9847", nfValue: 130800, issue: "2026-08-30", due: "2026-09-29", status: "Pagamento programado" },
    { id: "PAY-002", service: "Overhaul Bow Thruster #2", po: "450001300", poValue: 270000, nf: "5521", nfValue: 270000, issue: "2026-09-05", due: "2026-10-05", status: "NF validada" },
    { id: "PAY-003", service: "AVR Upgrade", po: "450001350", poValue: 190000, nf: "", nfValue: 0, issue: "", due: "", status: "Serviço executado" },
    { id: "PAY-004", service: "Overhaul SW Pump", po: "450001410", poValue: 165000, nf: "771", nfValue: 165000, issue: "2026-09-04", due: "2026-10-04", status: "Pago" },
    { id: "PAY-005", service: "Overhaul Motor Elétrico Thruster #4", po: "450001420", poValue: 420000, nf: "", nfValue: 0, issue: "", due: "", status: "PO emitida" },
];
const INITIAL_SERVICE_INVOICES = [
    { id: "INV-001", date: "2026-04-07", assunto: "Reparo da Rede - DG2", empresa: "Attech", md: "Sim", mdSentDate: "2026-04-07", diffDays: 0, daysOpenTotal: 121, rc: "10368488", serviceStatus: "Fechado", poContrato: "4500161492/4292564", medicao: "4306119", valorTotal: 94518.9, saldoPo: 0.0, obs: "Pagamento Programado para 06/08", statusPagamento: "Pago", dataPagamento: "2026-08-06" },
    { id: "INV-002", date: "2026-05-26", assunto: "Troca de Rede - DG2", empresa: "Attech", md: "Sim", mdSentDate: "2026-05-20", diffDays: 6, daysOpenTotal: 78, rc: "10378377", serviceStatus: "Fechado", poContrato: "4300387", medicao: "4306029", valorTotal: 80995.0, saldoPo: 0.0, obs: "Pagamento Programado para 06/08", statusPagamento: "Pago", dataPagamento: "2026-08-06" },
    { id: "INV-003", date: "2026-06-23", assunto: "Adequação da bomba do ROV", empresa: "Attech", md: "Sim", mdSentDate: "2026-06-02", diffDays: 21, daysOpenTotal: 85, rc: "10368489", serviceStatus: "Fechado", poContrato: "4324509", medicao: "4340561", valorTotal: 152083.58, saldoPo: 0.0, obs: "Medição 4323106 excluída. Aguardando aprovação do Alexandre Rosa", statusPagamento: "Aprovação Pendente", dataPagamento: "" },
    { id: "INV-004", date: "2026-06-23", assunto: "Fabricação e Instalação dos Olhais de Conves", empresa: "Attech", md: "Sim", mdSentDate: "2026-06-02", diffDays: 21, daysOpenTotal: 85, rc: "10386083", serviceStatus: "Fechado", poContrato: "4313001", medicao: "4336386", valorTotal: 81912.0, saldoPo: 0.0, obs: "Aguardando aprovação do Alexandre Rosa.", statusPagamento: "Aprovação Pendente", dataPagamento: "" },
    { id: "INV-005", date: "2026-07-09", assunto: "Fabricação e Instalação do Guarda Corpo", empresa: "Attech", md: "Sim", mdSentDate: "2026-06-02", diffDays: 37, daysOpenTotal: 85, rc: "10440171", serviceStatus: "Fechado", poContrato: "4500167162", medicao: "", valorTotal: 79312.0, saldoPo: 0.0, obs: "Pagamento Programado para 17/09", statusPagamento: "Pagamento Programado", dataPagamento: "" },
    { id: "INV-006", date: "2026-07-09", assunto: "Substituição do Cabo de Aço do Tugger Winch no TTS", empresa: "Attech", md: "Sim", mdSentDate: "2026-06-02", diffDays: 37, daysOpenTotal: 85, rc: "10406915", serviceStatus: "Fechado", poContrato: "4324509", medicao: "4340560", valorTotal: 69300.0, saldoPo: 0.0, obs: "Medição 4332488 excluída. Aguardando aprovação do Alexandre Rosa", statusPagamento: "Aprovação Pendente", dataPagamento: "" },
    { id: "INV-007", date: "2026-07-09", assunto: "Reparo das Defensas", empresa: "Attech", md: "Sim", mdSentDate: "2026-06-02", diffDays: 37, daysOpenTotal: 85, rc: "10406912", serviceStatus: "Fechado", poContrato: "4324509", medicao: "", valorTotal: 152400.0, saldoPo: 0.0, obs: "Contrato aprovado. Aguardando Medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-008", date: "2026-07-09", assunto: "Adequação de Drops", empresa: "Attech", md: "Sim", mdSentDate: "2026-06-02", diffDays: 37, daysOpenTotal: 85, rc: "10406914", serviceStatus: "Aberto", poContrato: "4324509", medicao: "", valorTotal: 0, saldoPo: 0.0, obs: "Contrato aprovado. Aguardando Medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-009", date: "2026-07-09", assunto: "Pintura e Tratamento do Guindaste TTS", empresa: "Attech", md: "Sim", mdSentDate: "2026-06-02", diffDays: 37, daysOpenTotal: 85, rc: "10368494", serviceStatus: "Aberto", poContrato: "4310832", medicao: "", valorTotal: 36860.0, saldoPo: 0.0, obs: "Contrato aprovado no ERP. Aguardando medição do fornecedor. Aguardando finalizar o serviço para envio da medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-010", date: "", assunto: "Fabricação e instalação das Redes dos Resfriadores nº 1 e nº 2 CuNiFe", empresa: "Attech", md: "Sim", mdSentDate: "2026-07-07", diffDays: -46210, daysOpenTotal: 50, rc: "10413165", serviceStatus: "Aberto", poContrato: "4324509", medicao: "", valorTotal: 137829.7, saldoPo: 0.0, obs: "Contrato aprovado. Aguardando Medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-011", date: "", assunto: "Instalação da escada do guindaste", empresa: "Attech", md: "Sim", mdSentDate: "2026-07-08", diffDays: -46211, daysOpenTotal: 49, rc: "10413186", serviceStatus: "Aberto", poContrato: "4327266", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Contrato aprovado. Aguardando Medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-012", date: "2026-08-04", assunto: "Fabricação e Instalação dos Olhais das Defensas", empresa: "Attech", md: "Sim", mdSentDate: "2026-07-27", diffDays: 8, daysOpenTotal: 30, rc: "10432124", serviceStatus: "Aberto", poContrato: "4324509", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Contrato aprovado. Aguardando Medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-013", date: "2026-08-04", assunto: "Limpeza e Pintura das Marcas de Calado", empresa: "Attech", md: "Sim", mdSentDate: "2026-07-27", diffDays: 8, daysOpenTotal: 30, rc: "10432142", serviceStatus: "Fechado", poContrato: "4324509", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Contrato aprovado. Aguardando Medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-014", date: "2026-08-04", assunto: "Adequação de Segurança do Guindaste", empresa: "Attech", md: "Sim", mdSentDate: "2026-07-27", diffDays: 8, daysOpenTotal: 30, rc: "10432145", serviceStatus: "Fechado", poContrato: "4324509", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Contrato aprovado. Aguardando Medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-015", date: "2026-08-04", assunto: "Suporte Resgatista IRATA para Manutenção no HiPAP", empresa: "Attech", md: "Sim", mdSentDate: "2026-07-27", diffDays: 8, daysOpenTotal: 30, rc: "10432150", serviceStatus: "Fechado", poContrato: "4324509", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Contrato aprovado. Aguardando Medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-016", date: "2026-05-26", assunto: "Overhaul 10.000 - Compressores de Ar de Partida", empresa: "Autocomp", md: "Sim", mdSentDate: "2026-05-01", diffDays: 25, daysOpenTotal: 76, rc: "10349772", serviceStatus: "Fechado", poContrato: "4292601", medicao: "4293869", valorTotal: 69976.0, saldoPo: 0.0, obs: "PO 4500158451. Pago 16.07.2026", statusPagamento: "Pago", dataPagamento: "2026-07-16" },
    { id: "INV-017", date: "2026-05-26", assunto: "Inspeção e Estudo - Trocador de Calor DGs", empresa: "Autocomp", md: "Sim", mdSentDate: "2026-05-01", diffDays: 25, daysOpenTotal: 76, rc: "10352766", serviceStatus: "Fechado", poContrato: "4292837", medicao: "4293854", valorTotal: 14400.0, saldoPo: 0.0, obs: "PO 4500158456. Pago 16.07.2026", statusPagamento: "Pago", dataPagamento: "2026-07-16" },
    { id: "INV-018", date: "2026-05-26", assunto: "Termografia - Equipamentos Críticos da Máquina", empresa: "Autocomp", md: "Sim", mdSentDate: "2026-05-01", diffDays: 25, daysOpenTotal: 117, rc: "10352762", serviceStatus: "Fechado", poContrato: "4302402", medicao: "4310201", valorTotal: 177216.0, saldoPo: 0.0, obs: "Pagamento Programado para 13/08", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-019", date: "2026-06-23", assunto: "Troca de Rolamentos dos Motores Elétricos dos Compressores", empresa: "Autocomp", md: "Sim", mdSentDate: "2026-06-02", diffDays: 21, daysOpenTotal: 65, rc: "10368491", serviceStatus: "Fechado", poContrato: "4302419", medicao: "4306409", valorTotal: 46973.0, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "2026-08-06" },
    { id: "INV-020", date: "2026-06-23", assunto: "Calibração e Certificação de Flowmeters", empresa: "Autocomp", md: "Sim", mdSentDate: "2026-05-28", diffDays: 26, daysOpenTotal: 90, rc: "10369462", serviceStatus: "Aberto", poContrato: "4300401", medicao: "", valorTotal: 22234.0, saldoPo: 0.0, obs: "Serviço não realizado, contrato cancelado", statusPagamento: "Cancelado", dataPagamento: "" },
    { id: "INV-021", date: "2026-06-23", assunto: "Calibração e Certificação de Manômetros", empresa: "Autocomp", md: "Sim", mdSentDate: "2026-05-28", diffDays: 26, daysOpenTotal: 90, rc: "10395954", serviceStatus: "Fechado", poContrato: "4315130", medicao: "4323952", valorTotal: 194595.28, saldoPo: 0.0, obs: "PO 4500171317. Aprovação Pendente de Bruno Tamiozo", statusPagamento: "Aprovação Pendente", dataPagamento: "" },
    { id: "INV-022", date: "2026-06-23", assunto: "Troca e Instalação das Válvulas Reguladoras do Sistema de Ar", empresa: "Autocomp", md: "Sim", mdSentDate: "2026-05-28", diffDays: 26, daysOpenTotal: 90, rc: "10386081", serviceStatus: "Fechado", poContrato: "4318114", medicao: "", valorTotal: 77588.0, saldoPo: 0.0, obs: "Foi Aprovado o contrato no ERP. Registro(s) SAP:5438 / Contrato:4600004652", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-023", date: "2026-04-07", assunto: "Análise Qualidade do Ar", empresa: "Cimartec", md: "Sim", mdSentDate: "2026-04-02", diffDays: 5, daysOpenTotal: 146, rc: "10327287", serviceStatus: "Fechado", poContrato: "4270312", medicao: "4279948", valorTotal: 7130.0, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-024", date: "2026-04-30", assunto: "Limpeza de Dutos Anual", empresa: "Cimartec", md: "Sim", mdSentDate: "2026-04-03", diffDays: 27, daysOpenTotal: 145, rc: "10327287", serviceStatus: "Fechado", poContrato: "4274069", medicao: "4287252", valorTotal: 7050.0, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-025", date: "2026-08-07", assunto: "Anual de AIS EPIRB SART GMDSS", empresa: "CWDMC", md: "Sim", mdSentDate: "2026-07-27", diffDays: 11, daysOpenTotal: 30, rc: "10431860", serviceStatus: "Fechado", poContrato: "4326240", medicao: "4339122", valorTotal: 5700.0, saldoPo: 0.0, obs: "Aguardando Aprovação do Alexandre Rosa.", statusPagamento: "Aprovação Pendente", dataPagamento: "" },
    { id: "INV-026", date: "2026-05-26", assunto: "Calibração Anual da Célula de 15 ppm", empresa: "Engeprime", md: "Sim", mdSentDate: "2026-05-20", diffDays: 6, daysOpenTotal: 98, rc: "10352768", serviceStatus: "Fechado", poContrato: "4292440", medicao: "4302771", valorTotal: 15400.67, saldoPo: 0.0, obs: "Pagamento Programado para 27/08", statusPagamento: "Pagamento Programado", dataPagamento: "" },
    { id: "INV-027", date: "2026-03-03", assunto: "Reparo das Defensas BE", empresa: "Evetec", md: "Sim", mdSentDate: "2026-02-17", diffDays: 14, daysOpenTotal: 190, rc: "10313395", serviceStatus: "Fechado", poContrato: "4264059", medicao: "4277252", valorTotal: 5752.0, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-028", date: "2026-03-03", assunto: "Reparo do Duto de Ventilação", empresa: "Evetec", md: "Sim", mdSentDate: "2026-01-26", diffDays: 36, daysOpenTotal: 193, rc: "10305941", serviceStatus: "Fechado", poContrato: "4264059", medicao: "4277225", valorTotal: 22544.0, saldoPo: 0.0, obs: "Pagamento Programado para 13/08", statusPagamento: "Pago", dataPagamento: "2026-08-07" },
    { id: "INV-029", date: "2026-04-30", assunto: "Pintura Calado", empresa: "Evetec", md: "Sim", mdSentDate: "2026-03-16", diffDays: 45, daysOpenTotal: 144, rc: "10313395", serviceStatus: "Fechado", poContrato: "4264059", medicao: "4277240", valorTotal: 28984.4, saldoPo: 0.0, obs: "Pagamento Programado para 13/08", statusPagamento: "Pago", dataPagamento: "2026-08-07" },
    { id: "INV-030", date: "2026-08-04", assunto: "Calibração da Bússola Magnética", empresa: "Gyromarsat", md: "Sim", mdSentDate: "2026-07-27", diffDays: 8, daysOpenTotal: 30, rc: "", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 12076.76, saldoPo: 0, obs: "Aguardando regularização de contrato", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-031", date: "2026-03-31", assunto: "Manutenção no Guindaste", empresa: "HighBras", md: "Sim", mdSentDate: "2026-03-06", diffDays: 25, daysOpenTotal: 160, rc: "10309413", serviceStatus: "Fechado", poContrato: "4500162889", medicao: "4303662", valorTotal: 74050.0, saldoPo: 0.0, obs: "Pagamento Programado para 20/08", statusPagamento: "Pago", dataPagamento: "2026-08-13" },
    { id: "INV-032", date: "2026-08-04", assunto: "Teste de Carga dos Olhais do Berço", empresa: "HighBras", md: "Sim", mdSentDate: "2026-07-27", diffDays: 8, daysOpenTotal: 30, rc: "10452300", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 8450.0, saldoPo: 0, obs: "Aguardando RC de regularização", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-033", date: "2026-02-03", assunto: "M&I Eletric - Painel, Disjuntor e Transformador", empresa: "M&I", md: "Sim", mdSentDate: "2026-02-03", diffDays: 0, daysOpenTotal: 204, rc: "Contrato", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 542146.0, saldoPo: 542146.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-034", date: "2026-03-17", assunto: "Sistema de Distribuição de Carga", empresa: "M&I", md: "Sim", mdSentDate: "2026-02-20", diffDays: 25, daysOpenTotal: 187, rc: "10313396", serviceStatus: "Fechado", poContrato: "4306809", medicao: "", valorTotal: 63380.0, saldoPo: 0.0, obs: "Será negociado o valor. Serviço não foi feito por eles e sim pela United Power. E-mail: [EXT] RES: [EXTERNO] ENC: Weekly meeting - Genesis comissioning", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-035", date: "2026-04-07", assunto: "Serviços de Manutenção de Equipamentos - BRUNVOLL", empresa: "Macnor", md: "Sim", mdSentDate: "2026-03-03", diffDays: 35, daysOpenTotal: 176, rc: "10327301", serviceStatus: "Fechado", poContrato: "4088695", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Aguardando Medição. E-mail enviado ao fornecedor para saber a situação atual.  Valor total disponível no contrato 1338301,37", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-036", date: "2026-06-23", assunto: "Instalação do Compressor de Provisões", empresa: "Macnor", md: "Sim", mdSentDate: "2026-05-11", diffDays: 43, daysOpenTotal: 107, rc: "10369468", serviceStatus: "Aberto", poContrato: "4088695", medicao: "", valorTotal: 0, saldoPo: 0, obs: "E-mail enviado. Solicitado a inclusão no mesmo contrato Valor total disponível no contrato R$ 1.338.301,37", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-037", date: "2026-08-07", assunto: "Anual do Bote Resgate e Davit", empresa: "Mapamar", md: "Sim", mdSentDate: "2026-07-27", diffDays: 11, daysOpenTotal: 30, rc: "10432177", serviceStatus: "Fechado", poContrato: "4500168422", medicao: "", valorTotal: 18100.0, saldoPo: 0.0, obs: "Acompanhar e-mail:RES: RES: RES: [EXT] RE: Solicitação de Orçamento – Inspeção Anual do Rescue Boat e Rescue Boat Davit", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-038", date: "2026-02-03", assunto: "Certificação Equipamentos da Enfermaria", empresa: "Measure", md: "Sim", mdSentDate: "2026-01-26", diffDays: 8, daysOpenTotal: 212, rc: "10305932", serviceStatus: "Fechado", poContrato: "4218799", medicao: "4263491", valorTotal: 12750.0, saldoPo: 0.0, obs: "Esse pagamento contempla:Certificação Equipamentos da Enfermaria e Certificação de Luvas de Borracha.", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-039", date: "", assunto: "Calibração e Certificação dos Multimetros", empresa: "Measure", md: "Sim", mdSentDate: "2026-07-07", diffDays: -46210, daysOpenTotal: 50, rc: "10410632", serviceStatus: "Aberto", poContrato: "4324031", medicao: "", valorTotal: 5500.0, saldoPo: 0.0, obs: "Contrato Criado (4324031). Aguardando Medição", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-040", date: "2026-01-23", assunto: "Acoplamento do Compressor", empresa: "Norpem", md: "Sim", mdSentDate: "2026-01-14", diffDays: 9, daysOpenTotal: 224, rc: "Contrato", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0.0, obs: "Serivço realizado via contrato", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-041", date: "2026-03-17", assunto: "Análise de Vibração nos BTT, Redutora, DGs", empresa: "Norpem", md: "Sim", mdSentDate: "2026-02-20", diffDays: 25, daysOpenTotal: 187, rc: "4263315", serviceStatus: "Fechado", poContrato: "4185241", medicao: "4299449", valorTotal: 61000.0, saldoPo: 0.0, obs: "Pagamento Programado para 03/09", statusPagamento: "Pagamento Programado", dataPagamento: "" },
    { id: "INV-042", date: "2026-05-26", assunto: "Limpeza do Hipap", empresa: "Northsub", md: "Sim", mdSentDate: "2026-05-26", diffDays: 0, daysOpenTotal: 92, rc: "10335199", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 24000.0, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-043", date: "2026-03-03", assunto: "Montagem de Andaime - Manutenção Tugger Winch", empresa: "Priner", md: "Sim", mdSentDate: "2026-01-14", diffDays: 48, daysOpenTotal: 224, rc: "10324367", serviceStatus: "Fechado", poContrato: "4190497", medicao: "", valorTotal: 0, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-044", date: "2026-01-23", assunto: "Detectores de Gases para Manutenção", empresa: "R2 Safety | Casa Offshore", md: "Sim", mdSentDate: "2026-01-14", diffDays: 9, daysOpenTotal: 224, rc: "Contrato", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 0.0, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-045", date: "2026-08-04", assunto: "Anual do VDR", empresa: "Radio Holland", md: "Sim", mdSentDate: "2026-07-27", diffDays: 8, daysOpenTotal: 30, rc: "10431861", serviceStatus: "Fechado", poContrato: "432567/ 4500166811", medicao: "", valorTotal: 6570.0, saldoPo: 0.0, obs: "PO Aprovada no SAP", statusPagamento: "Aguardando NF", dataPagamento: "" },
    { id: "INV-046", date: "2026-02-03", assunto: "Reparo Radar Banda X", empresa: "Radiomar", md: "Sim", mdSentDate: "2026-01-21", diffDays: 13, daysOpenTotal: 217, rc: "4600003659", serviceStatus: "Fechado", poContrato: "4215436", medicao: "4296004", valorTotal: 89880.3, saldoPo: 89880.31, obs: "PO 4500159000. medição aprovada.", statusPagamento: "Aguardando NF", dataPagamento: "" },
    { id: "INV-047", date: "2026-05-26", assunto: "Termografia", empresa: "Safe Trust", md: "Sim", mdSentDate: "2026-04-09", diffDays: 47, daysOpenTotal: 139, rc: "10369498", serviceStatus: "Fechado", poContrato: "4300384", medicao: "4304524", valorTotal: 12279.0, saldoPo: 0.0, obs: "Aguardando aprovação do Alexandre. Fornecedor incluiu os documentos solicitados", statusPagamento: "Aprovação Pendente", dataPagamento: "" },
    { id: "INV-048", date: "2026-02-16", assunto: "Manutenção Antena de TV", empresa: "Salestech", md: "Sim", mdSentDate: "2026-01-14", diffDays: 33, daysOpenTotal: 224, rc: "10311549", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-049", date: "2026-02-03", assunto: "Reparo Motor Elétrico do Bote de Resgate", empresa: "Sea Services", md: "Sim", mdSentDate: "2026-01-14", diffDays: 20, daysOpenTotal: 224, rc: "10303580", serviceStatus: "Fechado", poContrato: "4217060", medicao: "", valorTotal: 35515.64, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-050", date: "2026-02-16", assunto: "Equipe de Irata - Manutenção do Radares", empresa: "SETEC", md: "Sim", mdSentDate: "2026-01-14", diffDays: 33, daysOpenTotal: 224, rc: "10309582", serviceStatus: "Fechado", poContrato: "4233406", medicao: "", valorTotal: 17308.95, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-051", date: "2026-02-16", assunto: "Equipe de Irata -  Manutenção do Hipap", empresa: "SETEC", md: "Sim", mdSentDate: "2026-01-14", diffDays: 33, daysOpenTotal: 224, rc: "10309583", serviceStatus: "Fechado", poContrato: "4230444", medicao: "", valorTotal: 18673.59, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-052", date: "2026-02-20", assunto: "Inspeção e Reparo da Bomba de Alta Pressão", empresa: "Setec", md: "Sim", mdSentDate: "2026-01-19", diffDays: 32, daysOpenTotal: 219, rc: "10307578", serviceStatus: "Aberto", poContrato: "4232668", medicao: "", valorTotal: 4000.0, saldoPo: 0.0, obs: "Contrato SAP: 4600003799. Aguardando SETEC enviar medição.", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-053", date: "2026-06-23", assunto: "Certificação de PFOS nos LGS e Inspeção Semestral do CO2", empresa: "Sollax", md: "Sim", mdSentDate: "2026-05-18", diffDays: 36, daysOpenTotal: 100, rc: "10368492", serviceStatus: "Fechado", poContrato: "4500168289", medicao: "", valorTotal: 100000.0, saldoPo: 0.0, obs: "Aguardando envio da Medição por e-mail", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-054", date: "2026-04-07", assunto: "Teste Hidrostático PLT", empresa: "Survitec", md: "Sim", mdSentDate: "2026-03-25", diffDays: 13, daysOpenTotal: 154, rc: "10324355", serviceStatus: "Fechado", poContrato: "4269494", medicao: "", valorTotal: 4095.17, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-055", date: "2026-06-23", assunto: "NRs 12/ 13 e 35", empresa: "Tekee", md: "Sim", mdSentDate: "2026-05-04", diffDays: 50, daysOpenTotal: 114, rc: "10337500", serviceStatus: "Aberto", poContrato: "4273049", medicao: "", valorTotal: 160700.0, saldoPo: 0.0, obs: "Aguardando Medição", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-056", date: "2026-01-23", assunto: "Elaboração PGR", empresa: "Traume", md: "Sim", mdSentDate: "2026-01-15", diffDays: 8, daysOpenTotal: 223, rc: "10300888", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 35800.0, saldoPo: 35800.0, obs: "", statusPagamento: "Pago", dataPagamento: "" },
    { id: "INV-057", date: "2026-05-26", assunto: "Survey Instalação Sistema Monitoramento", empresa: "TWS", md: "Não", mdSentDate: "2026-05-26", diffDays: 0, daysOpenTotal: 65, rc: "10364210", serviceStatus: "Fechado", poContrato: "4500156779", medicao: "", valorTotal: 26206.0, saldoPo: 0.0, obs: "", statusPagamento: "Pago", dataPagamento: "2026-07-30" },
    { id: "INV-058", date: "2026-08-04", assunto: "Estudo dos Quadros Elétricos e Upgrade do AVR", empresa: "United Power", md: "Sim", mdSentDate: "2026-07-27", diffDays: 8, daysOpenTotal: 30, rc: "10432173", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Aguardando envio de orçamento no Portal", statusPagamento: "Aguardando Orçamento", dataPagamento: "" },
    { id: "INV-059", date: "2026-08-04", assunto: "Inspeção dos Sistemas de Filtragem CJC (Thrusters)", empresa: "United Power", md: "Sim", mdSentDate: "2026-07-27", diffDays: 8, daysOpenTotal: 30, rc: "10432174", serviceStatus: "Fechado", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Aguardando envio de orçamento no Portal", statusPagamento: "Aguardando Orçamento", dataPagamento: "" },
    { id: "INV-060", date: "2026-03-17", assunto: "Main Reduction - Inspeção", empresa: "Wartsila", md: "Sim", mdSentDate: "2026-02-04", diffDays: 41, daysOpenTotal: 203, rc: "Contrato", serviceStatus: "Aberto", poContrato: "4167832", medicao: "4319978", valorTotal: 99173.0, saldoPo: 99173.0, obs: "Medição aceita", statusPagamento: "Aguardando NF", dataPagamento: "" },
    { id: "INV-061", date: "2026-03-03", assunto: "Cilindros para Recarga", empresa: "Wilhelmsen", md: "Sim", mdSentDate: "2026-03-03", diffDays: 0, daysOpenTotal: 176, rc: "Contrato", serviceStatus: "Fechado", poContrato: "4133218", medicao: "", valorTotal: 192281.8, saldoPo: 0, obs: "Aguardando aprovação da empresa. Total dos Faturamentos para aprovar:	R$ 53.379,20", statusPagamento: "Aguardando Medição", dataPagamento: "" },
    { id: "INV-062", date: "", assunto: "Estudo de Instalação de Refrigeração Quadros Eletricos", empresa: "", md: "Sim", mdSentDate: "2026-03-03", diffDays: -46084, daysOpenTotal: 176, rc: "10432173", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Nova RC gerada. Aguardando envio do orçamento.", statusPagamento: "Aguardando Orçamento", dataPagamento: "" },
    { id: "INV-063", date: "", assunto: "Projeto de Construção de Almoxarifado dentro de um silo", empresa: "", md: "Sim", mdSentDate: "2026-07-07", diffDays: -46210, daysOpenTotal: 50, rc: "10413164", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Cotação expirada no Portal", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-064", date: "", assunto: "Comissionamento e Teste Operacional do Sistema BULK", empresa: "", md: "Sim", mdSentDate: "2026-07-07", diffDays: -46210, daysOpenTotal: 50, rc: "10413167", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Em cotação.", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-065", date: "", assunto: "Reparo dos Agitadores dos Tanques de Lama", empresa: "", md: "Sim", mdSentDate: "2026-07-07", diffDays: -46210, daysOpenTotal: 50, rc: "10413174", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Em cotação.", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-066", date: "", assunto: "Certificação dos Olhais da Praça de Máquinas e Ponte Rolante", empresa: "", md: "Sim", mdSentDate: "2026-07-07", diffDays: -46210, daysOpenTotal: 50, rc: "10413179", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Em cotação.", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-067", date: "", assunto: "Substituição da Rede de Bilge da Sala do Azimutal", empresa: "", md: "Sim", mdSentDate: "2026-07-07", diffDays: -46210, daysOpenTotal: 50, rc: "10413183", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Em cotação.", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-068", date: "", assunto: "Recuperação e Comissionamento das Bombas do Sistema MUD", empresa: "", md: "Sim", mdSentDate: "2026-07-07", diffDays: -46210, daysOpenTotal: 50, rc: "10413185", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Cotação expirada no Portal", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-069", date: "", assunto: "Tratamento e Reforma das Tampas do Moonpool (Superior e Inferior)", empresa: "", md: "Sim", mdSentDate: "2026-07-27", diffDays: -46230, daysOpenTotal: 30, rc: "10432115", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Status no Portal: Rejeitado por já ter contrato", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-070", date: "", assunto: "Certificação dos Olhais da Praça de Máquinas", empresa: "", md: "Sim", mdSentDate: "2026-07-27", diffDays: -46230, daysOpenTotal: 30, rc: "10432162", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Status no Portal: Em Agendamento", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-071", date: "", assunto: "Substituição da Gate Valve do Sistema HiPAP", empresa: "", md: "Sim", mdSentDate: "2026-07-27", diffDays: -46230, daysOpenTotal: 30, rc: "10432169", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Status no Portal: Aguardando Proposta", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
    { id: "INV-072", date: "", assunto: "Reparo do Detector Multigás Modelo 4X", empresa: "", md: "Sim", mdSentDate: "2026-07-27", diffDays: -46230, daysOpenTotal: 30, rc: "10432170", serviceStatus: "Aberto", poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0, obs: "Status no Portal: Em Agendamento", statusPagamento: "Aguardando Suprimentos", dataPagamento: "" },
];
const INITIAL_PORT_CALL_META = {};
const INITIAL_OP_CATEGORIES = [
"Manobras", "Troca de Turma", "Visitantes", "Manutenção", "Inspeção", "Base Açu", "Load", "Backload",
];
const INITIAL_EXCHANGE_RATE = 5.30;

export default function Root() {
  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState(null);

  const [users, setUsers] = useState(DEFAULT_USERS);
  const [currentUser, setCurrentUser] = useState(null);

  const [workPackages, setWorkPackages] = useState(INITIAL_WORK_PACKAGES);
  const [materials, setMaterials] = useState(INITIAL_MATERIALS);
  const [payments, setPayments] = useState(INITIAL_PAYMENTS);
  const [serviceInvoices, setServiceInvoices] = useState(INITIAL_SERVICE_INVOICES);
  const [portCallMeta, setPortCallMeta] = useState(INITIAL_PORT_CALL_META);
  const [opCategories, setOpCategories] = useState(INITIAL_OP_CATEGORIES);
  const [exchangeRate, setExchangeRate] = useState(INITIAL_EXCHANGE_RATE);

  /* carrega o estado salvo assim que o site abre — igual para qualquer usuário que entrar */
  React.useEffect(() => {
    fetch("/api/state")
      .then((r) => r.json())
      .then((res) => {
        const d = res && res.data;
        if (d) {
          if (d.users) setUsers(d.users);
          if (d.workPackages) setWorkPackages(d.workPackages);
          if (d.materials) setMaterials(d.materials);
          if (d.payments) setPayments(d.payments);
          if (d.serviceInvoices) setServiceInvoices(d.serviceInvoices);
          if (d.portCallMeta) setPortCallMeta(d.portCallMeta);
          if (d.opCategories) setOpCategories(d.opCategories);
          if (typeof d.exchangeRate === "number") setExchangeRate(d.exchangeRate);
        }
        setLoaded(true);
      })
      .catch(() => { setLoadError("Não foi possível conectar ao servidor — trabalhando localmente por enquanto."); setLoaded(true); });
  }, []);

  /* salva no servidor (com um pequeno atraso) toda vez que qualquer coisa muda — adicionar, editar,
     excluir linhas em qualquer aba — assim fica salvo automaticamente para todo mundo que acessar,
     independente de qual usuário fez a alteração */
  React.useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => {
      fetch("/api/state", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          data: { users, workPackages, materials, payments, serviceInvoices, portCallMeta, opCategories, exchangeRate },
        }),
      }).catch(() => setLoadError("Não foi possível salvar no servidor agora. Suas alterações ficam só neste navegador até a conexão voltar."));
    }, 700);
    return () => clearTimeout(t);
  }, [loaded, users, workPackages, materials, payments, serviceInvoices, portCallMeta, opCategories, exchangeRate]);

  if (!loaded) {
    return (
      <div className="genesis g-login-wrap"><Theme />
        <div style={{ color: "var(--text-dim)", fontFamily: "var(--mono)", fontSize: 13 }}>Carregando dados…</div>
      </div>
    );
  }

  if (!currentUser) return <LoginScreen users={users} onLogin={(u) => setCurrentUser(u)} />;

  return (
    <Genesis
      currentUser={currentUser}
      onLogout={() => setCurrentUser(null)}
      users={users} setUsers={setUsers}
      workPackages={workPackages} setWorkPackages={setWorkPackages}
      materials={materials} setMaterials={setMaterials}
      payments={payments} setPayments={setPayments}
      serviceInvoices={serviceInvoices} setServiceInvoices={setServiceInvoices}
      portCallMeta={portCallMeta} setPortCallMeta={setPortCallMeta}
      opCategories={opCategories} setOpCategories={setOpCategories}
      exchangeRate={exchangeRate} setExchangeRate={setExchangeRate}
      loadError={loadError}
    />
  );
}

/* ============================================================
   MAIN APP
   ============================================================ */
function Genesis({ currentUser, onLogout, users, setUsers,
  workPackages, setWorkPackages, materials, setMaterials, payments, setPayments,
  serviceInvoices, setServiceInvoices, portCallMeta, setPortCallMeta,
  opCategories, setOpCategories, exchangeRate, setExchangeRate, loadError }) {
  const [tab, setTab] = useState("dashboard");
  const [expandedWp, setExpandedWp] = useState(null);
  const [paySubTab, setPaySubTab] = useState("total"); // "total" | "status" | "dashboard"
  const [importMsg, setImportMsg] = useState(null);
  const fileInputRef = useRef(null);

  /* global period filter — present on every page */
  const [period, setPeriod] = useState({
    mode: "mes", // "mes" | "periodo" | "ano"
    month: 8,
    year: 2026,
    start: "2026-08-25",
    end: "2026-09-10",
  });

  /* workPackages vem de props (compartilhado via backend) */

  
  /* materials vem de props (compartilhado via backend) */

  /* payments vem de props (compartilhado via backend) */

  /* status de pagamento por serviço — importado da planilha "Pagamento Pendente (Serviços)" */
  /* serviceInvoices vem de props (compartilhado via backend) */

  /* ---------- generic row update/add/remove ---------- */
  const upd = (setter) => (idx, field, value) =>
    setter((rows) => rows.map((r, i) => (i === idx ? { ...r, [field]: value } : r)));
  const rem = (setter) => (idx) => setter((rows) => rows.filter((_, i) => i !== idx));

  const updWp = upd(setWorkPackages), remWp = rem(setWorkPackages);
  const updMat = upd(setMaterials), remMat = rem(setMaterials);
  const updPay = upd(setPayments), remPay = rem(setPayments);
  const updInv = upd(setServiceInvoices), remInv = rem(setServiceInvoices);

  const addWp = () => setWorkPackages((r) => [...r, {
    id: uid("PC-2026-08"), name: "Novo serviço", discipline: "Marine", group: "Sem categoria", empresa: "", md: "Não", rc: "", obs: "",
    budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-25T08:00", end: "2026-08-26T17:00",
    status: "Planejamento", progress: 0,
  }]);

  /* serviço não concluído na data planejada → gera uma NOVA linha com data em branco para reagendar,
     mantendo o registro antigo intacto (com seu desvio/histórico) e um vínculo entre as duas */
  const repeatWp = (original) => {
    const newId = uid("PC-2026-08");
    setWorkPackages((r) => [...r, {
      id: newId, name: original.name, discipline: original.discipline, group: original.group,
      ganttCategory: original.ganttCategory, empresa: original.empresa, md: original.md, rc: original.rc, obs: "",
      budget: 0, committed: 0, actual: 0, forecast: 0, start: "", end: "",
      status: "Não iniciado", progress: 0, repeatOf: original.id,
    }]);
    return newId;
  };

  /* ---------- effective period range from the global filter ---------- */
  const filterRange = useMemo(() => {
    let start, end;
    if (period.mode === "mes") {
      start = new Date(period.year, period.month - 1, 1, 0, 0, 0);
      end = new Date(period.year, period.month, 0, 23, 59, 59);
    } else if (period.mode === "ano") {
      start = new Date(period.year, 0, 1, 0, 0, 0);
      const yearEnd = new Date(period.year, 11, 31, 23, 59, 59);
      const now = new Date();
      end = now < yearEnd ? now : yearEnd;
    } else {
      start = new Date((period.start || "2026-01-01") + "T00:00:00");
      end = new Date((period.end || "2026-12-31") + "T23:59:59");
    }
    return { start, end };
  }, [period]);

  /* ---------- Port Call = data de início da atividade (nunca um texto fixo/sequencial) ---------- */
  const dateKeyOf = (dtStr) => (dtStr ? dtStr.slice(0, 10) : null);

  /* metadados extras de cada Port Call (duração em dias e local) — permite cadastrar um Port Call
     antes mesmo de ter atividades vinculadas a ele; portCallMeta/setPortCallMeta vêm de props (backend) */
  const addPortCallRecord = (dateKey, duration, local) => {
    setPortCallMeta((m) => ({ ...m, [dateKey]: { duration: Number(duration) || 24, local: local || "" } }));
  };

  const portCallLabel = (dateKey) => {
    const meta = portCallMeta[dateKey];
    const start = new Date(`${dateKey}T00:00:00`);
    const hours = meta?.duration || 24;
    const end = new Date(start.getTime() + hours * 3600000);
    let label;
    if (hours > 24) {
      const lastDay = new Date(end.getTime() - 1);
      label = `Port Call ${pad2(start.getDate())}/${pad2(start.getMonth() + 1)} — ${pad2(lastDay.getDate())}/${pad2(lastDay.getMonth() + 1)}`;
    } else {
      label = `Port Call ${pad2(start.getDate())}/${pad2(start.getMonth() + 1)}`;
    }
    if (meta?.local) label += ` · ${meta.local}`;
    label += ` · ${hours}h`;
    return label;
  };

  /* todas as datas de Port Call existentes — vindas de atividades já cadastradas OU de Port Calls
     criados antecipadamente (com data/duração/local) mas ainda sem nenhuma atividade */
  const allPortCallDates = useMemo(() => {
    const set = new Set([
      ...workPackages.map((w) => dateKeyOf(w.start)).filter(Boolean),
      ...Object.keys(portCallMeta),
    ]);
    return [...set].sort();
  }, [workPackages, portCallMeta]);

  /* somente as datas de Port Call cujo início cai dentro do período/mês selecionado no filtro global */
  const visiblePortCallDates = useMemo(() => {
    return allPortCallDates.filter((dk) => {
      const d = new Date(`${dk}T12:00:00`);
      return d >= filterRange.start && d <= filterRange.end;
    });
  }, [allPortCallDates, filterRange]);

  /* agrupa dias consecutivos que pertencem ao mesmo Port Call (quando a duração passa de 24h) em um
     único "span" — assim um Port Call de 01/09 até 03/09 aparece como um bloco só no Gantt, com seu
     próprio eixo de horas, em vez de três blocos separados */
  const portCallSpans = useMemo(() => {
    const sorted = [...visiblePortCallDates].sort();
    const spans = [];
    let i = 0;
    while (i < sorted.length) {
      const dk = sorted[i];
      const meta = portCallMeta[dk];
      const start = new Date(`${dk}T00:00:00`);
      const hours = meta?.duration || 24;
      const end = new Date(start.getTime() + hours * 3600000);
      spans.push({ startKey: dk, start, end, hours, local: meta?.local || "" });
      i++;
      while (i < sorted.length) {
        const nextStart = new Date(`${sorted[i]}T00:00:00`);
        if (nextStart < end) i++; else break;
      }
    }
    return spans;
  }, [visiblePortCallDates, portCallMeta]);

  const removePortCall = (dateKey) => {
    const meta = portCallMeta[dateKey];
    const hours = meta?.duration || 24;
    const start = new Date(`${dateKey}T00:00:00`);
    const end = new Date(start.getTime() + hours * 3600000);
    const affected = workPackages.filter((w) => { const d = new Date(w.start); return d >= start && d < end; });
    if (affected.length > 0 && !window.confirm(`Remover o ${portCallLabel(dateKey)} também vai remover ${affected.length} atividade(s) vinculada(s). Continuar?`)) return;
    setWorkPackages((wps) => wps.filter((w) => !affected.includes(w)));
    setPortCallMeta((m) => { const n = { ...m }; delete n[dateKey]; return n; });
  };
  const addWpOnDate = (dateKey, category) => setWorkPackages((r) => [...r, {
    id: uid("PC-2026-08"), name: "Nova atividade", discipline: "Marine", group: "Outros",
    ganttCategory: category || "Manutenção", empresa: "", md: "Não", rc: "", obs: "",
    budget: 0, committed: 0, actual: 0, forecast: 0,
    start: `${dateKey}T08:00`, end: `${dateKey}T17:00`,
    status: "Planejamento", progress: 0,
  }]);

  /* categorias operacionais do cronograma — globais, editáveis (renomear/adicionar/remover) */
  /* opCategories vem de props (compartilhado via backend) */
  const catOf = (w) => w.ganttCategory || "Manutenção";
  const addOpCategory = () => {
    let name = "Nova categoria";
    let n = 1;
    while (opCategories.includes(name)) { n += 1; name = `Nova categoria ${n}`; }
    setOpCategories((c) => [...c, name]);
  };
  const renameOpCategory = (oldName, newName) => {
    setOpCategories((c) => c.map((x) => (x === oldName ? newName : x)));
    setWorkPackages((wps) => wps.map((w) => (catOf(w) === oldName ? { ...w, ganttCategory: newName } : w)));
  };
  const removeOpCategory = (name) => {
    const count = workPackages.filter((w) => catOf(w) === name).length;
    if (count > 0 && !window.confirm(`Remover a categoria "${name}" vai deixar ${count} atividade(s) sem categoria. Continuar?`)) return;
    setOpCategories((c) => c.filter((x) => x !== name));
    setWorkPackages((wps) => wps.map((w) => (catOf(w) === name ? { ...w, ganttCategory: "" } : w)));
  };

  /* "PORT CALL" selecionado no topo — null = "Todos os Port Calls do período"; um valor = filtra só aquele dia,
     sobrepondo o filtro de Mês/Período/Ano em todas as telas (Gantt, Serviços, Custos, Dashboard) */
  const [selectedPortCallDate, setSelectedPortCallDate] = useState(null);
  React.useEffect(() => {
    if (selectedPortCallDate && !visiblePortCallDates.includes(selectedPortCallDate)) {
      setSelectedPortCallDate(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [visiblePortCallDates.join("|")]);
  const selectedPortCallLabel = selectedPortCallDate ? portCallLabel(selectedPortCallDate) : "Todos os Port Calls";

  /* datas efetivamente exibidas: se um Port Call específico foi escolhido, mostra só ele; senão, todos do período */
  const effectivePortCallDates = selectedPortCallDate ? [selectedPortCallDate] : visiblePortCallDates;
  const effectivePortCallSpans = selectedPortCallDate
    ? portCallSpans.filter((s) => s.startKey === selectedPortCallDate)
    : portCallSpans;

  /* intervalo de tempo efetivo: um único dia (Port Call escolhido) ou o período/mês/ano inteiro selecionado —
     usado pelo eixo do Gantt, total de horas/dias, e pelo resumo de datas no canto superior direito */
  const effectiveRange = useMemo(() => {
    if (selectedPortCallDate) {
      const span = portCallSpans.find((s) => s.startKey === selectedPortCallDate);
      if (span) return { start: span.start, end: span.end };
      return {
        start: new Date(`${selectedPortCallDate}T00:00:00`),
        end: new Date(`${selectedPortCallDate}T23:59:59`),
      };
    }
    return filterRange;
  }, [selectedPortCallDate, filterRange, portCallSpans]);

  const addMat = () => setMaterials((r) => [...r, {
    id: uid("MAT"), wp: "", tmMaster: "", departamento: "", sap: "", descricao: "Novo item",
    quantidade: 1, priority: "Média", dataSolicitacao: todayISO(), dataNecessidade: "", reserva: "", rc: "", po: "", linhaPo: "",
    valor: 0, eta: "", obs: "", dataRecebimento: "", status: "Solicitado",
  }]);
  const addPay = () => setPayments((r) => [...r, {
    id: uid("PAY"), service: "—", po: "", poValue: 0,
    nf: "", nfValue: 0, issue: "", due: "", status: "Orçamento",
  }]);
  const addInv = () => setServiceInvoices((r) => [...r, {
    id: uid("INV"), date: todayISO(), assunto: "Novo serviço", empresa: "", md: "Não", mdSentDate: "",
    diffDays: 0, daysOpenTotal: 0, rc: "", serviceStatus: "Aberto", poContrato: "", medicao: "",
    valorTotal: 0, saldoPo: 0, obs: "", statusPagamento: "Aguardando Medição", dataPagamento: "",
  }]);

  const portCallRange = useMemo(() => {
    if (workPackages.length === 0) return filterRange;
    const starts = workPackages.map((w) => new Date(w.start).getTime());
    const ends = workPackages.map((w) => new Date(w.end).getTime());
    return { start: new Date(Math.min(...starts)), end: new Date(Math.max(...ends)) };
  }, [workPackages, filterRange]);

  const overlapsWp = (w, range) => new Date(w.start) <= range.end && new Date(w.end) >= range.start;

  /* conjunto de serviços que pertencem ao(s) Port Call(s) efetivamente selecionado(s) — usado por
     Dashboard, Serviços e Custos para manter tudo sincronizado com o mesmo critério de filtragem */
  const filteredWorkPackages = useMemo(() => {
    return workPackages.filter((w) => effectivePortCallDates.includes(dateKeyOf(w.start)));
  }, [workPackages, effectivePortCallDates]);

  /* ---------- derived KPIs (per the requested dashboard spec) ---------- */
  const kpis = useMemo(() => {
    const inMonth = filteredWorkPackages;
    const inPortCall = workPackages.filter((w) => overlapsWp(w, portCallRange));

    const budgetMes = inMonth.reduce((s, w) => s + w.budget, 0);
    const utilizadoMes = inMonth.reduce((s, w) => s + w.committed, 0);
    const realizadoMes = inMonth.reduce((s, w) => s + w.actual, 0);

    const concluidosTotal = workPackages.filter((w) => w.status === "Concluído").length;
    const concluidosMes = inMonth.filter((w) => w.status === "Concluído").length;
    const concluidosPortCall = inPortCall.filter((w) => w.status === "Concluído").length;
    const emAndamento = workPackages.filter((w) => w.status === "Em andamento").length;
    const planejados = workPackages.filter((w) => w.status === "Planejamento").length;

    const requisicoesAbertas = materials.filter((m) => !["Recebido", "Entregue a bordo"].includes(m.status)).length;
    const materiaisUrgentes = materials.filter((m) =>
      ["Alta", "Crítica"].includes(m.priority) && !["Recebido", "Entregue a bordo"].includes(m.status)
    );

    const pagos = payments.filter((p) => paymentSituation(p) === "Pago");
    const pendentes = payments.filter((p) => paymentSituation(p) === "Pendente");
    const atrasados = payments.filter((p) => paymentSituation(p) === "Atrasado");
    const sum = (arr) => arr.reduce((s, p) => s + (p.nfValue || p.poValue), 0);
    const totalDiasAtraso = atrasados.reduce((s, p) => s + daysLate(p), 0);

    return {
      budgetMes, utilizadoMes, realizadoMes,
      concluidosTotal, concluidosMes, concluidosPortCall, emAndamento, planejados,
      requisicoesAbertas, materiaisUrgentes,
      pagosCount: pagos.length, pagosSum: sum(pagos),
      pendentesCount: pendentes.length, pendentesSum: sum(pendentes),
      atrasados, totalDiasAtraso,
    };
  }, [workPackages, filteredWorkPackages, materials, payments, portCallRange]);

  const disciplineCosts = useMemo(() => {
    const map = {};
    CATEGORIES.forEach((d) => {
      const orcadoUsd = CATEGORY_BUDGET_USD[d] || 0;
      map[d] = { discipline: d, orcadoUsd, orcadoBrl: orcadoUsd * exchangeRate, actual: 0 };
    });
    workPackages.forEach((w) => {
      if (!map[w.discipline]) map[w.discipline] = { discipline: w.discipline, orcadoUsd: 0, orcadoBrl: 0, actual: 0 };
      map[w.discipline].actual += w.actual;
    });
    return Object.values(map);
  }, [workPackages, exchangeRate]);

  /* ---------- EXPORT: full workbook (.xlsx) ---------- */
  const handleExportXlsx = () => {
    const wb = XLSX.utils.book_new();
    const resumo = [
      { Indicador: "Port Call", Valor: selectedPortCallLabel },
      { Indicador: "Período selecionado", Valor: `${fmtPeriodDate(filterRange.start)} — ${fmtPeriodDate(filterRange.end)}` },
      { Indicador: "Budget no período", Valor: kpis.budgetMes },
      { Indicador: "Utilizado (comprometido)", Valor: kpis.utilizadoMes },
      { Indicador: "Realizado", Valor: kpis.realizadoMes },
      { Indicador: "Serviços concluídos (total)", Valor: kpis.concluidosTotal },
      { Indicador: "Serviços concluídos (no período)", Valor: kpis.concluidosMes },
      { Indicador: "Serviços concluídos (no port call)", Valor: kpis.concluidosPortCall },
      { Indicador: "Em andamento", Valor: kpis.emAndamento },
      { Indicador: "Planejados", Valor: kpis.planejados },
      { Indicador: "Requisições abertas", Valor: kpis.requisicoesAbertas },
      { Indicador: "Pagamentos pagos (R$)", Valor: kpis.pagosSum },
      { Indicador: "Pagamentos pendentes (R$)", Valor: kpis.pendentesSum },
      { Indicador: "Total de dias em atraso (soma)", Valor: kpis.totalDiasAtraso },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumo), "Resumo");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(workPackages, WP_COLS)), "Servicos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(materials, MAT_COLS)), "Materiais");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(payments, PAY_COLS)), "Pagamentos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(serviceInvoices, INV_COLS)), "StatusPagamentos");
    XLSX.writeFile(wb, `genesis-${selectedPortCallLabel.replace(/\s+/g, "-").toLowerCase()}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleImportClick = () => fileInputRef.current?.click();
  const handleImportFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        let imported = [];
        if (wb.SheetNames.includes("Servicos")) {
          const json = XLSX.utils.sheet_to_json(wb.Sheets["Servicos"], { defval: "" });
          setWorkPackages(sheetToRows(json, WP_COLS));
          imported.push(`${json.length} serviços`);
        }
        if (wb.SheetNames.includes("Materiais")) {
          const json = XLSX.utils.sheet_to_json(wb.Sheets["Materiais"], { defval: "" });
          setMaterials(sheetToRows(json, MAT_COLS));
          imported.push(`${json.length} materiais`);
        }
        if (wb.SheetNames.includes("Pagamentos")) {
          const json = XLSX.utils.sheet_to_json(wb.Sheets["Pagamentos"], { defval: "" });
          setPayments(sheetToRows(json, PAY_COLS));
          imported.push(`${json.length} pagamentos`);
        }
        if (wb.SheetNames.includes("StatusPagamentos")) {
          const json = XLSX.utils.sheet_to_json(wb.Sheets["StatusPagamentos"], { defval: "" });
          setServiceInvoices(sheetToRows(json, INV_COLS));
          imported.push(`${json.length} status de pagamento`);
        }
        setImportMsg(imported.length ? `Importado: ${imported.join(", ")}.` : "Nenhuma aba reconhecida (esperado: Servicos, Materiais, Pagamentos, StatusPagamentos).");
      } catch (err) {
        setImportMsg("Erro ao ler o arquivo. Confira se é um .xlsx exportado por este sistema.");
      }
      setTimeout(() => setImportMsg(null), 5000);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const handleExportReport = () => {
    const today = new Date().toLocaleDateString("pt-BR");
    const critical = workPackages.filter((w) => w.status === "Não iniciado");
    const row = (cells) => `<tr>${cells.map((c) => `<td style="padding:6px 10px;border-bottom:1px solid #e3e3e3;font-size:12px;">${c}</td>`).join("")}</tr>`;
    const head = (cells) => `<tr>${cells.map((c) => `<th style="text-align:left;padding:6px 10px;font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:#666;border-bottom:2px solid #222;">${c}</th>`).join("")}</tr>`;

    const html = `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><title>Relatório — ${selectedPortCallLabel}</title>
<style>
  body { font-family: 'IBM Plex Sans', Arial, sans-serif; color:#1a1a1a; margin:40px; }
  h1 { font-size:20px; margin-bottom:2px; }
  h2 { font-size:14px; text-transform:uppercase; letter-spacing:.5px; border-bottom:2px solid #222; padding-bottom:6px; margin-top:34px; }
  .sub { color:#666; font-size:12px; margin-bottom:24px; }
  table { width:100%; border-collapse:collapse; margin-top:10px; }
  .kpis { display:flex; flex-wrap:wrap; gap:12px; margin-top:14px; }
  .kpi { border:1px solid #ddd; border-radius:4px; padding:10px 14px; min-width:150px; }
  .kpi-label { font-size:10px; text-transform:uppercase; color:#777; letter-spacing:.5px; }
  .kpi-value { font-size:18px; font-weight:700; margin-top:4px; }
  .alert { background:#fdecea; border:1px solid #f3b4ac; color:#8a2e22; padding:8px 12px; border-radius:4px; margin-bottom:8px; font-size:12.5px; }
  @media print { body { margin: 15mm; } }
</style></head><body>
  <h1>GENESIS I — Relatório de Manutenção &amp; Port Call</h1>
  <div class="sub">${selectedPortCallLabel} · Período: ${fmtPeriodDate(filterRange.start)} — ${fmtPeriodDate(filterRange.end)} · Gerado em ${today}</div>
  <h2>Indicadores do período</h2>
  <div class="kpis">
    <div class="kpi"><div class="kpi-label">Budget no período</div><div class="kpi-value">${fmt(kpis.budgetMes)}</div></div>
    <div class="kpi"><div class="kpi-label">Utilizado</div><div class="kpi-value">${fmt(kpis.utilizadoMes)}</div></div>
    <div class="kpi"><div class="kpi-label">Realizado</div><div class="kpi-value">${fmt(kpis.realizadoMes)}</div></div>
    <div class="kpi"><div class="kpi-label">Concluídos (total)</div><div class="kpi-value">${kpis.concluidosTotal}</div></div>
    <div class="kpi"><div class="kpi-label">Concluídos (período)</div><div class="kpi-value">${kpis.concluidosMes}</div></div>
    <div class="kpi"><div class="kpi-label">Pagos</div><div class="kpi-value">${fmt(kpis.pagosSum)}</div></div>
    <div class="kpi"><div class="kpi-label">Pendentes</div><div class="kpi-value">${fmt(kpis.pendentesSum)}</div></div>
    <div class="kpi"><div class="kpi-label">Dias em atraso (soma)</div><div class="kpi-value">${kpis.totalDiasAtraso}</div></div>
  </div>
  <h2>Serviços não iniciados</h2>
  ${critical.length === 0 ? "<p>Nenhum serviço pendente de início no momento.</p>" : critical.map((w) => `<div class="alert"><strong>${w.id}</strong> — ${w.name}: ${w.status}, ${w.progress}% concluído, início previsto ${fmtDateTime(w.start)}.</div>`).join("")}
  <h2>Todos os serviços</h2>
  <table>${head(["ID", "Serviço", "Disciplina", "Budget", "Realizado", "Status", "Progresso"])}
  ${workPackages.map((w) => row([w.id, w.name, w.discipline, fmt(w.budget), fmt(w.actual), w.status, w.progress + "%"])).join("")}</table>
  <h2>Pagamentos em atraso</h2>
  ${kpis.atrasados.length === 0 ? "<p>Nenhum pagamento em atraso.</p>" : `<table>${head(["ID", "Serviço", "Valor", "Vencimento", "Dias em atraso"])}
  ${kpis.atrasados.map((p) => row([p.id, p.service, fmt(p.nfValue || p.poValue), fmtDate(p.due), daysLate(p)])).join("")}</table>`}
  <p style="margin-top:40px;color:#999;font-size:11px;">Para salvar como PDF: abra este arquivo no navegador e use Imprimir → Salvar como PDF.</p>
</body></html>`;

    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `relatorio-${selectedPortCallLabel.replace(/\s+/g, "-").toLowerCase()}-${new Date().toISOString().slice(0, 10)}.html`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const navItems = [
    { key: "dashboard", label: "Dashboard", icon: LayoutGrid },
    { key: "gantt", label: "Port Call", icon: Ship },
    { key: "services", label: "Serviços", icon: Wrench },
    { key: "materials", label: "Materiais", icon: Package },
    { key: "payments", label: "Pagamentos", icon: Wallet },
    { key: "costs", label: "Custos", icon: Calculator },
    { key: "settings", label: "Configurações", icon: Settings },
  ];

  return (
    <div className="genesis">
      <Theme />

      {/* Horizontal top nav */}
      <div className="g-topnav">
        <span className="g-brand-mark">GENESIS I</span>
        <div className="g-nav-row">
          {navItems.map((n) => (
            <div key={n.key} className={`g-nav-item ${tab === n.key ? "active" : ""}`} onClick={() => setTab(n.key)}>
              <n.icon size={14} />{n.label}
            </div>
          ))}
        </div>
        {currentUser && <span className="g-muted" style={{ fontSize: 12, whiteSpace: "nowrap" }}>{currentUser.name}</span>}
        {onLogout && <div className="g-logout" onClick={onLogout}><LogOut size={14} />Sair</div>}
      </div>

      {/* Global period filter — presente em todas as páginas, exceto Pagamentos (que tem seu próprio filtro por período) */}
      {tab !== "payments" && tab !== "materials" && tab !== "costs" && tab !== "services" && (
      <div className="g-filterbar">
        <div className="g-field">
          <label>Port Call</label>
          {visiblePortCallDates.length === 0 ? (
            <select disabled style={{ minWidth: 170, opacity: 0.6 }}>
              <option>Nenhum Port Call</option>
            </select>
          ) : (
            <select
              value={selectedPortCallDate || ""}
              onChange={(e) => setSelectedPortCallDate(e.target.value || null)}
              style={{ minWidth: 170 }}
            >
              <option value="">Todos os Port Calls</option>
              {visiblePortCallDates.map((dk) => <option key={dk} value={dk}>{portCallLabel(dk)}</option>)}
            </select>
          )}
        </div>
        <div className="g-field">
          <label>Filtrar por</label>
          <div className="g-mode-toggle">
            <button className={period.mode === "mes" ? "active" : ""} onClick={() => setPeriod((p) => ({ ...p, mode: "mes" }))}>Mês</button>
            <button className={period.mode === "periodo" ? "active" : ""} onClick={() => setPeriod((p) => ({ ...p, mode: "periodo" }))}>Período</button>
            <button className={period.mode === "ano" ? "active" : ""} onClick={() => setPeriod((p) => ({ ...p, mode: "ano" }))}>Ano</button>
          </div>
        </div>
        {period.mode === "mes" && (
          <>
            <div className="g-field">
              <label>Mês</label>
              <select value={period.month} onChange={(e) => setPeriod((p) => ({ ...p, month: Number(e.target.value) }))}>
                {MONTH_NAMES.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
              </select>
            </div>
            <div className="g-field">
              <label>Ano</label>
              <input type="number" value={period.year} onChange={(e) => setPeriod((p) => ({ ...p, year: Number(e.target.value) }))} style={{ width: 80 }} />
            </div>
          </>
        )}
        {period.mode === "periodo" && (
          <>
            <div className="g-field">
              <label>Data inicial</label>
              <input type="date" value={period.start} onChange={(e) => setPeriod((p) => ({ ...p, start: e.target.value }))} />
            </div>
            <div className="g-field">
              <label>Data final</label>
              <input type="date" value={period.end} onChange={(e) => setPeriod((p) => ({ ...p, end: e.target.value }))} />
            </div>
          </>
        )}
        {period.mode === "ano" && (
          <div className="g-field">
            <label>Ano (até hoje)</label>
            <input type="number" value={period.year} onChange={(e) => setPeriod((p) => ({ ...p, year: Number(e.target.value) }))} style={{ width: 80 }} />
          </div>
        )}
        <div className="g-filter-spacer" />
        <div className="g-filter-summary">{fmtPeriodDate(effectiveRange.start)} → {fmtPeriodDate(effectiveRange.end)}</div>
      </div>
      )}

      <div className="g-pageactions">
        <div>
          <div className="g-title">{navItems.find((n) => n.key === tab)?.label}</div>
          <div className="g-title-sub">{selectedPortCallLabel.toUpperCase()}</div>
        </div>
        <div className="g-flex" style={{ gap: 8, flexWrap: "wrap" }}>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImportFile} />
          <button className="g-btn" onClick={handleImportClick} title="Importar planilha (.xlsx)"><Upload size={14} />Importar</button>
          <button className="g-btn" onClick={handleExportXlsx} title="Exportar todos os dados como planilha (.xlsx)"><Download size={14} />Exportar planilha</button>
          <button className="g-btn" onClick={handleExportReport} title="Exportar relatório (HTML, imprimível como PDF)"><FileText size={14} />Exportar relatório</button>
          {tab === "services" && <button className="g-btn primary" onClick={addWp}><Plus size={14} />Novo serviço</button>}
          {tab === "materials" && <button className="g-btn primary" onClick={addMat}><Plus size={14} />Nova requisição</button>}
          {tab === "payments" && <button className="g-btn primary" onClick={addInv}><Plus size={14} />Novo registro</button>}
        </div>
      </div>

      {importMsg && (
        <div style={{ margin: "10px 22px 0 22px" }}>
          <div className="g-alert" style={{ background: "rgba(63,193,201,0.08)", borderColor: "rgba(63,193,201,0.35)", color: "var(--teal)" }}>
            {importMsg}
          </div>
        </div>
      )}

      <div className="g-body">
        {tab === "dashboard" && (
          <DashboardView kpis={kpis} workPackages={workPackages} disciplineCosts={disciplineCosts}
            serviceInvoices={serviceInvoices}
            exchangeRate={exchangeRate} setExchangeRate={setExchangeRate} />
        )}

        {tab === "gantt" && (
          <GanttView workPackages={workPackages} portCallName={selectedPortCallLabel}
            spans={effectivePortCallSpans} portCallLabel={portCallLabel}
            updWp={updWp} remWp={remWp} removePortCall={removePortCall} addWpOnDate={addWpOnDate}
            addPortCallRecord={addPortCallRecord} filterRange={effectiveRange}
            opCategories={opCategories} catOf={catOf} addOpCategory={addOpCategory}
            renameOpCategory={renameOpCategory} removeOpCategory={removeOpCategory} />
        )}

        {tab === "services" && (
          <ServicesView workPackages={workPackages} updWp={updWp} remWp={remWp} repeatWp={repeatWp}
            expandedWp={expandedWp} setExpandedWp={setExpandedWp} />
        )}

        {tab === "materials" && <MaterialsView materials={materials} updMat={updMat} remMat={remMat} workPackages={workPackages} />}

        {tab === "payments" && (
          <PaymentsSection
            paySubTab={paySubTab} setPaySubTab={setPaySubTab}
            serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} addInv={addInv}
          />
        )}

        {tab === "costs" && (
          <CostsView serviceInvoices={serviceInvoices} updInv={updInv}
            exchangeRate={exchangeRate} setExchangeRate={setExchangeRate} />
        )}

        {tab === "settings" && (
          <SettingsView currentUser={currentUser} users={users} setUsers={setUsers} />
        )}
      </div>
    </div>
  );
}

/* ============================================================
   DASHBOARD — exact KPI set requested
   ============================================================ */
function DashboardView({ kpis, workPackages, disciplineCosts, serviceInvoices, exchangeRate, setExchangeRate }) {
  const fmtUsd = (n) => "US$ " + Number(n || 0).toLocaleString("en-US");
  const upcomingMaintenance = workPackages
    .filter((w) => w.status !== "Concluído")
    .slice()
    .sort((a, b) => new Date(a.start) - new Date(b.start));

  /* métricas de pagamento a partir dos dados reais de Status de Pagamento (planilha importada) */
  const invPagos = serviceInvoices.filter((r) => r.statusPagamento === "Pago");
  const invAbertos = serviceInvoices.filter((r) => r.statusPagamento !== "Pago" && r.statusPagamento !== "Cancelado");
  const invSum = (arr) => arr.reduce((s, r) => s + Number(r.valorTotal || 0), 0);
  const invTotalDiasAberto = invAbertos.reduce((s, r) => s + Number(r.daysOpenTotal || 0), 0);
  const topOpenInvoices = serviceInvoices
    .filter((r) => r.statusPagamento === "Aprovação Pendente")
    .slice()
    .sort((a, b) => Number(b.daysOpenTotal || 0) - Number(a.daysOpenTotal || 0))
    .slice(0, 8);

  /* último Port Call já concluído — derivado direto dos serviços reais cadastrados (aba Serviços) */
  const dateKeyOf = (dt) => (dt ? dt.slice(0, 10) : null);
  const pastDates = [...new Set(workPackages.map((w) => dateKeyOf(w.start)).filter(Boolean))]
    .filter((dk) => dk <= todayISO())
    .sort();
  const lastPortCallDate = pastDates[pastDates.length - 1] || null;
  const lastPortCallServices = lastPortCallDate
    ? workPackages.filter((w) => dateKeyOf(w.start) === lastPortCallDate)
    : [];
  const lastPortCallLabel = lastPortCallDate
    ? `Port Call ${lastPortCallDate.slice(8, 10)}/${lastPortCallDate.slice(5, 7)}`
    : "—";

  const financeiro = [
    { label: "Budget disponível no período", value: fmt(kpis.budgetMes), color: "var(--teal)" },
    { label: "Utilizado (comprometido)", value: fmt(kpis.utilizadoMes), color: "var(--warn)" },
    { label: "Realizado", value: fmt(kpis.realizadoMes), color: "var(--ok)" },
  ];
  const servicos = [
    { label: "Serviços concluídos (total)", value: kpis.concluidosTotal, color: "var(--ok)" },
    { label: "Serviços concluídos no período", value: kpis.concluidosMes, color: "var(--ok)" },
    { label: "Serviços concluídos no Port Call", value: kpis.concluidosPortCall, color: "var(--ok)" },
    { label: "Em andamento", value: kpis.emAndamento, color: "var(--teal)" },
    { label: "Serviços planejados", value: kpis.planejados, color: "var(--text-dim)" },
    { label: "Requisições abertas", value: kpis.requisicoesAbertas, color: "var(--warn)" },
  ];
  const pagamentos = [
    { label: "Pagos", value: `${invPagos.length} · ${fmt(invSum(invPagos))}`, color: "var(--ok)" },
    { label: "Em aberto", value: `${invAbertos.length} · ${fmt(invSum(invAbertos))}`, color: "var(--warn)" },
    { label: "Total de dias em aberto (soma)", value: `${invTotalDiasAberto} dias`, color: "var(--crit)" },
  ];

  return (
    <>
      <div className="g-section-label">Financeiro do período</div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(3, 1fr)" }}>
        {financeiro.map((k) => (
          <div className="g-kpi" style={{ "--kpi-accent": k.color }} key={k.label}>
            <div className="g-kpi-label">{k.label}</div>
            <div className="g-kpi-value">{k.value}</div>
          </div>
        ))}
      </div>

      <div className="g-section-label">Serviços</div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(3, 1fr)" }}>
        {servicos.map((k) => (
          <div className="g-kpi" style={{ "--kpi-accent": k.color }} key={k.label}>
            <div className="g-kpi-label">{k.label}</div>
            <div className="g-kpi-value small" style={{ color: k.color }}>{k.value}</div>
          </div>
        ))}
      </div>

      <div className="g-section-label">Pagamentos</div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(3, 1fr)" }}>
        {pagamentos.map((k) => (
          <div className="g-kpi" style={{ "--kpi-accent": k.color }} key={k.label}>
            <div className="g-kpi-label">{k.label}</div>
            <div className="g-kpi-value small" style={{ color: k.color }}>{k.value}</div>
          </div>
        ))}
      </div>

      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Aprovação Pendente há mais tempo (Status de Pagamentos)</span></div>
        {topOpenInvoices.length === 0 && <div className="g-muted">Nenhum pagamento com aprovação pendente no momento.</div>}
        {topOpenInvoices.map((r) => (
          <div className="g-list-item" key={r.id}>
            <span>{r.assunto} · {r.empresa} · {fmt(r.valorTotal)}</span>
            <span className="g-flex">
              <span style={{ color: "var(--crit)", fontFamily: "var(--mono)", fontWeight: 600 }}>{r.daysOpenTotal} dias</span>
              <Pill status={r.statusPagamento} />
            </span>
          </div>
        ))}
      </div>

      <div className="g-grid-2">
        <div className="g-panel">
          <div className="g-panel-head"><span className="g-panel-title">Próximas manutenções previstas para serem realizadas ({upcomingMaintenance.length})</span></div>
          {upcomingMaintenance.length === 0 && <div className="g-muted">Nenhuma manutenção prevista no momento.</div>}
          {upcomingMaintenance.map((w) => (
            <div className="g-list-item" key={w.id}>
              <span>
                {w.name} · {w.discipline} · início previsto {fmtDateTime(w.start)}
              </span>
              <Pill status={w.status} />
            </div>
          ))}
        </div>

        <div className="g-panel">
          <div className="g-panel-head">
            <span className="g-panel-title">{lastPortCallLabel}</span>
          </div>
          <div style={{ fontSize: 10.5, color: "var(--text-faint)", fontFamily: "var(--mono)", marginBottom: 8, textTransform: "uppercase", letterSpacing: ".5px" }}>
            Serviços do último Port Call registrado no sistema
          </div>
          {lastPortCallServices.length === 0 && <div className="g-muted">Nenhum Port Call registrado ainda.</div>}
          <table className="g-table">
            <thead><tr><th>Data</th><th>Serviço</th><th>Empresa</th><th>Status</th></tr></thead>
            <tbody>
              {lastPortCallServices.map((w) => (
                <tr className="g-row" key={w.id}>
                  <td style={{ fontFamily: "var(--mono)", fontSize: 12 }}>{fmtDate(dateKeyOf(w.start))}</td>
                  <td style={{ minWidth: 160 }}>{w.name}</td>
                  <td style={{ minWidth: 120 }}>{w.empresa || "—"}</td>
                  <td><Pill status={w.status} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="g-grid-2">
        <div className="g-panel">
          <div className="g-panel-head">
            <span className="g-panel-title">Custo por categoria (Orçado × Realizado)</span>
            <span className="g-flex" style={{ fontSize: 11 }}>
              <span className="g-muted" style={{ fontFamily: "var(--mono)" }}>Câmbio US$→R$</span>
              <input type="number" step="0.01" value={exchangeRate} onChange={(e) => setExchangeRate(Number(e.target.value))}
                style={{ width: 64, background: "var(--panel-raised)", border: "1px solid var(--border)", color: "var(--text)", fontFamily: "var(--mono)", fontSize: 11, padding: "4px 6px", borderRadius: 3 }} />
            </span>
          </div>
          <div style={{ width: "100%", height: 220 }}>
            <ResponsiveContainer>
              <BarChart data={disciplineCosts} margin={{ left: 0, right: 8, top: 4, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                <XAxis dataKey="discipline" tick={{ fill: "var(--text-faint)", fontSize: 9 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} interval={0} angle={-30} textAnchor="end" height={60} />
                <YAxis tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={40} tickFormatter={(v) => `${Math.round(v / 1000)}k`} />
                <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} formatter={(v) => fmt(v)} />
                <Bar dataKey="orcadoBrl" name="Orçado (R$)" radius={[3, 3, 0, 0]} fill="var(--border)" />
                <Bar dataKey="actual" name="Realizado (R$)" radius={[3, 3, 0, 0]} fill="var(--accent)" />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <table className="g-table" style={{ marginTop: 10 }}>
            <thead><tr><th>Categoria</th><th>Orçado (US$)</th><th>Orçado (R$)</th><th>Realizado (R$)</th></tr></thead>
            <tbody>
              {disciplineCosts.map((d) => (
                <tr key={d.discipline}>
                  <td>{d.discipline}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{fmtUsd(d.orcadoUsd)}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{fmt(d.orcadoBrl)}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{fmt(d.actual)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div className="g-panel">
          <div className="g-panel-head"><span className="g-panel-title">Materiais urgentes ({kpis.materiaisUrgentes.length})</span></div>
          {kpis.materiaisUrgentes.length === 0 && <div className="g-muted">Nenhum material urgente em aberto.</div>}
          {kpis.materiaisUrgentes.map((m) => (
            <div className="g-list-item" key={m.id}>
              <span>{m.descricao} · necessário {fmtDate(m.dataNecessidade)}</span>
              <span className="g-flex">
                <span className="g-pill" style={{ background: "var(--panel-raised)" }}>
                  <span className="g-dot" style={{ background: m.priority === "Crítica" ? "var(--crit)" : "var(--warn)" }} />{m.priority}
                </span>
                <Pill status={m.status} />
              </span>
            </div>
          ))}
        </div>
      </div>
    </>
  );
}

/* ============================================================
   GANTT / PORT CALL — hour precision, day grid, driven by the
   global period filter (mês ou período)
   ============================================================ */
function GanttView({ workPackages, filterRange, portCallName,
  spans, portCallLabel, updWp, remWp, removePortCall, addWpOnDate, addPortCallRecord,
  opCategories, catOf, addOpCategory, renameOpCategory, removeOpCategory }) {
  const [expandedRow, setExpandedRow] = useState(null);
  const isoMin = filterRange.start.toISOString().slice(0, 10);
  const isoMax = filterRange.end.toISOString().slice(0, 10);
  const [newPc, setNewPc] = useState({ date: isoMin, endDate: isoMin, duration: 24, local: "" });

  /* calculadora bidirecional: mudar a Data Fim recalcula a Duração (horas), e vice-versa */
  const setNewPcStart = (date) => {
    const end = new Date(`${newPc.endDate || date}T23:59:59`);
    const start = new Date(`${date}T00:00:00`);
    const hours = Math.max(1, Math.round((end - start) / 3600000));
    setNewPc((p) => ({ ...p, date, duration: end >= start ? hours : 24 }));
  };
  const setNewPcEnd = (endDate) => {
    const start = new Date(`${newPc.date}T00:00:00`);
    const end = new Date(`${endDate}T23:59:59`);
    const hours = Math.max(1, Math.round((end - start) / 3600000));
    setNewPc((p) => ({ ...p, endDate, duration: hours }));
  };
  const setNewPcDuration = (hours) => {
    const start = new Date(`${newPc.date}T00:00:00`);
    const end = new Date(start.getTime() + Number(hours || 0) * 3600000);
    setNewPc((p) => ({ ...p, duration: hours, endDate: end.toISOString().slice(0, 10) }));
  };

  const handleStatusChange = (i, v) => {
    updWp(i, "status", v);
    if (v === "Concluído") updWp(i, "progress", 100);
  };

  /* uma linha de dias + barras posicionadas em horas, calculada a partir do próprio intervalo do Port Call */
  const renderSpanTimeline = (span, activities) => {
    const totalHours = Math.max(1, span.hours);
    const days = Math.max(1, Math.ceil(span.hours / 24));
    const dayLabels = Array.from({ length: days }, (_, i) => {
      const d = new Date(span.start);
      d.setDate(d.getDate() + i);
      return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}`;
    });
    const hourOffset = (dtStr) => (new Date(dtStr) - span.start) / 3600000;
    const renderBar = (w) => {
      const startH = Math.max(0, hourOffset(w.start));
      const endH = Math.min(totalHours, hourOffset(w.end));
      const durationH = Math.max(0, endH - startH);
      if (endH < 0 || startH > totalHours) return <div className="g-gantt-track" />;
      const left = (startH / totalHours) * 100;
      const width = Math.max(0.6, (durationH / totalHours) * 100);
      return (
        <div className="g-gantt-track">
          <div className="g-gantt-bar" style={{ left: `${left}%`, width: `${width}%`, background: statusColor(w.status) }}
            title={`${w.name} · ${fmtDateTime(w.start)} → ${fmtDateTime(w.end)} · ${Math.round(durationH)}h · ${w.progress}%`}>
            <div className="g-gantt-bar-fill" style={{ width: `${w.progress}%` }} />
            <span style={{ position: "relative" }}>{Math.round(durationH)}h · {w.progress}%</span>
          </div>
        </div>
      );
    };
    return { dayLabels, days, renderBar };
  };

  return (
    <div className="g-panel">
      {/* ---- criar novo Port Call: data início, data fim (ou duração em horas — calculadas uma a partir da outra) e local ---- */}
      <div className="g-period-bar" style={{ marginBottom: 16 }}>
        <div className="g-field">
          <label>Novo Port Call — Início</label>
          <input type="date" value={newPc.date} onChange={(e) => setNewPcStart(e.target.value)} />
        </div>
        <div className="g-field">
          <label>Fim</label>
          <input type="date" value={newPc.endDate} onChange={(e) => setNewPcEnd(e.target.value)} />
        </div>
        <div className="g-field">
          <label>ou Duração (horas)</label>
          <input type="number" min="1" value={newPc.duration} onChange={(e) => setNewPcDuration(e.target.value)} style={{ width: 90 }} />
        </div>
        <div className="g-field">
          <label>Local</label>
          <input type="text" value={newPc.local} onChange={(e) => setNewPc((p) => ({ ...p, local: e.target.value }))} style={{ minWidth: 140 }} placeholder="ex: Base Açu" />
        </div>
        <button className="g-btn primary" onClick={() => addPortCallRecord(newPc.date, newPc.duration, newPc.local)}>
          <Plus size={14} />Adicionar Port Call
        </button>
      </div>

      <div className="g-panel-head">
        <span className="g-muted" style={{ fontFamily: "var(--mono)", fontSize: 12 }}>{spans.length} Port Call(s) no período selecionado</span>
        <button className="g-btn" onClick={addOpCategory}><Plus size={13} />Nova categoria</button>
      </div>

      <div className="g-gantt-wrap">
        <div className="g-gantt">
          {spans.length === 0 && (
            <div className="g-gantt-empty" style={{ marginLeft: 0, padding: "18px 0" }}>
              Nenhum Port Call encontrado para o período selecionado.
            </div>
          )}

          {spans.map((span) => {
            const pcTitle = portCallLabel(span.startKey);
            const pcActivities = workPackages.filter((w) => {
              const d = new Date(w.start);
              return d >= span.start && d < span.end;
            });
            const { dayLabels, days, renderBar } = renderSpanTimeline(span, pcActivities);
            return (
              <div key={span.startKey} style={{ marginBottom: 20 }}>
                {/* ---- título do Port Call + seu próprio eixo de horas/dias ---- */}
                <div className="g-gantt-group-row" style={{ background: "var(--panel-raised)", borderLeftColor: "var(--accent)" }}>
                  <span className="g-gantt-group-title" style={{ cursor: "default", fontSize: 12.5 }}>{pcTitle}</span>
                  <span className="g-flex" style={{ gap: 4 }}>
                    <span className="g-btn ghost" title="Remover este Port Call" onClick={() => removePortCall(span.startKey)}><Trash2 size={13} /></span>
                  </span>
                </div>
                <div className="g-gantt-header" style={{ "--gantt-days": days }}>
                  {dayLabels.map((d, i) => <div className="g-gantt-day" key={i}>{d}</div>)}
                </div>

                {pcActivities.length === 0 && (
                  <div className="g-gantt-empty">Nenhuma atividade neste Port Call ainda.</div>
                )}

                {/* ---- categorias operacionais dentro do Port Call ---- */}
                {opCategories.map((cat) => {
                  const rows = pcActivities.filter((w) => catOf(w) === cat);
                  return (
                    <div key={cat} style={{ marginLeft: 14, marginTop: 4 }}>
                      <div className="g-gantt-group-row" style={{ padding: "5px 10px", background: "transparent", border: "1px solid var(--border-soft)" }}>
                        <input
                          className="g-gantt-group-title"
                          style={{ fontSize: 10.5, color: "var(--text-dim)" }}
                          value={cat}
                          onChange={(e) => renameOpCategory(cat, e.target.value)}
                        />
                        <span className="g-flex" style={{ gap: 4 }}>
                          <span className="g-btn ghost" title="Adicionar atividade nesta categoria" onClick={() => addWpOnDate(span.startKey, cat)}><Plus size={13} /></span>
                          <span className="g-btn ghost danger" title="Remover categoria (global)" onClick={() => removeOpCategory(cat)}><Trash2 size={13} /></span>
                        </span>
                      </div>

                      {rows.length === 0 && (
                        <div className="g-gantt-empty" style={{ fontSize: 10.5 }}>Nenhuma linha aqui ainda — use o + acima.</div>
                      )}

                      {rows.map((w) => {
                        const i = workPackages.indexOf(w);
                        const isOpen = expandedRow === w.id;
                        return (
                          <div key={w.id} style={{ "--gantt-days": days }}>
                            <div className="g-gantt-row" style={{ borderLeft: `3px solid ${WP_STATUS_COLOR[w.status] || "#F2C94C"}` }}>
                              <div className="g-gantt-taskinfo">
                                <div className="g-flex" style={{ gap: 10, alignItems: "center", flexWrap: "nowrap" }}>
                                  <span className="g-btn ghost" style={{ padding: 2, flexShrink: 0 }} onClick={() => setExpandedRow(isOpen ? null : w.id)}>
                                    {isOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                                  </span>
                                  <textarea className="g-edit-wrap g-gantt-name-edit" rows={1} value={w.name}
                                    onChange={(e) => updWp(i, "name", e.target.value)}
                                    style={{ flex: "1 1 170px", minWidth: 150, fontWeight: 600 }} />
                                  <div style={{ flexShrink: 0 }}>
                                    <div className="g-gantt-mini-label">Início</div>
                                    <input type="datetime-local" className="g-gantt-mini-dt" value={w.start || ""}
                                      onChange={(e) => updWp(i, "start", e.target.value)} />
                                  </div>
                                  <div style={{ flexShrink: 0 }}>
                                    <div className="g-gantt-mini-label">Término</div>
                                    <input type="datetime-local" className="g-gantt-mini-dt" value={w.end || ""}
                                      onChange={(e) => updWp(i, "end", e.target.value)} />
                                  </div>
                                  <div style={{ flexShrink: 0, textAlign: "center", minWidth: 40 }}>
                                    <div className="g-gantt-mini-label">Duração</div>
                                    <div style={{ fontFamily: "var(--mono)", fontSize: 11, fontWeight: 700, color: "var(--text)" }}>
                                      {w.start && w.end ? `${Math.max(0, Math.round((new Date(w.end) - new Date(w.start)) / 3600000))}h` : "—"}
                                    </div>
                                  </div>
                                  <input type="text" className="g-edit" placeholder="Empresa" value={w.empresa || ""}
                                    onChange={(e) => updWp(i, "empresa", e.target.value)}
                                    style={{ width: 90, flexShrink: 0, fontSize: 10.5, background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 3, padding: "5px 6px" }} />
                                  <div style={{ flexShrink: 0 }}><StatusServicoSelect value={w.status} onChange={(v) => handleStatusChange(i, v)} /></div>
                                  <div className="g-flex" style={{ gap: 4, flexShrink: 0 }}>
                                    <div className="g-bar-bg" style={{ width: 36 }}><div className="g-bar-fg" style={{ width: `${w.progress}%`, background: statusColor(w.status) }} /></div>
                                    <input type="number" min="0" max="100" className="g-edit num" style={{ width: 36, fontSize: 10.5, padding: "2px 3px" }}
                                      value={w.progress} onChange={(e) => updWp(i, "progress", Number(e.target.value))} />
                                    <span style={{ fontSize: 9, color: "var(--text-faint)" }}>%</span>
                                  </div>
                                </div>
                              </div>
                              <div style={{ flex: 1, alignSelf: "center" }}>{renderBar(w)}</div>
                              <span className="g-btn ghost danger" style={{ marginLeft: 6, alignSelf: "center" }} onClick={() => remWp(i)}><Trash2 size={12} /></span>
                            </div>
                            {isOpen && (
                              <div className="g-gantt-editrow">
                                <div className="g-field"><label>Categoria (custo)</label><ESelect value={w.discipline} onChange={(v) => updWp(i, "discipline", v)} options={CATEGORIES} /></div>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      </div>
      <div className="g-flex" style={{ marginTop: 14, gap: 16, flexWrap: "wrap" }}>
        {WP_STATUS.map((s) => (
          <span key={s} className="g-flex" style={{ fontSize: 11 }}>
            <span className="g-dot" style={{ background: statusColor(s), marginRight: 5 }} />{s}
          </span>
        ))}
      </div>
    </div>
  );
}

/* ============================================================
   SERVICES / WORK PACKAGES
   ============================================================ */
/* dropdown de Status de Serviço com destaque de cor forte, no mesmo padrão usado em Pagamentos */
const WP_STATUS_COLOR = {
  "Planejamento": "#5D6E8C",
  "Não iniciado": "#8D9BB5",
  "Em andamento": "#3FC1C9",
  "Concluído": "#35D399",
  "Cancelado": "#6B7280",
};
const StatusServicoSelect = ({ value, onChange }) => {
  const color = WP_STATUS_COLOR[value] || "#F2C94C";
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{
        width: "100%", minWidth: 160, fontWeight: 700, fontSize: 12.5, cursor: "pointer",
        color, background: `${color}20`, border: `1.5px solid ${color}`, borderRadius: 5,
        padding: "6px 8px", fontFamily: "var(--sans)",
      }}
    >
      {WP_STATUS.map((o) => <option key={o} value={o} style={{ color: "#000" }}>{o}</option>)}
    </select>
  );
};

function ServicesView({ workPackages, updWp, remWp, repeatWp, expandedWp, setExpandedWp }) {
  const [sort, setSort] = useState({ key: null, dir: 1 });
  const [sf, setSf] = useState({ empresa: "", rc: "", manutencao: "", status: "Todos", portCall: "", dataInicio: "", dataFim: "" });
  const hasActiveFilter = sf.empresa || sf.rc || sf.manutencao || sf.status !== "Todos" || sf.portCall || sf.dataInicio || sf.dataFim;

  const dateKeyOf = (dt) => (dt ? dt.slice(0, 10) : null);
  const portCallLabel = (dk) => {
    const d = new Date(`${dk}T12:00:00`);
    return `Port Call ${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}`;
  };
  const portCallOptions = useMemo(() => [...new Set(workPackages.map((w) => dateKeyOf(w.start)).filter(Boolean))].sort(), [workPackages]);

  const setDateKeepTime = (i, w, newDate) => {
    const startTime = w.start ? w.start.slice(11) : "08:00";
    const endTime = w.end ? w.end.slice(11) : "17:00";
    updWp(i, "start", `${newDate}T${startTime}`);
    updWp(i, "end", `${newDate}T${endTime}`);
  };

  /* trocar o status sempre reajusta o progresso automaticamente para o padrão daquele status —
     mesmo indo "para trás" (ex: Concluído -> Não iniciado volta de 100% para 0%). O progresso
     continua podendo ser ajustado manualmente depois, até a próxima troca de status. */
  const handleStatusChange = (i, v) => {
    updWp(i, "status", v);
    updWp(i, "progress", WP_STATUS_DEFAULT_PROGRESS[v] ?? 0);
  };

  /* rastreabilidade: nem toda manutenção planejada para uma data é de fato executada nela —
     dataRealInicio/dataRealFim registram quando ela realmente aconteceu, e o desvio mostra
     a diferença em dias em relação à data planejada (w.start), pra não perder o histórico. */
  const desvioDias = (w) => {
    if (!w.dataRealInicio) return null;
    return Math.round((new Date(w.dataRealInicio) - new Date(dateKeyOf(w.start))) / 86400000);
  };

  const filtered = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return workPackages.filter((w) => {
      const dk = dateKeyOf(w.start);
      return (
        (!sf.empresa || norm(w.empresa).includes(norm(sf.empresa))) &&
        (!sf.rc || norm(w.rc).includes(norm(sf.rc))) &&
        (!sf.manutencao || norm(w.name).includes(norm(sf.manutencao))) &&
        (sf.status === "Todos" || w.status === sf.status) &&
        (!sf.portCall || dk === sf.portCall) &&
        (!sf.dataInicio || !dk || dk >= sf.dataInicio) &&
        (!sf.dataFim || !dk || dk <= sf.dataFim)
      );
    });
  }, [workPackages, sf]);
  const sorted = useMemo(() => sortRows(filtered, sort), [filtered, sort]);

  const concluidos = filtered.filter((w) => w.status === "Concluído").length;
  const emAndamento = filtered.filter((w) => w.status === "Em andamento").length;
  const naoIniciados = filtered.filter((w) => w.status === "Não iniciado").length;
  const cancelados = filtered.filter((w) => w.status === "Cancelado").length;
  const naoCancelados = filtered.length - cancelados;
  const taxaConclusao = naoCancelados ? Math.round((concluidos / naoCancelados) * 100) : 0;
  const comDesvio = filtered.filter((w) => { const d = desvioDias(w); return d !== null && d !== 0; });
  const desvioMedio = comDesvio.length ? Math.round(comDesvio.reduce((s, w) => s + Math.abs(desvioDias(w)), 0) / comDesvio.length) : 0;

  const bigKpi = (label, value, color, Icon) => (
    <div className="g-kpi" style={{ "--kpi-accent": color, padding: "18px 16px" }}>
      <div className="g-flex" style={{ gap: 6, marginBottom: 6 }}>
        {Icon && <Icon size={14} style={{ color, flexShrink: 0 }} />}
        <div className="g-kpi-label" style={{ fontSize: 11 }}>{label}</div>
      </div>
      <div className="g-kpi-value" style={{ fontSize: 24, color }}>{value}</div>
    </div>
  );

  return (
    <>
      {/* filtros locais — Portcall e Período agora são só desta aba, junto com Manutenção/Empresa/RC/Status */}
      <div className="g-filterbar" style={{ padding: "12px 0", marginBottom: 14, borderRadius: 4 }}>
        <div className="g-field">
          <label>Manutenção</label>
          <input type="text" value={sf.manutencao} onChange={(e) => setSf((p) => ({ ...p, manutencao: e.target.value }))} placeholder="digitar..." style={{ minWidth: 150 }} />
        </div>
        <div className="g-field">
          <label>Empresa</label>
          <input type="text" value={sf.empresa} onChange={(e) => setSf((p) => ({ ...p, empresa: e.target.value }))} placeholder="digitar..." style={{ minWidth: 120 }} />
        </div>
        <div className="g-field">
          <label>RC</label>
          <input type="text" value={sf.rc} onChange={(e) => setSf((p) => ({ ...p, rc: e.target.value }))} placeholder="digitar..." style={{ minWidth: 100 }} />
        </div>
        <div className="g-field">
          <label>Status</label>
          <select value={sf.status} onChange={(e) => setSf((p) => ({ ...p, status: e.target.value }))}>
            <option>Todos</option>
            {WP_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <div className="g-field">
          <label>Port Call</label>
          <select value={sf.portCall} onChange={(e) => setSf((p) => ({ ...p, portCall: e.target.value }))} style={{ minWidth: 140 }}>
            <option value="">Todos</option>
            {portCallOptions.map((dk) => <option key={dk} value={dk}>{portCallLabel(dk)}</option>)}
          </select>
        </div>
        <div className="g-field">
          <label>Período — de</label>
          <input type="date" value={sf.dataInicio} onChange={(e) => setSf((p) => ({ ...p, dataInicio: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>Período — até</label>
          <input type="date" value={sf.dataFim} onChange={(e) => setSf((p) => ({ ...p, dataFim: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>&nbsp;</label>
          <button className="g-btn" onClick={() => setSf({ empresa: "", rc: "", manutencao: "", status: "Todos", portCall: "", dataInicio: "", dataFim: "" })}
            disabled={!hasActiveFilter} style={{ opacity: hasActiveFilter ? 1 : 0.5 }}>
            <X size={13} />Limpar filtro
          </button>
        </div>
      </div>

      {/* KPIs de análise dos serviços */}
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
        {bigKpi("Total de Serviços", filtered.length, "var(--teal)", Wrench)}
        {bigKpi("Concluídos", concluidos, "var(--ok)", Wrench)}
        {bigKpi("Em Andamento", emAndamento, "var(--teal)", Wrench)}
        {bigKpi("Não Iniciados", naoIniciados, "var(--text-dim)", Wrench)}
        {bigKpi("Taxa de Conclusão", `${taxaConclusao}%`, "var(--warn)", LayoutGrid)}
      </div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(3, 1fr)" }}>
        {bigKpi("Cancelados", cancelados, "var(--text-dim)", X)}
        {bigKpi("Serviços com Desvio de Execução", comDesvio.length, "var(--crit)", AlertTriangle)}
        {bigKpi("Desvio Médio (dias)", desvioMedio, "var(--crit)", AlertTriangle)}
      </div>

      {filtered.length === 0 ? (
        <div className="g-panel">
          <div className="g-muted">Nenhum serviço encontrado com esses filtros.</div>
        </div>
      ) : (
      <div className="g-panel">
        <div className="g-table-wrap">
        <table className="g-table">
          <thead>
            <tr>
              <th></th>
              <SortTh sortKey="start" sort={sort} setSort={setSort}>Data</SortTh>
              <SortTh sortKey="group" sort={sort} setSort={setSort} style={{ minWidth: 130 }}>Categoria</SortTh>
              <SortTh sortKey="name" sort={sort} setSort={setSort} style={{ minWidth: 220 }}>Manutenção</SortTh>
              <SortTh sortKey="empresa" sort={sort} setSort={setSort} style={{ minWidth: 130 }}>Empresa</SortTh>
              <SortTh sortKey="md" sort={sort} setSort={setSort}>MD</SortTh>
              <SortTh sortKey="rc" sort={sort} setSort={setSort} style={{ minWidth: 100 }}>RC</SortTh>
              <SortTh sortKey="status" sort={sort} setSort={setSort} style={{ minWidth: 170 }}>Status</SortTh>
              <SortTh sortKey="progress" sort={sort} setSort={setSort} style={{ minWidth: 150 }}>Progresso</SortTh>
              <th style={{ minWidth: 80 }}>Desvio</th>
              <th style={{ minWidth: 200 }}>Observação</th><th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((w) => {
              const i = workPackages.indexOf(w);
              const isOpen = expandedWp === w.id;
              return (
                <React.Fragment key={w.id}>
                  <tr className="g-row" style={w.status === "Cancelado" ? { opacity: 0.5 } : undefined}>
                    <td>
                      <span className="g-btn ghost" onClick={() => setExpandedWp(isOpen ? null : w.id)}>
                        {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                      </span>
                    </td>
                    <td><EDate value={(w.start || "").slice(0, 10)} onChange={(v) => setDateKeepTime(i, w, v)} /></td>
                    <td style={{ minWidth: 130 }}><EText value={w.group || ""} onChange={(v) => updWp(i, "group", v)} /></td>
                    <td style={{ minWidth: 220 }}>
                      <div className="g-flex" style={{ gap: 4 }}>
                        {w.repeatOf && <span title="Esta linha é uma repetição de um serviço não concluído anteriormente" style={{ fontSize: 13, flexShrink: 0 }}>🔁</span>}
                        {w.status === "Cancelado" && <span title="Cancelado — não é mais necessário" style={{ fontSize: 13, flexShrink: 0 }}>🚫</span>}
                        <EText value={w.name} onChange={(v) => updWp(i, "name", v)} />
                      </div>
                    </td>
                    <td style={{ minWidth: 130 }}><EText value={w.empresa || ""} onChange={(v) => updWp(i, "empresa", v)} /></td>
                    <td><ESelect value={w.md || "Não"} onChange={(v) => updWp(i, "md", v)} options={["Sim", "Não"]} /></td>
                    <td style={{ minWidth: 100 }}><EText value={w.rc || ""} onChange={(v) => updWp(i, "rc", v)} mono /></td>
                    <td style={{ minWidth: 170 }}><StatusServicoSelect value={w.status} onChange={(v) => handleStatusChange(i, v)} /></td>
                    <td style={{ minWidth: 150 }}>
                      <div className="g-flex" style={{ gap: 6 }}>
                        <input type="range" min="0" max="100" value={w.progress}
                          onChange={(e) => updWp(i, "progress", Number(e.target.value))} style={{ width: 70 }} />
                        <input type="number" min="0" max="100" className="g-edit num" style={{ width: 48 }} value={w.progress}
                          onChange={(e) => updWp(i, "progress", Number(e.target.value))} />
                        <span style={{ fontSize: 10, color: "var(--text-faint)" }}>%</span>
                      </div>
                    </td>
                    {(() => {
                      const d = desvioDias(w);
                      return (
                        <td style={{ minWidth: 80, textAlign: "center" }}>
                          {d === null ? (
                            <span className="g-muted" style={{ fontSize: 11 }}>—</span>
                          ) : (
                            <span style={{
                              fontFamily: "var(--mono)", fontSize: 11, fontWeight: 700, padding: "3px 8px", borderRadius: 12,
                              color: d === 0 ? "var(--ok)" : "var(--crit)",
                              background: d === 0 ? "rgba(53,211,153,0.12)" : "rgba(242,104,91,0.12)",
                            }} title={d === 0 ? "Executado na data planejada" : d > 0 ? `Executado ${d} dia(s) depois do planejado` : `Executado ${Math.abs(d)} dia(s) antes do planejado`}>
                              {d > 0 ? `+${d}d` : `${d}d`}
                            </span>
                          )}
                        </td>
                      );
                    })()}
                    <td style={{ minWidth: 200 }}><EText value={w.obs || ""} onChange={(v) => updWp(i, "obs", v)} /></td>
                    <td><span className="g-btn ghost danger" onClick={() => remWp(i)}><Trash2 size={13} /></span></td>
                  </tr>
                  {isOpen && (
                    <tr className="g-expand-row">
                      <td></td>
                      <td colSpan={11} style={{ padding: "10px 8px 16px 8px" }}>
                        <div className="g-panel-title" style={{ marginBottom: 8 }}>Planejamento</div>
                        <div className="g-flex" style={{ flexWrap: "wrap", gap: 14, marginBottom: 14 }}>
                          <div className="g-field"><label>Categoria (custo)</label><ESelect value={w.discipline} onChange={(v) => updWp(i, "discipline", v)} options={CATEGORIES} /></div>
                          <div className="g-field"><label>Início</label><EDateTime value={w.start} onChange={(v) => updWp(i, "start", v)} /></div>
                          <div className="g-field"><label>Fim</label><EDateTime value={w.end} onChange={(v) => updWp(i, "end", v)} /></div>
                        </div>
                        <div className="g-panel-title" style={{ marginBottom: 8 }}>
                          Execução real <span className="g-muted" style={{ fontWeight: 400, textTransform: "none", letterSpacing: 0 }}>— preencha quando a manutenção acontecer numa data diferente da planejada, para manter o histórico</span>
                        </div>
                        <div className="g-flex" style={{ flexWrap: "wrap", gap: 14, marginBottom: 14 }}>
                          <div className="g-field"><label>Data real de início</label><EDate value={w.dataRealInicio} onChange={(v) => updWp(i, "dataRealInicio", v)} /></div>
                          <div className="g-field"><label>Data real de conclusão</label><EDate value={w.dataRealFim} onChange={(v) => updWp(i, "dataRealFim", v)} /></div>
                          {desvioDias(w) !== null && (
                            <div className="g-field">
                              <label>Desvio calculado</label>
                              <div style={{ fontFamily: "var(--mono)", fontSize: 13, fontWeight: 700, padding: "6px 0", color: desvioDias(w) === 0 ? "var(--ok)" : "var(--crit)" }}>
                                {desvioDias(w) > 0 ? `+${desvioDias(w)}` : desvioDias(w)} dia(s) {desvioDias(w) === 0 ? "(dentro do planejado)" : desvioDias(w) > 0 ? "depois do planejado" : "antes do planejado"}
                              </div>
                            </div>
                          )}
                        </div>

                        <div className="g-panel-title" style={{ marginBottom: 8 }}>Rastreabilidade</div>
                        {w.repeatOf && (() => {
                          const orig = workPackages.find((o) => o.id === w.repeatOf);
                          return (
                            <div className="g-alert" style={{ background: "rgba(63,193,201,0.08)", borderColor: "rgba(63,193,201,0.35)", color: "var(--teal)", marginBottom: 10 }}>
                              🔁 Esta linha é uma repetição de <strong>{orig ? orig.name : "um serviço anterior"}</strong>
                              {orig && ` — planejado originalmente para ${fmtDate(orig.start?.slice(0, 10))}, ficou como "${orig.status}"`}.
                            </div>
                          );
                        })()}
                        {(() => {
                          const repeats = workPackages.filter((o) => o.repeatOf === w.id);
                          return repeats.length > 0 && (
                            <div className="g-alert" style={{ background: "rgba(242,169,59,0.08)", borderColor: "rgba(242,169,59,0.35)", color: "var(--accent)", marginBottom: 10 }}>
                              Este serviço foi reagendado em {repeats.length} nova(s) linha(s): {repeats.map((r) => r.name).join(", ")}.
                            </div>
                          );
                        })()}
                        {w.status !== "Concluído" && w.status !== "Cancelado" && (
                          <div className="g-flex" style={{ gap: 8 }}>
                            <button className="g-btn" onClick={() => repeatWp(w)}>
                              🔁 Não concluído — repetir como nova linha (data em branco)
                            </button>
                            <button className="g-btn" onClick={() => updWp(i, "status", "Cancelado")}>
                              🚫 Não é mais necessário — cancelar
                            </button>
                          </div>
                        )}
                        {w.status === "Cancelado" && (
                          <div className="g-muted" style={{ fontSize: 12 }}>🚫 Este serviço foi marcado como não sendo mais necessário.</div>
                        )}
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
        </div>
      </div>
      )}
    </>
  );
}

/* ============================================================
   MATERIALS — fully editable, including ID
   ============================================================ */
function MaterialsView({ materials, updMat, remMat, workPackages }) {
  const [expandedRow, setExpandedRow] = useState(null);
  const [sort, setSort] = useState({ key: null, dir: 1 });
  const [mf, setMf] = useState({ tmMaster: "", sap: "", descricao: "", rc: "", reserva: "", po: "", status: "Todos", priority: "Todos" });
  const hasActiveFilter = mf.tmMaster || mf.sap || mf.descricao || mf.rc || mf.reserva || mf.po || mf.status !== "Todos" || mf.priority !== "Todos";

  const filtered = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return materials.filter((m) =>
      (!mf.tmMaster || norm(m.tmMaster).includes(norm(mf.tmMaster))) &&
      (!mf.sap || norm(m.sap).includes(norm(mf.sap))) &&
      (!mf.descricao || norm(m.descricao).includes(norm(mf.descricao))) &&
      (!mf.rc || norm(m.rc).includes(norm(mf.rc))) &&
      (!mf.reserva || norm(m.reserva).includes(norm(mf.reserva))) &&
      (!mf.po || norm(m.po).includes(norm(mf.po))) &&
      (mf.status === "Todos" || m.status === mf.status) &&
      (mf.priority === "Todos" || m.priority === mf.priority)
    );
  }, [materials, mf]);
  const sorted = useMemo(() => sortRows(filtered, sort), [filtered, sort]);

  const naoRecebido = (m) => !["Recebido", "Entregue a bordo"].includes(m.status);
  const urgentes = filtered.filter((m) => ["Alta", "Crítica"].includes(m.priority) && naoRecebido(m));
  const abertas = filtered.filter(naoRecebido);
  const semEta = filtered.filter((m) => !m.eta && naoRecebido(m));

  const bigKpi = (label, value, color, Icon) => (
    <div className="g-kpi" style={{ "--kpi-accent": color, padding: "18px 16px" }}>
      <div className="g-flex" style={{ gap: 6, marginBottom: 6 }}>
        {Icon && <Icon size={14} style={{ color, flexShrink: 0 }} />}
        <div className="g-kpi-label" style={{ fontSize: 11 }}>{label}</div>
      </div>
      <div className="g-kpi-value" style={{ fontSize: 24, color }}>{value}</div>
    </div>
  );

  return (
    <>
      {/* filtros da aba Materiais — digitáveis + selecionáveis */}
      <div className="g-filterbar" style={{ padding: "12px 0", marginBottom: 14, borderRadius: 4 }}>
        <div className="g-field">
          <label>TM Master</label>
          <input type="text" value={mf.tmMaster} onChange={(e) => setMf((p) => ({ ...p, tmMaster: e.target.value }))} placeholder="digitar..." style={{ minWidth: 110 }} />
        </div>
        <div className="g-field">
          <label>SAP</label>
          <input type="text" value={mf.sap} onChange={(e) => setMf((p) => ({ ...p, sap: e.target.value }))} placeholder="digitar..." style={{ minWidth: 100 }} />
        </div>
        <div className="g-field">
          <label>Descrição</label>
          <input type="text" value={mf.descricao} onChange={(e) => setMf((p) => ({ ...p, descricao: e.target.value }))} placeholder="digitar..." style={{ minWidth: 150 }} />
        </div>
        <div className="g-field">
          <label>RC</label>
          <input type="text" value={mf.rc} onChange={(e) => setMf((p) => ({ ...p, rc: e.target.value }))} placeholder="digitar..." style={{ minWidth: 100 }} />
        </div>
        <div className="g-field">
          <label>Reserva</label>
          <input type="text" value={mf.reserva} onChange={(e) => setMf((p) => ({ ...p, reserva: e.target.value }))} placeholder="digitar..." style={{ minWidth: 100 }} />
        </div>
        <div className="g-field">
          <label>PO</label>
          <input type="text" value={mf.po} onChange={(e) => setMf((p) => ({ ...p, po: e.target.value }))} placeholder="digitar..." style={{ minWidth: 100 }} />
        </div>
        <div className="g-field">
          <label>Status</label>
          <select value={mf.status} onChange={(e) => setMf((p) => ({ ...p, status: e.target.value }))}>
            <option>Todos</option>
            {MAT_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <div className="g-field">
          <label>Prioridade</label>
          <select value={mf.priority} onChange={(e) => setMf((p) => ({ ...p, priority: e.target.value }))}>
            <option>Todos</option>
            {PRIORITY.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
        </div>
        <div className="g-field">
          <label>&nbsp;</label>
          <button className="g-btn" onClick={() => setMf({ tmMaster: "", sap: "", descricao: "", rc: "", reserva: "", po: "", status: "Todos", priority: "Todos" })}
            disabled={!hasActiveFilter} style={{ opacity: hasActiveFilter ? 1 : 0.5 }}>
            <X size={13} />Limpar filtro
          </button>
        </div>
      </div>

      {/* KPIs de análise de materiais */}
      <div className="g-kpi-row">
        {bigKpi("Total de Materiais", filtered.length, "var(--teal)", Package)}
        {bigKpi("Materiais Urgentes", urgentes.length, "var(--crit)", AlertTriangle)}
        {bigKpi("Requisições Abertas", abertas.length, "var(--warn)", AlertTriangle)}
        {bigKpi("Sem ETA", semEta.length, "var(--crit)", AlertTriangle)}
      </div>

      <div className="g-panel">
        <div className="g-table-wrap">
        <table className="g-table">
          <thead>
            <tr>
              <th></th>
              <SortTh sortKey="tmMaster" sort={sort} setSort={setSort} style={{ minWidth: 100 }}>TM Master</SortTh>
              <SortTh sortKey="departamento" sort={sort} setSort={setSort} style={{ minWidth: 120 }}>Departamento</SortTh>
              <SortTh sortKey="sap" sort={sort} setSort={setSort} style={{ minWidth: 100 }}>SAP</SortTh>
              <SortTh sortKey="descricao" sort={sort} setSort={setSort} style={{ minWidth: 190 }}>Descrição</SortTh>
              <SortTh sortKey="quantidade" sort={sort} setSort={setSort}>Quantidade</SortTh>
              <SortTh sortKey="priority" sort={sort} setSort={setSort}>Prioridade</SortTh>
              <SortTh sortKey="dataSolicitacao" sort={sort} setSort={setSort}>Data da solicitação</SortTh>
              <SortTh sortKey="dataNecessidade" sort={sort} setSort={setSort}>Data da Necessidade</SortTh>
              <SortTh sortKey="reserva" sort={sort} setSort={setSort} style={{ minWidth: 90 }}>Reserva</SortTh>
              <SortTh sortKey="rc" sort={sort} setSort={setSort} style={{ minWidth: 100 }}>RC</SortTh>
              <SortTh sortKey="po" sort={sort} setSort={setSort} style={{ minWidth: 100 }}>PO</SortTh>
              <SortTh sortKey="linhaPo" sort={sort} setSort={setSort}>Linha da PO</SortTh>
              <SortTh sortKey="valor" sort={sort} setSort={setSort}>Valor</SortTh>
              <SortTh sortKey="eta" sort={sort} setSort={setSort}>ETA</SortTh>
              <SortTh sortKey="obs" sort={sort} setSort={setSort} style={{ minWidth: 160 }}>Observação</SortTh>
              <SortTh sortKey="dataRecebimento" sort={sort} setSort={setSort}>Data de Recebimento</SortTh>
              <SortTh sortKey="status" sort={sort} setSort={setSort} style={{ minWidth: 180 }}>Status</SortTh>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((m) => {
              const i = materials.indexOf(m);
              const isOpen = expandedRow === i;
              return (
                <React.Fragment key={i}>
                  <tr className="g-row">
                    <td>
                      <span className="g-btn ghost" onClick={() => setExpandedRow(isOpen ? null : i)}>
                        {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                      </span>
                    </td>
                    <td style={{ minWidth: 100 }}><EText value={m.tmMaster} onChange={(v) => updMat(i, "tmMaster", v)} mono /></td>
                    <td style={{ minWidth: 120 }}><EText value={m.departamento} onChange={(v) => updMat(i, "departamento", v)} /></td>
                    <td style={{ minWidth: 100 }}><EText value={m.sap} onChange={(v) => updMat(i, "sap", v)} mono /></td>
                    <td style={{ minWidth: 190 }}><EText value={m.descricao} onChange={(v) => updMat(i, "descricao", v)} /></td>
                    <td><ENum value={m.quantidade} onChange={(v) => updMat(i, "quantidade", v)} /></td>
                    <td><ESelect value={m.priority} onChange={(v) => updMat(i, "priority", v)} options={PRIORITY} /></td>
                    <td><EDate value={m.dataSolicitacao} onChange={(v) => updMat(i, "dataSolicitacao", v)} /></td>
                    <td><EDate value={m.dataNecessidade} onChange={(v) => updMat(i, "dataNecessidade", v)} /></td>
                    <td style={{ minWidth: 90 }}><EText value={m.reserva} onChange={(v) => updMat(i, "reserva", v)} mono /></td>
                    <td style={{ minWidth: 100 }}><EText value={m.rc} onChange={(v) => updMat(i, "rc", v)} mono /></td>
                    <td style={{ minWidth: 100 }}><EText value={m.po} onChange={(v) => updMat(i, "po", v)} mono /></td>
                    <td><EText value={m.linhaPo} onChange={(v) => updMat(i, "linhaPo", v)} mono /></td>
                    <td><ENum value={m.valor} onChange={(v) => updMat(i, "valor", v)} /></td>
                    <td>
                      {m.eta
                        ? <EDate value={m.eta} onChange={(v) => updMat(i, "eta", v)} />
                        : <span className="g-flex"><span style={{ color: "var(--crit)", fontSize: 11 }}>sem ETA</span>
                            <input type="date" className="g-edit mono" onChange={(e) => updMat(i, "eta", e.target.value)} /></span>}
                    </td>
                    <td style={{ minWidth: 160 }}><EText value={m.obs} onChange={(v) => updMat(i, "obs", v)} /></td>
                    <td><EDate value={m.dataRecebimento} onChange={(v) => updMat(i, "dataRecebimento", v)} /></td>
                    <td style={{ minWidth: 180 }}><ESelect value={m.status} onChange={(v) => updMat(i, "status", v)} options={MAT_STATUS} /></td>
                    <td><span className="g-btn ghost danger" onClick={() => remMat(i)}><Trash2 size={13} /></span></td>
                  </tr>
                  {isOpen && (
                    <tr className="g-expand-row">
                      <td></td>
                      <td colSpan={18} style={{ padding: "10px 8px 16px 8px" }}>
                        <div className="g-field" style={{ maxWidth: 320 }}>
                          <label>Vincular a um serviço (opcional)</label>
                          <select className="g-edit" value={m.wp || ""} onChange={(e) => updMat(i, "wp", e.target.value)}>
                            <option value="">— nenhum —</option>
                            {workPackages.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
                          </select>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
        </div>
        {filtered.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum material encontrado com esses filtros.</div>}
      </div>
    </>
  );
}

/* ============================================================
   PAYMENTS — Pago / Pendente / Atrasado
   ============================================================ */
/* ============================================================
   PAYMENTS SECTION — two pages: full Dashboard, and Status view
   ============================================================ */
function MultiSelectStatus({ options, selected, onChange }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  React.useEffect(() => {
    const onDocClick = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);
  const toggle = (opt) => {
    onChange(selected.includes(opt) ? selected.filter((s) => s !== opt) : [...selected, opt]);
  };
  const label = selected.length === 0 ? "Todos" : selected.length === 1 ? selected[0] : `${selected.length} selecionados`;
  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button type="button" className="g-btn" onClick={() => setOpen((o) => !o)} style={{ minWidth: 170, justifyContent: "space-between", fontFamily: "var(--mono)", fontSize: 12 }}>
        {label} <ChevronDown size={13} />
      </button>
      {open && (
        <div style={{
          position: "absolute", top: "calc(100% + 4px)", left: 0, zIndex: 20, minWidth: 220,
          background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4,
          padding: 8, boxShadow: "0 12px 30px rgba(0,0,0,0.4)",
        }}>
          {options.map((opt) => (
            <label key={opt} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 4px", fontSize: 12, cursor: "pointer" }}>
              <input type="checkbox" checked={selected.includes(opt)} onChange={() => toggle(opt)} />
              {opt}
            </label>
          ))}
          {selected.length > 0 && (
            <div className="g-btn ghost" style={{ marginTop: 4, fontSize: 11, justifyContent: "center" }} onClick={() => onChange([])}>Limpar seleção</div>
          )}
        </div>
      )}
    </div>
  );
}

const emptyPayFilter = { statuses: [], servico: "", po: "", rc: "", empresa: "", dataInicio: "", dataFim: "" };

function PaymentsSection({ paySubTab, setPaySubTab, serviceInvoices, updInv, remInv, addInv }) {
  const [f, setF] = useState(emptyPayFilter);
  const hasActiveFilter = f.statuses.length > 0 || f.servico || f.po || f.rc || f.empresa || f.dataInicio || f.dataFim;

  /* seleção de linhas (checkboxes) compartilhada entre as três páginas — os KPIs de cada página
     recalculam com base só nos serviços selecionados, quando houver alguma seleção */
  const [selectedIds, setSelectedIds] = useState(new Set());
  const toggleSelect = (id) => setSelectedIds((prev) => {
    const next = new Set(prev);
    next.has(id) ? next.delete(id) : next.add(id);
    return next;
  });
  const clearSelection = () => setSelectedIds(new Set());

  return (
    <>
      <div className="g-mode-toggle" style={{ marginBottom: 16, width: "fit-content" }}>
        <button className={paySubTab === "total" ? "active" : ""} onClick={() => setPaySubTab("total")}>Dashboard Total</button>
        <button className={paySubTab === "status" ? "active" : ""} onClick={() => setPaySubTab("status")}>Status dos Pagamentos</button>
        <button className={paySubTab === "dashboard" ? "active" : ""} onClick={() => setPaySubTab("dashboard")}>Dashboard de Valores</button>
      </div>

      {/* filtro compartilhado — presente em todas as páginas da aba Pagamentos */}
      <div className="g-filterbar" style={{ padding: "12px 0", marginBottom: 14, borderRadius: 4 }}>
        <div className="g-field">
          <label>Status (múltipla escolha)</label>
          <MultiSelectStatus options={STATUS_PAGAMENTO_OPTIONS} selected={f.statuses} onChange={(v) => setF((p) => ({ ...p, statuses: v }))} />
        </div>
        <div className="g-field">
          <label>Serviço</label>
          <input type="text" value={f.servico} onChange={(e) => setF((p) => ({ ...p, servico: e.target.value }))} placeholder="digitar..." style={{ minWidth: 140 }} />
        </div>
        <div className="g-field">
          <label>PO</label>
          <input type="text" value={f.po} onChange={(e) => setF((p) => ({ ...p, po: e.target.value }))} placeholder="digitar..." style={{ minWidth: 110 }} />
        </div>
        <div className="g-field">
          <label>RC</label>
          <input type="text" value={f.rc} onChange={(e) => setF((p) => ({ ...p, rc: e.target.value }))} placeholder="digitar..." style={{ minWidth: 100 }} />
        </div>
        <div className="g-field">
          <label>Empresa</label>
          <input type="text" value={f.empresa} onChange={(e) => setF((p) => ({ ...p, empresa: e.target.value }))} placeholder="digitar..." style={{ minWidth: 130 }} />
        </div>
        <div className="g-field">
          <label>Período — de</label>
          <input type="date" value={f.dataInicio} onChange={(e) => setF((p) => ({ ...p, dataInicio: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>Período — até</label>
          <input type="date" value={f.dataFim} onChange={(e) => setF((p) => ({ ...p, dataFim: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>&nbsp;</label>
          <button className="g-btn" onClick={() => setF(emptyPayFilter)} disabled={!hasActiveFilter} style={{ opacity: hasActiveFilter ? 1 : 0.5 }}>
            <X size={13} />Limpar filtro
          </button>
        </div>
      </div>

      {selectedIds.size > 0 && (
        <div className="g-alert" style={{ background: "rgba(242,169,59,0.1)", borderColor: "rgba(242,169,59,0.4)", color: "var(--accent)", justifyContent: "space-between", display: "flex", alignItems: "center" }}>
          <span><strong>{selectedIds.size}</strong> serviço(s) selecionado(s) — os KPIs abaixo refletem só a seleção.</span>
          <span className="g-btn ghost" onClick={clearSelection} style={{ color: "var(--accent)" }}><X size={13} />Limpar seleção</span>
        </div>
      )}

      {paySubTab === "total" && <PaymentsTotalView serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} f={f} selectedIds={selectedIds} toggleSelect={toggleSelect} />}
      {paySubTab === "status" && <PaymentsStatusView serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} f={f} setF={setF} selectedIds={selectedIds} toggleSelect={toggleSelect} />}
      {paySubTab === "dashboard" && <PaymentsValoresView serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} f={f} selectedIds={selectedIds} toggleSelect={toggleSelect} />}
    </>
  );
}

/* ---------- Página 1: Dashboard Total (todas as colunas da planilha + filtros + métricas de prazo) ---------- */
function PaymentsTotalView({ serviceInvoices, updInv, remInv, f, selectedIds, toggleSelect }) {
  const [sort, setSort] = useState({ key: null, dir: 1 });
  const daysBetween = (a, b) => Math.round((new Date(b) - new Date(a)) / 86400000);
  const execToPayDays = (r) => {
    if (!r.date) return null;
    const end = r.dataPagamento || todayISO();
    return daysBetween(r.date, end);
  };
  const mdToExecDays = (r) => {
    if (!r.mdSentDate || !r.date) return typeof r.diffDays === "number" ? r.diffDays : null;
    return daysBetween(r.mdSentDate, r.date);
  };

  const filtered = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return serviceInvoices.filter((r) =>
      (f.statuses.length === 0 || f.statuses.includes(r.statusPagamento)) &&
      (!f.servico || norm(r.assunto).includes(norm(f.servico))) &&
      (!f.po || norm(r.poContrato).includes(norm(f.po))) &&
      (!f.rc || norm(r.rc).includes(norm(f.rc))) &&
      (!f.empresa || norm(r.empresa).includes(norm(f.empresa))) &&
      (!f.dataInicio || !r.date || r.date >= f.dataInicio) &&
      (!f.dataFim || !r.date || r.date <= f.dataFim)
    );
  }, [serviceInvoices, f]);
  const sorted = useMemo(() => sortRows(filtered, sort), [filtered, sort]);

  /* quando há seleção, os KPIs refletem só os serviços selecionados (dentro do filtro atual) */
  const activeRows = selectedIds.size > 0 ? filtered.filter((r) => selectedIds.has(r.id)) : filtered;

  const avg = (arr) => (arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0);
  const execPayVals = activeRows.map(execToPayDays).filter((v) => v !== null);
  /* valores negativos (MD enviada depois da execução, inconsistência) não entram na média */
  const mdExecVals = activeRows.map(mdToExecDays).filter((v) => v !== null && v >= 0);
  const totalDiasAberto = activeRows.reduce((s, r) => s + Number(r.daysOpenTotal || 0), 0);
  const valorTotalSum = activeRows.reduce((s, r) => s + Number(r.valorTotal || 0), 0);
  const emAtraso = activeRows.filter((r) => Number(r.daysOpenTotal || 0) > 60).length;

  const bigKpi = (label, value, color, Icon) => (
    <div className="g-kpi" style={{ "--kpi-accent": color, padding: "18px 16px" }}>
      <div className="g-flex" style={{ gap: 6, marginBottom: 6 }}>
        {Icon && <Icon size={14} style={{ color, flexShrink: 0 }} />}
        <div className="g-kpi-label" style={{ fontSize: 11 }}>{label}</div>
      </div>
      <div className="g-kpi-value" style={{ fontSize: 24, color }}>{value}</div>
    </div>
  );

  return (
    <>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(3, 1fr)" }}>
        {bigKpi("Registros", activeRows.length, "var(--teal)", LayoutGrid)}
        {bigKpi("Valor Total", fmt(valorTotalSum), "var(--ok)", Wallet)}
        {bigKpi("Serviços em Atraso (+60d)", emAtraso, "var(--crit)", AlertTriangle)}
      </div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(3, 1fr)" }}>
        {bigKpi("Média Execução → Pagamento", `${avg(execPayVals).toFixed(1)} dias`, "var(--teal)")}
        {bigKpi("Média MD → Execução", `${avg(mdExecVals).toFixed(1)} dias`, "var(--teal)")}
        {bigKpi("Total de Dias em Aberto", `${totalDiasAberto} dias`, "var(--crit)", AlertTriangle)}
      </div>

      <div className="g-panel">
        <div className="g-table-wrap">
        <table className="g-table">
          <thead>
            <tr>
              <th></th>
              <SortTh sortKey="date" sort={sort} setSort={setSort}>Data</SortTh>
              <SortTh sortKey="assunto" sort={sort} setSort={setSort} style={{ minWidth: 260 }}>Manutenção</SortTh>
              <SortTh sortKey="empresa" sort={sort} setSort={setSort}>Empresa</SortTh>
              <SortTh sortKey="md" sort={sort} setSort={setSort}>MD</SortTh>
              <SortTh sortKey="mdSentDate" sort={sort} setSort={setSort}>Envio MD</SortTh>
              <th>MD→Exec (d)</th>
              <SortTh sortKey="daysOpenTotal" sort={sort} setSort={setSort}>Dias Aberto Total</SortTh>
              <SortTh sortKey="rc" sort={sort} setSort={setSort}>RC</SortTh>
              <SortTh sortKey="serviceStatus" sort={sort} setSort={setSort}>Status Serviço</SortTh>
              <SortTh sortKey="poContrato" sort={sort} setSort={setSort}>PO/Contrato</SortTh>
              <SortTh sortKey="medicao" sort={sort} setSort={setSort}>Medição</SortTh>
              <SortTh sortKey="valorTotal" sort={sort} setSort={setSort}>Valor Total</SortTh>
              <SortTh sortKey="saldoPo" sort={sort} setSort={setSort}>Saldo PO</SortTh>
              <th style={{ minWidth: 240 }}>Observações</th>
              <SortTh sortKey="statusPagamento" sort={sort} setSort={setSort} style={{ minWidth: 210 }}>Status Pagamento</SortTh>
              <SortTh sortKey="dataPagamento" sort={sort} setSort={setSort}>Data Pagamento</SortTh>
              <th>Exec→Pgto (d)</th><th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((r) => {
              const i = serviceInvoices.indexOf(r);
              return (
                <tr className="g-row" key={r.id} style={selectedIds.has(r.id) ? { background: "rgba(242,169,59,0.06)" } : undefined}>
                  <td><input type="checkbox" checked={selectedIds.has(r.id)} onChange={() => toggleSelect(r.id)} /></td>
                  <td><EDate value={r.date} onChange={(v) => updInv(i, "date", v)} /></td>
                  <td style={{ minWidth: 260, whiteSpace: "normal", verticalAlign: "top" }}><ETextArea value={r.assunto} onChange={(v) => updInv(i, "assunto", v)} /></td>
                  <td style={{ minWidth: 120 }}><EText value={r.empresa} onChange={(v) => updInv(i, "empresa", v)} /></td>
                  <td><ESelect value={r.md} onChange={(v) => updInv(i, "md", v)} options={["Sim", "Não"]} /></td>
                  <td><EDate value={r.mdSentDate} onChange={(v) => updInv(i, "mdSentDate", v)} /></td>
                  <td style={{ fontFamily: "var(--mono)", textAlign: "right", color: mdToExecDays(r) < 0 ? "var(--crit)" : undefined }} title={mdToExecDays(r) < 0 ? "Negativo — não entra na média" : ""}>
                    {mdToExecDays(r) ?? "—"}
                  </td>
                  <td><ENum value={r.daysOpenTotal} onChange={(v) => updInv(i, "daysOpenTotal", v)} /></td>
                  <td style={{ minWidth: 90 }}><EText value={r.rc} onChange={(v) => updInv(i, "rc", v)} mono /></td>
                  <td><ESelect value={r.serviceStatus} onChange={(v) => updInv(i, "serviceStatus", v)} options={["Aberto", "Fechado"]} /></td>
                  <td style={{ minWidth: 120 }}><EText value={r.poContrato} onChange={(v) => updInv(i, "poContrato", v)} mono /></td>
                  <td style={{ minWidth: 90 }}><EText value={r.medicao} onChange={(v) => updInv(i, "medicao", v)} mono /></td>
                  <td><ENum value={r.valorTotal} onChange={(v) => updInv(i, "valorTotal", v)} /></td>
                  <td><ENum value={r.saldoPo} onChange={(v) => updInv(i, "saldoPo", v)} /></td>
                  <td style={{ minWidth: 240, whiteSpace: "normal", verticalAlign: "top" }}><ETextArea value={r.obs} onChange={(v) => updInv(i, "obs", v)} /></td>
                  <td style={{ minWidth: 210 }}><StatusPagamentoSelect value={r.statusPagamento} onChange={(v) => updInv(i, "statusPagamento", v)} /></td>
                  <td><EDate value={r.dataPagamento} onChange={(v) => updInv(i, "dataPagamento", v)} /></td>
                  <td style={{ fontFamily: "var(--mono)", textAlign: "right" }}>{execToPayDays(r) ?? "—"}</td>
                  <td><span className="g-btn ghost danger" onClick={() => remInv(i)}><Trash2 size={13} /></span></td>
                </tr>
              );
            })}
          </tbody>
        </table>
        </div>
        {filtered.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum registro encontrado com esses filtros.</div>}
      </div>
    </>
  );
}

/* ---------- Página 2: Status dos Pagamentos (baseada na planilha "Pagamento Pendente") ---------- */
function PaymentsStatusView({ serviceInvoices, updInv, remInv, f, setF, selectedIds, toggleSelect }) {
  const [sort, setSort] = useState({ key: null, dir: 1 });
  const kpiStatuses = ["Aguardando Medição", "Aguardando Suprimentos", "Aprovação Pendente", "Aguardando NF"];
  const statusColorMap = STATUS_PAGAMENTO_COLOR;
  const toggleStatus = (s) => setF((p) => ({
    ...p, statuses: p.statuses.includes(s) ? p.statuses.filter((x) => x !== s) : [...p.statuses, s],
  }));

  const filtered = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return serviceInvoices.filter((r) =>
      (f.statuses.length === 0 || f.statuses.includes(r.statusPagamento)) &&
      (!f.servico || norm(r.assunto).includes(norm(f.servico))) &&
      (!f.po || norm(r.poContrato).includes(norm(f.po))) &&
      (!f.rc || norm(r.rc).includes(norm(f.rc))) &&
      (!f.empresa || norm(r.empresa).includes(norm(f.empresa))) &&
      (!f.dataInicio || !r.date || r.date >= f.dataInicio) &&
      (!f.dataFim || !r.date || r.date <= f.dataFim)
    );
  }, [serviceInvoices, f]);
  const sorted = useMemo(() => sortRows(filtered, sort), [filtered, sort]);

  /* KPIs refletem a seleção de linhas quando houver alguma */
  const activeRows = selectedIds.size > 0 ? filtered.filter((r) => selectedIds.has(r.id)) : filtered;
  const countOf = (s) => activeRows.filter((r) => r.statusPagamento === s).length;

  return (
    <>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
        <div className={`g-kpi clickable ${f.statuses.length === 0 ? "active" : ""}`}
          style={{ "--kpi-accent": "var(--text-dim)", padding: "18px 16px" }} onClick={() => setF((p) => ({ ...p, statuses: [] }))}>
          <div className="g-kpi-label" style={{ fontSize: 11 }}>Todos</div>
          <div className="g-kpi-value" style={{ fontSize: 26 }}>{activeRows.length}</div>
        </div>
        {kpiStatuses.map((s) => {
          const color = statusColorMap[s];
          const active = f.statuses.includes(s);
          return (
            <div key={s} className={`g-kpi clickable ${active ? "active" : ""}`}
              style={{ "--kpi-accent": color, padding: "18px 16px", background: active ? "var(--panel-alt)" : undefined }}
              onClick={() => toggleStatus(s)}>
              <div className="g-flex" style={{ gap: 6, marginBottom: 6 }}>
                <AlertTriangle size={14} style={{ color, flexShrink: 0 }} />
                <div className="g-kpi-label" style={{ fontSize: 11 }}>{s}</div>
              </div>
              <div className="g-kpi-value" style={{ fontSize: 28, color }}>{countOf(s)}</div>
            </div>
          );
        })}
      </div>

      <div className="g-panel">
        <div className="g-table-wrap">
        <table className="g-table">
          <thead>
            <tr>
              <th></th>
              <SortTh sortKey="assunto" sort={sort} setSort={setSort} style={{ minWidth: 260 }}>Manutenção</SortTh>
              <SortTh sortKey="empresa" sort={sort} setSort={setSort}>Empresa</SortTh>
              <SortTh sortKey="poContrato" sort={sort} setSort={setSort}>PO/Contrato</SortTh>
              <SortTh sortKey="medicao" sort={sort} setSort={setSort}>Medição</SortTh>
              <SortTh sortKey="daysOpenTotal" sort={sort} setSort={setSort} style={{ minWidth: 64 }}>Dias<br />Aberto</SortTh>
              <SortTh sortKey="statusPagamento" sort={sort} setSort={setSort} style={{ minWidth: 210 }}>Status Pagamento</SortTh>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((r) => {
              const i = serviceInvoices.indexOf(r);
              const color = statusColorMap[r.statusPagamento];
              const days = Number(r.daysOpenTotal || 0);
              const daysColor = days > 90 ? "var(--crit)" : days > 30 ? "var(--warn)" : "var(--text-dim)";
              return (
                <tr className="g-row" key={r.id} style={color ? { borderLeft: `3px solid ${color}`, background: selectedIds.has(r.id) ? "rgba(242,169,59,0.08)" : "rgba(255,255,255,0.015)" } : (selectedIds.has(r.id) ? { background: "rgba(242,169,59,0.08)" } : undefined)}>
                  <td><input type="checkbox" checked={selectedIds.has(r.id)} onChange={() => toggleSelect(r.id)} /></td>
                  <td style={{ minWidth: 260, whiteSpace: "normal", verticalAlign: "top" }}><ETextArea value={r.assunto} onChange={(v) => updInv(i, "assunto", v)} /></td>
                  <td style={{ minWidth: 140 }}><EText value={r.empresa} onChange={(v) => updInv(i, "empresa", v)} /></td>
                  <td style={{ minWidth: 130 }}><EText value={r.poContrato} onChange={(v) => updInv(i, "poContrato", v)} mono /></td>
                  <td style={{ minWidth: 100 }}><EText value={r.medicao} onChange={(v) => updInv(i, "medicao", v)} mono /></td>
                  <td style={{ minWidth: 56, maxWidth: 64 }}>
                    <div className="g-flex" style={{ gap: 3 }}>
                      <input type="number" className="g-edit num" style={{ width: 42, padding: "3px 3px" }} value={r.daysOpenTotal}
                        onChange={(e) => updInv(i, "daysOpenTotal", Number(e.target.value))} />
                      <span style={{ fontSize: 9, fontFamily: "var(--mono)", color: daysColor, fontWeight: 700 }}>d</span>
                    </div>
                  </td>
                  <td style={{ minWidth: 210 }}><StatusPagamentoSelect value={r.statusPagamento} onChange={(v) => updInv(i, "statusPagamento", v)} /></td>
                  <td><span className="g-btn ghost danger" onClick={() => remInv(i)}><Trash2 size={13} /></span></td>
                </tr>
              );
            })}
          </tbody>
        </table>
        </div>
        {filtered.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum registro com esse status.</div>}
      </div>
    </>
  );
}

/* situação derivada do status de pagamento, no mesmo padrão Pago/Pendente/Atrasado usado no resto do sistema */
const invoiceSituation = (r) => {
  if (r.statusPagamento === "Pago") return "Pago";
  if (r.statusPagamento === "Cancelado") return "Cancelado";
  return Number(r.daysOpenTotal || 0) > 60 ? "Atrasado" : "Pendente";
};
const invoiceSituationColor = { Pago: "var(--ok)", Pendente: "var(--warn)", Atrasado: "var(--crit)", Cancelado: "var(--text-faint)" };
const InvoiceSituationPill = ({ situation }) => (
  <span className="g-pill" style={{ background: "var(--panel-raised)", color: "var(--text)" }}>
    <span className="g-dot" style={{ background: invoiceSituationColor[situation] }} />{situation}
  </span>
);

/* ---------- Página 3: Dashboard de Valores — visão enxuta puxando os mesmos dados do Dashboard Total ---------- */
function PaymentsValoresView({ serviceInvoices, updInv, remInv, f, selectedIds, toggleSelect }) {
  const [situationFilter, setSituationFilter] = useState("Todos");
  const [sort, setSort] = useState({ key: null, dir: 1 });

  const withSituation = useMemo(() => serviceInvoices.map((r) => ({ ...r, _situation: invoiceSituation(r) })), [serviceInvoices]);

  const filtered = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return withSituation.filter((r) =>
      (situationFilter === "Todos" || r._situation === situationFilter) &&
      (f.statuses.length === 0 || f.statuses.includes(r.statusPagamento)) &&
      (!f.servico || norm(r.assunto).includes(norm(f.servico))) &&
      (!f.po || norm(r.poContrato).includes(norm(f.po))) &&
      (!f.empresa || norm(r.empresa).includes(norm(f.empresa))) &&
      (!f.dataInicio || !r.date || r.date >= f.dataInicio) &&
      (!f.dataFim || !r.date || r.date <= f.dataFim)
    );
  }, [withSituation, situationFilter, f]);
  const sorted = useMemo(() => sortRows(filtered, sort), [filtered, sort]);

  /* KPIs refletem a seleção de linhas quando houver alguma */
  const activeRows = selectedIds.size > 0 ? withSituation.filter((r) => selectedIds.has(r.id)) : withSituation;
  const pagos = activeRows.filter((r) => r._situation === "Pago");
  const pendentes = activeRows.filter((r) => r._situation === "Pendente");
  const atrasados = activeRows.filter((r) => r._situation === "Atrasado");
  const sum = (arr) => arr.reduce((s, r) => s + Number(r.valorTotal || 0), 0);

  const cards = [
    { key: "Todos", label: "Todos", value: activeRows.length, color: "var(--text-dim)", icon: LayoutGrid },
    { key: "Pago", label: "Pago", value: `${pagos.length} · ${fmt(sum(pagos))}`, color: "var(--ok)", icon: Wallet },
    { key: "Pendente", label: "Pendente", value: `${pendentes.length} · ${fmt(sum(pendentes))}`, color: "var(--warn)", icon: AlertTriangle },
    { key: "Atrasado", label: "Atrasado", value: `${atrasados.length} · ${fmt(sum(atrasados))}`, color: "var(--crit)", icon: AlertTriangle },
  ];

  return (
    <>
      <div className="g-kpi-row">
        {cards.map((c) => (
          <div key={c.key} className={`g-kpi clickable ${situationFilter === c.key ? "active" : ""}`}
            style={{ "--kpi-accent": c.color, padding: "18px 16px" }} onClick={() => setSituationFilter(c.key)}>
            <div className="g-flex" style={{ gap: 6, marginBottom: 6 }}>
              <c.icon size={14} style={{ color: c.color, flexShrink: 0 }} />
              <div className="g-kpi-label" style={{ fontSize: 11 }}>{c.label}</div>
            </div>
            <div className="g-kpi-value" style={{ fontSize: 22, color: c.color }}>{c.value}</div>
          </div>
        ))}
      </div>

      <div className="g-panel">
        <div className="g-table-wrap">
        <table className="g-table">
          <thead>
            <tr>
              <th></th>
              <SortTh sortKey="assunto" sort={sort} setSort={setSort} style={{ minWidth: 220 }}>Serviço</SortTh>
              <SortTh sortKey="empresa" sort={sort} setSort={setSort}>Empresa</SortTh>
              <SortTh sortKey="poContrato" sort={sort} setSort={setSort}>PO</SortTh>
              <SortTh sortKey="valorTotal" sort={sort} setSort={setSort}>Valor</SortTh>
              <SortTh sortKey="daysOpenTotal" sort={sort} setSort={setSort}>Dias em atraso</SortTh>
              <SortTh sortKey="statusPagamento" sort={sort} setSort={setSort} style={{ minWidth: 210 }}>Status</SortTh>
              <th>Situação</th><th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((r) => {
              const i = serviceInvoices.indexOf(r);
              const days = Number(r.daysOpenTotal || 0);
              return (
                <tr className="g-row" key={r.id} style={selectedIds.has(r.id) ? { background: "rgba(242,169,59,0.06)" } : undefined}>
                  <td><input type="checkbox" checked={selectedIds.has(r.id)} onChange={() => toggleSelect(r.id)} /></td>
                  <td style={{ minWidth: 220 }}><EText value={r.assunto} onChange={(v) => updInv(i, "assunto", v)} /></td>
                  <td style={{ minWidth: 130 }}><EText value={r.empresa} onChange={(v) => updInv(i, "empresa", v)} /></td>
                  <td style={{ minWidth: 110 }}><EText value={r.poContrato} onChange={(v) => updInv(i, "poContrato", v)} mono /></td>
                  <td><ENum value={r.valorTotal} onChange={(v) => updInv(i, "valorTotal", v)} /></td>
                  <td style={{ fontFamily: "var(--mono)", textAlign: "right", color: days > 60 ? "var(--crit)" : days > 30 ? "var(--warn)" : undefined }}>
                    <ENum value={r.daysOpenTotal} onChange={(v) => updInv(i, "daysOpenTotal", v)} />
                  </td>
                  <td style={{ minWidth: 210 }}><StatusPagamentoSelect value={r.statusPagamento} onChange={(v) => updInv(i, "statusPagamento", v)} /></td>
                  <td><InvoiceSituationPill situation={r._situation} /></td>
                  <td><span className="g-btn ghost danger" onClick={() => remInv(i)}><Trash2 size={13} /></span></td>
                </tr>
              );
            })}
          </tbody>
        </table>
        </div>
        {filtered.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum registro encontrado com esses filtros.</div>}
      </div>
    </>
  );
}

/* ============================================================
   COSTS
   ============================================================ */
function CostsView({ serviceInvoices, updInv, exchangeRate, setExchangeRate }) {
  const [expandedRow, setExpandedRow] = useState(null);
  const defaultPeriod = useMemo(() => {
    const now = new Date();
    const start = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-01`;
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    const end = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(lastDay)}`;
    return { start, end };
  }, []);
  const [cf, setCf] = useState({ statuses: [], servico: "", empresa: "", categorias: [], dataInicio: defaultPeriod.start, dataFim: defaultPeriod.end });
  const hasActiveFilter = cf.statuses.length > 0 || cf.servico || cf.empresa || cf.categorias.length > 0 || cf.dataInicio !== defaultPeriod.start || cf.dataFim !== defaultPeriod.end;

  /* esta aba só considera serviços que ainda não estão como "Pago" na aba Pagamentos */
  const naoPagos = useMemo(() => serviceInvoices.filter((r) => r.statusPagamento !== "Pago"), [serviceInvoices]);
  const statusOptions = useMemo(() => STATUS_PAGAMENTO_OPTIONS.filter((s) => s !== "Pago"), []);

  const allocationsOf = (r) => r.allocations || [];
  const allocatedSum = (r) => allocationsOf(r).reduce((s, a) => s + Number(a.valor || 0), 0);

  const filtered = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return naoPagos.filter((r) => {
      const inRange = !r.date || (!cf.dataInicio || r.date >= cf.dataInicio) && (!cf.dataFim || r.date <= cf.dataFim);
      const inStatus = cf.statuses.length === 0 || cf.statuses.includes(r.statusPagamento);
      const inServico = !cf.servico || norm(r.assunto).includes(norm(cf.servico));
      const inEmpresa = !cf.empresa || norm(r.empresa).includes(norm(cf.empresa));
      const inCategoria = cf.categorias.length === 0 || allocationsOf(r).some((a) => cf.categorias.includes(a.category));
      return inRange && inStatus && inServico && inEmpresa && inCategoria;
    });
  }, [naoPagos, cf]);

  const addAllocation = (i, r) => updInv(i, "allocations", [...allocationsOf(r), { category: CATEGORIES[0], valor: 0 }]);
  const updAllocation = (i, r, ai, field, value) => {
    const next = allocationsOf(r).map((a, idx) => (idx === ai ? { ...a, [field]: value } : a));
    updInv(i, "allocations", next);
  };
  const remAllocation = (i, r, ai) => updInv(i, "allocations", allocationsOf(r).filter((_, idx) => idx !== ai));

  /* resumo por categoria: Orçado é um valor mensal fixo (não acumula entre meses); Realizado é sempre
     recalculado apenas a partir do período selecionado — então ao trocar de mês, o "gasto" zera e o
     Orçado volta a aparecer inteiro, disponível de novo */
  const categoryCosts = useMemo(() => {
    return CATEGORIES.map((cat) => {
      const orcadoUsd = CATEGORY_BUDGET_USD[cat] || 0;
      const orcadoBrl = orcadoUsd * exchangeRate;
      const realizado = filtered.reduce((s, r) => s + allocationsOf(r).filter((a) => a.category === cat).reduce((s2, a) => s2 + Number(a.valor || 0), 0), 0);
      return { category: cat, orcadoUsd, orcadoBrl, realizado, disponivel: orcadoBrl - realizado };
    });
  }, [filtered, exchangeRate]);

  const totalOrcadoBrl = categoryCosts.reduce((s, c) => s + c.orcadoBrl, 0);
  const totalRealizado = categoryCosts.reduce((s, c) => s + c.realizado, 0);
  const totalDisponivel = totalOrcadoBrl - totalRealizado;
  const pctConsumido = totalOrcadoBrl ? Math.round((totalRealizado / totalOrcadoBrl) * 100) : 0;
  const semRateioCompleto = filtered.filter((r) => Math.round(allocatedSum(r)) !== Math.round(Number(r.valorTotal || 0))).length;

  const bigKpi = (label, value, color, Icon) => (
    <div className="g-kpi" style={{ "--kpi-accent": color, padding: "18px 16px" }}>
      <div className="g-flex" style={{ gap: 6, marginBottom: 6 }}>
        {Icon && <Icon size={14} style={{ color, flexShrink: 0 }} />}
        <div className="g-kpi-label" style={{ fontSize: 11 }}>{label}</div>
      </div>
      <div className="g-kpi-value" style={{ fontSize: 22, color }}>{value}</div>
    </div>
  );

  return (
    <>
      <div className="g-alert" style={{ background: "rgba(63,193,201,0.08)", borderColor: "rgba(63,193,201,0.35)", color: "var(--teal)" }}>
        Esta aba mostra apenas serviços que ainda <strong>não</strong> estão marcados como "Pago" na aba Pagamentos.
        Os valores de orçamento por categoria são mensais — ao trocar o período para outro mês, o realizado zera e o orçado volta inteiro.
      </div>
      <div className="g-filterbar" style={{ padding: "12px 0", marginBottom: 14, borderRadius: 4 }}>
        <div className="g-field">
          <label>Status de pagamento (múltipla escolha)</label>
          <MultiSelectStatus options={statusOptions} selected={cf.statuses} onChange={(v) => setCf((p) => ({ ...p, statuses: v }))} />
        </div>
        <div className="g-field">
          <label>Serviço</label>
          <input type="text" value={cf.servico} onChange={(e) => setCf((p) => ({ ...p, servico: e.target.value }))} placeholder="digitar..." style={{ minWidth: 150 }} />
        </div>
        <div className="g-field">
          <label>Empresa</label>
          <input type="text" value={cf.empresa} onChange={(e) => setCf((p) => ({ ...p, empresa: e.target.value }))} placeholder="digitar..." style={{ minWidth: 130 }} />
        </div>
        <div className="g-field">
          <label>Categoria (múltipla escolha)</label>
          <MultiSelectStatus options={CATEGORIES} selected={cf.categorias} onChange={(v) => setCf((p) => ({ ...p, categorias: v }))} />
        </div>
        <div className="g-field">
          <label>Período — de</label>
          <input type="date" value={cf.dataInicio} onChange={(e) => setCf((p) => ({ ...p, dataInicio: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>Período — até</label>
          <input type="date" value={cf.dataFim} onChange={(e) => setCf((p) => ({ ...p, dataFim: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>&nbsp;</label>
          <button className="g-btn" onClick={() => setCf({ statuses: [], servico: "", empresa: "", categorias: [], dataInicio: defaultPeriod.start, dataFim: defaultPeriod.end })}
            disabled={!hasActiveFilter} style={{ opacity: hasActiveFilter ? 1 : 0.5 }}>
            <X size={13} />Limpar filtro
          </button>
        </div>
      </div>

      {/* KPIs de análise */}
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
        {bigKpi("Total Realizado (rateado)", fmt(totalRealizado), "var(--teal)", Wallet)}
        {bigKpi("Total Orçado", fmt(totalOrcadoBrl), "var(--text-dim)", Calculator)}
        {bigKpi("Saldo Disponível", fmt(totalDisponivel), totalDisponivel < 0 ? "var(--crit)" : "var(--ok)", Wallet)}
        {bigKpi("% do Orçamento Consumido", `${pctConsumido}%`, pctConsumido > 100 ? "var(--crit)" : "var(--warn)", AlertTriangle)}
        {bigKpi("Serviços sem Rateio Completo", semRateioCompleto, "var(--crit)", AlertTriangle)}
      </div>

      {/* Custo por categoria — Orçado (US$/R$) × Realizado × Disponível */}
      <div className="g-panel">
        <div className="g-panel-head">
          <span className="g-panel-title">Custo por categoria — Orçado × Realizado × Disponível</span>
          <span className="g-flex" style={{ fontSize: 11 }}>
            <span className="g-muted" style={{ fontFamily: "var(--mono)" }}>Câmbio US$→R$</span>
            <input type="number" step="0.01" value={exchangeRate} onChange={(e) => setExchangeRate(Number(e.target.value))}
              style={{ width: 64, background: "var(--panel-raised)", border: "1px solid var(--border)", color: "var(--text)", fontFamily: "var(--mono)", fontSize: 11, padding: "4px 6px", borderRadius: 3 }} />
          </span>
        </div>
        <div className="g-table-wrap">
        <table className="g-table">
          <thead>
            <tr><th>Categoria</th><th>Orçado (US$)</th><th>Orçado (R$)</th><th>Realizado (R$)</th><th>Disponível (R$)</th></tr>
          </thead>
          <tbody>
            {categoryCosts.map((c) => (
              <tr key={c.category}>
                <td>{c.category}</td>
                <td style={{ fontFamily: "var(--mono)" }}>{"US$ " + c.orcadoUsd.toLocaleString("en-US")}</td>
                <td style={{ fontFamily: "var(--mono)" }}>{fmt(c.orcadoBrl)}</td>
                <td style={{ fontFamily: "var(--mono)" }}>{fmt(c.realizado)}</td>
                <td style={{ fontFamily: "var(--mono)", color: c.disponivel < 0 ? "var(--crit)" : "var(--ok)", fontWeight: 700 }}>{fmt(c.disponivel)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        </div>
      </div>

      {/* Tabela de serviços com rateio por categoria */}
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Serviços — rateio de custo por categoria</span></div>
        <div className="g-table-wrap">
        <table className="g-table">
          <thead>
            <tr>
              <th></th><th style={{ minWidth: 220 }}>Serviço</th><th>Empresa</th><th>Valor</th><th>PO</th>
              <th style={{ minWidth: 210 }}>Status de pagamento</th><th>Rateado</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((r) => {
              const i = serviceInvoices.indexOf(r);
              const isOpen = expandedRow === r.id;
              const alocado = allocatedSum(r);
              const total = Number(r.valorTotal || 0);
              const completo = Math.round(alocado) === Math.round(total);
              return (
                <React.Fragment key={r.id}>
                  <tr className="g-row">
                    <td>
                      <span className="g-btn ghost" onClick={() => setExpandedRow(isOpen ? null : r.id)}>
                        {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                      </span>
                    </td>
                    <td style={{ minWidth: 220, whiteSpace: "normal" }}>{r.assunto}</td>
                    <td style={{ minWidth: 130 }}>{r.empresa}</td>
                    <td style={{ fontFamily: "var(--mono)" }}>{fmt(r.valorTotal)}</td>
                    <td style={{ fontFamily: "var(--mono)" }}>{r.poContrato}</td>
                    <td style={{ minWidth: 210 }}><StatusPagamentoSelect value={r.statusPagamento} onChange={(v) => updInv(i, "statusPagamento", v)} /></td>
                    <td style={{ fontFamily: "var(--mono)", fontSize: 11, color: completo ? "var(--ok)" : "var(--warn)" }}>
                      {fmt(alocado)} / {fmt(total)}
                    </td>
                  </tr>
                  {isOpen && (
                    <tr className="g-expand-row">
                      <td></td>
                      <td colSpan={6} style={{ padding: "10px 8px 16px 8px" }}>
                        <div className="g-panel-title" style={{ marginBottom: 8 }}>
                          Rateio por categoria — {fmt(alocado)} alocado de {fmt(total)}
                          {!completo && <span style={{ color: "var(--warn)", marginLeft: 8, fontWeight: 400, fontSize: 11 }}>(ainda não bate com o valor total)</span>}
                        </div>
                        {allocationsOf(r).map((a, ai) => (
                          <div className="g-flex" key={ai} style={{ gap: 10, marginBottom: 6 }}>
                            <div style={{ minWidth: 200 }}>
                              <select className="g-edit" value={a.category} onChange={(e) => updAllocation(i, r, ai, "category", e.target.value)}>
                                {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                              </select>
                            </div>
                            <input type="number" className="g-edit num" style={{ width: 120 }} value={a.valor}
                              onChange={(e) => updAllocation(i, r, ai, "valor", Number(e.target.value))} />
                            <span className="g-btn ghost danger" onClick={() => remAllocation(i, r, ai)}><Trash2 size={13} /></span>
                          </div>
                        ))}
                        <button className="g-btn" onClick={() => addAllocation(i, r)}><Plus size={13} />Adicionar categoria</button>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
        </div>
        {filtered.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum serviço encontrado com esses filtros.</div>}
      </div>
    </>
  );
}

/* ============================================================
   SETTINGS — user/login management (client-side demo accounts)
   ============================================================ */
function SettingsView({ currentUser, users, setUsers }) {
  const updUser = (i, field, value) => setUsers((r) => r.map((u, idx) => (idx === i ? { ...u, [field]: value } : u)));
  const remUser = (i) => {
    if (users.length <= 1) { window.alert("É preciso manter pelo menos um usuário cadastrado."); return; }
    setUsers((r) => r.filter((_, idx) => idx !== i));
  };
  const addUser = () => setUsers((r) => [...r, { id: uid("USR"), name: "Novo usuário", username: `usuario${r.length + 1}`, password: "GenesisI" }]);

  return (
    <>
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Conta atual</span></div>
        <div className="g-muted">Conectado como <strong style={{ color: "var(--text)" }}>{currentUser?.name}</strong> (usuário: {currentUser?.username})</div>
      </div>

      <div className="g-panel">
        <div className="g-panel-head">
          <span className="g-panel-title">Usuários e senhas de acesso</span>
          <button className="g-btn primary" onClick={addUser}><Plus size={14} />Novo usuário</button>
        </div>
        <div className="g-alert" style={{ background: "rgba(63,193,201,0.08)", borderColor: "rgba(63,193,201,0.35)", color: "var(--teal)" }}>
          <Lock size={14} style={{ marginTop: 1 }} />
          Login local ao navegador, apenas para separar o acesso entre as pessoas — não é uma autenticação segura de servidor.
        </div>
        <table className="g-table">
          <thead>
            <tr><th>Nome</th><th>Usuário</th><th>Senha</th><th></th></tr>
          </thead>
          <tbody>
            {users.map((u, i) => (
              <tr className="g-row" key={u.id}>
                <td style={{ minWidth: 160 }}><EText value={u.name} onChange={(v) => updUser(i, "name", v)} /></td>
                <td style={{ minWidth: 140 }}><EText value={u.username} onChange={(v) => updUser(i, "username", v)} mono /></td>
                <td style={{ minWidth: 140 }}><EText value={u.password} onChange={(v) => updUser(i, "password", v)} mono /></td>
                <td><span className="g-btn ghost danger" onClick={() => remUser(i)}><Trash2 size={13} /></span></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}
