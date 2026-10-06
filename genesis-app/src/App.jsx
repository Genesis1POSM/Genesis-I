import React, { useState, useMemo, useRef } from "react";
import {
  LayoutGrid, Ship, Wrench, Package, Wallet, Calculator,
  Plus, Trash2, ChevronDown, ChevronUp, ChevronRight, AlertTriangle,
  Download, Upload, FileText, LogOut, Lock, User, X, Settings, DollarSign, Clock, ClipboardList, Gauge
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, LabelList, PieChart, Pie
} from "recharts";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";

/* ============================================================
   THEME — offshore ops console: deep marine dark, amber beacon
   accent, teal secondary, mono for IDs/data, sans for UI.
   ============================================================ */
const Theme = () => (
  <style>{`
    .genesis {
      --bg: #F4F5F9;
      --panel: #FFFFFF;
      --panel-alt: #F0F2F7;
      --panel-raised: #F7F8FB;
      --border: #E7E9F0;
      --border-soft: #EFF1F5;
      --text: #1F2430;
      --text-dim: #5B6273;
      --text-faint: #9499A8;
      --accent: #3B82F6;
      --accent-dim: #E8F0FE;
      --teal: #2B6CB0;
      --ok: #22C55E;
      --warn: #F5A623;
      --crit: #EF4444;
      --navbar: #FFFFFF;
      --navbar-alt: #EAF1FF;
      --navbar-text: #5B6273;
      --navbar-accent: #3B82F6;
      --sans: system-ui, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      --mono: system-ui, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;

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
      background: #F4F5F9;
    }
    .g-login-card {
      width: 340px; background: var(--panel); border: 1px solid var(--border);
      border-radius: 6px; padding: 30px 28px; box-shadow: 0 12px 32px rgba(20,30,45,0.10);
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
      background: var(--navbar); border-bottom: 1px solid var(--border);
    }
    .g-topnav { box-shadow: 0 1px 8px rgba(20,30,45,0.06); height: 60px !important; }
    .g-brand { display: flex; align-items: center; gap: 11px; flex-shrink: 0; }
    .g-brand-logo { height: 34px; width: auto; display: block; }
    .g-brand-text { display: flex; flex-direction: column; line-height: 1.1; }
    .g-brand-name { font-size: 17px; font-weight: 800; letter-spacing: 3px; color: #12203A; }
    .g-brand-name b { color: #E11D2E; font-weight: 800; }
    .g-brand-tag { font-size: 9px; letter-spacing: 1.6px; text-transform: uppercase; color: var(--text-faint); margin-top: 3px; font-weight: 600; }
    .g-brand-sep { width: 1px; height: 30px; background: var(--border); flex-shrink: 0; }
    .g-nav-item { border-radius: 8px !important; }
    .g-nav-item.active { border-bottom-color: transparent !important; background: var(--accent-dim) !important; font-weight: 600; }
    .g-login-logo { display: block; width: 190px; height: auto; margin: 0 auto 10px auto; }
    .g-brand-mark { font-family: var(--mono); font-weight: 700; font-size: 15px; color: var(--navbar-accent); letter-spacing: .5px; white-space: nowrap; }
    .g-nav-row { display: flex; align-items: center; gap: 2px; flex: 1; overflow-x: auto; }
    .g-nav-item {
      display: flex; align-items: center; gap: 7px;
      padding: 8px 13px; border-radius: 4px; color: var(--navbar-text); cursor: pointer;
      font-size: 12.5px; font-weight: 500; white-space: nowrap;
      border-bottom: 2px solid transparent;
    }
    .g-nav-item:hover { background: var(--navbar-alt); color: var(--navbar-accent); }
    .g-nav-item.active { color: var(--navbar-accent); border-bottom: 2px solid var(--navbar-accent); background: var(--navbar-alt); }
    .g-logout { display: flex; align-items: center; gap: 6px; color: var(--navbar-text); cursor: pointer; font-size: 12px; white-space: nowrap; }
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
    .g-mode-toggle button.active { background: var(--accent); color: #FFFFFF; font-weight: 600; }
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
      color: #FFFFFF;
      font-weight: 600;
    }
    .g-btn.primary:hover { filter: brightness(1.12); color: #FFFFFF; }
    .g-btn.ghost { background: transparent; border-color: transparent; padding: 4px 6px; }
    .g-btn.danger:hover { border-color: var(--crit); color: var(--crit); }

    /* ---------- KPI cards ---------- */
    .g-kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 14px; }
    .g-kpi {
      background: color-mix(in srgb, var(--kpi-accent, var(--teal)) 12%, var(--panel));
      border: 1px solid color-mix(in srgb, var(--kpi-accent, var(--teal)) 22%, var(--border));
      border-radius: 8px;
      padding: 13px 14px;
      position: relative;
      overflow: hidden;
      box-shadow: 0 1px 3px rgba(20,30,45,0.04);
    }
    .g-kpi.clickable { cursor: pointer; transition: filter .12s, border-color .12s; }
    .g-kpi.clickable:hover { filter: brightness(0.97); }
    .g-kpi.active { border-color: var(--kpi-accent, var(--teal)); box-shadow: 0 0 0 1px var(--kpi-accent, var(--teal)); }
    .g-kpi-label {
      font-size: 10px; text-transform: uppercase; letter-spacing: .8px;
      color: var(--text-faint); font-family: var(--mono); margin-bottom: 6px;
    }
    .g-kpi-value { font-size: 19px; font-weight: 600; font-family: var(--mono); }
    .g-kpi-value.small { font-size: 16px; }
    /* ---------- Tabelas modernas (Serviços / Pagamentos) ---------- */
    .tbl-toolbar { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 12px 16px; border-bottom: 1px solid var(--border); background: var(--panel); flex-wrap: wrap; }
    .tbl-check { display: inline-flex; align-items: center; gap: 7px; font-size: 12px; font-weight: 600; color: var(--text-dim); cursor: pointer; }
    .tbl-modern th { padding: 12px 14px; }
    .tbl-modern td { padding: 12px 14px; border-bottom: 1px solid var(--border-soft); }
    .tbl-modern tr.g-row:nth-child(even) { background: transparent !important; }
    .tbl-row td.tbl-first { border-left: 4px solid var(--c, transparent); }
    .tbl-chips { display: flex; flex-wrap: wrap; gap: 5px; padding: 2px 6px 0 6px; }
    .tbl-chip { font-size: 10.5px; background: #F0F2F7; color: var(--text-dim); border-radius: 6px; padding: 2px 7px; white-space: nowrap; }
    .tbl-sub { font-size: 11px; color: var(--text-faint); padding: 2px 6px 0 6px; }
    .tbl-days { display: inline-block; font-weight: 700; padding: 3px 10px; border-radius: 8px; font-size: 12px; white-space: nowrap; }
    .tbl-days.d1 { background: #E3F8EA; color: #15803D; }
    .tbl-days.d2 { background: #FFF3DC; color: #B45309; }
    .tbl-days.d3 { background: #FEE7E7; color: #B91C1C; }
    .tbl-progress { flex: 1; min-width: 60px; height: 7px; background: #EEF0F4; border-radius: 99px; overflow: hidden; }
    .tbl-progress > div { height: 100%; background: linear-gradient(90deg, #3B82F6, #22C55E); border-radius: 99px; transition: width .3s; }
    .tbl-group { cursor: pointer; }
    .tbl-group td { background: #F1F5FF !important; padding: 10px 14px !important; font-weight: 700; color: #12203A; border-bottom: 1px solid #E1E8FA !important; }
    .tbl-group:hover td { background: #E8EFFF !important; }
    .tbl-group-arrow { display: inline-block; width: 18px; color: var(--accent); }
    .tbl-group-cnt { background: #fff; border: 1px solid #D6E2FB; color: var(--accent); border-radius: 99px; padding: 1px 9px; font-size: 11px; margin-left: 10px; font-weight: 600; }
    .tbl-group-tot { float: right; color: var(--text-dim); font-weight: 600; font-size: 12px; }
    .tbl-detail td { background: var(--panel-alt) !important; padding: 14px 14px 18px 14px !important; }
    .tbl-detail-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 12px 16px; }
    /* ---------- Pagamentos (novo visual) ---------- */
    .pay-hero { display: grid; grid-template-columns: 1.6fr repeat(4, 1fr); gap: 14px; margin-bottom: 14px; }
    .pay-main { padding: 18px 20px; border-top: 3px solid var(--accent); }
    .pay-mini { padding: 16px 18px; }
    .pay-num { font-size: 26px; font-weight: 700; margin-top: 8px; letter-spacing: -.5px; }
    .pay-num small { font-size: 12px; font-weight: 500; color: var(--text-faint); letter-spacing: 0; }
    .pay-statusbar { display: flex; height: 12px; border-radius: 99px; overflow: hidden; gap: 2px; background: var(--border-soft); }
    .pay-view .g-mode-toggle { background: var(--panel); border: 1px solid var(--border); border-radius: 10px; padding: 3px; }
    .pay-view > .g-panel, .pay-view .g-panel { border-radius: 12px; box-shadow: 0 1px 4px rgba(20,30,45,0.05); }
    .pay-view .g-table th { background: var(--panel-raised); font-size: 10px; padding: 10px 8px; }
    .pay-view .g-table tr.g-row:nth-child(even) { background: #FAFBFD; }
    .pay-view .g-table tr.g-row:hover { background: var(--accent-dim); }
    .pay-view .g-table td { padding: 8px 8px; }
    @media (max-width: 1100px) { .pay-hero { grid-template-columns: 1fr 1fr; } }
    /* ---------- Dashboard (novo visual) ---------- */
    .dsh-period { display: flex; align-items: flex-end; gap: 12px; flex-wrap: wrap; background: var(--panel); border: 1px solid var(--border); border-radius: 10px; padding: 10px 14px; margin-bottom: 14px; }
    .dsh-period-note { font-size: 11px; color: var(--text-faint); }
    .dsh-hero { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 14px; }
    .dsh-row2 { display: grid; grid-template-columns: repeat(2, 1fr); gap: 14px; margin-bottom: 14px; }
    .dsh-card { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px; box-shadow: 0 1px 4px rgba(20,30,45,0.05); min-width: 0; }
    .dsh-fin { padding: 18px 20px; }
    .dsh-label { font-size: 11px; text-transform: uppercase; letter-spacing: .8px; color: var(--text-faint); font-weight: 600; }
    .dsh-big { font-size: 28px; font-weight: 700; margin: 6px 0 8px 0; letter-spacing: -.5px; }
    .dsh-sub { font-size: 11.5px; color: var(--text-dim); display: flex; align-items: center; gap: 8px; }
    .dsh-fx { width: 58px; border: 1px solid var(--border); border-radius: 5px; padding: 2px 5px; font-size: 11px; background: var(--panel-raised); color: var(--text); }
    .dsh-bar { height: 8px; background: var(--border-soft); border-radius: 99px; overflow: hidden; margin: 4px 0 8px 0; }
    .dsh-bar > div { height: 100%; border-radius: 99px; transition: width .4s; }
    .dsh-title { font-size: 13.5px; font-weight: 600; margin-bottom: 10px; display: flex; align-items: center; gap: 8px; }
    .dsh-count { background: var(--accent-dim); color: var(--accent); border-radius: 99px; font-size: 11px; padding: 1px 8px; font-weight: 600; }
    .dsh-donut { display: flex; align-items: center; gap: 14px; }
    .dsh-donut-chart { position: relative; width: 150px; height: 150px; flex: none; }
    .dsh-donut-center { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; pointer-events: none; }
    .dsh-donut-center b { font-size: 22px; }
    .dsh-donut-center span { font-size: 10px; color: var(--text-faint); text-transform: uppercase; letter-spacing: .6px; }
    .dsh-legend { display: flex; flex-direction: column; gap: 6px; font-size: 12px; color: var(--text-dim); flex: 1; }
    .dsh-legend div { display: flex; align-items: center; gap: 7px; }
    .dsh-legend i { width: 9px; height: 9px; border-radius: 99px; display: inline-block; }
    .dsh-legend b { margin-left: auto; color: var(--text); }
    .dsh-chips { display: flex; gap: 8px; margin-top: 12px; }
    .dsh-chips div { flex: 1; background: var(--panel-raised); border-radius: 8px; padding: 8px 10px; font-size: 10.5px; color: var(--text-faint); text-transform: uppercase; letter-spacing: .5px; min-width: 0; }
    .dsh-chips b { display: block; font-size: 13px; color: var(--text); margin-bottom: 2px; text-transform: none; letter-spacing: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .dsh-scroll { max-height: 320px; overflow-y: auto; }
    .dsh-item { display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 9px 4px; border-bottom: 1px solid var(--border-soft); }
    .dsh-item:last-child { border-bottom: none; }
    .dsh-item-main { font-size: 12.5px; font-weight: 500; }
    .dsh-item-sub { font-size: 11px; color: var(--text-faint); margin-top: 1px; }
    .dsh-days { color: var(--crit); font-weight: 700; font-size: 12px; white-space: nowrap; }
    .dsh-details { margin-top: 10px; font-size: 12px; }
    .dsh-details summary { cursor: pointer; color: var(--accent); font-weight: 500; }
    @media (max-width: 1100px) { .dsh-hero, .dsh-row2 { grid-template-columns: 1fr; } }
    .g-section-label {
      font-size: 10.5px; text-transform: uppercase; letter-spacing: 1px; color: var(--text-faint);
      margin: 18px 0 8px 2px; font-family: var(--mono);
    }
    .g-section-label:first-child { margin-top: 0; }

    /* ---------- Panels / sections ---------- */
    .g-panel {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 16px;
      margin-bottom: 16px;
      box-shadow: 0 1px 3px rgba(20,30,45,0.05);
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

    /* ---------- Visual geral renovado (todas as páginas) ---------- */
    .g-body { background: var(--bg); }
    .g-pageactions { padding: 18px 22px 2px 22px; }
    .g-title { font-size: 21px; font-weight: 700; letter-spacing: -.2px; color: #12203A; display: flex; align-items: center; gap: 10px; }
    .g-title::before { content: ""; width: 4px; height: 22px; border-radius: 4px; background: linear-gradient(180deg, #3B82F6, #E11D2E); display: inline-block; }
    .g-btn { border-radius: 8px; padding: 8px 14px; background: var(--panel); box-shadow: 0 1px 2px rgba(20,30,45,0.05); transition: border-color .12s, box-shadow .12s, transform .05s; }
    .g-btn:hover { border-color: var(--accent); box-shadow: 0 2px 6px rgba(59,130,246,0.15); }
    .g-btn:active { transform: translateY(1px); }
    .g-btn.primary { background: linear-gradient(180deg, #4C8FF7, #3B82F6); border-color: #3B82F6; color: #fff; box-shadow: 0 2px 6px rgba(59,130,246,0.35); }
    .g-btn.ghost { box-shadow: none; background: transparent; }
    .g-panel { border-radius: 12px; padding: 18px; box-shadow: 0 1px 4px rgba(20,30,45,0.05); }
    .g-panel-title { text-transform: none; letter-spacing: 0; font-size: 14px; font-weight: 700; color: #12203A; font-family: var(--sans); }
    .g-kpi { border-radius: 12px; background: var(--panel); border: 1px solid var(--border); padding: 16px 18px; box-shadow: 0 1px 4px rgba(20,30,45,0.05); }
    .g-kpi::before { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: var(--kpi-accent, var(--teal)); }
    .g-kpi.active { background: color-mix(in srgb, var(--kpi-accent, var(--teal)) 8%, var(--panel)); }
    .g-kpi-label { font-family: var(--sans); font-weight: 600; letter-spacing: .6px; }
    .g-kpi-value { font-weight: 700; letter-spacing: -.3px; }
    .g-section-label { font-family: var(--sans); font-weight: 700; font-size: 11px; letter-spacing: 1.2px; color: var(--text-dim); display: flex; align-items: center; gap: 8px; }
    .g-section-label::before { content: ""; width: 14px; height: 3px; border-radius: 3px; background: var(--accent); }
    .g-filterbar { border: 1px solid var(--border); border-radius: 12px !important; background: var(--panel); box-shadow: 0 1px 4px rgba(20,30,45,0.04); }
    .g-field label { font-family: var(--sans); font-weight: 600; }
    .g-field input, .g-field select { border-radius: 7px; background: var(--panel); padding: 7px 10px; }
    .g-field input:focus, .g-field select:focus { box-shadow: 0 0 0 3px var(--accent-dim); }
    .g-mode-toggle { background: var(--panel); border-radius: 10px; padding: 3px; gap: 2px; }
    .g-mode-toggle button { background: transparent; border-radius: 8px; padding: 7px 14px; }
    .g-mode-toggle button.active { background: var(--accent); box-shadow: 0 2px 6px rgba(59,130,246,0.3); }
    .g-table th { background: var(--panel-raised); font-family: var(--sans); padding: 10px 8px; }
    .g-table tr.g-row:nth-child(even) { background: #FAFBFD; }
    .g-table tr.g-row:hover { background: var(--accent-dim); }
    .g-table td { padding: 8px; }
    .g-table-wrap { border-radius: 8px; }
    .g-alert { border-radius: 10px; }
    /* ---------- WP tag (signature element) ---------- */
    .g-tag {
      display: inline-flex; align-items: center;
      font-family: var(--mono); font-size: 10.5px; font-weight: 600;
      color: var(--accent);
      background: rgba(59,130,246,0.08);
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
      position: sticky; top: 0; background: var(--panel); z-index: 3;
    }
    .g-table td {
      padding: 6px 8px; border-bottom: 1px solid var(--border-soft);
      font-size: 12px; vertical-align: middle; white-space: nowrap;
    }
    .g-table tr.g-row:hover { background: var(--panel-alt); }
    @keyframes g-row-flash-anim {
      0% { background: color-mix(in srgb, var(--accent) 35%, var(--panel)); }
      100% { background: transparent; }
    }
    .g-row-flash { animation: g-row-flash-anim 2.2s ease-out; scroll-margin-top: 90px; }
    .g-panel-flash { animation: g-row-flash-anim 2.2s ease-out; scroll-margin-top: 90px; }
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
    .g-gantt-row { display: flex; align-items: center; min-height: 68px; padding: 8px 10px; margin-bottom: 4px; border-radius: 6px; background: var(--panel); border: 1px solid var(--border-soft); transition: border-color .12s, background .12s; }
    .g-gantt-row:hover { border-color: var(--border); background: var(--panel-alt); }
    .g-gantt-taskinfo { width: 620px; flex-shrink: 0; padding-right: 14px; }

    /* ---------- grade de dados estilo MS Project (coluna fixa) + linha do tempo (coluna scrollável) ---------- */
    .g-gantt-datagrid-wrap {
      display: flex; align-items: flex-start; border: 1px solid var(--border);
      border-radius: 6px; overflow: hidden; margin-top: 4px; background: var(--panel);
    }
    .g-gantt-grid-col { flex-shrink: 0; border-right: 1px solid var(--border); }
    .g-gantt-timeline-col { flex: 1; overflow-x: auto; min-width: 260px; }
    .g-gantt-grid-headrow {
      display: flex; align-items: center; height: 30px; background: var(--panel-raised);
      border-bottom: 1px solid var(--border); font-family: var(--mono); font-size: 8.5px;
      text-transform: uppercase; letter-spacing: .4px; color: var(--text-faint); padding: 0 6px;
    }
    .g-gantt-gridrow {
      display: flex; align-items: center; height: 34px; padding: 0 6px 0 8px;
      border-bottom: 1px solid var(--border-soft); transition: background .12s;
    }
    .g-gantt-gridrow:nth-child(even) { background: rgba(10,14,20,0.02); }
    .g-gantt-gridrow:hover { background: var(--panel-alt); }
    .g-gantt-timeline-row { display: flex; align-items: center; height: 34px; border-bottom: 1px solid var(--border-soft); padding: 0 6px; position: relative; }
    .g-gantt-timeline-row:nth-child(even) { background: rgba(10,14,20,0.02); }
    .g-gantt-col-num { width: 20px; flex-shrink: 0; text-align: center; font-family: var(--mono); font-size: 9.5px; color: var(--text-faint); }
    .g-gantt-col-task { width: 190px; flex-shrink: 0; padding: 0 6px; overflow: hidden; }
    .g-gantt-col-empresa { width: 84px; flex-shrink: 0; padding: 0 4px; overflow: hidden; }
    .g-gantt-col-dur { width: 50px; flex-shrink: 0; padding: 0 4px; }
    .g-gantt-col-date { width: 122px; flex-shrink: 0; padding: 0 4px; }
    .g-gantt-col-status { width: 148px; flex-shrink: 0; padding: 0 4px; }
    .g-gantt-col-progress { width: 82px; flex-shrink: 0; padding: 0 4px; }
    .g-gantt-col-action { width: 24px; flex-shrink: 0; text-align: center; }
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
      background: rgba(239,68,68,0.08); border: 1px solid rgba(239,68,68,0.35);
      color: #F6A79E; padding: 8px 10px; border-radius: 4px; font-size: 11.5px; margin-bottom: 8px;
    }
  `}</style>
);

/* ============================================================
   HELPERS
   ============================================================ */
const pad2 = (n) => String(n).padStart(2, "0");
const LOGO_MARK = "data:image/png;base64,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";
const LOGO_FULL = "data:image/png;base64,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";
const fmt = (n) => "R$ " + Number(n || 0).toLocaleString("pt-BR");
const fmtUsd = (n) => "US$ " + Number(n || 0).toLocaleString("en-US");
/* para categorias sem orçamento definido (ex: CAPEX), mostra "—" em vez de "US$ 0" / "R$ 0" */
const fmtBudgetUsd = (usd) => (usd ? fmtUsd(usd) : "—");
const fmtBudgetBrl = (usd, brl) => (usd ? fmt(brl) : "—");

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

/* ============================================================
   PDF REPORT HELPERS — cada aba registra seu próprio gerador de
   relatório (via setReportFn), refletindo exatamente o que está
   sendo mostrado/filtrado naquela página no momento do clique.
   ============================================================ */
const PDF_NAVY = [22, 24, 29];
const PDF_ACCENT = [59, 130, 246];
const PDF_TEXT = [30, 34, 42];
const PDF_MUTED = [120, 128, 140];

const pdfHeader = (doc, pageTitle, subtitle) => {
  // faixa de marca no topo
  doc.setFillColor(...PDF_NAVY);
  doc.rect(0, 0, 210, 26, "F");
  doc.setFillColor(...PDF_ACCENT);
  doc.rect(0, 26, 210, 1.1, "F");

  doc.setFontSize(16);
  doc.setFont(undefined, "bold");
  doc.setTextColor(255, 255, 255);
  doc.text("GENESIS I", 14, 12);

  doc.setFontSize(8.5);
  doc.setFont(undefined, "normal");
  doc.setTextColor(200, 205, 215);
  doc.text("Maintenance & Port Call", 14, 17.5);

  doc.setFontSize(11);
  doc.setFont(undefined, "bold");
  doc.setTextColor(...PDF_ACCENT);
  doc.text(pageTitle, 196, 12, { align: "right" });

  doc.setFontSize(8);
  doc.setFont(undefined, "normal");
  doc.setTextColor(210, 210, 210);
  doc.text(subtitle, 196, 18, { align: "right", maxWidth: 140 });

  doc.setTextColor(...PDF_TEXT);
  return 36;
};

/* grade de KPIs com destaque visual: número grande, faixa colorida à esquerda, rótulo abaixo */
const pdfKpis = (doc, y, kpis) => {
  const perRow = 4;
  const gap = 4;
  const cardW = (196 - 14 - gap * (perRow - 1)) / perRow;
  const valueFontSize = 11;
  const labelFontSize = 6.5;
  const lineH = 4.4;

  doc.setFontSize(valueFontSize);
  doc.setFont(undefined, "bold");
  const wrapped = kpis.map((k) => doc.splitTextToSize(String(k.value), cardW - 6));
  const maxLines = Math.max(1, ...wrapped.map((w) => w.length));
  const cardH = 6 + maxLines * lineH + 6;
  const rows = Math.ceil(kpis.length / perRow);
  const blockHeight = rows * (cardH + 4) + 8;

  /* nunca deixa o bloco de KPIs nascer perto demais do rodapé — se não sobrar espaço pra ele
     inteiro, pula pra próxima página automaticamente */
  if (y + blockHeight > 283) { doc.addPage(); y = 15; }

  kpis.forEach((k, idx) => {
    const col = idx % perRow;
    const row = Math.floor(idx / perRow);
    const x = 14 + col * (cardW + gap);
    const cardY = y + row * (cardH + 4);

    doc.setFillColor(248, 249, 251);
    doc.roundedRect(x, cardY, cardW, cardH, 1.2, 1.2, "F");
    doc.setFillColor(...PDF_ACCENT);
    doc.rect(x, cardY, 1.2, cardH, "F");

    doc.setFontSize(valueFontSize);
    doc.setFont(undefined, "bold");
    doc.setTextColor(...PDF_TEXT);
    const lines = wrapped[idx];
    lines.forEach((line, li) => {
      doc.text(line, x + 4, cardY + 6 + li * lineH);
    });

    doc.setFontSize(labelFontSize);
    doc.setFont(undefined, "normal");
    doc.setTextColor(...PDF_MUTED);
    doc.text(k.label.toUpperCase(), x + 4, cardY + 6 + lines.length * lineH + 2, { maxWidth: cardW - 6 });
  });
  doc.setTextColor(...PDF_TEXT);
  return y + blockHeight;
};

const pdfSectionTitle = (doc, y, title) => {
  /* nunca deixa o título nascer perto demais do rodapé — se não sobrar espaço, pula pra
     próxima página automaticamente, sem precisar de checagem manual espalhada pelo relatório */
  if (y > 268) { doc.addPage(); y = 15; }
  doc.setFillColor(...PDF_ACCENT);
  doc.rect(14, y - 3.2, 2.2, 4.2, "F");
  doc.setFontSize(10.5);
  doc.setFont(undefined, "bold");
  doc.setTextColor(...PDF_TEXT);
  doc.text(title, 19, y);
  doc.setFont(undefined, "normal");
  doc.setTextColor(0, 0, 0);
  return y + 5;
};

const pdfTable = (doc, y, columns, rows, opts = {}) => {
  autoTable(doc, {
    startY: y,
    head: [columns],
    body: rows,
    styles: { fontSize: 7.5, cellPadding: 2.2, textColor: PDF_TEXT, overflow: "linebreak" },
    headStyles: { fillColor: PDF_NAVY, textColor: 255, fontStyle: "bold", fontSize: 7.5 },
    alternateRowStyles: { fillColor: [246, 247, 249] },
    margin: { left: 14, right: 14 },
    columnStyles: opts.columnStyles || {},
  });
  return doc.lastAutoTable.finalY + 10;
};

/* rodapé com numeração de página e marca d'água discreta em todas as páginas */
const pdfSave = (doc, filenamePrefix) => {
  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setDrawColor(225, 225, 225);
    doc.line(14, 285, 196, 285);
    doc.setFontSize(7);
    doc.setFont(undefined, "normal");
    doc.setTextColor(...PDF_MUTED);
    doc.text(`GENESIS I — Maintenance & Port Call · Gerado em ${new Date().toLocaleString("pt-BR")}`, 14, 290);
    doc.text(`Página ${i} de ${pageCount}`, 196, 290, { align: "right" });
  }
  doc.save(`${filenamePrefix}-${todayISO()}.pdf`);
};

/* mini-Gantt visual no PDF: uma barra colorida por tarefa, posicionada por hora dentro do
   intervalo do Port Call — dá o efeito de cronograma tipo MS Project no relatório */
const pdfMiniGantt = (doc, y, span, activities) => {
  const chartX = 14, chartW = 130;
  const totalHours = Math.max(1, span.hours);
  const rowH = 5.2;
  if (y + activities.length * rowH + 10 > 280) { doc.addPage(); y = 15; }
  doc.setFontSize(7);
  doc.setTextColor(...PDF_MUTED);
  doc.text(fmtPeriodDate(span.start), chartX, y);
  doc.text(fmtPeriodDate(span.end), chartX + chartW, y, { align: "right" });
  y += 3;
  doc.setDrawColor(225, 225, 225);
  doc.rect(chartX, y, chartW, activities.length * rowH, "S");
  activities.forEach((w, idx) => {
    const rowY = y + idx * rowH;
    const startH = Math.max(0, (new Date(w.start) - span.start) / 3600000);
    const endH = Math.min(totalHours, (new Date(w.end) - span.start) / 3600000);
    const durH = Math.max(0, endH - startH);
    const barX = chartX + (startH / totalHours) * chartW;
    const barW = Math.max(1.2, (durH / totalHours) * chartW);
    const hex = WP_STATUS_COLOR[w.status] || "#F2C94C";
    const rgb = [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];
    doc.setFillColor(...rgb);
    doc.roundedRect(barX, rowY + 0.6, barW, rowH - 1.4, 0.6, 0.6, "F");
    doc.setFontSize(6);
    doc.setTextColor(60, 60, 60);
    const label = w.name.length > 26 ? w.name.slice(0, 24) + "…" : w.name;
    doc.text(label, chartX + chartW + 2, rowY + rowH - 1.6);
  });
  return y + activities.length * rowH + 8;
};

const todayISO = () => new Date().toISOString().slice(0, 10);

const MONTH_NAMES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

const WP_STATUS = ["Planejamento", "Não iniciado", "Em andamento", "Concluído", "Cancelado"];
const WP_STATUS_DEFAULT_PROGRESS = { "Planejamento": 0, "Não iniciado": 0, "Em andamento": 50, "Concluído": 100, "Cancelado": 0 };
const MAT_STATUS = ["Solicitado", "Em aprovação", "Cotação", "Cotação recebida", "Em aprovação comercial", "PO emitida", "Em fabricação", "Em trânsito", "Recebido", "Entregue a bordo", "Dentro do Prazo", "Fora do Prazo"];
const PAY_STATUS = ["Orçamento", "Aprovado", "PO emitida", "Serviço executado", "Medição aprovada", "NF recebida", "NF validada", "Pagamento programado", "Pago"];
const PRIORITY = ["Baixa", "Média", "Alta", "Crítica", "Importante", "Emergencial", "Sobressalente crítico"];
const IMPACT_LEVELS = ["Baixo", "Médio", "Alto", "Crítico"];
const IMPACT_COLOR = { "Baixo": "#8D9BB5", "Médio": "#F2C94C", "Alto": "#F2A93B", "Crítico": "#E0483E" };
const PLAN_STATUS = ["A Executar", "Em Andamento", "Concluído", "Cancelado"];
const PLAN_STATUS_PADRAO = PLAN_STATUS.filter((x) => x !== "Concluído"); /* filtro automático: tudo menos Concluído */
const mesmoConjunto = (a, b) => a.length === b.length && a.every((x) => b.includes(x));
const PLAN_STATUS_COLOR = { "A Executar": "#8D9BB5", "Em Andamento": "#3FC1C9", "Concluído": "#35D399", "Cancelado": "#6B7280" };
/* categories + Orçado (USD) exactly as in the uploaded drill-down report */
const CATEGORIES = ["Elétrica", "Hse", "Hull & Structure", "Integridade", "Lubrificantes", "Marine", "Mecânica", "R&R Elétrica", "R&R Mecânica", "CAPEX"];
const CATEGORY_BUDGET_USD = {
  "Elétrica": 27900,
  "Hse": 6200,
  "Hull & Structure": 12772,
  "Integridade": 12400,
  "Lubrificantes": 31000,
  "Marine": 21700,
  "Mecânica": 27900,
  "R&R Elétrica": 18600,
  "R&R Mecânica": 18600,
  /* CAPEX não tem orçamento mensal fixo definido — fica sem valor mesmo */
};
const DISCIPLINES = CATEGORIES;

/* Código "Ordem" interno de referência para "Compra de Serviços" por categoria — vem direto da
   planilha de referência da empresa. Associado automaticamente a cada categoria. */
const CATEGORY_ADP_SERVICOS = {
  "Hse": "307156",
  "Hull & Structure": "305203",
  "Marine": "307157",
  "Mecânica": "307164",
  "Elétrica": "307158",
  "Integridade": "306505",
  "Lubrificantes": "307161",
  "R&R Mecânica": "307155",
  "R&R Elétrica": "307165",
};
const adpServicosLabel = (categoria) => CATEGORY_ADP_SERVICOS[categoria] || "—";

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
  ["planoAcao", "Plano de Ação"], ["precisaMaterial", "Precisa de Material"], ["materialNecessario", "Material Necessário"],
  ["impacto", "Impacto"], ["novaDataPrevista", "Nova Data Prevista"],
];
const PLAN_COLS = [
  ["nome", "Nome"], ["departamento", "Departamento"], ["empresa", "Empresa"], ["descricaoProblema", "Descrição do Problema"],
  ["planoAcao", "Plano de Ação"], ["rc", "RC"], ["obs", "Observação"],
  ["precisaMaterial", "Precisa de Material"], ["materialNecessario", "Material Necessário"], ["poMaterial", "PO"],
  ["impacto", "Impacto"], ["status", "Status"], ["dataExecucao", "Data de Execução"],
];
const DOC_COLS = [
  ["nome", "Nome"], ["localizacao", "Localização"], ["tipoPeriodo", "Tipo"], ["empresa", "Empresa"], ["planoAcao", "Plano de Ação"],
  ["necessitaMaterial", "Necessita de Material"], ["poRelacionada", "PO Relacionada"],
  ["previsaoExecucao", "Previsão de Execução"], ["dataConclusao", "Data de Conclusão"], ["status", "Status"],
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
const TM_DUE_COLS = [
  ["code", "Code"], ["component", "Component"], ["jobType", "Job type"], ["jobNo", "Job no"], ["status", "Status"],
  ["jobName", "Job name"], ["interval", "Int"], ["hours", "Hours"], ["dueRaw", "Due"], ["diffRaw", "Diff"],
  ["pri", "Pri"], ["department", "Department"], ["estimatedDue", "EstimatedDue"],
  ["lastDoneDate", "LastDoneDate"], ["lastDoneHours", "LastDoneHours"],
  ["planoAcao", "Plano de Ação"], ["linkedPlanId", "Adicionado ao Planejamento"],
];
const TM_HISTORY_COLS = [
  ["jobHistoryNumber", "Job History Number"], ["componentCode", "ComponentCode"], ["componentName", "ComponentName"],
  ["dateDone", "DateDone"], ["jobType", "JobType"], ["jobNo", "JobNo"], ["jobName", "JobName"],
  ["doneByName", "DoneByName"], ["serviceReport", "ServiceReport"], ["remarks", "Remarks"], ["reason", "Reason"],
  ["jobPriority", "JobPriority"], ["dateSigned", "Date signed"], ["hoursDone", "Hours done"],
  ["dueHours", "Due hours"], ["dueDate", "Due date"], ["interval", "Interval"], ["signedBy", "Signed by"],
];
const STATUS_PAGAMENTO_OPTIONS = [
  "Aguardando Orçamento", "Aguardando Suprimentos", "Aguardando Execução", "Aguardando Medição", "Aprovação Pendente",
  "Aguardando NF", "Pagamento Programado", "On Hold", "Pago", "Cancelado",
];
const STATUS_PAGAMENTO_COLOR = {
  "Aguardando Orçamento": "#8D9BB5",
  "Aguardando Suprimentos": "#F2C94C",
  "Aguardando Execução": "#4FA8D8",
  "Aguardando Medição": "#3FC1C9",
  "Aprovação Pendente": "#F2685B",
  "Aguardando NF": "#F2A93B",
  "Pagamento Programado": "#9B8CF2",
  "On Hold": "#6B7280",
  "Pago": "#35D399",
  "Cancelado": "#5D6E8C",
};
const INV_COLS = [
  ["id", "ID"], ["date", "Data"], ["assunto", "Manutenção"], ["empresa", "Empresa"], ["md", "MD"],
  ["mdSentDate", "Data de Envio da MD"], ["diffDays", "Diferença de Dias"], ["daysOpenTotal", "Dias em Aberto Total"],
  ["rc", "RC"], ["serviceStatus", "Status do Serviço"], ["poContrato", "PO / Contrato"], ["medicao", "Medição"],
  ["valorTotal", "Valor Total"], ["saldoPo", "Saldo PO"], ["obs", "Observações"],
  ["statusPagamento", "Status Pagamento"], ["dataPagamento", "Data de Pagamento"], ["previsaoMes", "Previsão (mês)"], ["justificativaGeral", "Justificativa Geral do Rateio"],
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
const ETextArea = ({ value, onChange, rows = 2 }) => {
  const ref = useRef(null);
  const autoSize = (el) => {
    if (!el) return;
    el.style.height = "auto";
    el.style.height = el.scrollHeight + "px";
  };
  React.useEffect(() => { autoSize(ref.current); }, [value]);
  return (
    <textarea
      ref={ref}
      className="g-edit-wrap"
      rows={rows}
      value={value || ""}
      onChange={(e) => { autoSize(e.target); onChange(e.target.value); }}
      style={{ overflow: "hidden", resize: "none" }}
    />
  );
};
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
/* KPI "cartão grande" compartilhado por todas as abas — versão única (antes estava duplicada
   localmente em vários componentes, o que impedia componentes novos de usá-la) */
const bigKpi = (label, value, color, Icon) => (
  <div className="g-kpi" style={{ "--kpi-accent": color, padding: "16px" }}>
    <div className="g-flex" style={{ justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
      <div className="g-kpi-label" style={{ fontSize: 11 }}>{label}</div>
      {Icon && (
        <div style={{ width: 28, height: 28, borderRadius: 7, background: color, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Icon size={14} style={{ color: "#fff" }} />
        </div>
      )}
    </div>
    <div className="g-kpi-value" style={{ fontSize: 24, color: "var(--text)" }}>{value}</div>
  </div>
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
  const aEmpty = av === "" || av === null || av === undefined;
  const bEmpty = bv === "" || bv === null || bv === undefined;
  /* valores em branco sempre vão para o final, independente da direção da ordenação */
  if (aEmpty && bEmpty) return 0;
  if (aEmpty) return 1;
  if (bEmpty) return -1;
  const an = parseFloat(av), bn = parseFloat(bv);
  const numeric = !isNaN(an) && !isNaN(bn);
  if (numeric) return (an - bn) * dir;
  av = av.toString().toLowerCase();
  bv = bv.toString().toLowerCase();
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
        <img src={LOGO_FULL} alt="OSM Thome" className="g-login-logo" />
        <div className="g-login-brand" style={{ textAlign: "center", letterSpacing: 4 }}>GENESIS I</div>
        <div className="g-login-sub" style={{ textAlign: "center" }}>Maintenance &amp; Port Call</div>
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
/* Planejamento — lista independente de "mapeados para execução": nasce aqui, sem data de execução
   ainda, e só passa a existir também como um Serviço de verdade (na aba Serviços/Gantt) quando
   ganhar uma data de execução */
const INITIAL_PLANNING_ITEMS = [];
/* Docagem — itens de Machinery Items da DNV que precisam ser vistoriados/feitos durante a docagem */
const INITIAL_DOCAGEM_ITEMS = [];
/* TM Master — importado semanalmente do sistema de manutenção. "Due" é sempre substituído por
   completo a cada importação (é uma fotografia do que está em aberto agora); "History" é mesclado
   por "Job History Number" (é um histórico cumulativo desde o início do ano) */
const INITIAL_TM_DUE = [];
const INITIAL_TM_HISTORY = [];
/* histórico de "fotografias" do tamanho do backlog Due a cada importação semanal — permite ver a
   tendência (o backlog está crescendo ou diminuindo semana a semana) */
const INITIAL_TM_DUE_SNAPSHOTS = [];

export default function Root() {
  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState(null);
  /* trava de segurança: só habilita o auto-save quando o GET /api/state realmente teve sucesso.
     Se a conexão falhar (backend reiniciando, erro de rede, etc.), NUNCA deixamos o app salvar —
     senão o estado local (ainda nos valores INITIAL_* vazios) seria gravado por cima dos dados
     reais do servidor 700ms depois, apagando tudo. Ver efeito de load logo abaixo. */
  const [saveEnabled, setSaveEnabled] = useState(false);
  const [loadAttempt, setLoadAttempt] = useState(0);

  const [users, setUsers] = useState(DEFAULT_USERS);
  const [currentUser, setCurrentUser] = useState(null);

  const [workPackages, setWorkPackages] = useState(INITIAL_WORK_PACKAGES);
  const [materials, setMaterials] = useState(INITIAL_MATERIALS);
  const [payments, setPayments] = useState(INITIAL_PAYMENTS);
  const [serviceInvoices, setServiceInvoices] = useState(INITIAL_SERVICE_INVOICES);
  const [portCallMeta, setPortCallMeta] = useState(INITIAL_PORT_CALL_META);
  const [opCategories, setOpCategories] = useState(INITIAL_OP_CATEGORIES);
  const [exchangeRate, setExchangeRate] = useState(INITIAL_EXCHANGE_RATE);
  const [planningItems, setPlanningItems] = useState(INITIAL_PLANNING_ITEMS);
  const [docagemItems, setDocagemItems] = useState(INITIAL_DOCAGEM_ITEMS);
  const [tmDue, setTmDue] = useState(INITIAL_TM_DUE);
  const [tmHistory, setTmHistory] = useState(INITIAL_TM_HISTORY);
  const [tmDueSnapshots, setTmDueSnapshots] = useState(INITIAL_TM_DUE_SNAPSHOTS);

  /* carrega o estado salvo assim que o site abre — igual para qualquer usuário que entrar.
     saveEnabled só vira true dentro do "then" de sucesso (mesmo que o servidor não tenha
     nenhum dado ainda, o que é normal na primeiríssima vez). Em caso de falha (rede caiu,
     backend reiniciando/redeployando, resposta inesperada), NÃO liberamos o app: mostramos uma
     tela de erro com "Tentar novamente" em vez de deixar o usuário mexer em cima de um estado
     vazio que seria salvo por cima do que já existe no servidor. */
  React.useEffect(() => {
    let cancelled = false;
    setLoaded(false);
    setLoadError(null);
    fetch("/api/state")
      .then((r) => {
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        return r.json();
      })
      .then((res) => {
        if (cancelled) return;
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
          if (d.planningItems) setPlanningItems(d.planningItems);
          if (d.docagemItems) setDocagemItems(d.docagemItems);
          if (d.tmDue) setTmDue(d.tmDue);
          if (d.tmHistory) setTmHistory(d.tmHistory);
          if (d.tmDueSnapshots) setTmDueSnapshots(d.tmDueSnapshots);
        }
        setSaveEnabled(true);
        setLoaded(true);
      })
      .catch(() => {
        if (cancelled) return;
        setLoadError("Não foi possível carregar os dados do servidor. Suas alterações NÃO seriam salvas com segurança agora, então o acesso ficou bloqueado até a conexão voltar — clique em \"Tentar novamente\".");
        setSaveEnabled(false);
        setLoaded(true);
      });
    return () => { cancelled = true; };
  }, [loadAttempt]);

  /* salva no servidor (com um pequeno atraso) toda vez que qualquer coisa muda — adicionar, editar,
     excluir linhas em qualquer aba — assim fica salvo automaticamente para todo mundo que acessar,
     independente de qual usuário fez a alteração. Só roda depois de um load bem-sucedido
     (saveEnabled) — nunca com dados ainda não confirmados vindos do servidor. */
  React.useEffect(() => {
    if (!loaded || !saveEnabled) return;
    const t = setTimeout(() => {
      fetch("/api/state", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          data: { users, workPackages, materials, payments, serviceInvoices, portCallMeta, opCategories, exchangeRate, planningItems, docagemItems, tmDue, tmHistory, tmDueSnapshots },
        }),
      }).catch(() => setLoadError("Não foi possível salvar no servidor agora. Suas alterações ficam só neste navegador até a conexão voltar."));
    }, 700);
    return () => clearTimeout(t);
  }, [loaded, saveEnabled, users, workPackages, materials, payments, serviceInvoices, portCallMeta, opCategories, exchangeRate, planningItems, docagemItems, tmDue, tmHistory, tmDueSnapshots]);

  if (!loaded) {
    return (
      <div className="genesis g-login-wrap"><Theme />
        <div style={{ color: "var(--text-dim)", fontFamily: "var(--mono)", fontSize: 13 }}>Carregando dados…</div>
      </div>
    );
  }

  /* falha ao carregar: tela bloqueante (não deixa entrar no app com dados vazios) */
  if (loadError && !saveEnabled) {
    return (
      <div className="genesis g-login-wrap"><Theme />
        <div style={{ maxWidth: 420, textAlign: "center", display: "flex", flexDirection: "column", gap: 14, alignItems: "center" }}>
          <AlertTriangle size={32} color="var(--crit)" />
          <div style={{ color: "var(--text)", fontSize: 14, fontWeight: 600 }}>Não foi possível carregar os dados</div>
          <div style={{ color: "var(--text-dim)", fontSize: 12.5, lineHeight: 1.5 }}>{loadError}</div>
          <button className="g-btn primary" onClick={() => setLoadAttempt((n) => n + 1)}>Tentar novamente</button>
        </div>
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
      planningItems={planningItems} setPlanningItems={setPlanningItems}
      docagemItems={docagemItems} setDocagemItems={setDocagemItems}
      tmDue={tmDue} setTmDue={setTmDue} tmHistory={tmHistory} setTmHistory={setTmHistory}
      tmDueSnapshots={tmDueSnapshots} setTmDueSnapshots={setTmDueSnapshots}
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
  opCategories, setOpCategories, exchangeRate, setExchangeRate,
  planningItems, setPlanningItems, docagemItems, setDocagemItems,
  tmDue, setTmDue, tmHistory, setTmHistory, tmDueSnapshots, setTmDueSnapshots, loadError }) {
  const [tab, setTab] = useState("dashboard");
  React.useEffect(() => { setReportFn(null); setExportXlsxFn(null); }, [tab]);
  const [newRowId, setNewRowId] = useState(null);
  /* usado sempre que uma linha nova é criada (novo serviço, novo material, novo registro de pagamento):
     guarda o id por alguns segundos pra a linha poder ser destacada e "scrollada" até a visão do usuário */
  const flashNewRow = (id) => {
    setNewRowId(id);
    setTimeout(() => setNewRowId((cur) => (cur === id ? null : cur)), 2500);
  };
  const [expandedWp, setExpandedWp] = useState(null);
  const [paySubTab, setPaySubTab] = useState("total"); // "total" | "status" | "dashboard"
  const [planSubTab, setPlanSubTab] = useState("mapeados"); // "mapeados" | "docagem" | "board"
  const [tmSubTab, setTmSubTab] = useState("due"); // "due" | "history" | "metricas"
  const [importMsg, setImportMsg] = useState(null);
  const [reportFn, setReportFn] = useState(null);
  /* mesmo princípio do reportFn (PDF), mas pra "Exportar planilha": cada aba pode registrar sua
     própria exportação em Excel, refletindo exatamente o que está filtrado ali. Abas que não
     registram nada continuam usando o exportador global (handleExportXlsx, com tudo do sistema). */
  const [exportXlsxFn, setExportXlsxFn] = useState(null);
  const fileInputRef = useRef(null);

  /* global period filter — present on every page */
  const [period, setPeriod] = useState({
    mode: "periodo", // "mes" | "periodo" | "ano"
    month: 8,
    year: 2026,
    start: "2020-01-01",
    end: "2030-12-31",
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
  const updPlan = upd(setPlanningItems), remPlan = rem(setPlanningItems);
  const updDocagem = upd(setDocagemItems), remDocagem = rem(setDocagemItems);
  const updTmDue = upd(setTmDue), remTmDue = rem(setTmDue);
  const updTmHistory = upd(setTmHistory), remTmHistory = rem(setTmHistory);


  /* ---------- TM Master: parsing helpers ---------- */
  const parseTmDiff = (diff) => {
    const s = (diff || "").toString().trim();
    const m = s.match(/^(-?\d+(?:\.\d+)?)\s*([DH])$/i);
    if (!m) return { value: null, unit: null };
    return { value: parseFloat(m[1]), unit: m[2].toUpperCase() };
  };
  const parseTmDate = (v) => {
    if (!v) return "";
    if (v instanceof Date) return `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}`;
    const s = v.toString().trim();
    const m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    return ""; // valores baseados em horas de máquina (ex.: "43000H") não são uma data de calendário
  };

  /* "Due" é uma fotografia do que está em aberto agora — cada importação SUBSTITUI a lista inteira,
     senão itens já resolvidos ficariam presos pra sempre */
  const handleImportTmDue = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        const rows = json.map((row) => {
          const diffRaw = row["Diff"];
          const { value: diffValue, unit: diffUnit } = parseTmDiff(diffRaw);
          const dueRaw = row["Due"];
          return {
            id: uid("TMD"),
            code: (row["Code"] || "").toString(),
            component: (row["Component"] || "").toString(),
            jobType: (row["Job type"] || "").toString().trim(),
            jobNo: row["Job no"],
            status: (row["Status"] || "").toString(),
            jobName: (row["Job name"] || "").toString(),
            interval: (row["Int"] || "").toString(),
            hours: row["Hours"],
            dueRaw: dueRaw ? dueRaw.toString() : "",
            dueDate: parseTmDate(dueRaw),
            diffRaw: diffRaw ? diffRaw.toString() : "",
            diffValue, diffUnit,
            pri: (row["Pri"] || "").toString().trim() || "Não definida",
            department: (row["Department"] || "").toString().trim() || "Não definido",
            estimatedDue: parseTmDate(row["EstimatedDue"]),
            lastDoneDate: parseTmDate(row["LastDoneDate"]),
            lastDoneHours: row["LastDoneHours"],
          };
        }).filter((r) => r.code || r.jobName);
        /* a importação do Due vem fresca da planilha e gera ids novos a cada vez — pra não perder o
           "Plano de Ação" preenchido nem o vínculo "Adicionado ao Planejamento" quando o usuário
           reimporta uma planilha atualizada, casa pelo par Code+Job no com a lista anterior e
           preserva esses dois campos (e o id) quando encontra o mesmo item */
        setTmDue((prev) => {
          const byKey = new Map();
          prev.forEach((r) => { if (r.code) byKey.set(`${r.code}|${r.jobNo}`, r); });
          return rows.map((r) => {
            const old = r.code ? byKey.get(`${r.code}|${r.jobNo}`) : null;
            return old
              ? { ...r, id: old.id, planoAcao: old.planoAcao || "", linkedPlanId: old.linkedPlanId || null }
              : { ...r, planoAcao: "", linkedPlanId: null };
          });
        });
        const vencidasCount = rows.filter((r) => r.diffValue !== null && r.diffValue < 0).length;
        const ateVencer40Count = rows.filter((r) => r.diffUnit === "D" && r.diffValue !== null && r.diffValue >= 0 && r.diffValue <= 40).length;
        setTmDueSnapshots((prev) => {
          const today = todayISO();
          /* uma fotografia por dia — se importar mais de uma vez no mesmo dia, atualiza a mesma entrada
             em vez de acumular várias fotografias do mesmo dia */
          const withoutToday = prev.filter((s) => s.date !== today);
          return [...withoutToday, { date: today, total: rows.length, vencidas: vencidasCount, ateVencer40: ateVencer40Count }].sort((a, b) => a.date.localeCompare(b.date));
        });
        setImportMsg(`TM Master - Due importado: ${rows.length} itens (lista substituída pela mais recente).`);
      } catch (err) {
        setImportMsg("Erro ao ler a planilha TM Master - Due. Confira se as colunas seguem o padrão esperado.");
      }
      setTimeout(() => setImportMsg(null), 6000);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  /* "History" é cumulativo desde o início do ano — mescla por "Job History Number" (chave única de
     cada job concluído), então reimportar a planilha atualizada só adiciona o que é novo */
  const handleImportTmHistory = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        const parsedRows = json.map((row) => ({
          jobHistoryNumber: (row["Job History Number"] || "").toString().trim(),
          componentCode: (row["ComponentCode"] || "").toString(),
          componentName: (row["ComponentName"] || "").toString(),
          dateDone: parseTmDate(row["DateDone"]),
          jobType: (row["JobType"] || "").toString().trim(),
          jobNo: row["JobNo"],
          jobName: (row["JobName"] || "").toString(),
          doneByName: (row["DoneByName"] || "").toString(),
          serviceReport: (row["ServiceReport"] || "").toString(),
          remarks: (row["Remarks"] || "").toString(),
          reason: (row["Reason"] || "").toString(),
          jobPriority: (row["JobPriority"] || "").toString(),
          dateSigned: parseTmDate(row["Date signed"]),
          hoursDone: row["Hours done"],
          dueHours: row["Due hours"],
          dueDate: parseTmDate(row["Due date"]),
          interval: (row["Interval"] || "").toString(),
          signedBy: (row["Signed by"] || "").toString(),
        })).filter((r) => r.jobHistoryNumber);

        setTmHistory((prev) => {
          const byKey = new Map(prev.map((h, idx) => [h.jobHistoryNumber, idx]));
          const next = [...prev];
          let added = 0, updated = 0;
          parsedRows.forEach((row) => {
            if (byKey.has(row.jobHistoryNumber)) {
              const idx = byKey.get(row.jobHistoryNumber);
              next[idx] = { ...next[idx], ...row };
              updated++;
            } else {
              next.push(row);
              added++;
            }
          });
          setImportMsg(`TM Master - History importado: ${added} novo(s), ${updated} atualizado(s).`);
          return next;
        });
      } catch (err) {
        setImportMsg("Erro ao ler a planilha TM Master - History. Confira se as colunas seguem o padrão esperado.");
      }
      setTimeout(() => setImportMsg(null), 6000);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const addDocagemItem = () => {
    const id = uid("DOC");
    setDocagemItems((r) => [...r, {
      id, nome: "Novo item", localizacao: "", tipoPeriodo: "", empresa: "", planoAcao: "",
      necessitaMaterial: false, poRelacionada: "", previsaoExecucao: "", dataConclusao: "",
      status: "A Executar", linkedServiceId: null,
    }]);
    flashNewRow(id);
  };

  /* Importação da planilha de Machinery Items da DNV (colunas: Name, ItemLocation, PeriodType) —
     mescla por Nome, então reimportar a lista atualizada do site da DNV não duplica itens já
     cadastrados (só atualiza localização/tipo se tiverem mudado, preservando o que já foi preenchido
     manualmente: plano de ação, material, PO, datas e status) */
  const handleImportDocagem = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        const norm = (s) => (s || "").toString().trim();
        const parsedRows = [];
        wb.SheetNames.forEach((sheetName) => {
          const json = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
          json.forEach((row) => {
            const nome = norm(row["Name"] || row["Nome"]);
            if (!nome) return;
            parsedRows.push({
              nome,
              localizacao: norm(row["ItemLocation"] || row["Localização"]),
              tipoPeriodo: norm(row["PeriodType"] || row["Tipo"]),
            });
          });
        });

        setDocagemItems((prev) => {
          const byNome = new Map(prev.map((d, idx) => [d.nome.toLowerCase(), idx]));
          const next = [...prev];
          let added = 0, updated = 0;
          parsedRows.forEach((row) => {
            const key = row.nome.toLowerCase();
            if (byNome.has(key)) {
              const idx = byNome.get(key);
              next[idx] = { ...next[idx], localizacao: row.localizacao, tipoPeriodo: row.tipoPeriodo };
              updated++;
            } else {
              next.push({
                id: uid("DOC"), ...row, empresa: "", planoAcao: "", necessitaMaterial: false,
                poRelacionada: "", previsaoExecucao: "", dataConclusao: "", status: "A Executar", linkedServiceId: null,
              });
              added++;
            }
          });
          setImportMsg(`Docagem importada: ${added} novo(s), ${updated} atualizado(s).`);
          return next;
        });
      } catch (err) {
        setImportMsg("Erro ao ler a planilha de Docagem. Confira se as colunas seguem o padrão esperado (Name/ItemLocation/PeriodType).");
      }
      setTimeout(() => setImportMsg(null), 6000);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };
  const addPlanningItem = () => {
    const id = uid("PLAN");
    setPlanningItems((r) => [...r, {
      id, nome: "Novo mapeamento", departamento: "", empresa: "",
      descricaoProblema: "", planoAcao: "", rc: "", obs: "", status: "A Executar",
      precisaMaterial: false, materialNecessario: "", poMaterial: "", impacto: "Baixo",
      dataExecucao: "", linkedServiceId: null,
    }]);
    flashNewRow(id);
  };

  /* TM Master → coluna "Adicionar ao Planejamento": cria (ou remove) um mapeamento correspondente
     na aba Planejamento a partir de um item do Due, mantendo o vínculo pelos ids (linkedPlanId no
     item do Due / linkedFromTmDueId no item de Planejamento), pra não duplicar se marcar de novo. */
  const onToggleAddToPlanning = (dueRow, checked) => {
    if (checked) {
      if (dueRow.linkedPlanId) return;
      const planId = uid("PLAN");
      setPlanningItems((prev) => [...prev, {
        id: planId,
        nome: dueRow.jobName || dueRow.component || "Item TM Master",
        departamento: dueRow.department || "", empresa: "",
        descricaoProblema: [dueRow.component, dueRow.jobName].filter(Boolean).join(" — "),
        planoAcao: dueRow.planoAcao || "", rc: "",
        obs: dueRow.code ? `TM Master: ${dueRow.code}` : "",
        status: "A Executar", precisaMaterial: false, materialNecessario: "", poMaterial: "",
        impacto: dueRow.pri === "High" ? "Alto" : dueRow.pri === "Medium" ? "Médio" : "Baixo",
        dataExecucao: "", linkedServiceId: null, linkedFromTmDueId: dueRow.id,
      }]);
      setTmDue((prev) => prev.map((r) => (r.id === dueRow.id ? { ...r, linkedPlanId: planId } : r)));
      setImportMsg("Item adicionado à aba Planejamento.");
      flashNewRow(planId);
    } else {
      const planIdToRemove = dueRow.linkedPlanId;
      if (planIdToRemove) setPlanningItems((prev) => prev.filter((p) => p.id !== planIdToRemove));
      setTmDue((prev) => prev.map((r) => (r.id === dueRow.id ? { ...r, linkedPlanId: null } : r)));
      setImportMsg("Item removido da aba Planejamento.");
    }
    setTimeout(() => setImportMsg(null), 4000);
  };

  /* Importação de planilhas de mapeamento (ex.: "Saúde do Ativo", "Planejamento de Manutenção") —
     lê TODAS as abas do arquivo, usando o mesmo padrão de colunas (Equipamento/Título/Área/
     Prioridade/Status/Descrição/MD/RC/Observação). Só traz itens ainda pendentes (não traz o que já
     está Concluído/Realizado). Mescla por Título, pra reimportações não duplicarem. */
  const handleImportPlanejamento = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        const norm = (s) => (s || "").toString().trim();
        const mapImpacto = (p) => {
          const s = norm(p).toLowerCase();
          if (s.startsWith("crít")) return "Crítico";
          if (s.startsWith("alt")) return "Alto";
          if (s.startsWith("méd") || s.startsWith("med")) return "Médio";
          if (s.startsWith("bai")) return "Baixo";
          return "Médio";
        };
        const mapStatus = (s) => {
          const v = norm(s);
          if (v === "Realizado") return "Concluído";
          return PLAN_STATUS.includes(v) ? v : "A Executar";
        };
        const parsedRows = [];
        wb.SheetNames.forEach((sheetName) => {
          const json = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
          json.forEach((row) => {
            const titulo = norm(row["Título"] || row["Titulo"]);
            if (!titulo) return;
            const rc = row["RC"];
            parsedRows.push({
              nome: titulo,
              departamento: norm(row["Área"] || row["Area"]),
              empresa: "",
              descricaoProblema: norm(row["Descrição"] || row["Descricao"]),
              planoAcao: "",
              rc: rc ? String(rc) : "",
              obs: norm(row["Observação"] || row["Observacao"]),
              precisaMaterial: !!rc,
              materialNecessario: "",
              poMaterial: "",
              impacto: mapImpacto(row["Prioridade"]),
              status: mapStatus(row["Status"]),
              dataExecucao: "",
            });
          });
        });

        setPlanningItems((prev) => {
          const byNome = new Map(prev.map((p, idx) => [p.nome.toLowerCase(), idx]));
          const next = [...prev];
          let added = 0, updated = 0;
          parsedRows.forEach((row) => {
            const key = row.nome.toLowerCase();
            if (byNome.has(key)) {
              const idx = byNome.get(key);
              next[idx] = { ...next[idx], ...row, id: next[idx].id, dataExecucao: next[idx].dataExecucao, linkedServiceId: next[idx].linkedServiceId };
              updated++;
            } else {
              next.push({ id: uid("PLAN"), linkedServiceId: null, ...row });
              added++;
            }
          });
          setImportMsg(`Planejamento importado: ${added} novo(s), ${updated} atualizado(s).`);
          return next;
        });
      } catch (err) {
        setImportMsg("Erro ao ler a planilha de planejamento. Confira se as colunas seguem o padrão esperado.");
      }
      setTimeout(() => setImportMsg(null), 6000);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  /* Assim que um item mapeado ganha uma Data de Execução, ele "vira" um Serviço de verdade — cria
     automaticamente uma linha correspondente em Serviços (aparece no Gantt/Port Call normalmente),
     e marca o vínculo pra não duplicar se a data for editada de novo depois */
  React.useEffect(() => {
    const paraGraduar = planningItems.filter((p) => p.dataExecucao && !p.linkedServiceId);
    if (paraGraduar.length === 0) return;
    const linkByPlanId = {};
    const novosServicos = paraGraduar.map((p) => {
      const serviceId = uid("PC-2026-08");
      linkByPlanId[p.id] = serviceId;
      const start = `${p.dataExecucao}T08:00`;
      const end = `${p.dataExecucao}T17:00`;
      return {
        id: serviceId, name: p.nome, discipline: "Marine", group: "Sem categoria",
        ganttCategory: "Manutenção", empresa: p.empresa || "", md: "Não", rc: "", obs: "",
        budget: 0, committed: 0, actual: 0, forecast: 0, start, end,
        status: "Planejamento", progress: 0, createdAt: new Date().toISOString(),
        planoAcao: p.planoAcao || "", precisaMaterial: p.precisaMaterial || false,
        materialNecessario: p.materialNecessario || "", impacto: p.impacto || "Baixo",
      };
    });
    setWorkPackages((prev) => [...prev, ...novosServicos]);
    setPlanningItems((prev) => prev.map((p) => (linkByPlanId[p.id] ? { ...p, linkedServiceId: linkByPlanId[p.id] } : p)));
  }, [planningItems]);

  /* mesma lógica de graduação automática, agora pros itens de Docagem (Machinery Items DNV) — usa
     a Previsão de Execução como gatilho (não a Data de Conclusão, que só é preenchida depois) */
  React.useEffect(() => {
    const paraGraduar = docagemItems.filter((d) => d.previsaoExecucao && !d.linkedServiceId);
    if (paraGraduar.length === 0) return;
    const linkByDocId = {};
    const novosServicos = paraGraduar.map((d) => {
      const serviceId = uid("PC-2026-08");
      linkByDocId[d.id] = serviceId;
      const start = `${d.previsaoExecucao}T08:00`;
      const end = `${d.previsaoExecucao}T17:00`;
      return {
        id: serviceId, name: d.nome, discipline: "Marine", group: "Sem categoria",
        ganttCategory: "Manutenção", empresa: d.empresa || "", md: "Não", rc: "", obs: "",
        budget: 0, committed: 0, actual: 0, forecast: 0, start, end,
        status: "Planejamento", progress: 0, createdAt: new Date().toISOString(),
        planoAcao: d.planoAcao || "", precisaMaterial: d.necessitaMaterial || false,
        materialNecessario: "", impacto: "Baixo",
      };
    });
    setWorkPackages((prev) => [...prev, ...novosServicos]);
    setDocagemItems((prev) => prev.map((d) => (linkByDocId[d.id] ? { ...d, linkedServiceId: linkByDocId[d.id] } : d)));
  }, [docagemItems]);

  const addWp = () => {
    const id = uid("PC-2026-08");
    setWorkPackages((r) => [...r, {
      id, name: "Novo serviço", discipline: "Marine", group: "Sem categoria", empresa: "", md: "Não", rc: "", obs: "",
      budget: 0, committed: 0, actual: 0, forecast: 0, start: "2026-08-25T08:00", end: "2026-08-26T17:00",
      status: "Planejamento", progress: 0, createdAt: new Date().toISOString(),
    }]);
    flashNewRow(id);
  };

  /* serviço não concluído na data planejada → gera uma NOVA linha com data em branco para reagendar,
     mantendo o registro antigo intacto (com seu desvio/histórico) e um vínculo entre as duas */
  const repeatWp = (original) => {
    const newId = uid("PC-2026-08");
    setWorkPackages((r) => [...r, {
      id: newId, name: original.name, discipline: original.discipline, group: original.group,
      ganttCategory: original.ganttCategory, empresa: original.empresa, md: original.md, rc: original.rc, obs: "",
      budget: 0, committed: 0, actual: 0, forecast: 0, start: "", end: "",
      status: original.status === "Concluído" || original.status === "Cancelado" ? "Não iniciado" : original.status,
      progress: original.progress || 0, repeatOf: original.id, createdAt: new Date().toISOString(),
    }]);
    flashNewRow(newId);
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
    status: "Planejamento", progress: 0, createdAt: new Date().toISOString(),
  }]);

  /* SINCRONIZAÇÃO AUTOMÁTICA Serviços → Pagamentos: só para linhas criadas a partir de agora
     (identificadas pela marca "createdAt", que os serviços importados/antigos não têm) e só quando
     o status chega a "Concluído". Cada serviço só gera um lançamento em Pagamentos uma única vez
     (marcado via linkedInvoiceId), mesmo que o status mude de novo depois. */
  React.useEffect(() => {
    const toSync = workPackages.filter((w) => w.createdAt && w.status === "Concluído" && !w.linkedInvoiceId);
    if (toSync.length === 0) return;
    const linkByWpId = {};
    const newInvoices = toSync.map((w) => {
      const invId = uid("INV");
      linkByWpId[w.id] = invId;
      return {
        id: invId,
        date: (w.start || "").slice(0, 10) || todayISO(),
        assunto: w.name, empresa: w.empresa || "", md: w.md || "Não", mdSentDate: "",
        diffDays: 0, daysOpenTotal: 0, rc: w.rc || "", serviceStatus: "Fechado",
        poContrato: "", medicao: "", valorTotal: 0, saldoPo: 0,
        obs: "Gerado automaticamente a partir de um serviço concluído na aba Serviços.",
        statusPagamento: "Aguardando Orçamento", dataPagamento: "",
      };
    });
    setServiceInvoices((prev) => [...prev, ...newInvoices]);
    setWorkPackages((prev) => prev.map((w) => (linkByWpId[w.id] ? { ...w, linkedInvoiceId: linkByWpId[w.id] } : w)));
  }, [workPackages]);

  /* mantém o RC sincronizado continuamente: se o RC for editado em Serviços depois de já ter
     gerado o lançamento em Pagamentos, o valor é atualizado lá também */
  React.useEffect(() => {
    const linked = workPackages.filter((w) => w.linkedInvoiceId);
    if (linked.length === 0) return;
    setServiceInvoices((prev) => {
      let changed = false;
      const next = prev.map((inv) => {
        const wp = linked.find((w) => w.linkedInvoiceId === inv.id);
        if (wp && inv.rc !== (wp.rc || "")) {
          changed = true;
          return { ...inv, rc: wp.rc || "" };
        }
        return inv;
      });
      return changed ? next : prev;
    });
  }, [workPackages]);

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

  const addMat = () => {
    const id = uid("MAT");
    setMaterials((r) => [...r, {
      id, wp: "", tmMaster: "", departamento: "", sap: "", descricao: "Novo item",
      quantidade: 1, priority: "Média", dataSolicitacao: todayISO(), dataNecessidade: "", reserva: "", rc: "", po: "", linhaPo: "",
      valor: 0, eta: "", obs: "", dataRecebimento: "", status: "Solicitado",
    }]);
    flashNewRow(id);
  };
  const addPay = () => setPayments((r) => [...r, {
    id: uid("PAY"), service: "—", po: "", poValue: 0,
    nf: "", nfValue: 0, issue: "", due: "", status: "Orçamento",
  }]);
  const addInv = () => {
    const id = uid("INV");
    setServiceInvoices((r) => [...r, {
      id, date: todayISO(), assunto: "Novo serviço", empresa: "", md: "Não", mdSentDate: "",
      diffDays: 0, daysOpenTotal: 0, rc: "", serviceStatus: "Aberto", poContrato: "", medicao: "",
      valorTotal: 0, saldoPo: 0, obs: "", statusPagamento: "Aguardando Medição", dataPagamento: "",
    }]);
    flashNewRow(id);
  };

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
      ["Alta", "Crítica", "Emergencial", "Sobressalente crítico"].includes(m.priority) && !["Recebido", "Entregue a bordo"].includes(m.status)
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
    /* Custo por categoria — mês vigente, com a coluna de referência "Ordem" ao lado de cada categoria */
    const mesAtual = `${new Date().getFullYear()}-${pad2(new Date().getMonth() + 1)}`;
    const inPeriodCusto = serviceInvoices.filter((r) => r.statusPagamento !== "Pago" && r.previsaoMes === mesAtual);
    const custoPorCategoria = CATEGORIES.map((cat) => {
      const orcadoUsd = CATEGORY_BUDGET_USD[cat] || 0;
      const orcadoBrl = orcadoUsd * exchangeRate;
      const realizado = inPeriodCusto.reduce((s, r) => s + (r.allocations || []).filter((a) => a.category === cat).reduce((s2, a) => s2 + Number(a.valor || 0), 0), 0);
      return {
        Categoria: cat,
        "Ordem (Compra de Serviços)": adpServicosLabel(cat),
        "Orçado (US$)": orcadoUsd,
        "Orçado (R$)": orcadoBrl,
        "Realizado (R$)": realizado,
        "Disponível (R$)": orcadoBrl - realizado,
      };
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(custoPorCategoria), "CustoPorCategoria");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(workPackages, WP_COLS)), "Servicos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(materials, MAT_COLS)), "Materiais");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(payments, PAY_COLS)), "Pagamentos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(serviceInvoices, INV_COLS)), "StatusPagamentos");
    XLSX.writeFile(wb, `genesis-${selectedPortCallLabel.replace(/\s+/g, "-").toLowerCase()}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleImportClick = () => fileInputRef.current?.click();
  /* ---------- Importação semanal: "Pedidos Emergenciais" (planilha externa, mesmas colunas de
     Materiais, mas com pequenas diferenças de formato: quantidade vem como texto "N Unidade",
     valor às vezes vem em branco, e prioridade/status usam termos próprios). Mescla por SAP, então
     rodar essa importação toda semana atualiza os itens já existentes em vez de duplicar. ---------- */
  const handleImportEmergenciais = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        /* Normaliza cabeçalhos removendo acentos, pontuação e preposições ("de"/"da"/"do"/"dos"/"das"),
           para reconhecer planilhas próprias/externas que não seguem exatamente os nomes deste sistema
           (ex.: "Descrição do Material", "Qtd", "Data de Solicitação" etc.) */
        const stripAccents = (s) => s.normalize("NFD").replace(/[̀-ͯ]/g, "");
        const normHeader = (h) =>
          stripAccents(h.toString())
            .toLowerCase()
            .replace(/[.:º°]/g, "")
            .split(/\s+/)
            .filter((w) => w && !["de", "da", "do", "dos", "das"].includes(w))
            .join(" ")
            .trim();
        const HEADER_MAP = {
          "tm master": "tmMaster", "tm": "tmMaster",
          "departamento": "departamento", "depto": "departamento", "area": "departamento", "setor": "departamento",
          "sap": "sap", "codigo sap": "sap", "cod sap": "sap", "material sap": "sap", "codigo material": "sap", "codigo": "sap", "cod": "sap", "cod material": "sap",
          "descricao": "descricao", "descricao material": "descricao", "descricao item": "descricao", "material": "descricao", "item": "descricao", "produto": "descricao", "descricao do item": "descricao", "descricao produto": "descricao",
          "quantidade": "quantidade", "qtd": "quantidade", "qde": "quantidade", "quant": "quantidade", "qtde": "quantidade",
          "prioridade": "priority", "urgencia": "priority", "prioridade solicitacao": "priority",
          "data solicitacao": "dataSolicitacao", "data pedido": "dataSolicitacao", "solicitacao": "dataSolicitacao", "data abertura": "dataSolicitacao",
          "data necessidade": "dataNecessidade", "necessidade": "dataNecessidade", "data limite": "dataNecessidade", "data desejada": "dataNecessidade",
          "reserva": "reserva", "no reserva": "reserva", "numero reserva": "reserva", "nº reserva": "reserva",
          "rc": "rc", "requisicao compra": "rc", "requisicao": "rc", "no rc": "rc",
          "po": "po", "pedido compra": "po", "ordem compra": "po", "no po": "po",
          "linha po": "linhaPo", "item po": "linhaPo", "linha da po": "linhaPo",
          "valor": "valor", "valor unitario": "valor", "valor total": "valor", "preco": "valor", "preco unitario": "valor", "vlr": "valor",
          "eta": "eta", "previsao chegada": "eta", "chegada prevista": "eta", "previsao entrega": "eta",
          "obs": "obs", "observacao": "obs", "observacoes": "obs", "obs gerais": "obs",
          "data recebimento": "dataRecebimento", "recebido em": "dataRecebimento", "data entrega": "dataRecebimento",
          "status": "status", "situacao": "status",
        };
        const rowKeyMap = {};
        if (json.length > 0) {
          Object.keys(json[0]).forEach((h) => {
            const key = HEADER_MAP[normHeader(h)];
            if (key) rowKeyMap[h] = key;
          });
        }

        const parseQty = (v) => {
          const m = String(v).match(/\d+([.,]\d+)?/);
          return m ? Number(m[0].replace(",", ".")) : 0;
        };
        const parseValor = (v) => {
          const n = Number(v);
          return isNaN(n) ? 0 : n;
        };
        const parseDate = (v) => {
          if (!v) return "";
          if (v instanceof Date) return `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}`;
          return String(v);
        };
        const normPriority = (v) => {
          const s = (v || "").toString().trim();
          return PRIORITY.includes(s) ? s : (s || "Média");
        };
        const normStatus = (v) => {
          const s = (v || "").toString().trim();
          return MAT_STATUS.includes(s) ? s : (s || "Solicitado");
        };

        const parsedRows = json.map((row) => {
          const r = {};
          Object.keys(row).forEach((h) => { if (rowKeyMap[h]) r[rowKeyMap[h]] = row[h]; });
          return {
            tmMaster: (r.tmMaster || "").toString(),
            departamento: (r.departamento || "").toString(),
            sap: (r.sap || "").toString(),
            descricao: (r.descricao || "").toString(),
            quantidade: parseQty(r.quantidade),
            priority: normPriority(r.priority),
            dataSolicitacao: parseDate(r.dataSolicitacao),
            dataNecessidade: parseDate(r.dataNecessidade),
            reserva: (r.reserva || "").toString(),
            rc: (r.rc || "").toString(),
            po: (r.po || "").toString(),
            linhaPo: (r.linhaPo || "").toString(),
            valor: parseValor(r.valor),
            eta: parseDate(r.eta),
            obs: (r.obs || "").toString().trim(),
            dataRecebimento: parseDate(r.dataRecebimento),
            status: normStatus(r.status),
          };
        }).filter((r) => r.descricao || r.sap || r.tmMaster);

        if (parsedRows.length === 0) {
          const headersFound = json.length > 0 ? Object.keys(json[0]).join(", ") : "(planilha vazia)";
          setImportMsg(`Nenhuma linha reconhecida nessa planilha. Cabeçalhos encontrados: ${headersFound}. Renomeie as colunas para algo como "Descrição"/"Material", "Quantidade", "Status" etc., ou avise quais são esses nomes para eu ajustar a importação.`);
          setTimeout(() => setImportMsg(null), 12000);
          e.target.value = "";
          return;
        }

        setMaterials((prev) => {
          const bySap = new Map();
          prev.forEach((m, idx) => { if (m.sap) bySap.set(String(m.sap), idx); });
          const next = [...prev];
          let updated = 0, added = 0;
          parsedRows.forEach((row) => {
            const key = row.sap;
            if (key && bySap.has(key)) {
              const idx = bySap.get(key);
              next[idx] = { ...next[idx], ...row, id: next[idx].id, wp: next[idx].wp };
              updated++;
            } else {
              next.push({ id: uid("MAT"), wp: "", ...row });
              added++;
            }
          });
          setImportMsg(`Pedidos Emergenciais importado: ${added} novo(s), ${updated} atualizado(s).`);
          return next;
        });
      } catch (err) {
        setImportMsg("Erro ao ler a planilha de Pedidos Emergenciais. Confira se o formato de colunas não mudou.");
      }
      setTimeout(() => setImportMsg(null), 6000);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

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
          const invoiceRows = sheetToRows(json, INV_COLS);
          /* aba opcional "Alocacoes" — permite reconstruir o rateio por categoria (Custos) junto
             com a importação, casando pelo par Serviço (assunto) + Empresa. Sem essa aba, o
             rateio simplesmente fica vazio (do jeito que já era antes). Colunas esperadas:
             Serviço, Empresa, Categoria, Valor. */
          if (wb.SheetNames.includes("Alocacoes")) {
            const allocJson = XLSX.utils.sheet_to_json(wb.Sheets["Alocacoes"], { defval: "" });
            const norm = (s) => (s || "").toString().trim().toLowerCase();
            allocJson.forEach((row) => {
              const servico = norm(row["Serviço"] ?? row["Servico"]);
              const empresa = norm(row["Empresa"]);
              const categoria = (row["Categoria"] || "").toString().trim();
              const valor = Number(row["Valor"]) || 0;
              if (!servico || !categoria) return;
              const inv = invoiceRows.find((r) => norm(r.assunto) === servico && norm(r.empresa) === empresa);
              if (inv) {
                if (!Array.isArray(inv.allocations)) inv.allocations = [];
                inv.allocations.push({ category: categoria, valor });
              }
            });
            imported.push(`${allocJson.length} alocação(ões) de rateio`);
          }
          setServiceInvoices(invoiceRows);
          imported.push(`${json.length} status de pagamento`);
        }
        setImportMsg(imported.length ? `Importado: ${imported.join(", ")}.` : "Nenhuma aba reconhecida (esperado: Servicos, Materiais, Pagamentos, StatusPagamentos, Alocacoes).");
      } catch (err) {
        setImportMsg("Erro ao ler o arquivo. Confira se é um .xlsx exportado por este sistema.");
      }
      setTimeout(() => setImportMsg(null), 5000);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  /* "Exportar relatório" agora dispara o gerador de PDF registrado pela página atualmente aberta
     (cada aba registra o seu via setReportFn, refletindo exatamente o que está sendo mostrado nela) */

  const navItems = [
    { key: "dashboard", label: "Dashboard", icon: LayoutGrid },
    { key: "services", label: "Serviços", icon: Wrench },
    { key: "planejamento", label: "Planejamento", icon: ClipboardList },
    { key: "tmmaster", label: "TM Master", icon: Gauge },
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
        <div className="g-brand">
          <img src={LOGO_MARK} alt="OSM Thome" className="g-brand-logo" />
          <div className="g-brand-text">
            <span className="g-brand-name">GENESIS <b>I</b></span>
            <span className="g-brand-tag">Maintenance &amp; Port Call</span>
          </div>
        </div>
        <div className="g-brand-sep" />
        <div className="g-nav-row">
          {navItems.map((n) => (
            <div key={n.key} className={`g-nav-item ${tab === n.key ? "active" : ""}`} onClick={() => setTab(n.key)}>
              <n.icon size={14} />{n.label}
            </div>
          ))}
        </div>
        {currentUser && <span style={{ fontSize: 12, whiteSpace: "nowrap", color: "var(--navbar-text)" }}>{currentUser.name}</span>}
        {onLogout && <div className="g-logout" onClick={onLogout}><LogOut size={14} />Sair</div>}
      </div>

      {/* Global period filter — removido conforme solicitado; workPackages/Dashboard agora mostram
          todos os dados por padrão (período fixado num intervalo amplo em vez de restringir por mês) */}
      {false && (
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
        </div>
        <div className="g-flex" style={{ gap: 8, flexWrap: "wrap" }}>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImportFile} />
          {tab !== "planejamento" && tab !== "tmmaster" && <button className="g-btn" onClick={handleImportClick} title="Importar planilha (.xlsx) — reconhece abas Servicos, Materiais, Pagamentos, StatusPagamentos e Alocacoes"><Upload size={14} />Importar</button>}
          <button className="g-btn" onClick={() => (exportXlsxFn ? exportXlsxFn() : handleExportXlsx())} title="Exportar em planilha (.xlsx) o conteúdo desta página, já filtrado"><Download size={14} />Exportar planilha</button>
          <button className="g-btn" onClick={() => reportFn && reportFn()} disabled={!reportFn}
            title="Exportar relatório em PDF, com o conteúdo exato da página aberta"><FileText size={14} />Exportar relatório</button>
          {tab === "services" && <button className="g-btn primary" onClick={addWp}><Plus size={14} />Novo serviço</button>}
          {tab === "planejamento" && planSubTab === "mapeados" && <button className="g-btn primary" onClick={addPlanningItem}><Plus size={14} />Novo mapeamento</button>}
          {tab === "materials" && <button className="g-btn primary" onClick={addMat}><Plus size={14} />Nova requisição</button>}
          {tab === "payments" && <button className="g-btn primary" onClick={addInv}><Plus size={14} />Novo registro</button>}
        </div>
      </div>

      {importMsg && (
        <div style={{ margin: "10px 22px 0 22px" }}>
          <div className="g-alert" style={{ background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>
            {importMsg}
          </div>
        </div>
      )}

      <div className="g-body">
        {tab === "dashboard" && (
          <DashboardView kpis={kpis} workPackages={workPackages} disciplineCosts={disciplineCosts}
            serviceInvoices={serviceInvoices} setReportFn={setReportFn} tmDue={tmDue}
            exchangeRate={exchangeRate} setExchangeRate={setExchangeRate} />
        )}

        {tab === "gantt" && (
          <GanttView workPackages={workPackages} portCallName={selectedPortCallLabel}
            spans={effectivePortCallSpans} portCallLabel={portCallLabel} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn}
            updWp={updWp} remWp={remWp} removePortCall={removePortCall} addWpOnDate={addWpOnDate}
            addPortCallRecord={addPortCallRecord} filterRange={effectiveRange}
            opCategories={opCategories} catOf={catOf} addOpCategory={addOpCategory}
            renameOpCategory={renameOpCategory} removeOpCategory={removeOpCategory} />
        )}

        {tab === "services" && (
          <ServicesView workPackages={workPackages} updWp={updWp} remWp={remWp} repeatWp={repeatWp}
            expandedWp={expandedWp} setExpandedWp={setExpandedWp} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} newRowId={newRowId} />
        )}

        {tab === "planejamento" && (
          <PlanejamentoView workPackages={workPackages} updWp={updWp} materials={materials} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn}
            allPortCallDates={allPortCallDates} portCallLabel={portCallLabel} addWpOnDate={addWpOnDate}
            planningItems={planningItems} updPlan={updPlan} remPlan={remPlan} addPlanningItem={addPlanningItem}
            docagemItems={docagemItems} updDocagem={updDocagem} remDocagem={remDocagem} addDocagemItem={addDocagemItem}
            handleImportDocagem={handleImportDocagem} planSubTab={planSubTab} setPlanSubTab={setPlanSubTab}
            newRowId={newRowId} handleImportPlanejamento={handleImportPlanejamento} />
        )}

        {tab === "tmmaster" && (
          <TmMasterView tmDue={tmDue} tmHistory={tmHistory} tmDueSnapshots={tmDueSnapshots} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn}
            handleImportTmDue={handleImportTmDue} handleImportTmHistory={handleImportTmHistory}
            tmSubTab={tmSubTab} setTmSubTab={setTmSubTab} updTmDue={updTmDue} onToggleAddToPlanning={onToggleAddToPlanning} />
        )}

        {tab === "materials" && <MaterialsView materials={materials} updMat={updMat} remMat={remMat} workPackages={workPackages} setReportFn={setReportFn} handleImportEmergenciais={handleImportEmergenciais} newRowId={newRowId} />}

        {tab === "payments" && (
          <PaymentsSection
            paySubTab={paySubTab} setPaySubTab={setPaySubTab}
            serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} addInv={addInv}
            setReportFn={setReportFn} newRowId={newRowId}
          />
        )}

        {tab === "costs" && (
          <CostsView serviceInvoices={serviceInvoices} updInv={updInv} setReportFn={setReportFn}
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
function DashboardView({ kpis, workPackages, disciplineCosts, serviceInvoices, exchangeRate, setExchangeRate, setReportFn, tmDue = [] }) {

  /* filtro de período do Dashboard — por padrão, o mês vigente */
  const defaultDashPeriod = useMemo(() => {
    const now = new Date();
    const start = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-01`;
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    const end = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(lastDay)}`;
    return { start, end };
  }, []);
  const [dp, setDp] = useState({ dataInicio: defaultDashPeriod.start, dataFim: defaultDashPeriod.end });
  const isDefaultDashPeriod = dp.dataInicio === defaultDashPeriod.start && dp.dataFim === defaultDashPeriod.end;

  /* próximas manutenções — só o que ainda está em aberto de fato (não concluído nem cancelado) */
  const upcomingMaintenance = workPackages
    .filter((w) => w.status === "Planejamento")
    .slice()
    .sort((a, b) => (a.start ? new Date(a.start).getTime() : Infinity) - (b.start ? new Date(b.start).getTime() : Infinity));

  /* ---------- Serviços: espelha exatamente os mesmos números da aba Serviços ---------- */
  const totalServicos = workPackages.length;
  const concluidosServicos = workPackages.filter((w) => w.status === "Concluído").length;
  const emAndamentoServicos = workPackages.filter((w) => w.status === "Em andamento").length;
  const naoIniciadosServicos = workPackages.filter((w) => w.status === "Não iniciado").length;
  const canceladosServicos = workPackages.filter((w) => w.status === "Cancelado").length;
  const naoCanceladosServicos = totalServicos - canceladosServicos;
  const taxaConclusaoServicos = naoCanceladosServicos ? Math.round((concluidosServicos / naoCanceladosServicos) * 100) : 0;

  /* ---------- Pagamentos: espelha exatamente a mesma classificação Pago/Pendente/Atrasado da aba Pagamentos ---------- */
  const pagosDash = serviceInvoices.filter((r) => invoiceSituation(r) === "Pago");
  const pendentesDash = serviceInvoices.filter((r) => invoiceSituation(r) === "Pendente");
  const atrasadosDash = serviceInvoices.filter((r) => invoiceSituation(r) === "Atrasado");
  const sumVal = (arr) => arr.reduce((s, r) => s + Number(r.valorTotal || 0), 0);
  const invTotalDiasAberto = [...pendentesDash, ...atrasadosDash].reduce((s, r) => s + Number(r.daysOpenTotal || 0), 0);
  const topOpenInvoices = serviceInvoices
    .filter((r) => r.statusPagamento === "Aprovação Pendente")
    .slice()
    .sort((a, b) => Number(b.daysOpenTotal || 0) - Number(a.daysOpenTotal || 0))
    .slice(0, 8);

  /* ---------- Financeiro: espelha exatamente a mesma lógica de Orçado/Realizado/Disponível da aba Custos
     (categoria × alocação de rateio, serviços ainda não pagos), escopado ao período do Dashboard ---------- */
  const categoryCostsDash = useMemo(() => {
    const inicioMes = dp.dataInicio ? dp.dataInicio.slice(0, 7) : null;
    const fimMes = dp.dataFim ? dp.dataFim.slice(0, 7) : null;
    /* usa o mês PROVISIONADO de cada serviço (não a data de execução), para que pagamentos de
       backlog (executados num mês anterior, mas provisionados para o mês vigente) entrem certinho
       no cálculo do mês selecionado */
    const inPeriod = serviceInvoices.filter((r) =>
      r.statusPagamento !== "Pago" && r.previsaoMes &&
      (!inicioMes || r.previsaoMes >= inicioMes) && (!fimMes || r.previsaoMes <= fimMes)
    );
    return CATEGORIES.map((cat) => {
      const orcadoUsd = CATEGORY_BUDGET_USD[cat] || 0;
      const orcadoBrl = orcadoUsd * exchangeRate;
      const realizado = inPeriod.reduce((s, r) => s + (r.allocations || []).filter((a) => a.category === cat).reduce((s2, a) => s2 + Number(a.valor || 0), 0), 0);
      return { category: cat, orcadoUsd, orcadoBrl, realizado, disponivel: orcadoBrl - realizado };
    });
  }, [serviceInvoices, dp, exchangeRate]);
  const totalOrcadoDash = categoryCostsDash.reduce((s, c) => s + c.orcadoBrl, 0);
  const totalRealizadoDash = categoryCostsDash.reduce((s, c) => s + c.realizado, 0);
  const totalDisponivelDash = totalOrcadoDash - totalRealizadoDash;

  /* ---------- Gastos por mês provisionado: soma do Valor Total, agrupado por r.previsaoMes (não pela data de execução) ---------- */
  const gastosPorMes = useMemo(() => {
    const map = {};
    serviceInvoices.forEach((r) => {
      if (!r.previsaoMes) return;
      map[r.previsaoMes] = (map[r.previsaoMes] || 0) + Number(r.valorTotal || 0);
    });
    return Object.keys(map).sort().map((key) => {
      const [y, m] = key.split("-");
      return { mes: `${MONTH_NAMES[Number(m) - 1].slice(0, 3)}/${y.slice(2)}`, valor: map[key] };
    });
  }, [serviceInvoices]);

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
    { label: "Total Orçado (mês)", value: fmt(totalOrcadoDash), color: "var(--text-dim)" },
    { label: "Total Realizado (mês)", value: fmt(totalRealizadoDash), color: "var(--teal)" },
    { label: "Saldo Disponível (mês)", value: fmt(totalDisponivelDash), color: totalDisponivelDash < 0 ? "var(--crit)" : "var(--ok)" },
  ];
  const servicos = [
    { label: "Total de Serviços", value: totalServicos, color: "var(--teal)" },
    { label: "Concluídos", value: concluidosServicos, color: "var(--ok)" },
    { label: "Em Andamento", value: emAndamentoServicos, color: "var(--teal)" },
    { label: "Não Iniciados", value: naoIniciadosServicos, color: "var(--text-dim)" },
    { label: "Cancelados", value: canceladosServicos, color: "var(--text-dim)" },
    { label: "Taxa de Conclusão", value: `${taxaConclusaoServicos}%`, color: "var(--warn)" },
  ];
  const pagamentos = [
    { label: "Pago", value: `${pagosDash.length} · ${fmt(sumVal(pagosDash))}`, color: "var(--ok)" },
    { label: "Pendente", value: `${pendentesDash.length} · ${fmt(sumVal(pendentesDash))}`, color: "var(--warn)" },
    { label: "Atrasado", value: `${atrasadosDash.length} · ${fmt(sumVal(atrasadosDash))}`, color: "var(--crit)" },
  ];

  React.useEffect(() => {
    if (!setReportFn) return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      let y = pdfHeader(doc, "Relatório do Dashboard",
        `Período: ${fmtPeriodDate(dp.dataInicio)} — ${fmtPeriodDate(dp.dataFim)} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [
        ...financeiro.map((k) => ({ label: k.label, value: k.value })),
        ...servicos.map((k) => ({ label: k.label, value: k.value })),
        ...pagamentos.map((k) => ({ label: k.label, value: k.value })),
      ]);
      y = pdfSectionTitle(doc, y, "Custo por categoria (mês selecionado)");
      y = pdfTable(doc, y,
        ["Categoria", "Orçado (US$)", "Orçado (R$)", "Realizado (R$)"],
        categoryCostsDash.map((d) => [d.category, fmtBudgetUsd(d.orcadoUsd), fmtBudgetBrl(d.orcadoUsd, d.orcadoBrl), fmt(d.realizado)]),
        { columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" } } }
      );
      if (y > 220) { doc.addPage(); y = 15; }
      y = pdfSectionTitle(doc, y, "Pagamentos com Aprovação Pendente há mais tempo");
      pdfTable(doc, y,
        ["Data", "Serviço", "Empresa", "Valor", "Dias em Aberto"],
        topOpenInvoices.map((r) => [fmtDate(r.date), r.assunto, r.empresa, fmt(r.valorTotal), r.daysOpenTotal]),
        { columnStyles: { 3: { halign: "right" } } }
      );
      pdfSave(doc, "relatorio-dashboard");
    });
  }, [financeiro, servicos, pagamentos, categoryCostsDash, gastosPorMes, topOpenInvoices, dp, setReportFn]);

  const pctRealizado = totalOrcadoDash > 0 ? Math.min(100, Math.round((totalRealizadoDash / totalOrcadoDash) * 100)) : 0;
  const pctReal = totalOrcadoDash > 0 ? Math.round((totalRealizadoDash / totalOrcadoDash) * 100) : 0;
  const donutServicos = [
    { name: "Concluídos", value: concluidosServicos, color: "#22C55E" },
    { name: "Em andamento", value: emAndamentoServicos, color: "#3B82F6" },
    { name: "Não iniciados", value: naoIniciadosServicos, color: "#9499A8" },
    { name: "Cancelados", value: canceladosServicos, color: "#EF4444" },
  ].filter((d) => d.value > 0);
  const donutPagamentos = [
    { name: "Pago", value: pagosDash.length, color: "#22C55E" },
    { name: "Pendente", value: pendentesDash.length, color: "#F5A623" },
    { name: "Atrasado", value: atrasadosDash.length, color: "#EF4444" },
  ].filter((d) => d.value > 0);
  const tipStyle = { background: "#fff", border: "1px solid var(--border)", borderRadius: 6, fontSize: 11 };
  /* TM Master (Due): situação de cada manutenção por departamento — mesma regra da aba TM Master */
  const tmVencida = (r) => r.diffValue !== null && r.diffValue < 0;
  const tmAte40 = (r) => r.diffUnit === "D" && r.diffValue !== null && r.diffValue >= 0 && r.diffValue <= 40;
  const tmVencidas = tmDue.filter(tmVencida).length;
  const tmAte40Dias = tmDue.filter(tmAte40).length;
  const tmNormais = tmDue.length - tmVencidas - tmAte40Dias;
  const tmPorDepto = useMemo(() => {
    const map = {};
    tmDue.forEach((d) => {
      const k = d.department || "Sem departamento";
      if (!map[k]) map[k] = { departamento: k, Vencidas: 0, "Até 40 dias": 0, "No prazo": 0, total: 0 };
      if (tmVencida(d)) map[k].Vencidas++; else if (tmAte40(d)) map[k]["Até 40 dias"]++; else map[k]["No prazo"]++;
      map[k].total++;
    });
    return Object.values(map).sort((a, b) => b.total - a.total).slice(0, 10);
  }, [tmDue]);

  const Donut = ({ data, center, centerLabel }) => (
    <div className="dsh-donut">
      <div className="dsh-donut-chart">
        <ResponsiveContainer>
          <PieChart>
            <Pie data={data.length ? data : [{ name: "Sem dados", value: 1, color: "#E7E9F0" }]} dataKey="value" innerRadius={52} outerRadius={74} paddingAngle={data.length > 1 ? 2 : 0} stroke="none">
              {(data.length ? data : [{ color: "#E7E9F0" }]).map((d, i) => <Cell key={i} fill={d.color} />)}
            </Pie>
            {data.length > 0 && <Tooltip contentStyle={tipStyle} />}
          </PieChart>
        </ResponsiveContainer>
        <div className="dsh-donut-center"><b>{center}</b><span>{centerLabel}</span></div>
      </div>
      <div className="dsh-legend">
        {data.map((d) => (
          <div key={d.name}><i style={{ background: d.color }} />{d.name}<b>{d.value}</b></div>
        ))}
      </div>
    </div>
  );

  return (
    <>
      <div className="dsh-period">
        <div className="g-field">
          <label>De</label>
          <input type="date" value={dp.dataInicio} onChange={(e) => setDp((p) => ({ ...p, dataInicio: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>Até</label>
          <input type="date" value={dp.dataFim} onChange={(e) => setDp((p) => ({ ...p, dataFim: e.target.value }))} />
        </div>
        <button className="g-btn" onClick={() => setDp(defaultDashPeriod)} disabled={isDefaultDashPeriod} style={{ opacity: isDefaultDashPeriod ? 0.5 : 1 }}>
          <X size={13} />Mês vigente
        </button>
        <div className="g-filter-spacer" />
        <div className="dsh-period-note">Financeiro pelo mês provisionado de cada serviço</div>
      </div>

      {/* ---------- Financeiro ---------- */}
      <div className="dsh-hero">
        <div className="dsh-card dsh-fin">
          <div className="dsh-label">Orçado no período</div>
          <div className="dsh-big">{fmt(totalOrcadoDash)}</div>
          <div className="dsh-sub">Câmbio US$ {exchangeRate}
            <input type="number" step="0.01" value={exchangeRate} onChange={(e) => setExchangeRate(Number(e.target.value))} className="dsh-fx" title="Câmbio US$ → R$" />
          </div>
        </div>
        <div className="dsh-card dsh-fin">
          <div className="dsh-label">Realizado</div>
          <div className="dsh-big" style={{ color: "var(--teal)" }}>{fmt(totalRealizadoDash)}</div>
          <div className="dsh-bar"><div style={{ width: `${pctRealizado}%`, background: pctReal > 100 ? "var(--crit)" : pctReal > 85 ? "var(--warn)" : "var(--accent)" }} /></div>
          <div className="dsh-sub">{pctReal}% do orçado</div>
        </div>
        <div className="dsh-card dsh-fin" style={{ borderTop: `3px solid ${totalDisponivelDash < 0 ? "var(--crit)" : "var(--ok)"}` }}>
          <div className="dsh-label">Saldo disponível</div>
          <div className="dsh-big" style={{ color: totalDisponivelDash < 0 ? "var(--crit)" : "var(--ok)" }}>{fmt(totalDisponivelDash)}</div>
          <div className="dsh-sub">{totalDisponivelDash < 0 ? "Acima do orçamento" : "Dentro do orçamento"}</div>
        </div>
      </div>

      {/* ---------- Serviços + Pagamentos ---------- */}
      <div className="dsh-row2">
        <div className="dsh-card">
          <div className="dsh-title">Serviços</div>
          <Donut data={donutServicos} center={`${taxaConclusaoServicos}%`} centerLabel="concluído" />
          <div className="dsh-chips">
            <div><b>{totalServicos}</b>Total</div>
            <div><b style={{ color: "var(--ok)" }}>{concluidosServicos}</b>Concluídos</div>
            <div><b style={{ color: "var(--accent)" }}>{emAndamentoServicos}</b>Andamento</div>
          </div>
        </div>
        <div className="dsh-card">
          <div className="dsh-title">Pagamentos</div>
          <Donut data={donutPagamentos} center={serviceInvoices.length} centerLabel="registros" />
          <div className="dsh-chips">
            <div><b style={{ color: "var(--ok)" }}>{fmt(sumVal(pagosDash))}</b>Pago</div>
            <div><b style={{ color: "var(--warn)" }}>{fmt(sumVal(pendentesDash))}</b>Pendente</div>
            <div><b style={{ color: "var(--crit)" }}>{fmt(sumVal(atrasadosDash))}</b>Atrasado</div>
          </div>
        </div>
      </div>

      {/* ---------- TM Master ---------- */}
      <div className="dsh-card" style={{ marginBottom: 14 }}>
        <div className="dsh-title">TM Master — Manutenções Vencendo <span className="dsh-count">{tmDue.length}</span></div>
        {tmDue.length === 0 ? (
          <div className="g-muted">Nenhum dado de TM Master importado ainda.</div>
        ) : (
          <>
            <div className="dsh-chips" style={{ marginBottom: 12 }}>
              <div><b style={{ color: "var(--crit)" }}>{tmVencidas}</b>Vencidas</div>
              <div><b style={{ color: "var(--warn)" }}>{tmAte40Dias}</b>Vencem em até 40 dias</div>
              <div><b style={{ color: "var(--ok)" }}>{tmNormais}</b>No prazo</div>
            </div>
            <div style={{ width: "100%", height: Math.max(180, tmPorDepto.length * 34 + 50) }}>
              <ResponsiveContainer>
                <BarChart data={tmPorDepto} layout="vertical" margin={{ left: 10, right: 16, top: 4, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" horizontal={false} />
                  <XAxis type="number" allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="departamento" width={120} tick={{ fill: "var(--text-dim)", fontSize: 11 }} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={tipStyle} cursor={{ fill: "rgba(59,130,246,0.06)" }} />
                  <Bar dataKey="Vencidas" stackId="a" fill="#EF4444" />
                  <Bar dataKey="Até 40 dias" stackId="a" fill="#F5A623" />
                  <Bar dataKey="No prazo" stackId="a" fill="#22C55E" radius={[0, 5, 5, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="dsh-legend" style={{ flexDirection: "row", gap: 16, marginTop: 6 }}>
              <div><i style={{ background: "#EF4444" }} />Vencidas</div>
              <div><i style={{ background: "#F5A623" }} />Até 40 dias</div>
              <div><i style={{ background: "#22C55E" }} />No prazo</div>
            </div>
          </>
        )}
      </div>

      {/* ---------- Manutenções + Port Call ---------- */}
      <div className="g-grid-2">
        <div className="dsh-card">
          <div className="dsh-title">Próximas Manutenções (Planejamento) <span className="dsh-count">{upcomingMaintenance.length}</span></div>
          <div className="dsh-scroll">
            {upcomingMaintenance.length === 0 && <div className="g-muted">Nenhuma manutenção prevista no momento.</div>}
            {upcomingMaintenance.map((w) => (
              <div className="dsh-item" key={w.id}>
                <div>
                  <div className="dsh-item-main">{w.name}</div>
                  <div className="dsh-item-sub">{w.discipline} · início {w.start ? fmtDateTime(w.start) : "sem data"}</div>
                </div>
                <Pill status={w.status} />
              </div>
            ))}
          </div>
        </div>

        <div className="dsh-card">
          <div className="dsh-title">{lastPortCallLabel} <span className="dsh-count">{lastPortCallServices.length}</span></div>
          <div className="dsh-item-sub" style={{ marginBottom: 8 }}>Serviços do Último Port Call Registrado</div>
          <div className="dsh-scroll">
            {lastPortCallServices.length === 0 && <div className="g-muted">Nenhum Port Call registrado ainda.</div>}
            {lastPortCallServices.map((w) => (
              <div className="dsh-item" key={w.id}>
                <div>
                  <div className="dsh-item-main">{w.name}</div>
                  <div className="dsh-item-sub">{w.empresa || "—"} · {fmtDate(dateKeyOf(w.start))}</div>
                </div>
                <Pill status={w.status} />
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ---------- Atenção: aprovações + materiais ---------- */}
      <div className="g-grid-2" style={{ marginTop: 14 }}>
        <div className="dsh-card">
          <div className="dsh-title">Aprovação Pendente Há Mais Tempo <span className="dsh-count">{topOpenInvoices.length}</span></div>
          <div className="dsh-scroll">
            {topOpenInvoices.length === 0 && <div className="g-muted">Nenhum pagamento com aprovação pendente.</div>}
            {topOpenInvoices.map((r) => (
              <div className="dsh-item" key={r.id}>
                <div>
                  <div className="dsh-item-main">{r.assunto}</div>
                  <div className="dsh-item-sub">{r.empresa} · {fmt(r.valorTotal)}</div>
                </div>
                <span className="dsh-days">{r.daysOpenTotal} dias</span>
              </div>
            ))}
          </div>
        </div>

        <div className="dsh-card">
          <div className="dsh-title">Materiais Urgentes <span className="dsh-count">{kpis.materiaisUrgentes.length}</span></div>
          <div className="dsh-scroll">
            {kpis.materiaisUrgentes.length === 0 && <div className="g-muted">Nenhum material urgente em aberto.</div>}
            {kpis.materiaisUrgentes.map((m) => (
              <div className="dsh-item" key={m.id}>
                <div>
                  <div className="dsh-item-main">{m.descricao}</div>
                  <div className="dsh-item-sub">Necessário {fmtDate(m.dataNecessidade)}</div>
                </div>
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
  opCategories, catOf, addOpCategory, renameOpCategory, removeOpCategory, setReportFn, setExportXlsxFn }) {
  const [collapsedCats, setCollapsedCats] = useState(new Set());
  const toggleCollapse = (key) => setCollapsedCats((prev) => {
    const next = new Set(prev);
    next.has(key) ? next.delete(key) : next.add(key);
    return next;
  });
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

  /* essas categorias são operações do próprio navio — não precisam de campo Empresa */
  const NO_EMPRESA_CATEGORIES = ["Manobras", "Troca de Turma", "Base Açu", "Load", "Backload"];

  /* mudar a Duração (horas) recalcula o Término automaticamente a partir do Início —
     e mudar o Término, como já acontece, recalcula a Duração exibida. As duas direções funcionam. */
  const setTaskDuration = (i, w, hours) => {
    if (!w.start) return;
    const start = new Date(w.start);
    const end = new Date(start.getTime() + Math.max(0, Number(hours) || 0) * 3600000);
    const endStr = `${end.getFullYear()}-${pad2(end.getMonth() + 1)}-${pad2(end.getDate())}T${pad2(end.getHours())}:${pad2(end.getMinutes())}`;
    updWp(i, "end", endStr);
  };

  /* filtro de período próprio da aba Port Call — em branco (padrão) mostra todos os Port Calls */
  const [pcPeriod, setPcPeriod] = useState({ start: "", end: "" });
  const hasActivePcPeriod = !!(pcPeriod.start || pcPeriod.end);
  const visibleSpans = useMemo(() => {
    if (!hasActivePcPeriod) return spans;
    return spans.filter((s) => (!pcPeriod.start || s.startKey >= pcPeriod.start) && (!pcPeriod.end || s.startKey <= pcPeriod.end));
  }, [spans, pcPeriod, hasActivePcPeriod]);

  React.useEffect(() => {
    if (!setReportFn) return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      let y = pdfHeader(doc, "Relatório de Port Call — Cronograma",
        `${portCallName} · ${visibleSpans.length} Port Call(s) no período · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      visibleSpans.forEach((span) => {
        const activities = workPackages.filter((w) => {
          const d = new Date(w.start);
          return d >= span.start && d < span.end;
        }).sort((a, b) => new Date(a.start) - new Date(b.start));
        if (y > 250) { doc.addPage(); y = 15; }
        y = pdfSectionTitle(doc, y, `${portCallLabel(span.startKey)} — ${activities.length} atividade(s) · ${Math.round(span.hours)}h`);
        /* colunas no padrão de um cronograma tipo MS Project: Tarefa / Início / Término / Duração / % Completo / Status / Recurso */
        y = pdfTable(doc, y,
          ["Nome da Tarefa", "Início", "Término", "Duração", "% Completo", "Status", "Recurso (Empresa)"],
          activities.map((w) => [
            w.name, fmtDateTime(w.start), fmtDateTime(w.end),
            `${Math.max(0, Math.round((new Date(w.end) - new Date(w.start)) / 3600000))}h`,
            `${w.progress}%`, w.status, w.empresa || "—",
          ]),
          { columnStyles: { 4: { halign: "right" } } }
        );
        if (activities.length > 0) {
          y = pdfSectionTitle(doc, y, "Cronograma visual");
          y = pdfMiniGantt(doc, y, span, activities);
        }
      });
      pdfSave(doc, "relatorio-portcall");
    });
  }, [visibleSpans, workPackages, portCallName, setReportFn]);

  React.useEffect(() => {
    if (!setExportXlsxFn) return;
    setExportXlsxFn(() => () => {
      const wb = XLSX.utils.book_new();
      const rows = [];
      visibleSpans.forEach((span) => {
        workPackages.filter((w) => {
          const d = new Date(w.start);
          return d >= span.start && d < span.end;
        }).forEach((w) => rows.push(w));
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(rows, WP_COLS)), "PortCall");
      XLSX.writeFile(wb, `genesis-portcall-${todayISO()}.xlsx`);
    });
  }, [visibleSpans, workPackages, setExportXlsxFn]);

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

      <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
        <div className="g-field">
          <label>Período — de</label>
          <input type="date" value={pcPeriod.start} onChange={(e) => setPcPeriod((p) => ({ ...p, start: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>Período — até</label>
          <input type="date" value={pcPeriod.end} onChange={(e) => setPcPeriod((p) => ({ ...p, end: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>&nbsp;</label>
          <button className="g-btn" onClick={() => setPcPeriod({ start: "", end: "" })} disabled={!hasActivePcPeriod} style={{ opacity: hasActivePcPeriod ? 1 : 0.5 }}>
            <X size={13} />Ver todos os Port Calls
          </button>
        </div>
      </div>

      <div className="g-panel-head">
        <span className="g-muted" style={{ fontFamily: "var(--mono)", fontSize: 12 }}>{visibleSpans.length} Port Call(s) no período selecionado</span>
        <button className="g-btn" onClick={addOpCategory}><Plus size={13} />Nova categoria</button>
      </div>

      <div className="g-gantt-wrap">
        <div className="g-gantt">
          {visibleSpans.length === 0 && (
            <div className="g-gantt-empty" style={{ marginLeft: 0, padding: "18px 0" }}>
              Nenhum Port Call encontrado para o período selecionado.
            </div>
          )}

          {visibleSpans.map((span) => {
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

                {pcActivities.length === 0 && (
                  <div className="g-gantt-empty">Nenhuma atividade neste Port Call ainda.</div>
                )}

                {/* ---- categorias operacionais dentro do Port Call ---- */}
                {opCategories.map((cat) => {
                  const rows = pcActivities.filter((w) => catOf(w) === cat);
                  const collapseKey = `${span.startKey}__${cat}`;
                  const isCollapsed = collapsedCats.has(collapseKey);
                  return (
                    <div key={cat} style={{ marginLeft: 14, marginTop: 4 }}>
                      <div className="g-gantt-group-row" style={{ padding: "5px 10px", background: "transparent", border: "1px solid var(--border-soft)" }}>
                        <span className="g-flex" style={{ gap: 6, flex: 1 }}>
                          <span className="g-btn ghost" style={{ padding: 2 }} onClick={() => toggleCollapse(collapseKey)} title={isCollapsed ? "Expandir" : "Recolher"}>
                            {isCollapsed ? <ChevronRight size={12} /> : <ChevronDown size={12} />}
                          </span>
                          <input
                            className="g-gantt-group-title"
                            style={{ fontSize: 10.5, color: "var(--text-dim)" }}
                            value={cat}
                            onChange={(e) => renameOpCategory(cat, e.target.value)}
                          />
                        </span>
                        <span className="g-flex" style={{ gap: 4 }}>
                          <span className="g-muted" style={{ fontFamily: "var(--mono)", fontSize: 9.5 }}>{rows.length}</span>
                          <span className="g-btn ghost" title="Adicionar atividade nesta categoria" onClick={() => addWpOnDate(span.startKey, cat)}><Plus size={13} /></span>
                          <span className="g-btn ghost danger" title="Remover categoria (global)" onClick={() => removeOpCategory(cat)}><Trash2 size={13} /></span>
                        </span>
                      </div>

                      {!isCollapsed && rows.length === 0 && (
                        <div className="g-gantt-empty" style={{ fontSize: 10.5 }}>Nenhuma linha aqui ainda — use o + acima.</div>
                      )}

                      {!isCollapsed && rows.length > 0 && (
                        <div className="g-gantt-datagrid-wrap">
                          {/* ---- coluna fixa: grade de dados, estilo planilha de projeto ---- */}
                          <div className="g-gantt-grid-col">
                            <div className="g-gantt-grid-headrow">
                              <div className="g-gantt-col-num">#</div>
                              <div className="g-gantt-col-task">Tarefa</div>
                              {!NO_EMPRESA_CATEGORIES.includes(cat) && <div className="g-gantt-col-empresa">Empresa</div>}
                              <div className="g-gantt-col-dur">Dur. (h)</div>
                              <div className="g-gantt-col-date">Início</div>
                              <div className="g-gantt-col-date">Término</div>
                              <div className="g-gantt-col-status">Status</div>
                              <div className="g-gantt-col-progress">% Compl.</div>
                              <div className="g-gantt-col-action"></div>
                            </div>
                            {rows.map((w, idx) => {
                              const i = workPackages.indexOf(w);
                              return (
                                <div className="g-gantt-gridrow" key={w.id} style={{ borderLeft: `3px solid ${WP_STATUS_COLOR[w.status] || "#F2C94C"}` }}>
                                  <div className="g-gantt-col-num">{idx + 1}</div>
                                  <div className="g-gantt-col-task">
                                    <textarea className="g-edit-wrap g-gantt-name-edit" rows={1} value={w.name}
                                      onChange={(e) => updWp(i, "name", e.target.value)} />
                                  </div>
                                  {!NO_EMPRESA_CATEGORIES.includes(cat) && (
                                    <div className="g-gantt-col-empresa">
                                      <input type="text" className="g-edit" placeholder="—" value={w.empresa || ""}
                                        onChange={(e) => updWp(i, "empresa", e.target.value)} style={{ fontSize: 10.5 }} />
                                    </div>
                                  )}
                                  <div className="g-gantt-col-dur">
                                    <input type="number" min="0" className="g-gantt-mini-dt" style={{ width: "100%", textAlign: "center" }}
                                      value={w.start && w.end ? Math.max(0, Math.round((new Date(w.end) - new Date(w.start)) / 3600000)) : ""}
                                      onChange={(e) => setTaskDuration(i, w, e.target.value)} />
                                  </div>
                                  <div className="g-gantt-col-date">
                                    <input type="datetime-local" className="g-gantt-mini-dt" style={{ width: "100%" }} value={w.start || ""}
                                      onChange={(e) => updWp(i, "start", e.target.value)} />
                                  </div>
                                  <div className="g-gantt-col-date">
                                    <input type="datetime-local" className="g-gantt-mini-dt" style={{ width: "100%" }} value={w.end || ""}
                                      onChange={(e) => updWp(i, "end", e.target.value)} />
                                  </div>
                                  <div className="g-gantt-col-status">
                                    <StatusServicoSelect value={w.status} onChange={(v) => handleStatusChange(i, v)} />
                                  </div>
                                  <div className="g-gantt-col-progress">
                                    <div className="g-flex" style={{ gap: 4 }}>
                                      <div className="g-bar-bg" style={{ width: 30 }}><div className="g-bar-fg" style={{ width: `${w.progress}%`, background: statusColor(w.status) }} /></div>
                                      <input type="number" min="0" max="100" className="g-edit num" style={{ width: 32, fontSize: 10, padding: "2px 3px" }}
                                        value={w.progress} onChange={(e) => updWp(i, "progress", Number(e.target.value))} />
                                    </div>
                                  </div>
                                  <div className="g-gantt-col-action">
                                    <span className="g-btn ghost danger" onClick={() => remWp(i)}><Trash2 size={12} /></span>
                                  </div>
                                </div>
                              );
                            })}
                          </div>

                          {/* ---- coluna scrollável: linha do tempo visual (só as barras) ---- */}
                          <div className="g-gantt-timeline-col">
                            <div className="g-gantt-grid-headrow" style={{ borderLeft: "none" }}>
                              {dayLabels.map((d, di) => <div className="g-gantt-day" key={di} style={{ flex: 1, textAlign: "center" }}>{d}</div>)}
                            </div>
                            {rows.map((w) => (
                              <div className="g-gantt-timeline-row" key={w.id}>{renderBar(w)}</div>
                            ))}
                          </div>
                        </div>
                      )}
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

function ServicesView({ workPackages, updWp, remWp, repeatWp, expandedWp, setExpandedWp, setReportFn, setExportXlsxFn, newRowId }) {
  React.useEffect(() => {
    if (!newRowId) return;
    const el = document.getElementById(`row-${newRowId}`);
    if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
  }, [newRowId]);
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
  const [groupByPc, setGroupByPc] = useState(true);
  const [collapsedWp, setCollapsedWp] = useState(new Set());
  const toggleWpGroup = (k) => setCollapsedWp((p) => { const n = new Set(p); n.has(k) ? n.delete(k) : n.add(k); return n; });
  const wpGroups = useMemo(() => {
    if (!groupByPc) return [{ key: "todos", label: "Todos", rows: sorted }];
    const map = new Map();
    sorted.forEach((w) => {
      const k = dateKeyOf(w.start) || "sem-data";
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(w);
    });
    return [...map.entries()].sort((x, y) => (x[0] === "sem-data") - (y[0] === "sem-data") || x[0].localeCompare(y[0]))
      .map(([key, rows]) => ({ key, rows, label: key === "sem-data" ? "Sem data definida" : portCallLabel(key) }));
  }, [sorted, groupByPc]);

  const concluidos = filtered.filter((w) => w.status === "Concluído").length;
  const emAndamento = filtered.filter((w) => w.status === "Em andamento").length;
  const naoIniciados = filtered.filter((w) => w.status === "Não iniciado").length;
  const cancelados = filtered.filter((w) => w.status === "Cancelado").length;
  const naoCancelados = filtered.length - cancelados;
  const taxaConclusao = naoCancelados ? Math.round((concluidos / naoCancelados) * 100) : 0;
  const comDesvio = filtered.filter((w) => { const d = desvioDias(w); return d !== null && d !== 0; });
  const desvioMedio = comDesvio.length ? Math.round(comDesvio.reduce((s, w) => s + Math.abs(desvioDias(w)), 0) / comDesvio.length) : 0;

  React.useEffect(() => {
    if (!setExportXlsxFn) return;
    setExportXlsxFn(() => () => {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(filtered, WP_COLS)), "Servicos");
      XLSX.writeFile(wb, `genesis-servicos-${todayISO()}.xlsx`);
    });
  }, [filtered, setExportXlsxFn]);

  React.useEffect(() => {
    if (!setReportFn) return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      let y = pdfHeader(doc, "Relatório de Serviços", `${filtered.length} serviço(s) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [
        { label: "Total de Serviços", value: filtered.length },
        { label: "Concluídos", value: concluidos },
        { label: "Em Andamento", value: emAndamento },
        { label: "Não Iniciados", value: naoIniciados },
        { label: "Cancelados", value: cancelados },
        { label: "Taxa de Conclusão", value: `${taxaConclusao}%` },
        { label: "Serviços com Desvio de Execução", value: comDesvio.length },
        { label: "Desvio Médio (dias)", value: desvioMedio },
      ]);
      y = pdfSectionTitle(doc, y, "Serviços");
      pdfTable(doc, y,
        ["Data", "Categoria", "Manutenção", "Empresa", "RC", "Status", "Progresso", "Desvio"],
        filtered.map((w) => {
          const d = desvioDias(w);
          return [fmtDate((w.start || "").slice(0, 10)), w.group || "—", w.name, w.empresa || "—", w.rc || "—", w.status, `${w.progress}%`, d === null ? "—" : (d > 0 ? `+${d}d` : `${d}d`)];
        })
      );
      pdfSave(doc, "relatorio-servicos");
    });
  }, [filtered, concluidos, emAndamento, naoIniciados, cancelados, taxaConclusao, comDesvio, desvioMedio, setReportFn]);

  return (
    <>
      {/* filtros locais — Portcall e Período agora são só desta aba, junto com Manutenção/Empresa/RC/Status */}
      <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
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
      <div className="g-panel" style={{ padding: 0, overflow: "hidden" }}>
        <div className="tbl-toolbar">
          <label className="tbl-check"><input type="checkbox" checked={groupByPc} onChange={(e) => setGroupByPc(e.target.checked)} />Agrupar por Port Call</label>
          <span className="g-muted" style={{ fontSize: 11 }}>{filtered.length} serviço(s) · clique na seta de uma linha para editar data, categoria, MD, RC e observação</span>
        </div>
        <div className="g-table-wrap">
        <table className="g-table tbl-modern">
          <thead>
            <tr>
              <th style={{ width: 34 }}></th>
              <SortTh sortKey="name" sort={sort} setSort={setSort} style={{ minWidth: 320 }}>Manutenção</SortTh>
              <SortTh sortKey="empresa" sort={sort} setSort={setSort} style={{ minWidth: 140 }}>Empresa</SortTh>
              <SortTh sortKey="status" sort={sort} setSort={setSort} style={{ minWidth: 170 }}>Status</SortTh>
              <SortTh sortKey="progress" sort={sort} setSort={setSort} style={{ minWidth: 150 }}>Progresso</SortTh>
              <th style={{ minWidth: 80 }}>Desvio</th>
              <th style={{ width: 90 }}></th>
            </tr>
          </thead>
          <tbody>
            {wpGroups.map((g) => (
              <React.Fragment key={g.key}>
                {groupByPc && (
                  <tr className="tbl-group" onClick={() => toggleWpGroup(g.key)}>
                    <td colSpan={7}>
                      <span className="tbl-group-arrow">{collapsedWp.has(g.key) ? "▸" : "▾"}</span>
                      {g.label}
                      <span className="tbl-group-cnt">{g.rows.length} serviço(s)</span>
                      <span className="tbl-group-tot">{g.rows.filter((x) => x.status === "Concluído").length} de {g.rows.length} concluído(s)</span>
                    </td>
                  </tr>
                )}
                {!(groupByPc && collapsedWp.has(g.key)) && g.rows.map((w) => {
              const i = workPackages.indexOf(w);
              const isOpen = expandedWp === w.id;
              return (
                <React.Fragment key={w.id}>
                  <tr id={`row-${w.id}`} className={"g-row tbl-row" + (newRowId === w.id ? " g-row-flash" : "")} style={{ "--c": ({ "Concluído": "#22C55E", "Em andamento": "#3B82F6", "Planejamento": "#F5A623", "Não iniciado": "#9499A8", "Cancelado": "#CBD0DC" })[w.status] || "#9499A8", ...(w.status === "Cancelado" ? { opacity: 0.5 } : {}) }}>
                    <td className="tbl-first"></td>
                    <td style={{ minWidth: 320, whiteSpace: "normal", verticalAlign: "top" }}>
                      <div className="g-flex" style={{ gap: 4, alignItems: "flex-start" }}>
                        {w.repeatOf && <span title="Esta linha é uma repetição de um serviço não concluído anteriormente" style={{ fontSize: 13, flexShrink: 0, paddingTop: 5 }}>🔁</span>}
                        {w.status === "Cancelado" && <span title="Cancelado — não é mais necessário" style={{ fontSize: 13, flexShrink: 0, paddingTop: 5 }}>🚫</span>}
                        {w.linkedInvoiceId && <span title="Lançamento gerado automaticamente na aba Pagamentos" style={{ fontSize: 13, flexShrink: 0, paddingTop: 5 }}>💲</span>}
                        <ETextArea rows={1} value={w.name} onChange={(v) => updWp(i, "name", v)} />
                      </div>
                      <div className="tbl-chips">
                        {w.start && <span className="tbl-chip">📅 {fmtDate((w.start || "").slice(0, 10))}</span>}
                        {w.group && <span className="tbl-chip">{w.group}</span>}
                        {w.rc && <span className="tbl-chip">RC {w.rc}</span>}
                        {w.md === "Sim" && <span className="tbl-chip">MD</span>}
                      </div>
                    </td>
                    <td style={{ minWidth: 140 }}><EText value={w.empresa || ""} onChange={(v) => updWp(i, "empresa", v)} /></td>
                    <td style={{ minWidth: 170 }}><StatusServicoSelect value={w.status} onChange={(v) => handleStatusChange(i, v)} /></td>
                    <td style={{ minWidth: 150 }}>
                      <div className="g-flex" style={{ gap: 8 }}>
                        <div className="tbl-progress"><div style={{ width: `${Math.max(0, Math.min(100, Number(w.progress) || 0))}%` }} /></div>
                        <input type="number" min="0" max="100" className="g-edit num" style={{ width: 52 }} value={w.progress}
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
                              fontSize: 11, fontWeight: 700, padding: "3px 8px", borderRadius: 12,
                              color: d === 0 ? "var(--ok)" : "var(--crit)",
                              background: d === 0 ? "rgba(34,197,94,0.12)" : "rgba(239,68,68,0.12)",
                            }} title={d === 0 ? "Executado na data planejada" : d > 0 ? `Executado ${d} dia(s) depois do planejado` : `Executado ${Math.abs(d)} dia(s) antes do planejado`}>
                              {d > 0 ? `+${d}d` : `${d}d`}
                            </span>
                          )}
                        </td>
                      );
                    })()}
                    <td style={{ whiteSpace: "nowrap" }}>
                      {w.status !== "Concluído" && w.status !== "Cancelado" && (
                        <span className="g-btn ghost" onClick={() => repeatWp(w)} title="Ainda não concluído — repetir esta linha para outra data">🔁</span>
                      )}
                      <span className="g-btn ghost" onClick={() => setExpandedWp(isOpen ? null : w.id)} title="Ver e editar todos os detalhes">
                        {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                      </span>
                      <span className="g-btn ghost danger" onClick={() => remWp(i)}><Trash2 size={13} /></span>
                    </td>
                  </tr>
                  {isOpen && (
                    <tr className="g-expand-row">
                      <td></td>
                      <td colSpan={6} style={{ padding: "10px 8px 16px 8px" }}>
                        <div className="g-panel-title" style={{ marginBottom: 8 }}>Dados Gerais</div>
                        <div className="g-flex" style={{ flexWrap: "wrap", gap: 14, marginBottom: 14 }}>
                          <div className="g-field"><label>Data</label><EDate value={(w.start || "").slice(0, 10)} onChange={(v) => setDateKeepTime(i, w, v)} /></div>
                          <div className="g-field" style={{ minWidth: 150 }}><label>Categoria</label><EText value={w.group || ""} onChange={(v) => updWp(i, "group", v)} /></div>
                          <div className="g-field"><label>MD</label><ESelect value={w.md || "Não"} onChange={(v) => updWp(i, "md", v)} options={["Sim", "Não"]} /></div>
                          <div className="g-field" style={{ minWidth: 120 }}><label>RC</label><EText value={w.rc || ""} onChange={(v) => updWp(i, "rc", v)} mono /></div>
                          <div className="g-field" style={{ minWidth: 260, flex: 1 }}><label>Observação</label><ETextArea rows={1} value={w.obs || ""} onChange={(v) => updWp(i, "obs", v)} /></div>
                        </div>
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
                            <div className="g-alert" style={{ background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)", marginBottom: 10 }}>
                              🔁 Esta linha é uma repetição de <strong>{orig ? orig.name : "um serviço anterior"}</strong>
                              {orig && ` — planejado originalmente para ${fmtDate(orig.start?.slice(0, 10))}, ficou como "${orig.status}"`}.
                            </div>
                          );
                        })()}
                        {(() => {
                          const repeats = workPackages.filter((o) => o.repeatOf === w.id);
                          return repeats.length > 0 && (
                            <div className="g-alert" style={{ background: "rgba(59,130,246,0.08)", borderColor: "rgba(59,130,246,0.35)", color: "var(--accent)", marginBottom: 10 }}>
                              Este serviço foi reagendado em {repeats.length} nova(s) linha(s): {repeats.map((r) => r.name).join(", ")}.
                            </div>
                          );
                        })()}
                        {w.status !== "Concluído" && w.status !== "Cancelado" && (
                          <div className="g-flex" style={{ gap: 8 }}>
                            <button className="g-btn" onClick={() => repeatWp(w)}>
                              🔁 Ainda não concluído — repetir em outra data
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
              </React.Fragment>
            ))}
          </tbody>
        </table>
        </div>
      </div>
      )}
    </>
  );
}

/* ============================================================
   PLANEJAMENTO — mapeia manutenções em atraso: plano de ação para
   concluir, necessidade de material, e impacto de cada atraso
   ============================================================ */
function PlanejamentoView({ workPackages, updWp, materials, setReportFn, setExportXlsxFn, allPortCallDates, portCallLabel, addWpOnDate,
  planningItems, updPlan, remPlan, addPlanningItem, newRowId, handleImportPlanejamento,
  docagemItems, updDocagem, remDocagem, addDocagemItem, handleImportDocagem, planSubTab, setPlanSubTab }) {
  const [showPast, setShowPast] = useState(false);
  const [expandedCard, setExpandedCard] = useState(null);
  const importFileRef = useRef(null);
  const docagemFileRef = useRef(null);
  const todayKey = todayISO();

  const dateKeyOf = (dt) => (dt ? dt.slice(0, 10) : null);
  const materiaisPendentesDe = (serviceId) => materials.filter((m) => m.wp === serviceId && !["Recebido", "Entregue a bordo"].includes(m.status));

  React.useEffect(() => {
    if (!newRowId) return;
    const el = document.getElementById(`row-${newRowId}`);
    if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
  }, [newRowId]);

  /* ========================================================
     MAPEADOS PARA EXECUÇÃO — lista independente da aba Serviços.
     Um item aqui só passa a existir também como um Serviço de verdade
     (no Gantt/Port Call) quando ganha uma Data de Execução.
     ======================================================== */
  const [mf, setMf] = useState({ busca: "", departamento: "", empresa: "", impacto: "Todos", statuses: PLAN_STATUS_PADRAO, execInicio: "", execFim: "" });
  const hasActiveFilterMap = mf.busca || mf.departamento || mf.empresa || mf.impacto !== "Todos" || !mesmoConjunto(mf.statuses, PLAN_STATUS_PADRAO) || mf.execInicio || mf.execFim;
  const filteredMapeados = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return planningItems.filter((p) => {
      const inBusca = !mf.busca || norm(p.nome).includes(norm(mf.busca));
      const inDep = !mf.departamento || norm(p.departamento).includes(norm(mf.departamento));
      const inEmpresa = !mf.empresa || norm(p.empresa).includes(norm(mf.empresa));
      const inImpacto = mf.impacto === "Todos" || (p.impacto || "Baixo") === mf.impacto;
      const inStatus = mf.statuses.length === 0 || mf.statuses.includes(p.status || "A Executar");
      const inExec = (!mf.execInicio || (p.dataExecucao && p.dataExecucao >= mf.execInicio)) && (!mf.execFim || (p.dataExecucao && p.dataExecucao <= mf.execFim));
      return inBusca && inDep && inEmpresa && inImpacto && inStatus && inExec;
    });
  }, [planningItems, mf]);

  const totalMapeados = planningItems.length;
  const semDataExecucao = planningItems.filter((p) => !p.dataExecucao).length;
  const jaExecutando = planningItems.filter((p) => !!p.linkedServiceId).length;
  const precisamMaterialMap = planningItems.filter((p) => p.precisaMaterial).length;

  /* ========================================================
     QUADRO POR PORT CALL — visão dos serviços já com data de execução (já "graduados"
     pra Serviços), organizados por Port Call, com opção de redistribuir entre eles
     ======================================================== */
  const pendentes = useMemo(() => workPackages.filter((w) => !["Concluído", "Cancelado"].includes(w.status)), [workPackages]);
  const colunas = useMemo(() => {
    const dates = [...new Set([...allPortCallDates, ...pendentes.map((w) => dateKeyOf(w.start)).filter(Boolean)])].sort();
    const visibleDates = showPast ? dates : dates.filter((dk) => dk >= todayKey);
    return visibleDates.map((dk) => ({
      dateKey: dk,
      label: portCallLabel(dk),
      itens: pendentes.filter((w) => dateKeyOf(w.start) === dk),
    }));
  }, [allPortCallDates, pendentes, showPast, todayKey]);

  const moverParaPortCall = (w, targetDateKey) => {
    const i = workPackages.indexOf(w);
    const durMs = w.start && w.end ? (new Date(w.end) - new Date(w.start)) : 8 * 3600000;
    const newStart = new Date(`${targetDateKey}T08:00:00`);
    const newEnd = new Date(newStart.getTime() + durMs);
    const pad = (n) => String(n).padStart(2, "0");
    const toLocal = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    updWp(i, "start", toLocal(newStart));
    updWp(i, "end", toLocal(newEnd));
  };

  /* ========================================================
     DOCAGEM — Itens de Machinery Items (DNV) que precisam ser vistoriados/feitos na docagem
     ======================================================== */
  const [df, setDf] = useState({ busca: "", localizacao: "", necessitaMaterial: "Todos", statuses: PLAN_STATUS_PADRAO });
  const hasActiveFilterDoc = df.busca || df.localizacao || df.necessitaMaterial !== "Todos" || !mesmoConjunto(df.statuses, PLAN_STATUS_PADRAO);
  const filteredDocagem = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return docagemItems.filter((d) => {
      const inBusca = !df.busca || norm(d.nome).includes(norm(df.busca));
      const inLoc = !df.localizacao || norm(d.localizacao).includes(norm(df.localizacao));
      const inMat = df.necessitaMaterial === "Todos" || (df.necessitaMaterial === "Sim" ? d.necessitaMaterial : !d.necessitaMaterial);
      const inStatus = df.statuses.length === 0 || df.statuses.includes(d.status || "A Executar");
      return inBusca && inLoc && inMat && inStatus;
    });
  }, [docagemItems, df]);
  const totalDocagem = docagemItems.length;
  const necessitamMaterialDoc = docagemItems.filter((d) => d.necessitaMaterial).length;
  const concluidosDoc = docagemItems.filter((d) => d.status === "Concluído").length;
  const pendentesDoc = docagemItems.filter((d) => d.status !== "Concluído" && d.status !== "Cancelado").length;

  const [groupMap, setGroupMap] = useState(true);
  const [collapsedMap, setCollapsedMap] = useState(new Set());
  const [expandedPlanRow, setExpandedPlanRow] = useState(null);
  const [groupDoc, setGroupDoc] = useState(true);
  const [collapsedDoc, setCollapsedDoc] = useState(new Set());
  const [expandedDocRow, setExpandedDocRow] = useState(null);
  const groupRows = (rows, on, keyFn, fallback) => {
    if (!on) return [{ key: "Todos", rows }];
    const map = new Map();
    rows.forEach((r) => { const k = (keyFn(r) || "").toString().trim() || fallback; if (!map.has(k)) map.set(k, []); map.get(k).push(r); });
    return [...map.entries()].sort((x, y) => (x[0] === fallback) - (y[0] === fallback) || x[0].localeCompare(y[0], "pt-BR")).map(([key, rs]) => ({ key, rows: rs }));
  };
  const mapGroups = useMemo(() => groupRows(filteredMapeados, groupMap, (p) => p.departamento, "Sem departamento"), [filteredMapeados, groupMap]);
  const docGroups = useMemo(() => groupRows(filteredDocagem, groupDoc, (d) => d.localizacao, "Sem localização"), [filteredDocagem, groupDoc]);

  React.useEffect(() => {
    if (!setExportXlsxFn) return;
    setExportXlsxFn(() => () => {
      const wb = XLSX.utils.book_new();
      if (planSubTab === "mapeados") {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(filteredMapeados, PLAN_COLS)), "Mapeados");
        XLSX.writeFile(wb, `genesis-planejamento-mapeados-${todayISO()}.xlsx`);
      } else if (planSubTab === "docagem") {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(filteredDocagem, DOC_COLS)), "Docagem");
        XLSX.writeFile(wb, `genesis-docagem-${todayISO()}.xlsx`);
      } else {
        const rows = [];
        colunas.forEach((col) => col.itens.forEach((w) => rows.push(w)));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(rows, WP_COLS)), "QuadroPortCall");
        XLSX.writeFile(wb, `genesis-planejamento-quadro-${todayISO()}.xlsx`);
      }
    });
  }, [planSubTab, filteredMapeados, filteredDocagem, colunas, setExportXlsxFn]);

  React.useEffect(() => {
    if (!setReportFn) return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      if (planSubTab === "mapeados") {
        let y = pdfHeader(doc, "Relatório de Planejamento — Mapeados para Execução",
          `${filteredMapeados.length} item(ns) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
        y = pdfKpis(doc, y, [
          { label: "Total Mapeado", value: totalMapeados },
          { label: "Sem Data de Execução", value: semDataExecucao },
          { label: "Já em Execução (Serviço criado)", value: jaExecutando },
          { label: "Precisam de Material", value: precisamMaterialMap },
        ]);
        y = pdfSectionTitle(doc, y, "Itens mapeados");
        pdfTable(doc, y,
          ["Nome", "Departamento", "Empresa", "Impacto", "Precisa Material", "PO", "Data de Execução", "Status"],
          filteredMapeados.map((p) => [
            p.nome, p.departamento || "—", p.empresa || "—", p.impacto || "Baixo",
            p.precisaMaterial ? "Sim" : "Não", p.precisaMaterial ? (p.poMaterial || "—") : "—",
            p.dataExecucao ? fmtDate(p.dataExecucao) : "—", p.status || "A Executar",
          ])
        );
        pdfSave(doc, "relatorio-planejamento-mapeados");
      } else if (planSubTab === "docagem") {
        let y = pdfHeader(doc, "Relatório de Docagem — Machinery Items (DNV)",
          `${filteredDocagem.length} item(ns) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
        y = pdfKpis(doc, y, [
          { label: "Total de Itens", value: totalDocagem },
          { label: "Necessitam de Material", value: necessitamMaterialDoc },
          { label: "Pendentes", value: pendentesDoc },
          { label: "Concluídos", value: concluidosDoc },
        ]);
        y = pdfSectionTitle(doc, y, "Itens de Docagem");
        pdfTable(doc, y,
          ["Nome", "Localização", "Tipo", "Empresa", "Plano de Ação", "Necessita Material", "PO", "Previsão Execução", "Data Conclusão", "Status"],
          filteredDocagem.map((d) => [
            d.nome, d.localizacao || "—", d.tipoPeriodo || "—", d.empresa || "—", d.planoAcao || "—",
            d.necessitaMaterial ? "Sim" : "Não", d.necessitaMaterial ? (d.poRelacionada || "—") : "—",
            d.previsaoExecucao ? fmtDate(d.previsaoExecucao) : "—", d.dataConclusao ? fmtDate(d.dataConclusao) : "—",
            d.status || "A Executar",
          ])
        );
        pdfSave(doc, "relatorio-docagem");
      } else {
        let y = pdfHeader(doc, "Relatório de Planejamento — Quadro por Port Call",
          `${pendentes.length} serviço(s) pendente(s) · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
        colunas.forEach((col) => {
          if (col.itens.length === 0) return;
          y = pdfSectionTitle(doc, y, `${col.label} — ${col.itens.length} serviço(s)`);
          y = pdfTable(doc, y,
            ["Serviço", "Empresa", "Impacto", "Plano de Ação"],
            col.itens.map((w) => [w.name, w.empresa || "—", w.impacto || "Baixo", w.planoAcao || "—"])
          );
        });
        pdfSave(doc, "relatorio-planejamento-quadro");
      }
    });
  }, [planSubTab, filteredMapeados, totalMapeados, semDataExecucao, jaExecutando, precisamMaterialMap, colunas, pendentes,
      filteredDocagem, totalDocagem, necessitamMaterialDoc, pendentesDoc, concluidosDoc, setReportFn]);

  return (
    <>
      <div className="g-mode-toggle" style={{ marginBottom: 16, width: "fit-content" }}>
        <button className={planSubTab === "mapeados" ? "active" : ""} onClick={() => setPlanSubTab("mapeados")}>Mapeados para Execução</button>
        <button className={planSubTab === "docagem" ? "active" : ""} onClick={() => setPlanSubTab("docagem")}>Docagem (Machinery Items - DNV)</button>
      </div>

      {planSubTab === "mapeados" ? (
        <>
          <div className="g-panel" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
            <div>
              <div className="g-panel-title" style={{ marginBottom: 4 }}>Importar planilha de mapeamento</div>
              <div className="g-muted" style={{ fontSize: 11.5 }}>
                Aceita arquivos no padrão Título/Área/Prioridade/Status/Descrição/RC/Observação (ex.: "Saúde do Ativo").
                Lê todas as abas do arquivo e importa tudo, independente do status. Mescla por Título — reimportar não duplica.
              </div>
            </div>
            <div>
              <input ref={importFileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImportPlanejamento} />
              <button className="g-btn primary" onClick={() => importFileRef.current?.click()}>
                <Upload size={14} />Importar planilha
              </button>
            </div>
          </div>

          <div className="g-alert" style={{ background: "rgba(37,104,160,0.08)", borderColor: "rgba(37,104,160,0.35)", color: "var(--accent)" }}>
            Cadastre aqui toda manutenção que você precisa mapear — mesmo sem saber ainda quando ela vai acontecer.
            Assim que você preencher a <strong>Data de Execução</strong> de um item, ele passa a aparecer também
            na aba Serviços e no Port Call automaticamente, como um serviço de verdade agendado.
          </div>

          <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
            <div className="g-field">
              <label>Nome</label>
              <input type="text" value={mf.busca} onChange={(e) => setMf((p) => ({ ...p, busca: e.target.value }))} placeholder="digitar..." style={{ minWidth: 160 }} />
            </div>
            <div className="g-field">
              <label>Departamento</label>
              <input type="text" value={mf.departamento} onChange={(e) => setMf((p) => ({ ...p, departamento: e.target.value }))} placeholder="digitar..." style={{ minWidth: 130 }} />
            </div>
            <div className="g-field">
              <label>Empresa</label>
              <input type="text" value={mf.empresa} onChange={(e) => setMf((p) => ({ ...p, empresa: e.target.value }))} placeholder="digitar..." style={{ minWidth: 130 }} />
            </div>
            <div className="g-field">
              <label>Impacto</label>
              <select value={mf.impacto} onChange={(e) => setMf((p) => ({ ...p, impacto: e.target.value }))}>
                <option>Todos</option>
                {IMPACT_LEVELS.map((l) => <option key={l} value={l}>{l}</option>)}
              </select>
            </div>
            <div className="g-field">
              <label>Status</label>
              <MultiSelectStatus options={PLAN_STATUS} selected={mf.statuses} onChange={(v) => setMf((p) => ({ ...p, statuses: v }))} />
            </div>
            <div className="g-field">
              <label>Período de Execução — de</label>
              <input type="date" value={mf.execInicio} onChange={(e) => setMf((p) => ({ ...p, execInicio: e.target.value }))} />
            </div>
            <div className="g-field">
              <label>Período de Execução — até</label>
              <input type="date" value={mf.execFim} onChange={(e) => setMf((p) => ({ ...p, execFim: e.target.value }))} />
            </div>
            <div className="g-field">
              <label>&nbsp;</label>
              <button className="g-btn" onClick={() => setMf({ busca: "", departamento: "", empresa: "", impacto: "Todos", statuses: PLAN_STATUS_PADRAO, execInicio: "", execFim: "" })}
                disabled={!hasActiveFilterMap} style={{ opacity: hasActiveFilterMap ? 1 : 0.5 }}>
                <X size={13} />Limpar filtro
              </button>
            </div>
          </div>

          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
            {bigKpi("Total Mapeado", totalMapeados, "var(--teal)", ClipboardList)}
            {bigKpi("Sem Data de Execução", semDataExecucao, "var(--warn)", Clock)}
            {bigKpi("Já em Execução", jaExecutando, "var(--ok)", ClipboardList)}
            {bigKpi("Precisam de Material", precisamMaterialMap, "var(--crit)", Package)}
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Itens mapeados para execução ({filteredMapeados.length})</span></div>
            <div className="tbl-toolbar" style={{ margin: "-18px -18px 0 -18px", borderRadius: "12px 12px 0 0" }}>
              <label className="tbl-check"><input type="checkbox" checked={groupMap} onChange={(e) => setGroupMap(e.target.checked)} />Agrupar por departamento</label>
              <span className="g-muted" style={{ fontSize: 11 }}>clique na seta de uma linha para editar problema, plano de ação, empresa e departamento</span>
            </div>
            <div className="g-table-wrap">
            <table className="g-table tbl-modern">
              <thead>
                <tr>
                  <th style={{ width: 34 }}></th>
                  <th style={{ minWidth: 300 }}>Item</th>
                  <th style={{ minWidth: 140 }}>Data de Execução</th>
                  <th>Impacto</th>
                  <th style={{ minWidth: 150 }}>Material</th>
                  <th style={{ minWidth: 150 }}>Status</th>
                  <th style={{ width: 70 }}></th>
                </tr>
              </thead>
              <tbody>
                {mapGroups.map((g) => (
                  <React.Fragment key={g.key}>
                    {groupMap && (
                      <tr className="tbl-group" onClick={() => setCollapsedMap((p) => { const n = new Set(p); n.has(g.key) ? n.delete(g.key) : n.add(g.key); return n; })}>
                        <td colSpan={7}>
                          <span className="tbl-group-arrow">{collapsedMap.has(g.key) ? "▸" : "▾"}</span>
                          {g.key}
                          <span className="tbl-group-cnt">{g.rows.length} item(ns)</span>
                          <span className="tbl-group-tot">{g.rows.filter((x) => !x.dataExecucao).length > 0 ? `${g.rows.filter((x) => !x.dataExecucao).length} sem data de execução` : "todos com data"}</span>
                        </td>
                      </tr>
                    )}
                    {!(groupMap && collapsedMap.has(g.key)) && g.rows.map((p) => {
                  const i = planningItems.indexOf(p);
                  const isOpen = expandedPlanRow === p.id;
                  return (
                    <React.Fragment key={p.id}>
                    <tr id={`row-${p.id}`} className={"g-row tbl-row" + (newRowId === p.id ? " g-row-flash" : "")} style={{ "--c": PLAN_STATUS_COLOR[p.status || "A Executar"] || "#9499A8" }}>
                      <td className="tbl-first"></td>
                      <td style={{ minWidth: 300, whiteSpace: "normal", verticalAlign: "top" }}>
                        <ETextArea rows={1} value={p.nome} onChange={(v) => updPlan(i, "nome", v)} />
                        <div className="tbl-chips">
                          {p.empresa && <span className="tbl-chip">{p.empresa}</span>}
                          {p.precisaMaterial && p.poMaterial && <span className="tbl-chip">PO {p.poMaterial}</span>}
                          {p.planoAcao && <span className="tbl-chip" title={p.planoAcao}>com plano de ação</span>}
                        </div>
                      </td>
                      <td style={{ minWidth: 140 }}>
                        {p.dataExecucao
                          ? <input type="date" value={p.dataExecucao} onChange={(e) => updPlan(i, "dataExecucao", e.target.value)}
                              style={{ background: "var(--panel-raised)", border: "1px solid var(--border)", color: "var(--text)", borderRadius: 6, padding: "5px 6px", fontSize: 11.5 }} />
                          : <div className="g-flex" style={{ gap: 6 }}><span style={{ color: "var(--text-faint)", fontSize: 11 }}>sem data</span>
                              <input type="date" value="" onChange={(e) => updPlan(i, "dataExecucao", e.target.value)}
                                style={{ background: "var(--panel-raised)", border: "1px solid var(--border)", color: "var(--text)", borderRadius: 6, padding: "5px 6px", fontSize: 11.5 }} /></div>}
                      </td>
                      <td>
                        <select value={p.impacto || "Baixo"} onChange={(e) => updPlan(i, "impacto", e.target.value)}
                          style={{ background: "var(--panel-raised)", border: "1px solid var(--border)", color: IMPACT_COLOR[p.impacto || "Baixo"], fontWeight: 700, borderRadius: 99, padding: "4px 8px", fontSize: 11.5 }}>
                          {IMPACT_LEVELS.map((l) => <option key={l} value={l}>{l}</option>)}
                        </select>
                      </td>
                      <td style={{ minWidth: 150 }}>
                        <label className="g-flex" style={{ gap: 6, fontSize: 11.5, cursor: "pointer" }}>
                          <input type="checkbox" checked={!!p.precisaMaterial} onChange={(e) => updPlan(i, "precisaMaterial", e.target.checked)} />
                          Precisa de material
                        </label>
                        {p.precisaMaterial && (
                          <input type="text" className="g-edit" placeholder="nº da PO..." value={p.poMaterial || ""} onChange={(e) => updPlan(i, "poMaterial", e.target.value)}
                            style={{ fontSize: 11, background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 6, padding: "4px 6px", width: "100%", marginTop: 4 }} />
                        )}
                      </td>
                      <td style={{ minWidth: 150 }}>
                        <select value={p.status || "A Executar"} onChange={(e) => updPlan(i, "status", e.target.value)}
                          style={{ background: "var(--panel-raised)", border: "1px solid var(--border)", color: PLAN_STATUS_COLOR[p.status || "A Executar"], fontWeight: 700, borderRadius: 6, padding: "5px 8px", fontSize: 11.5, width: "100%" }}>
                          {PLAN_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
                        </select>
                        {p.linkedServiceId && (
                          <div style={{ marginTop: 4, fontSize: 10, color: "var(--ok)" }}>🔗 Serviço criado</div>
                        )}
                      </td>
                      <td style={{ whiteSpace: "nowrap" }}>
                        <span className="g-btn ghost" onClick={() => setExpandedPlanRow(isOpen ? null : p.id)} title="Ver todos os detalhes">
                          {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                        </span>
                        <span className="g-btn ghost danger" onClick={() => remPlan(i)}><Trash2 size={13} /></span>
                      </td>
                    </tr>
                    {isOpen && (
                      <tr className="tbl-detail">
                        <td></td>
                        <td colSpan={6}>
                          <div className="tbl-detail-grid">
                            <div className="g-field"><label>Departamento</label><EText value={p.departamento} onChange={(v) => updPlan(i, "departamento", v)} /></div>
                            <div className="g-field"><label>Empresa</label><EText value={p.empresa} onChange={(v) => updPlan(i, "empresa", v)} /></div>
                            <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Descrição do Problema</label><ETextArea rows={1} value={p.descricaoProblema} onChange={(v) => updPlan(i, "descricaoProblema", v)} /></div>
                            <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Plano de Ação</label><ETextArea rows={1} value={p.planoAcao} onChange={(v) => updPlan(i, "planoAcao", v)} /></div>
                          </div>
                        </td>
                      </tr>
                    )}
                    </React.Fragment>
                  );
                })}
                  </React.Fragment>
                ))}
              </tbody>
            </table>
            </div>
            {filteredMapeados.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum item mapeado ainda — use o botão "Novo mapeamento" no topo da página.</div>}
          </div>
        </>
      ) : planSubTab === "docagem" ? (
        <>
          <div className="g-panel" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
            <div>
              <div className="g-panel-title" style={{ marginBottom: 4 }}>Importar planilha de Machinery Items (DNV)</div>
              <div className="g-muted" style={{ fontSize: 11.5 }}>
                Aceita o arquivo exportado do site da DNV, no padrão Name/ItemLocation/PeriodType. Mescla por Nome —
                reimportar a lista atualizada não duplica, só atualiza localização/tipo dos itens já cadastrados.
              </div>
            </div>
            <div className="g-flex" style={{ gap: 8 }}>
              <input ref={docagemFileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImportDocagem} />
              <button className="g-btn" onClick={() => docagemFileRef.current?.click()}>
                <Upload size={14} />Importar planilha DNV
              </button>
              <button className="g-btn primary" onClick={addDocagemItem}>
                <Plus size={14} />Novo item
              </button>
            </div>
          </div>

          <div className="g-alert" style={{ background: "rgba(37,104,160,0.08)", borderColor: "rgba(37,104,160,0.35)", color: "var(--accent)" }}>
            Itens classificados como <strong>Machinery Items</strong> pela DNV que precisam ser vistoriados e/ou feitos
            durante a docagem. Preencha o plano de ação, se precisa de material, PO relacionada, previsão de execução,
            data de conclusão e o status de cada item.
          </div>

          <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
            <div className="g-field">
              <label>Nome</label>
              <input type="text" value={df.busca} onChange={(e) => setDf((p) => ({ ...p, busca: e.target.value }))} placeholder="digitar..." style={{ minWidth: 160 }} />
            </div>
            <div className="g-field">
              <label>Localização</label>
              <input type="text" value={df.localizacao} onChange={(e) => setDf((p) => ({ ...p, localizacao: e.target.value }))} placeholder="digitar..." style={{ minWidth: 140 }} />
            </div>
            <div className="g-field">
              <label>Necessita de Material</label>
              <select value={df.necessitaMaterial} onChange={(e) => setDf((p) => ({ ...p, necessitaMaterial: e.target.value }))}>
                <option>Todos</option><option>Sim</option><option>Não</option>
              </select>
            </div>
            <div className="g-field">
              <label>Status</label>
              <MultiSelectStatus options={PLAN_STATUS} selected={df.statuses} onChange={(v) => setDf((p) => ({ ...p, statuses: v }))} />
            </div>
            <div className="g-field">
              <label>&nbsp;</label>
              <button className="g-btn" onClick={() => setDf({ busca: "", localizacao: "", necessitaMaterial: "Todos", statuses: PLAN_STATUS_PADRAO })}
                disabled={!hasActiveFilterDoc} style={{ opacity: hasActiveFilterDoc ? 1 : 0.5 }}>
                <X size={13} />Limpar filtro
              </button>
            </div>
          </div>

          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
            {bigKpi("Total de Itens", totalDocagem, "var(--teal)", ClipboardList)}
            {bigKpi("Necessitam de Material", necessitamMaterialDoc, "var(--warn)", Package)}
            {bigKpi("Pendentes", pendentesDoc, "var(--crit)", AlertTriangle)}
            {bigKpi("Concluídos", concluidosDoc, "var(--ok)", ClipboardList)}
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Itens de Docagem — Machinery Items DNV ({filteredDocagem.length})</span></div>
            <div className="tbl-toolbar" style={{ margin: "-18px -18px 0 -18px", borderRadius: "12px 12px 0 0" }}>
              <label className="tbl-check"><input type="checkbox" checked={groupDoc} onChange={(e) => setGroupDoc(e.target.checked)} />Agrupar por localização</label>
              <span className="g-muted" style={{ fontSize: 11 }}>clique na seta de uma linha para editar localização, tipo, empresa, plano de ação e conclusão</span>
            </div>
            <div className="g-table-wrap">
            <table className="g-table tbl-modern">
              <thead>
                <tr>
                  <th style={{ width: 34 }}></th>
                  <th style={{ minWidth: 300 }}>Item</th>
                  <th style={{ minWidth: 150 }}>Material</th>
                  <th style={{ minWidth: 140 }}>Previsão de Execução</th>
                  <th style={{ minWidth: 150 }}>Status</th>
                  <th style={{ width: 70 }}></th>
                </tr>
              </thead>
              <tbody>
                {docGroups.map((g) => (
                  <React.Fragment key={g.key}>
                    {groupDoc && (
                      <tr className="tbl-group" onClick={() => setCollapsedDoc((p) => { const n = new Set(p); n.has(g.key) ? n.delete(g.key) : n.add(g.key); return n; })}>
                        <td colSpan={6}>
                          <span className="tbl-group-arrow">{collapsedDoc.has(g.key) ? "▸" : "▾"}</span>
                          {g.key}
                          <span className="tbl-group-cnt">{g.rows.length} item(ns)</span>
                          <span className="tbl-group-tot">{g.rows.filter((x) => x.status === "Concluído").length} de {g.rows.length} concluído(s)</span>
                        </td>
                      </tr>
                    )}
                    {!(groupDoc && collapsedDoc.has(g.key)) && g.rows.map((d) => {
                  const i = docagemItems.indexOf(d);
                  const isOpen = expandedDocRow === d.id;
                  return (
                    <React.Fragment key={d.id}>
                    <tr id={`row-${d.id}`} className={"g-row tbl-row" + (newRowId === d.id ? " g-row-flash" : "")} style={{ "--c": PLAN_STATUS_COLOR[d.status || "A Executar"] || "#9499A8" }}>
                      <td className="tbl-first"></td>
                      <td style={{ minWidth: 300, whiteSpace: "normal", verticalAlign: "top" }}>
                        <ETextArea rows={1} value={d.nome} onChange={(v) => updDocagem(i, "nome", v)} />
                        <div className="tbl-chips">
                          {d.tipoPeriodo && <span className="tbl-chip">{d.tipoPeriodo}</span>}
                          {d.empresa && <span className="tbl-chip">{d.empresa}</span>}
                          {d.necessitaMaterial && d.poRelacionada && <span className="tbl-chip">PO {d.poRelacionada}</span>}
                          {d.dataConclusao && <span className="tbl-chip">Concluído {fmtDate(d.dataConclusao)}</span>}
                        </div>
                      </td>
                      <td style={{ minWidth: 150 }}>
                        <label className="g-flex" style={{ gap: 6, fontSize: 11.5, cursor: "pointer" }}>
                          <input type="checkbox" checked={!!d.necessitaMaterial} onChange={(e) => updDocagem(i, "necessitaMaterial", e.target.checked)} />
                          Necessita de material
                        </label>
                        {d.necessitaMaterial && (
                          <input type="text" className="g-edit" placeholder="nº da PO..." value={d.poRelacionada || ""} onChange={(e) => updDocagem(i, "poRelacionada", e.target.value)}
                            style={{ fontSize: 11, background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 6, padding: "4px 6px", width: "100%", marginTop: 4 }} />
                        )}
                      </td>
                      <td style={{ minWidth: 140 }}>
                        <input type="date" value={d.previsaoExecucao || ""} onChange={(e) => updDocagem(i, "previsaoExecucao", e.target.value)}
                          style={{ background: "var(--panel-raised)", border: "1px solid var(--border)", color: "var(--text)", borderRadius: 6, padding: "5px 6px", fontSize: 11.5 }} />
                      </td>
                      <td style={{ minWidth: 150 }}>
                        <select value={d.status || "A Executar"} onChange={(e) => updDocagem(i, "status", e.target.value)}
                          style={{ background: "var(--panel-raised)", border: "1px solid var(--border)", color: PLAN_STATUS_COLOR[d.status || "A Executar"], fontWeight: 700, borderRadius: 6, padding: "5px 8px", fontSize: 11.5, width: "100%" }}>
                          {PLAN_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
                        </select>
                        {d.linkedServiceId && (
                          <div style={{ marginTop: 4, fontSize: 10, color: "var(--ok)" }}>🔗 Serviço criado</div>
                        )}
                      </td>
                      <td style={{ whiteSpace: "nowrap" }}>
                        <span className="g-btn ghost" onClick={() => setExpandedDocRow(isOpen ? null : d.id)} title="Ver todos os detalhes">
                          {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                        </span>
                        <span className="g-btn ghost danger" onClick={() => remDocagem(i)}><Trash2 size={13} /></span>
                      </td>
                    </tr>
                    {isOpen && (
                      <tr className="tbl-detail">
                        <td></td>
                        <td colSpan={5}>
                          <div className="tbl-detail-grid">
                            <div className="g-field"><label>Localização</label><EText value={d.localizacao} onChange={(v) => updDocagem(i, "localizacao", v)} /></div>
                            <div className="g-field"><label>Tipo</label><EText value={d.tipoPeriodo} onChange={(v) => updDocagem(i, "tipoPeriodo", v)} /></div>
                            <div className="g-field"><label>Empresa</label><EText value={d.empresa} onChange={(v) => updDocagem(i, "empresa", v)} /></div>
                            <div className="g-field"><label>Data de Conclusão</label>
                              <input type="date" value={d.dataConclusao || ""} onChange={(e) => updDocagem(i, "dataConclusao", e.target.value)}
                                style={{ background: "var(--panel-raised)", border: "1px solid var(--border)", color: "var(--text)", borderRadius: 6, padding: "6px 8px", fontSize: 12 }} /></div>
                            <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Plano de Ação</label><ETextArea rows={1} value={d.planoAcao} onChange={(v) => updDocagem(i, "planoAcao", v)} /></div>
                          </div>
                        </td>
                      </tr>
                    )}
                    </React.Fragment>
                  );
                })}
                  </React.Fragment>
                ))}
              </tbody>
            </table>
            </div>
            {filteredDocagem.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum item de docagem ainda — importe a planilha da DNV ou use "Novo item".</div>}
          </div>
        </>
      ) : (
        <>
          <div className="g-alert" style={{ background: "rgba(37,104,160,0.08)", borderColor: "rgba(37,104,160,0.35)", color: "var(--accent)" }}>
            Aqui aparecem os serviços que já têm data de execução (vindos da aba Serviços, incluindo os que
            acabaram de "graduar" do mapeamento acima), organizados pelo Port Call em que estão agendados.
            Use "Mover para" pra redistribuir um serviço entre os Port Calls.
          </div>

          <div className="g-flex" style={{ justifyContent: "flex-end", marginBottom: 10 }}>
            <label className="g-flex" style={{ gap: 6, fontSize: 12, cursor: "pointer" }}>
              <input type="checkbox" checked={showPast} onChange={(e) => setShowPast(e.target.checked)} />
              Mostrar Port Calls passados também
            </label>
          </div>

          {colunas.length === 0 && (
            <div className="g-panel"><div className="g-muted">Nenhum Port Call encontrado. Cadastre um na aba Port Call primeiro.</div></div>
          )}

          <div style={{ display: "flex", gap: 12, overflowX: "auto", paddingBottom: 8 }}>
            {colunas.map((col) => (
              <div key={col.dateKey} style={{ minWidth: 280, maxWidth: 280, flexShrink: 0 }}>
                <div className="g-panel" style={{ marginBottom: 0, height: "100%" }}>
                  <div className="g-panel-head" style={{ marginBottom: 10 }}>
                    <span className="g-panel-title" style={{ fontSize: 12.5 }}>{col.label}</span>
                    <span className="g-muted" style={{ fontFamily: "var(--mono)", fontSize: 11 }}>{col.itens.length}</span>
                  </div>
                  {col.itens.length === 0 && <div className="g-muted" style={{ fontSize: 11.5 }}>Nada agendado aqui.</div>}
                  {col.itens.map((w) => {
                    const isOpen = expandedCard === w.id;
                    const precisaMat = w.precisaMaterial || materiaisPendentesDe(w.id).length > 0;
                    return (
                      <div key={w.id}
                        style={{
                          background: "var(--panel-raised)", border: "1px solid var(--border)", borderLeft: `3px solid ${WP_STATUS_COLOR[w.status] || "var(--border)"}`,
                          borderRadius: 5, padding: "8px 9px", marginBottom: 8, cursor: "pointer",
                        }}
                        onClick={() => setExpandedCard(isOpen ? null : w.id)}>
                        <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 4 }}>{w.name}</div>
                        <div className="g-flex" style={{ gap: 5, flexWrap: "wrap", marginBottom: 4 }}>
                          <Pill status={w.status} />
                          {w.impacto && <span className="g-pill" style={{ background: "var(--panel)" }}><span className="g-dot" style={{ background: IMPACT_COLOR[w.impacto] }} />{w.impacto}</span>}
                          {precisaMat && <span title="Precisa de material" style={{ fontSize: 12 }}>📦</span>}
                        </div>
                        <div style={{ fontSize: 10.5, color: "var(--text-faint)" }}>{w.empresa || "sem empresa definida"}</div>
                        <div onClick={(e) => e.stopPropagation()} style={{ marginTop: 6 }}>
                          <select value={col.dateKey} onChange={(e) => moverParaPortCall(w, e.target.value)}
                            style={{ width: "100%", fontSize: 10.5, background: "var(--panel)", border: "1px solid var(--border)", borderRadius: 3, padding: "3px 4px" }}>
                            {colunas.map((c) => <option key={c.dateKey} value={c.dateKey}>Mover para: {c.label}</option>)}
                          </select>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </>
      )}
    </>
  );
}

/* ============================================================
   TM MASTER — importação semanal do sistema de manutenção: "Due" (tudo em aberto,
   vencido ou a vencer) e "History" (tudo que já foi fechado desde o início do ano)
   ============================================================ */
function TmMasterView({ tmDue, tmHistory, tmDueSnapshots, setReportFn, setExportXlsxFn, handleImportTmDue, handleImportTmHistory, tmSubTab, setTmSubTab, updTmDue, onToggleAddToPlanning }) {
  const dueFileRef = useRef(null);
  const historyFileRef = useRef(null);

  const isVencida = (r) => r.diffValue !== null && r.diffValue < 0;
  const isAteVencer40 = (r) => r.diffUnit === "D" && r.diffValue !== null && r.diffValue >= 0 && r.diffValue <= 40;
  const isVencidaOuAte40 = (r) => isVencida(r) || isAteVencer40(r);
  const isCritica = (r) => r.pri === "High";
  const isCorretiva = (r) => (r.jobType || "").toUpperCase() === "ONE";
  const isPostergada = (h) => (h.remarks || "").toLowerCase().includes("postpon");

  /* mapa Código → Departamento, construído a partir do Due, usado pra descobrir o departamento
     de itens do History (que não tem essa coluna) cruzando pelo código do componente */
  const codeToDept = useMemo(() => {
    const map = new Map();
    tmDue.forEach((d) => { if (d.code && d.department) map.set(d.code, d.department); });
    return map;
  }, [tmDue]);
  const deptOfHistory = (h) => codeToDept.get(h.componentCode) || "Não identificado";

  /* ---------- filtros: Due ---------- */
  const [duef, setDuef] = useState({ busca: "", department: "Todos", pri: "Todos", jobType: "Todos", situacao: "Todos" });
  /* usado pelos gráficos da subaba Métricas: clicar numa barra já filtra e leva direto pra tabela
     correspondente, em vez de só mostrar o número no gráfico */
  const goToDueFiltered = (patch) => { setDuef((p) => ({ ...p, ...patch })); setTmSubTab("due"); };
  const goToHistoryFiltered = (patch) => { setHistf((p) => ({ ...p, ...patch })); setTmSubTab("history"); };
  const dueDepartments = useMemo(() => [...new Set(tmDue.map((d) => d.department).filter(Boolean))].sort(), [tmDue]);
  const dueJobTypes = useMemo(() => [...new Set(tmDue.map((d) => d.jobType).filter(Boolean))].sort(), [tmDue]);
  const hasActiveFilterDue = duef.busca || duef.department !== "Todos" || duef.pri !== "Todos" || duef.jobType !== "Todos" || duef.situacao !== "Todos";
  const filteredDue = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return tmDue.filter((d) => {
      const inBusca = !duef.busca || norm(d.jobName).includes(norm(duef.busca)) || norm(d.component).includes(norm(duef.busca)) || norm(d.code).includes(norm(duef.busca));
      const inDep = duef.department === "Todos" || d.department === duef.department;
      const inPri = duef.pri === "Todos" || d.pri === duef.pri;
      const inJobType = duef.jobType === "Todos" || d.jobType === duef.jobType;
      const inSituacao = duef.situacao === "Todos"
        || (duef.situacao === "Vencida" && isVencida(d))
        || (duef.situacao === "Até 40 dias" && isAteVencer40(d))
        || (duef.situacao === "Normal" && !isVencidaOuAte40(d));
      return inBusca && inDep && inPri && inJobType && inSituacao;
    });
  }, [tmDue, duef]);

  /* ---------- filtros: History ---------- */
  const [histf, setHistf] = useState({ busca: "", jobType: "Todos", department: "Todos", usuario: "Todos", soPostergadas: false });
  const histJobTypes = useMemo(() => [...new Set(tmHistory.map((h) => h.jobType).filter(Boolean))].sort(), [tmHistory]);
  const histDepartments = useMemo(() => [...new Set(tmHistory.map((h) => deptOfHistory(h)).filter(Boolean))].sort(), [tmHistory, codeToDept]);
  const histUsuarios = useMemo(() => [...new Set(tmHistory.map((h) => h.doneByName).filter(Boolean))].sort(), [tmHistory]);
  const hasActiveFilterHist = histf.busca || histf.jobType !== "Todos" || histf.department !== "Todos" || histf.usuario !== "Todos" || histf.soPostergadas;
  const filteredHistory = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return tmHistory.filter((h) => {
      const inBusca = !histf.busca || norm(h.jobName).includes(norm(histf.busca)) || norm(h.componentName).includes(norm(histf.busca)) || norm(h.doneByName).includes(norm(histf.busca));
      const inJobType = histf.jobType === "Todos" || h.jobType === histf.jobType;
      const inDep = histf.department === "Todos" || deptOfHistory(h) === histf.department;
      const inUsuario = histf.usuario === "Todos" || h.doneByName === histf.usuario;
      const inPost = !histf.soPostergadas || isPostergada(h);
      return inBusca && inJobType && inDep && inUsuario && inPost;
    });
  }, [tmHistory, histf, codeToDept]);

  const [groupDue, setGroupDue] = useState(true);
  const [collapsedDue, setCollapsedDue] = useState(new Set());
  const [groupHist, setGroupHist] = useState(true);
  const [collapsedHist, setCollapsedHist] = useState(new Set());
  const dueGroups = useMemo(() => {
    const vis = filteredDue.slice(0, 500);
    const urg = (d) => (d.diffValue === null || d.diffValue === undefined ? 1e9 : d.diffValue);
    if (!groupDue) return [{ key: "Todos", rows: [...vis].sort((x, y) => urg(x) - urg(y)) }];
    const map = new Map();
    vis.forEach((d) => { const k = (d.department || "").trim() || "Sem departamento"; if (!map.has(k)) map.set(k, []); map.get(k).push(d); });
    return [...map.entries()].sort((x, y) => x[0].localeCompare(y[0], "pt-BR")).map(([key, rows]) => ({ key, rows: rows.sort((p, q) => urg(p) - urg(q)) }));
  }, [filteredDue, groupDue]);
  const histGroups = useMemo(() => {
    const vis = filteredHistory.slice(0, 500);
    if (!groupHist) return [{ key: "todos", label: "Todos", rows: vis }];
    const map = new Map();
    vis.forEach((h) => { const k = h.dateDone ? h.dateDone.slice(0, 7) : "sem-data"; if (!map.has(k)) map.set(k, []); map.get(k).push(h); });
    return [...map.entries()].sort((x, y) => (x[0] === "sem-data") - (y[0] === "sem-data") || y[0].localeCompare(x[0])).map(([key, rows]) => ({
      key, rows: rows.sort((p, q) => (q.dateDone || "").localeCompare(p.dateDone || "")),
      label: key === "sem-data" ? "Sem data de fechamento" : `${MONTH_NAMES[Number(key.slice(5, 7)) - 1]} / ${key.slice(0, 4)}`,
    }));
  }, [filteredHistory, groupHist]);

  /* ---------- filtro de Mês/Ano específico da subaba Métricas — escopa todas as métricas abaixo
     (Due pela data de vencimento, History pela data de fechamento) ---------- */
  const [metricasFiltro, setMetricasFiltro] = useState({ mes: "Todos", ano: "Todos" });
  const anosDisponiveis = useMemo(() => {
    const anos = new Set();
    tmDue.forEach((d) => { if (d.dueDate) anos.add(d.dueDate.slice(0, 4)); });
    tmHistory.forEach((h) => { if (h.dateDone) anos.add(h.dateDone.slice(0, 4)); });
    return [...anos].sort();
  }, [tmDue, tmHistory]);
  const hasActiveFiltroMetricas = metricasFiltro.mes !== "Todos" || metricasFiltro.ano !== "Todos";
  const tmDueForMetrics = useMemo(() => {
    if (!hasActiveFiltroMetricas) return tmDue;
    return tmDue.filter((d) => {
      if (!d.dueDate) return false; // sem data de vencimento não dá pra encaixar num mês/ano específico
      const [y, m] = d.dueDate.split("-");
      const inMes = metricasFiltro.mes === "Todos" || Number(m) === Number(metricasFiltro.mes);
      const inAno = metricasFiltro.ano === "Todos" || y === metricasFiltro.ano;
      return inMes && inAno;
    });
  }, [tmDue, metricasFiltro, hasActiveFiltroMetricas]);
  const tmHistoryForMetrics = useMemo(() => {
    if (!hasActiveFiltroMetricas) return tmHistory;
    return tmHistory.filter((h) => {
      if (!h.dateDone) return false;
      const [y, m] = h.dateDone.split("-");
      const inMes = metricasFiltro.mes === "Todos" || Number(m) === Number(metricasFiltro.mes);
      const inAno = metricasFiltro.ano === "Todos" || y === metricasFiltro.ano;
      return inMes && inAno;
    });
  }, [tmHistory, metricasFiltro, hasActiveFiltroMetricas]);

  /* ---------- métricas: Due ---------- */
  const totalDue = tmDueForMetrics.length;
  const vencidas = tmDueForMetrics.filter(isVencida);
  const ateVencer40 = tmDueForMetrics.filter(isAteVencer40);
  const criticasVencidasOu40 = tmDueForMetrics.filter((d) => isCritica(d) && isVencidaOuAte40(d));
  const corretivasDue = tmDueForMetrics.filter(isCorretiva);

  const duePorMes = useMemo(() => {
    const map = {};
    tmDueForMetrics.forEach((d) => {
      if (!d.dueDate) return;
      const key = d.dueDate.slice(0, 7);
      map[key] = (map[key] || 0) + 1;
    });
    const totalComData = Object.values(map).reduce((s, v) => s + v, 0) || 1;
    return Object.keys(map).sort().map((k) => {
      const [y, m] = k.split("-");
      const count = map[k];
      return { mes: `${MONTH_NAMES[Number(m) - 1].slice(0, 3)}/${y.slice(2)}`, mesKey: k, count, pct: Math.round((count / totalComData) * 100) };
    });
  }, [tmDueForMetrics]);

  const urgenciaBuckets = useMemo(() => {
    const buckets = { "Vencida": 0, "0-10 dias": 0, "11-20 dias": 0, "21-30 dias": 0, "31-40 dias": 0 };
    tmDueForMetrics.forEach((d) => {
      if (isVencida(d)) buckets["Vencida"]++;
      else if (d.diffUnit === "D" && d.diffValue !== null) {
        if (d.diffValue <= 10) buckets["0-10 dias"]++;
        else if (d.diffValue <= 20) buckets["11-20 dias"]++;
        else if (d.diffValue <= 30) buckets["21-30 dias"]++;
        else if (d.diffValue <= 40) buckets["31-40 dias"]++;
      }
    });
    const total = Object.values(buckets).reduce((s, v) => s + v, 0) || 1;
    return Object.entries(buckets).map(([bucket, count]) => ({ bucket, count, pct: Math.round((count / total) * 100) }));
  }, [tmDueForMetrics]);

  const duePorDepartamento = useMemo(() => {
    const map = {};
    tmDueForMetrics.forEach((d) => { map[d.department] = (map[d.department] || 0) + 1; });
    return Object.entries(map).map(([departamento, count]) => ({ departamento, count })).sort((a, b) => b.count - a.count);
  }, [tmDueForMetrics]);

  const duePorComponente = useMemo(() => {
    const map = {};
    tmDueForMetrics.forEach((d) => { const k = d.component || "—"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([componente, count]) => ({ componente, count })).sort((a, b) => b.count - a.count).slice(0, 15);
  }, [tmDueForMetrics]);

  const vencidasPorJobType = useMemo(() => {
    const map = {};
    vencidas.forEach((d) => { const k = d.jobType || "—"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([tipo, count]) => ({ tipo, count })).sort((a, b) => b.count - a.count).slice(0, 12);
  }, [tmDueForMetrics]);

  const aVencerPorJobType = useMemo(() => {
    const map = {};
    ateVencer40.forEach((d) => { const k = d.jobType || "—"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([tipo, count]) => ({ tipo, count })).sort((a, b) => b.count - a.count).slice(0, 12);
  }, [tmDueForMetrics]);

  /* quantidade de jobs VENCIDOS (não conta "a vencer"), agrupados por Componente — usado no gráfico
     "Jobs vencidos por Componente" */
  const vencidasPorComponente = useMemo(() => {
    const map = {};
    vencidas.forEach((d) => { const k = d.component || "—"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([componente, count]) => ({ componente, count })).sort((a, b) => b.count - a.count).slice(0, 15);
  }, [tmDueForMetrics]);

  /* comparativo mensal: quantidade de itens do Due que estão vencidos ou a vencer (agrupados pelo mês
     de vencimento) x quantidade de itens do History fechados naquele mês (agrupados pela data de fechamento) */
  const abertosVsFechadosPorMes = useMemo(() => {
    const map = {};
    tmDueForMetrics.forEach((d) => {
      if (!d.dueDate || !isVencidaOuAte40(d)) return;
      const key = d.dueDate.slice(0, 7);
      if (!map[key]) map[key] = { mesKey: key, abertos: 0, fechados: 0 };
      map[key].abertos++;
    });
    tmHistoryForMetrics.forEach((h) => {
      if (!h.dateDone) return;
      const key = h.dateDone.slice(0, 7);
      if (!map[key]) map[key] = { mesKey: key, abertos: 0, fechados: 0 };
      map[key].fechados++;
    });
    return Object.values(map).sort((a, b) => a.mesKey.localeCompare(b.mesKey)).map((r) => {
      const [y, m] = r.mesKey.split("-");
      return { ...r, mes: `${MONTH_NAMES[Number(m) - 1].slice(0, 3)}/${y.slice(2)}` };
    });
  }, [tmDueForMetrics, tmHistoryForMetrics]);

  /* mesma comparação mensal acima, mas restrita só às CORRETIVAS (Job type = ONE): itens do Due
     vencidos/a vencer agrupados pelo mês de vencimento x itens do History fechados naquele mês */
  const corretivasPorMes = useMemo(() => {
    const map = {};
    tmDueForMetrics.forEach((d) => {
      if (!isCorretiva(d) || !d.dueDate || !isVencidaOuAte40(d)) return;
      const key = d.dueDate.slice(0, 7);
      if (!map[key]) map[key] = { mesKey: key, vencidasOuAVencer: 0, fechadas: 0 };
      map[key].vencidasOuAVencer++;
    });
    tmHistoryForMetrics.forEach((h) => {
      if (!isCorretiva(h) || !h.dateDone) return;
      const key = h.dateDone.slice(0, 7);
      if (!map[key]) map[key] = { mesKey: key, vencidasOuAVencer: 0, fechadas: 0 };
      map[key].fechadas++;
    });
    return Object.values(map).sort((a, b) => a.mesKey.localeCompare(b.mesKey)).map((r) => {
      const [y, m] = r.mesKey.split("-");
      return { ...r, mes: `${MONTH_NAMES[Number(m) - 1].slice(0, 3)}/${y.slice(2)}` };
    });
  }, [tmDueForMetrics, tmHistoryForMetrics]);

  /* ---------- métricas: History ---------- */
  const totalHistory = tmHistoryForMetrics.length;
  const corretivasFechadas = tmHistoryForMetrics.filter(isCorretiva);
  const postergadas = tmHistoryForMetrics.filter(isPostergada);

  const historyPorMes = useMemo(() => {
    const map = {};
    tmHistoryForMetrics.forEach((h) => {
      if (!h.dateDone) return;
      const key = h.dateDone.slice(0, 7);
      map[key] = (map[key] || 0) + 1;
    });
    const totalComData = Object.values(map).reduce((s, v) => s + v, 0) || 1;
    return Object.keys(map).sort().map((k) => {
      const [y, m] = k.split("-");
      const count = map[k];
      return { mes: `${MONTH_NAMES[Number(m) - 1].slice(0, 3)}/${y.slice(2)}`, mesKey: k, count, pct: Math.round((count / totalComData) * 100) };
    });
  }, [tmHistoryForMetrics]);

  const historyPorDepartamento = useMemo(() => {
    const map = {};
    tmHistoryForMetrics.forEach((h) => { const k = deptOfHistory(h); map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([departamento, count]) => ({ departamento, count })).sort((a, b) => b.count - a.count);
  }, [tmHistoryForMetrics, codeToDept]);

  const historyPorComponente = useMemo(() => {
    const map = {};
    tmHistoryForMetrics.forEach((h) => { const k = h.componentName || "—"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([componente, count]) => ({ componente, count })).sort((a, b) => b.count - a.count).slice(0, 15);
  }, [tmHistoryForMetrics]);

  const historyPorUsuario = useMemo(() => {
    const map = {};
    tmHistoryForMetrics.forEach((h) => { const k = h.doneByName || "Não informado"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([usuario, count]) => ({ usuario, count })).sort((a, b) => b.count - a.count).slice(0, 15);
  }, [tmHistoryForMetrics]);

  /* métrica extra: prazo médio de execução — compara Due date × Date signed. Positivo = feito
     antes do prazo; negativo = feito depois do prazo (atraso real na execução) */
  const prazoStats = useMemo(() => {
    const diffs = tmHistoryForMetrics
      .filter((h) => h.dueDate && h.dateSigned)
      .map((h) => Math.round((new Date(h.dueDate) - new Date(h.dateSigned)) / 86400000));
    if (diffs.length === 0) return { media: 0, noPrazo: 0, atrasado: 0, total: 0 };
    const media = Math.round(diffs.reduce((s, v) => s + v, 0) / diffs.length);
    const noPrazo = diffs.filter((v) => v >= 0).length;
    const atrasado = diffs.filter((v) => v < 0).length;
    return { media, noPrazo, atrasado, total: diffs.length };
  }, [tmHistoryForMetrics]);

  React.useEffect(() => {
    if (!setReportFn) return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      if (tmSubTab === "due" || tmSubTab === "metricas") {
        let y = pdfHeader(doc, "TM Master — Due (em aberto)",
          `${filteredDue.length} item(ns) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
        y = pdfKpis(doc, y, [
          { label: "Total em Aberto", value: totalDue },
          { label: "Vencidas", value: vencidas.length },
          { label: "A Vencer em 40 dias", value: ateVencer40.length },
          { label: "Críticas Vencidas/≤40d", value: criticasVencidasOu40.length },
          { label: "Corretivas (ONE)", value: corretivasDue.length },
        ]);
        y = pdfSectionTitle(doc, y, "Por departamento");
        y = pdfTable(doc, y, ["Departamento", "Quantidade"], duePorDepartamento.map((d) => [d.departamento, d.count]));
        y = pdfSectionTitle(doc, y, "Vencidas por tipo de serviço");
        y = pdfTable(doc, y, ["Tipo", "Quantidade"], vencidasPorJobType.map((d) => [d.tipo, d.count]));
        y = pdfSectionTitle(doc, y, "A vencer (até 40 dias) por tipo de serviço");
        y = pdfTable(doc, y, ["Tipo", "Quantidade"], aVencerPorJobType.map((d) => [d.tipo, d.count]));
        pdfSave(doc, "relatorio-tm-master-due");
      }
      if (tmSubTab === "history") {
        let y = pdfHeader(doc, "TM Master — History (fechados)",
          `${filteredHistory.length} item(ns) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
        y = pdfKpis(doc, y, [
          { label: "Total Fechado (ano)", value: totalHistory },
          { label: "Corretivas Fechadas", value: corretivasFechadas.length },
          { label: "Postergadas", value: postergadas.length },
          { label: "Prazo Médio (dias)", value: `${prazoStats.media}d` },
        ]);
        y = pdfSectionTitle(doc, y, "Por departamento");
        y = pdfTable(doc, y, ["Departamento", "Quantidade"], historyPorDepartamento.map((d) => [d.departamento, d.count]));
        y = pdfSectionTitle(doc, y, "Fechados por usuário");
        y = pdfTable(doc, y, ["Usuário", "Quantidade"], historyPorUsuario.map((d) => [d.usuario, d.count]));
        pdfSave(doc, "relatorio-tm-master-history");
      }
    });
  }, [tmSubTab, filteredDue, totalDue, vencidas, ateVencer40, criticasVencidasOu40, corretivasDue, duePorDepartamento,
      vencidasPorJobType, aVencerPorJobType, filteredHistory, totalHistory, corretivasFechadas, postergadas, prazoStats,
      historyPorDepartamento, historyPorUsuario, setReportFn]);

  /* "Exportar planilha" no cabeçalho segue a subaba/filtro ativo aqui — Due exporta o filtro de Due,
     History o filtro de History, e Métricas exporta os dois conjuntos já escopados por mês/ano */
  React.useEffect(() => {
    if (!setExportXlsxFn) return;
    setExportXlsxFn(() => () => {
      const wb = XLSX.utils.book_new();
      if (tmSubTab === "due") {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(filteredDue, TM_DUE_COLS)), "TM Master - Due");
        XLSX.writeFile(wb, `genesis-tm-master-due-${todayISO()}.xlsx`);
      } else if (tmSubTab === "history") {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(filteredHistory, TM_HISTORY_COLS)), "TM Master - History");
        XLSX.writeFile(wb, `genesis-tm-master-history-${todayISO()}.xlsx`);
      } else {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(tmDueForMetrics, TM_DUE_COLS)), "Due (filtrado)");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(tmHistoryForMetrics, TM_HISTORY_COLS)), "History (filtrado)");
        XLSX.writeFile(wb, `genesis-tm-master-metricas-${todayISO()}.xlsx`);
      }
    });
  }, [tmSubTab, filteredDue, filteredHistory, tmDueForMetrics, tmHistoryForMetrics, setExportXlsxFn]);

  return (
    <>
      <div className="g-mode-toggle" style={{ marginBottom: 16, width: "fit-content" }}>
        <button className={tmSubTab === "due" ? "active" : ""} onClick={() => setTmSubTab("due")}>Due (Em Aberto)</button>
        <button className={tmSubTab === "history" ? "active" : ""} onClick={() => setTmSubTab("history")}>History (Fechados)</button>
        <button className={tmSubTab === "metricas" ? "active" : ""} onClick={() => setTmSubTab("metricas")}>Métricas</button>
      </div>

      {tmSubTab === "due" && (
        <>
          <div className="g-panel" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
            <div>
              <div className="g-panel-title" style={{ marginBottom: 4 }}>Importar TM Master — Due</div>
              <div className="g-muted" style={{ fontSize: 11.5 }}>
                Suba a planilha semanalmente. Cada importação <strong>substitui por completo</strong> a lista de itens em aberto
                (é uma fotografia do momento — o que já foi resolvido não aparece mais na nova planilha).
              </div>
            </div>
            <div>
              <input ref={dueFileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImportTmDue} />
              <button className="g-btn primary" onClick={() => dueFileRef.current?.click()}><Upload size={14} />Importar planilha Due</button>
            </div>
          </div>

          <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
            <div className="g-field">
              <label>Buscar</label>
              <input type="text" value={duef.busca} onChange={(e) => setDuef((p) => ({ ...p, busca: e.target.value }))} placeholder="job, componente, código..." style={{ minWidth: 180 }} />
            </div>
            <div className="g-field">
              <label>Departamento</label>
              <select value={duef.department} onChange={(e) => setDuef((p) => ({ ...p, department: e.target.value }))}>
                <option>Todos</option>
                {dueDepartments.map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
            </div>
            <div className="g-field">
              <label>Prioridade</label>
              <select value={duef.pri} onChange={(e) => setDuef((p) => ({ ...p, pri: e.target.value }))}>
                <option>Todos</option><option>High</option><option>Medium</option><option>Low</option><option>Req</option>
              </select>
            </div>
            <div className="g-field">
              <label>Job Type</label>
              <select value={duef.jobType} onChange={(e) => setDuef((p) => ({ ...p, jobType: e.target.value }))}>
                <option>Todos</option>
                {dueJobTypes.map((j) => <option key={j} value={j}>{j}</option>)}
              </select>
            </div>
            <div className="g-field">
              <label>Situação</label>
              <select value={duef.situacao} onChange={(e) => setDuef((p) => ({ ...p, situacao: e.target.value }))}>
                <option>Todos</option><option>Vencida</option><option>Até 40 dias</option><option>Normal</option>
              </select>
            </div>
            <div className="g-field">
              <label>&nbsp;</label>
              <button className="g-btn" onClick={() => setDuef({ busca: "", department: "Todos", pri: "Todos", jobType: "Todos", situacao: "Todos" })}
                disabled={!hasActiveFilterDue} style={{ opacity: hasActiveFilterDue ? 1 : 0.5 }}>
                <X size={13} />Limpar filtro
              </button>
            </div>
          </div>

          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
            {bigKpi("Total em Aberto", totalDue, "var(--teal)", ClipboardList)}
            {bigKpi("Vencidas", vencidas.length, "var(--crit)", AlertTriangle)}
            {bigKpi("A Vencer em 40 dias", ateVencer40.length, "var(--warn)", Clock)}
            {bigKpi("Críticas Vencidas/≤40d", criticasVencidasOu40.length, "var(--crit)", AlertTriangle)}
            {bigKpi("Corretivas (ONE)", corretivasDue.length, "var(--text-dim)", Wrench)}
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Itens em aberto ({filteredDue.length})</span></div>
            <div className="tbl-toolbar" style={{ margin: "-18px -18px 0 -18px", borderRadius: "12px 12px 0 0" }}>
              <label className="tbl-check"><input type="checkbox" checked={groupDue} onChange={(e) => setGroupDue(e.target.checked)} />Agrupar por departamento</label>
              <span className="g-muted" style={{ fontSize: 11 }}>faixa vermelha = vencida · laranja = vence em até 40 dias · verde = no prazo</span>
            </div>
            <div className="g-table-wrap">
            <table className="g-table tbl-modern">
              <thead>
                <tr>
                  <th style={{ width: 8, padding: 0 }}></th>
                  <th style={{ minWidth: 320 }}>Manutenção</th>
                  <th style={{ minWidth: 150 }}>Prazo</th>
                  <th style={{ minWidth: 100 }}>Prioridade</th>
                  <th style={{ minWidth: 90 }}>Status</th>
                  <th style={{ minWidth: 220 }}>Plano de Ação</th>
                  <th style={{ minWidth: 130 }}>Planejamento</th>
                </tr>
              </thead>
              <tbody>
                {dueGroups.map((g) => (
                  <React.Fragment key={g.key}>
                    {groupDue && (
                      <tr className="tbl-group" onClick={() => setCollapsedDue((p) => { const n = new Set(p); n.has(g.key) ? n.delete(g.key) : n.add(g.key); return n; })}>
                        <td colSpan={7}>
                          <span className="tbl-group-arrow">{collapsedDue.has(g.key) ? "▸" : "▾"}</span>
                          {g.key}
                          <span className="tbl-group-cnt">{g.rows.length} item(ns)</span>
                          <span className="tbl-group-tot">{g.rows.filter(isVencida).length} vencida(s) · {g.rows.filter(isAteVencer40).length} em até 40 dias</span>
                        </td>
                      </tr>
                    )}
                    {!(groupDue && collapsedDue.has(g.key)) && g.rows.map((d) => {
                  const i = tmDue.indexOf(d);
                  const cor = isVencida(d) ? "#EF4444" : isAteVencer40(d) ? "#F5A623" : "#22C55E";
                  const pc = isVencida(d) ? "d3" : isAteVencer40(d) ? "d2" : "d1";
                  const priCor = d.pri === "High" ? "#B91C1C" : d.pri === "Medium" ? "#B45309" : "#5B6273";
                  const priBg = d.pri === "High" ? "#FEE7E7" : d.pri === "Medium" ? "#FFF3DC" : "#EEF0F4";
                  return (
                  <tr className="g-row tbl-row" key={d.id} style={{ "--c": cor }}>
                    <td className="tbl-first" style={{ padding: 0 }}></td>
                    <td style={{ minWidth: 320, whiteSpace: "normal" }}>
                      <div style={{ fontWeight: 600, color: "#12203A" }}>{d.component}</div>
                      <div className="tbl-sub" style={{ padding: "2px 0 0 0" }}>{d.jobName}</div>
                      <div className="tbl-chips" style={{ padding: "4px 0 0 0" }}>
                        {d.jobType && <span className="tbl-chip">{d.jobType}</span>}
                        {d.jobNo && <span className="tbl-chip">Job {d.jobNo}</span>}
                        {d.code && <span className="tbl-chip">Cód. {d.code}</span>}
                      </div>
                    </td>
                    <td style={{ minWidth: 150 }}>
                      <span className={`tbl-days ${pc}`}>{d.diffValue === null || d.diffValue === undefined ? (d.diffRaw || "—") : isVencida(d) ? `vencida há ${Math.abs(d.diffValue)} ${d.diffUnit === "H" ? "h" : "dias"}` : `vence em ${d.diffValue} ${d.diffUnit === "H" ? "h" : "dias"}`}</span>
                      {d.dueRaw && <div className="tbl-sub" style={{ padding: "4px 0 0 2px" }}>Due {d.dueRaw}</div>}
                    </td>
                    <td><span className="tbl-days" style={{ background: priBg, color: priCor }}>{d.pri || "—"}</span></td>
                    <td>{d.status}</td>
                    <td style={{ minWidth: 220, whiteSpace: "normal", verticalAlign: "top" }}>
                      <ETextArea rows={1} value={d.planoAcao || ""} onChange={(v) => updTmDue(i, "planoAcao", v)} />
                    </td>
                    <td style={{ minWidth: 130 }}>
                      <label className="g-flex" style={{ gap: 6, fontSize: 11.5, cursor: "pointer" }}>
                        <input type="checkbox" checked={!!d.linkedPlanId} onChange={(e) => onToggleAddToPlanning(d, e.target.checked)} />
                        {d.linkedPlanId ? "Adicionado" : "Adicionar"}
                      </label>
                    </td>
                  </tr>
                  );
                })}
                  </React.Fragment>
                ))}
              </tbody>
            </table>
            </div>
            {filteredDue.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum item encontrado — importe a planilha Due acima.</div>}
            {filteredDue.length > 500 && <div className="g-muted" style={{ marginTop: 10 }}>Mostrando os primeiros 500 de {filteredDue.length} — refine o filtro pra ver outros.</div>}
          </div>
        </>
      )}

      {tmSubTab === "history" && (
        <>
          <div className="g-panel" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
            <div>
              <div className="g-panel-title" style={{ marginBottom: 4 }}>Importar TM Master — History</div>
              <div className="g-muted" style={{ fontSize: 11.5 }}>
                Suba a planilha semanalmente. É uma lista cumulativa desde o início do ano — a importação
                <strong> mescla por número do histórico</strong>, então nada é duplicado.
              </div>
            </div>
            <div>
              <input ref={historyFileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImportTmHistory} />
              <button className="g-btn primary" onClick={() => historyFileRef.current?.click()}><Upload size={14} />Importar planilha History</button>
            </div>
          </div>

          <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
            <div className="g-field">
              <label>Buscar</label>
              <input type="text" value={histf.busca} onChange={(e) => setHistf((p) => ({ ...p, busca: e.target.value }))} placeholder="job, componente, responsável..." style={{ minWidth: 200 }} />
            </div>
            <div className="g-field">
              <label>Job Type</label>
              <select value={histf.jobType} onChange={(e) => setHistf((p) => ({ ...p, jobType: e.target.value }))}>
                <option>Todos</option>
                {histJobTypes.map((j) => <option key={j} value={j}>{j}</option>)}
              </select>
            </div>
            <div className="g-field">
              <label>Departamento</label>
              <select value={histf.department} onChange={(e) => setHistf((p) => ({ ...p, department: e.target.value }))}>
                <option>Todos</option>
                {histDepartments.map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
            </div>
            <div className="g-field">
              <label>Usuário</label>
              <select value={histf.usuario} onChange={(e) => setHistf((p) => ({ ...p, usuario: e.target.value }))}>
                <option>Todos</option>
                {histUsuarios.map((u) => <option key={u} value={u}>{u}</option>)}
              </select>
            </div>
            <div className="g-field">
              <label>&nbsp;</label>
              <label className="g-flex" style={{ gap: 6, fontSize: 12, cursor: "pointer", padding: "9px 0" }}>
                <input type="checkbox" checked={histf.soPostergadas} onChange={(e) => setHistf((p) => ({ ...p, soPostergadas: e.target.checked }))} />
                Só postergadas
              </label>
            </div>
            <div className="g-field">
              <label>&nbsp;</label>
              <button className="g-btn" onClick={() => setHistf({ busca: "", jobType: "Todos", department: "Todos", usuario: "Todos", soPostergadas: false })}
                disabled={!hasActiveFilterHist} style={{ opacity: hasActiveFilterHist ? 1 : 0.5 }}>
                <X size={13} />Limpar filtro
              </button>
            </div>
          </div>

          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
            {bigKpi("Total Fechado (ano)", totalHistory, "var(--ok)", ClipboardList)}
            {bigKpi("Corretivas Fechadas", corretivasFechadas.length, "var(--text-dim)", Wrench)}
            {bigKpi("Postergadas", postergadas.length, "var(--warn)", AlertTriangle)}
            {bigKpi("Prazo Médio (dias)", `${prazoStats.media}d`, prazoStats.media < 0 ? "var(--crit)" : "var(--ok)", Clock)}
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Jobs fechados ({filteredHistory.length})</span></div>
            <div className="tbl-toolbar" style={{ margin: "-18px -18px 0 -18px", borderRadius: "12px 12px 0 0" }}>
              <label className="tbl-check"><input type="checkbox" checked={groupHist} onChange={(e) => setGroupHist(e.target.checked)} />Agrupar por mês de fechamento</label>
              <span className="g-muted" style={{ fontSize: 11 }}>faixa laranja = postergada</span>
            </div>
            <div className="g-table-wrap">
            <table className="g-table tbl-modern">
              <thead>
                <tr>
                  <th style={{ width: 8, padding: 0 }}></th>
                  <th style={{ minWidth: 340 }}>Manutenção Realizada</th>
                  <th style={{ minWidth: 130 }}>Responsável</th>
                  <th style={{ minWidth: 110 }}>Fechado Em</th>
                  <th style={{ minWidth: 120 }}>Situação</th>
                  <th style={{ minWidth: 200 }}>Observações</th>
                </tr>
              </thead>
              <tbody>
                {histGroups.map((g) => (
                  <React.Fragment key={g.key}>
                    {groupHist && (
                      <tr className="tbl-group" onClick={() => setCollapsedHist((p) => { const n = new Set(p); n.has(g.key) ? n.delete(g.key) : n.add(g.key); return n; })}>
                        <td colSpan={6}>
                          <span className="tbl-group-arrow">{collapsedHist.has(g.key) ? "▸" : "▾"}</span>
                          {g.label}
                          <span className="tbl-group-cnt">{g.rows.length} fechamento(s)</span>
                          <span className="tbl-group-tot">{g.rows.filter(isPostergada).length} postergada(s)</span>
                        </td>
                      </tr>
                    )}
                    {!(groupHist && collapsedHist.has(g.key)) && g.rows.map((h) => (
                  <tr className="g-row tbl-row" key={h.jobHistoryNumber} style={{ "--c": isPostergada(h) ? "#F5A623" : "#22C55E" }}>
                    <td className="tbl-first" style={{ padding: 0 }}></td>
                    <td style={{ minWidth: 340, whiteSpace: "normal" }}>
                      <div style={{ fontWeight: 600, color: "#12203A" }}>{h.jobName}</div>
                      <div className="tbl-sub" style={{ padding: "2px 0 0 0" }}>{h.componentName}</div>
                      <div className="tbl-chips" style={{ padding: "4px 0 0 0" }}>
                        {h.jobType && <span className="tbl-chip">{h.jobType}</span>}
                        {h.jobNo && <span className="tbl-chip">Job {h.jobNo}</span>}
                        {deptOfHistory(h) && <span className="tbl-chip">{deptOfHistory(h)}</span>}
                      </div>
                    </td>
                    <td>{h.doneByName || "—"}</td>
                    <td>{fmtDate(h.dateDone)}</td>
                    <td>{isPostergada(h) ? <span className="tbl-days d2">Postergada</span> : <span className="tbl-days d1">Fechada</span>}</td>
                    <td style={{ minWidth: 200, whiteSpace: "normal" }}>{h.remarks || "—"}</td>
                  </tr>
                    ))}
                  </React.Fragment>
                ))}
              </tbody>
            </table>
            </div>
            {filteredHistory.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum item encontrado — importe a planilha History acima.</div>}
            {filteredHistory.length > 500 && <div className="g-muted" style={{ marginTop: 10 }}>Mostrando os primeiros 500 de {filteredHistory.length} — refine o filtro pra ver outros.</div>}
          </div>
        </>
      )}

      {tmSubTab === "metricas" && (
        <>
          <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
            <div className="g-field">
              <label>Mês</label>
              <select value={metricasFiltro.mes} onChange={(e) => setMetricasFiltro((p) => ({ ...p, mes: e.target.value }))}>
                <option value="Todos">Todos</option>
                {MONTH_NAMES.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
              </select>
            </div>
            <div className="g-field">
              <label>Ano</label>
              <select value={metricasFiltro.ano} onChange={(e) => setMetricasFiltro((p) => ({ ...p, ano: e.target.value }))}>
                <option value="Todos">Todos</option>
                {anosDisponiveis.map((a) => <option key={a} value={a}>{a}</option>)}
              </select>
            </div>
            <div className="g-field">
              <label>&nbsp;</label>
              <button className="g-btn" onClick={() => setMetricasFiltro({ mes: "Todos", ano: "Todos" })}
                disabled={!hasActiveFiltroMetricas} style={{ opacity: hasActiveFiltroMetricas ? 1 : 0.5 }}>
                <X size={13} />Ver todos os meses
              </button>
            </div>
            <div className="g-filter-spacer" />
            <div className="g-filter-summary">
              {hasActiveFiltroMetricas
                ? "Escopo: Due pela data de vencimento · History pela data de fechamento. Itens sem data de calendário (baseados em horas de máquina) ficam de fora quando um mês/ano é selecionado."
                : "Mostrando todos os períodos. Selecione um mês e/ou ano pra focar as métricas nele."}
            </div>
          </div>

          {tmDueSnapshots.length > 1 && (
            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Tendência do backlog — evolução semana a semana</span></div>
              <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 8 }}>
                Uma fotografia é registrada automaticamente a cada vez que você importa a planilha Due — isso mostra se o
                backlog de manutenção está crescendo ou diminuindo ao longo das semanas.
              </div>
              <div style={{ width: "100%", height: 220 }}>
                <ResponsiveContainer>
                  <BarChart data={tmDueSnapshots.map((s) => ({ ...s, dataLabel: fmtDate(s.date) }))} margin={{ left: 0, right: 8, top: 4, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                    <XAxis dataKey="dataLabel" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                    <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} />
                    <Bar dataKey="total" name="Total em Aberto" radius={[3, 3, 0, 0]} fill="var(--text-faint)" />
                    <Bar dataKey="vencidas" name="Vencidas" radius={[3, 3, 0, 0]} fill="var(--crit)" />
                    <Bar dataKey="ateVencer40" name="A Vencer ≤40d" radius={[3, 3, 0, 0]} fill="var(--warn)" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}

          {tmDueSnapshots.length > 1 && (
            <div className="g-panel">
              <div className="g-panel-head">
                <span className="g-panel-title">Backlog total de jobs vencidos</span>
                <span style={{ fontFamily: "var(--mono)", fontSize: 20, fontWeight: 700, color: "var(--crit)" }}>{vencidas.length}</span>
              </div>
              <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 8 }}>
                Somente os itens já vencidos (Diff negativo), a cada fotografia da planilha Due — mostra se o
                backlog vencido está aumentando ou sendo reduzido ao longo do tempo.
              </div>
              <div style={{ width: "100%", height: 200 }}>
                <ResponsiveContainer>
                  <BarChart data={tmDueSnapshots.map((s) => ({ ...s, dataLabel: fmtDate(s.date) }))} margin={{ left: 0, right: 8, top: 4, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                    <XAxis dataKey="dataLabel" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                    <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} />
                    <Bar dataKey="vencidas" name="Backlog Vencido" radius={[3, 3, 0, 0]} fill="var(--crit)">
                      <LabelList dataKey="vencidas" position="top" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Abertos (vencidos + a vencer) x Fechados por mês</span></div>
            <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 8 }}>
              Barra cinza: itens do Due vencidos ou a vencer, agrupados pelo mês de vencimento. Barra verde: itens do
              History fechados naquele mês, agrupados pela data de fechamento.
            </div>
            <div style={{ width: "100%", height: 240 }}>
              <ResponsiveContainer>
                <BarChart data={abertosVsFechadosPorMes} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                  <XAxis dataKey="mes" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                  <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                  <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} />
                  <Bar dataKey="abertos" name="Abertos (vencidos + a vencer)" radius={[3, 3, 0, 0]} fill="var(--text-faint)">
                    <LabelList dataKey="abertos" position="top" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                  </Bar>
                  <Bar dataKey="fechados" name="Fechados" radius={[3, 3, 0, 0]} fill="var(--ok)">
                    <LabelList dataKey="fechados" position="top" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            {abertosVsFechadosPorMes.length === 0 && <div className="g-muted" style={{ fontSize: 11.5 }}>Sem dados suficientes com data de vencimento/fechamento no filtro atual.</div>}
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Corretivas — vencidas/a vencer x fechadas por mês (Job type = ONE)</span></div>
            <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 8 }}>
              Mesma comparação acima, restrita só às manutenções corretivas (Job type ONE). Barra vermelha: corretivas
              vencidas ou a vencer, agrupadas pelo mês de vencimento. Barra verde: corretivas fechadas naquele mês.
            </div>
            <div style={{ width: "100%", height: 240 }}>
              <ResponsiveContainer>
                <BarChart data={corretivasPorMes} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                  <XAxis dataKey="mes" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                  <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                  <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} />
                  <Bar dataKey="vencidasOuAVencer" name="Vencidas / A Vencer" radius={[3, 3, 0, 0]} fill="var(--crit)">
                    <LabelList dataKey="vencidasOuAVencer" position="top" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                  </Bar>
                  <Bar dataKey="fechadas" name="Fechadas" radius={[3, 3, 0, 0]} fill="var(--accent)">
                    <LabelList dataKey="fechadas" position="top" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            {corretivasPorMes.length === 0 && <div className="g-muted" style={{ fontSize: 11.5 }}>Nenhuma corretiva (Job type ONE) com data de vencimento/fechamento no filtro atual.</div>}
          </div>

          <div className="g-section-label">Due — o que está em aberto</div>
          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
            {bigKpi("Total em Aberto", totalDue, "var(--teal)", ClipboardList)}
            {bigKpi("Vencidas", vencidas.length, "var(--crit)", AlertTriangle)}
            {bigKpi("A Vencer em 40 dias", ateVencer40.length, "var(--warn)", Clock)}
            {bigKpi("Críticas Vencidas/≤40d", criticasVencidasOu40.length, "var(--crit)", AlertTriangle)}
            {bigKpi("Corretivas (ONE)", corretivasDue.length, "var(--text-dim)", Wrench)}
          </div>

          <div className="g-grid-2">
            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Aderência por mês — itens a vencer</span></div>
              <div style={{ width: "100%", height: 240 }}>
                <ResponsiveContainer>
                  <BarChart data={duePorMes} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                    <XAxis dataKey="mes" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                    <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} formatter={(v, n, p) => [`${v} (${p.payload.pct}%)`, "Itens"]} />
                    <Bar dataKey="count" name="Itens" radius={[3, 3, 0, 0]} fill="var(--accent)">
                      <LabelList dataKey="count" position="top" formatter={(v, entry) => `${v}`} style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="g-muted" style={{ fontSize: 10.5, marginTop: 4 }}>% indica a fatia de cada mês sobre o total de itens com data de vencimento definida.</div>
            </div>

            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Urgência — vencidas ou a vencer em 40 dias (%)</span></div>
              <div style={{ width: "100%", height: 220 }}>
                <ResponsiveContainer>
                  <BarChart data={urgenciaBuckets} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                    <XAxis dataKey="bucket" tick={{ fill: "var(--text-faint)", fontSize: 9 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} interval={0} angle={-15} textAnchor="end" height={45} />
                    <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} formatter={(v, n, p) => [`${v} (${p.payload.pct}%)`, "Itens"]} />
                    <Bar dataKey="count" name="Itens" radius={[3, 3, 0, 0]} cursor="pointer"
                      onClick={(data) => goToDueFiltered({ situacao: data.bucket === "Vencida" ? "Vencida" : "Até 40 dias" })}>
                      <LabelList dataKey="pct" position="top" formatter={(v) => `${v}%`} style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                      <Cell fill="var(--crit)" /><Cell fill="var(--warn)" /><Cell fill="var(--warn)" /><Cell fill="var(--teal)" /><Cell fill="var(--teal)" />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="g-muted" style={{ fontSize: 10.5, marginTop: 4 }}>Clique numa barra pra ver os itens correspondentes na tabela Due.</div>
            </div>
          </div>

          <div className="g-grid-2">
            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Por departamento</span></div>
              <div style={{ width: "100%", height: 220 }}>
                <ResponsiveContainer>
                  <BarChart data={duePorDepartamento} layout="vertical" margin={{ left: 0, right: 24, top: 4, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" horizontal={false} />
                    <XAxis type="number" allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                    <YAxis type="category" dataKey="departamento" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={90} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} cursor={{ fill: "rgba(10,14,20,0.04)" }} />
                    <Bar dataKey="count" name="Itens" radius={[0, 3, 3, 0]} fill="var(--teal)" cursor="pointer" onClick={(data) => goToDueFiltered({ department: data.departamento })}>
                      <LabelList dataKey="count" position="right" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="g-muted" style={{ fontSize: 10.5, marginTop: 4 }}>Clique numa barra pra ver os itens desse departamento.</div>
            </div>

            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Top 15 equipamentos com mais itens em aberto</span></div>
              <div className="g-table-wrap" style={{ maxHeight: 240, overflowY: "auto" }}>
                <table className="g-table">
                  <thead><tr><th>Equipamento</th><th>Qtd</th></tr></thead>
                  <tbody>
                    {duePorComponente.map((d) => (
                      <tr key={d.componente}><td style={{ whiteSpace: "normal" }}>{d.componente}</td><td style={{ fontFamily: "var(--mono)" }}>{d.count}</td></tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          <div className="g-grid-2">
            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Vencidas por Job Type</span></div>
              <div style={{ width: "100%", height: 240 }}>
                <ResponsiveContainer>
                  <BarChart data={vencidasPorJobType} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                    <XAxis dataKey="tipo" tick={{ fill: "var(--text-faint)", fontSize: 9 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} interval={0} angle={-20} textAnchor="end" height={50} />
                    <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} />
                    <Bar dataKey="count" name="Vencidas" radius={[3, 3, 0, 0]} fill="var(--crit)" cursor="pointer"
                      onClick={(data) => goToDueFiltered({ jobType: data.tipo, situacao: "Vencida" })}>
                      <LabelList dataKey="count" position="top" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="g-muted" style={{ fontSize: 10.5, marginTop: 4 }}>Clique numa barra pra filtrar a tabela Due.</div>
            </div>

            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">A vencer (até 40 dias) por Job Type</span></div>
              <div style={{ width: "100%", height: 240 }}>
                <ResponsiveContainer>
                  <BarChart data={aVencerPorJobType} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                    <XAxis dataKey="tipo" tick={{ fill: "var(--text-faint)", fontSize: 9 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} interval={0} angle={-20} textAnchor="end" height={50} />
                    <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} />
                    <Bar dataKey="count" name="A vencer" radius={[3, 3, 0, 0]} fill="var(--warn)" cursor="pointer"
                      onClick={(data) => goToDueFiltered({ jobType: data.tipo, situacao: "Até 40 dias" })}>
                      <LabelList dataKey="count" position="top" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="g-muted" style={{ fontSize: 10.5, marginTop: 4 }}>Clique numa barra pra filtrar a tabela Due.</div>
            </div>
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Jobs vencidos por Componente</span></div>
            <div style={{ width: "100%", height: Math.max(240, vencidasPorComponente.length * 26) }}>
              <ResponsiveContainer>
                <BarChart data={vencidasPorComponente} layout="vertical" margin={{ left: 0, right: 24, top: 4, bottom: 4 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" horizontal={false} />
                  <XAxis type="number" allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                  <YAxis type="category" dataKey="componente" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={140} />
                  <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} cursor={{ fill: "rgba(10,14,20,0.04)" }} />
                  <Bar dataKey="count" name="Vencidas" radius={[0, 3, 3, 0]} fill="var(--crit)" cursor="pointer"
                    onClick={(data) => goToDueFiltered({ busca: data.componente, situacao: "Vencida" })}>
                    <LabelList dataKey="count" position="right" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            {vencidasPorComponente.length === 0 && <div className="g-muted" style={{ fontSize: 11.5 }}>Nenhum item vencido no filtro atual.</div>}
            <div className="g-muted" style={{ fontSize: 10.5, marginTop: 4 }}>Top 15 componentes com mais jobs vencidos · clique numa barra pra filtrar a tabela Due.</div>
          </div>

          <div className="g-section-label">History — o que já foi fechado</div>
          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
            {bigKpi("Total Fechado (ano)", totalHistory, "var(--ok)", ClipboardList)}
            {bigKpi("Corretivas Fechadas", corretivasFechadas.length, "var(--text-dim)", Wrench)}
            {bigKpi("Postergadas", postergadas.length, "var(--warn)", AlertTriangle)}
            {bigKpi("Prazo Médio de Execução", `${prazoStats.media}d`, prazoStats.media < 0 ? "var(--crit)" : "var(--ok)", Clock)}
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Aderência — jobs fechados por mês</span></div>
            <div style={{ width: "100%", height: 240 }}>
              <ResponsiveContainer>
                <BarChart data={historyPorMes} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                  <XAxis dataKey="mes" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                  <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                  <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} formatter={(v, n, p) => [`${v} (${p.payload.pct}%)`, "Fechados"]} />
                  <Bar dataKey="count" name="Fechados" radius={[3, 3, 0, 0]} fill="var(--ok)">
                    <LabelList dataKey="count" position="top" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="g-muted" style={{ fontSize: 10.5, marginTop: 4 }}>% indica a fatia de cada mês sobre o total fechado no ano.</div>
          </div>

          <div className="g-grid-2">
            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Fechados por departamento</span></div>
              <div style={{ width: "100%", height: 220 }}>
                <ResponsiveContainer>
                  <BarChart data={historyPorDepartamento} layout="vertical" margin={{ left: 0, right: 24, top: 4, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" horizontal={false} />
                    <XAxis type="number" allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                    <YAxis type="category" dataKey="departamento" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={100} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} cursor={{ fill: "rgba(10,14,20,0.04)" }} />
                    <Bar dataKey="count" name="Fechados" radius={[0, 3, 3, 0]} fill="var(--accent)" cursor="pointer" onClick={(data) => goToHistoryFiltered({ department: data.departamento })}>
                      <LabelList dataKey="count" position="right" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="g-muted" style={{ fontSize: 10.5, marginTop: 6 }}>Departamento identificado cruzando o código do componente com a planilha Due — itens sem correspondência aparecem como "Não identificado". Clique numa barra pra filtrar a tabela.</div>
            </div>

            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Top 15 equipamentos fechados</span></div>
              <div className="g-table-wrap" style={{ maxHeight: 240, overflowY: "auto" }}>
                <table className="g-table">
                  <thead><tr><th>Equipamento</th><th>Qtd</th></tr></thead>
                  <tbody>
                    {historyPorComponente.map((d) => (
                      <tr key={d.componente}><td style={{ whiteSpace: "normal" }}>{d.componente}</td><td style={{ fontFamily: "var(--mono)" }}>{d.count}</td></tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Top 15 usuários com mais jobs fechados</span></div>
            <div style={{ width: "100%", height: 240 }}>
              <ResponsiveContainer>
                <BarChart data={historyPorUsuario} layout="vertical" margin={{ left: 0, right: 24, top: 4, bottom: 4 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" horizontal={false} />
                  <XAxis type="number" allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                  <YAxis type="category" dataKey="usuario" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={120} />
                  <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} cursor={{ fill: "rgba(10,14,20,0.04)" }} />
                  <Bar dataKey="count" name="Jobs fechados" radius={[0, 3, 3, 0]} fill="var(--teal)" cursor="pointer" onClick={(data) => goToHistoryFiltered({ usuario: data.usuario })}>
                    <LabelList dataKey="count" position="right" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="g-muted" style={{ fontSize: 10.5, marginTop: 6 }}>Clique numa barra pra filtrar a tabela por esse usuário.</div>
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Cumprimento de prazo — Due date × Data assinada</span></div>
            <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 10 }}>
              Considera só os {prazoStats.total} jobs do History que têm Due date e Data assinada preenchidas.
              Prazo médio de {prazoStats.media} dia(s) {prazoStats.media >= 0 ? "de antecedência" : "de atraso"} em relação ao prazo.
            </div>
            <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(2, 1fr)" }}>
              {bigKpi("Feitos no Prazo ou Antes", prazoStats.noPrazo, "var(--ok)", ClipboardList)}
              {bigKpi("Feitos Após o Prazo", prazoStats.atrasado, "var(--crit)", AlertTriangle)}
            </div>
          </div>
        </>
      )}
    </>
  );
}

/* ============================================================
   MATERIALS — fully editable, including ID
   ============================================================ */
function MaterialsView({ materials, updMat, remMat, workPackages, setReportFn, handleImportEmergenciais, newRowId }) {
  React.useEffect(() => {
    if (!newRowId) return;
    const el = document.getElementById(`row-${newRowId}`);
    if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
  }, [newRowId]);
  const [matSubTab, setMatSubTab] = useState("requisicoes"); // "requisicoes" | "analises"
  const emergFileRef = useRef(null);
  const [expandedRow, setExpandedRow] = useState(null);
  const [sort, setSort] = useState({ key: null, dir: 1 });
  const [mf, setMf] = useState({ tmMaster: "", sap: "", descricao: "", rc: "", reserva: "", po: "", status: "Todos", priority: "Todos", dataInicio: "", dataFim: "" });
  const hasActiveFilter = mf.tmMaster || mf.sap || mf.descricao || mf.rc || mf.reserva || mf.po || mf.status !== "Todos" || mf.priority !== "Todos" || mf.dataInicio || mf.dataFim;

  const filtered = useMemo(() => {
    const norm = (s) => String(s ?? "").toLowerCase().trim();
    return materials.filter((m) =>
      (!mf.tmMaster || norm(m.tmMaster).includes(norm(mf.tmMaster))) &&
      (!mf.sap || norm(m.sap).includes(norm(mf.sap))) &&
      (!mf.descricao || norm(m.descricao).includes(norm(mf.descricao))) &&
      (!mf.rc || norm(m.rc).includes(norm(mf.rc))) &&
      (!mf.reserva || norm(m.reserva).includes(norm(mf.reserva))) &&
      (!mf.po || norm(m.po).includes(norm(mf.po))) &&
      (mf.status === "Todos" || norm(m.status) === norm(mf.status)) &&
      (mf.priority === "Todos" || norm(m.priority) === norm(mf.priority)) &&
      (!mf.dataInicio || !m.dataSolicitacao || m.dataSolicitacao >= mf.dataInicio) &&
      (!mf.dataFim || !m.dataSolicitacao || m.dataSolicitacao <= mf.dataFim)
    );
  }, [materials, mf]);
  const sorted = useMemo(() => sortRows(filtered, sort), [filtered, sort]);

  const naoRecebido = (m) => !["Recebido", "Entregue a bordo"].includes(m.status);
  const PRIO_COLOR = { "Crítica": "#EF4444", "Emergencial": "#EF4444", "Sobressalente crítico": "#EF4444", "Alta": "#F5A623", "Importante": "#F5A623", "Média": "#3B82F6", "Baixa": "#9499A8" };
  const [groupByPrio, setGroupByPrio] = useState(true);
  const [collapsedMat, setCollapsedMat] = useState(new Set());
  const toggleMatGroup = (k) => setCollapsedMat((p) => { const n = new Set(p); n.has(k) ? n.delete(k) : n.add(k); return n; });
  const matGroups = useMemo(() => {
    if (!groupByPrio) return [{ key: "Todas", rows: sorted }];
    const ordem = ["Emergencial", "Sobressalente crítico", "Crítica", "Alta", "Importante", "Média", "Baixa"];
    const map = new Map();
    sorted.forEach((m) => { const k = m.priority || "Sem prioridade"; if (!map.has(k)) map.set(k, []); map.get(k).push(m); });
    const pos = (k) => { const x = ordem.indexOf(k); return x < 0 ? 99 : x; };
    return [...map.entries()].sort((x, y) => pos(x[0]) - pos(y[0])).map(([key, rows]) => ({ key, rows }));
  }, [sorted, groupByPrio]);
  const urgentes = filtered.filter((m) => ["Alta", "Crítica", "Emergencial", "Sobressalente crítico"].includes(m.priority) && naoRecebido(m));
  const abertas = filtered.filter(naoRecebido);
  const semEta = filtered.filter((m) => !m.eta && naoRecebido(m));
  const semPo = filtered.filter((m) => !m.po && naoRecebido(m));

  /* dados para os gráficos — sempre em cima do conjunto já filtrado na página */
  const PRIORITY_COLOR = {
    "Baixa": "#8D9BB5", "Média": "#F2C94C", "Alta": "#F2A93B",
    "Crítica": "#F2685B", "Importante": "#F2A93B", "Emergencial": "#F2685B", "Sobressalente crítico": "#C0392B",
  };
  const porStatus = useMemo(() => {
    const map = {};
    filtered.forEach((m) => { const k = m.status || "—"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([status, count]) => ({ status, count })).sort((a, b) => b.count - a.count);
  }, [filtered]);
  const porPrioridade = useMemo(() => {
    const map = {};
    filtered.forEach((m) => { const k = m.priority || "—"; map[k] = (map[k] || 0) + 1; });
    return PRIORITY.filter((p) => map[p]).map((p) => ({ priority: p, count: map[p] }));
  }, [filtered]);
  const porDepartamento = useMemo(() => {
    const map = {};
    filtered.forEach((m) => { const k = m.departamento || "Sem departamento"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([departamento, count]) => ({ departamento, count })).sort((a, b) => b.count - a.count);
  }, [filtered]);

  React.useEffect(() => {
    if (!setReportFn) return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      let y = pdfHeader(doc, "Relatório de Materiais", `${filtered.length} material(is) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [
        { label: "Total de Materiais", value: filtered.length },
        { label: "Materiais Urgentes", value: urgentes.length },
        { label: "Requisições Abertas", value: abertas.length },
        { label: "Sem ETA", value: semEta.length },
        { label: "Sem PO", value: semPo.length },
      ]);
      y = pdfSectionTitle(doc, y, "Por Status");
      y = pdfTable(doc, y, ["Status", "Quantidade"], porStatus.map((d) => [d.status, d.count]), { columnStyles: { 1: { halign: "right" } } });
      y = pdfSectionTitle(doc, y, "Por Prioridade");
      y = pdfTable(doc, y, ["Prioridade", "Quantidade"], porPrioridade.map((d) => [d.priority, d.count]), { columnStyles: { 1: { halign: "right" } } });
      y = pdfSectionTitle(doc, y, "Por Departamento");
      y = pdfTable(doc, y, ["Departamento", "Quantidade"], porDepartamento.map((d) => [d.departamento, d.count]), { columnStyles: { 1: { halign: "right" } } });
      if (y > 220) { doc.addPage(); y = 15; }
      y = pdfSectionTitle(doc, y, "Materiais");
      pdfTable(doc, y,
        ["TM Master", "Descrição", "Qtd", "Prioridade", "Necessidade", "PO", "ETA", "Status"],
        filtered.map((m) => [m.tmMaster || "—", m.descricao, m.quantidade, m.priority, fmtDate(m.dataNecessidade), m.po || "—", m.eta ? fmtDate(m.eta) : "sem ETA", m.status])
      );
      pdfSave(doc, "relatorio-materiais");
    });
  }, [filtered, urgentes, abertas, semEta, semPo, porStatus, porPrioridade, porDepartamento, setReportFn]);

  return (
    <>
      <div className="g-mode-toggle" style={{ marginBottom: 16, width: "fit-content" }}>
        <button className={matSubTab === "requisicoes" ? "active" : ""} onClick={() => setMatSubTab("requisicoes")}>Requisições</button>
        <button className={matSubTab === "analises" ? "active" : ""} onClick={() => setMatSubTab("analises")}>Análises</button>
      </div>

      {/* filtros da aba Materiais — digitáveis + selecionáveis, compartilhados pelas duas sub-abas */}
      <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
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
          <label>Período (Data da Solicitação) — de</label>
          <input type="date" value={mf.dataInicio} onChange={(e) => setMf((p) => ({ ...p, dataInicio: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>Período — até</label>
          <input type="date" value={mf.dataFim} onChange={(e) => setMf((p) => ({ ...p, dataFim: e.target.value }))} />
        </div>
        <div className="g-field">
          <label>&nbsp;</label>
          <button className="g-btn" onClick={() => setMf({ tmMaster: "", sap: "", descricao: "", rc: "", reserva: "", po: "", status: "Todos", priority: "Todos", dataInicio: "", dataFim: "" })}
            disabled={!hasActiveFilter} style={{ opacity: hasActiveFilter ? 1 : 0.5 }}>
            <X size={13} />Limpar filtro
          </button>
        </div>
      </div>

      {/* KPIs de análise de materiais — visíveis nas duas sub-abas */}
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
        {bigKpi("Total de Materiais", filtered.length, "var(--teal)", Package)}
        {bigKpi("Materiais Urgentes", urgentes.length, "var(--crit)", AlertTriangle)}
        {bigKpi("Requisições Abertas", abertas.length, "var(--warn)", AlertTriangle)}
        {bigKpi("Sem ETA", semEta.length, "var(--crit)", AlertTriangle)}
        {bigKpi("Sem PO", semPo.length, "var(--warn)", AlertTriangle)}
      </div>

      {matSubTab === "analises" ? (
        <>
          <div className="g-grid-2">
            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Materiais por Status</span></div>
              <div style={{ width: "100%", height: 220 }}>
                <ResponsiveContainer>
                  <BarChart data={porStatus} layout="vertical" margin={{ left: 0, right: 16, top: 4, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" horizontal={false} />
                    <XAxis type="number" allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                    <YAxis type="category" dataKey="status" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={110} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} cursor={{ fill: "rgba(10,14,20,0.04)" }} />
                    <Bar dataKey="count" name="Materiais" radius={[0, 3, 3, 0]}>
                      {porStatus.map((d, idx) => <Cell key={idx} fill={statusColor(d.status)} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="g-panel">
              <div className="g-panel-head"><span className="g-panel-title">Materiais por Prioridade</span></div>
              <div style={{ width: "100%", height: 220 }}>
                <ResponsiveContainer>
                  <BarChart data={porPrioridade} margin={{ left: 0, right: 8, top: 4, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                    <XAxis dataKey="priority" tick={{ fill: "var(--text-faint)", fontSize: 9 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} interval={0} angle={-20} textAnchor="end" height={50} />
                    <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} cursor={{ fill: "rgba(10,14,20,0.04)" }} />
                    <Bar dataKey="count" name="Materiais" radius={[3, 3, 0, 0]}>
                      {porPrioridade.map((d, idx) => <Cell key={idx} fill={PRIORITY_COLOR[d.priority] || "var(--teal)"} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Materiais por Departamento</span></div>
            <div style={{ width: "100%", height: 220 }}>
              <ResponsiveContainer>
                <BarChart data={porDepartamento} margin={{ left: 0, right: 8, top: 4, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                  <XAxis dataKey="departamento" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} interval={0} angle={-20} textAnchor="end" height={50} />
                  <YAxis allowDecimals={false} tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
                  <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} cursor={{ fill: "rgba(10,14,20,0.04)" }} />
                  <Bar dataKey="count" name="Materiais" radius={[3, 3, 0, 0]} fill="var(--accent)" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Materiais urgentes em aberto ({urgentes.length})</span></div>
            {urgentes.length === 0 && <div className="g-muted">Nenhum material urgente em aberto no momento.</div>}
            {urgentes.slice(0, 15).map((m) => (
              <div className="g-list-item" key={m.id}>
                <span>{m.descricao} · {m.departamento || "sem departamento"} · necessário {fmtDate(m.dataNecessidade)}</span>
                <span className="g-flex">
                  <span className="g-pill" style={{ background: "var(--panel-raised)" }}>
                    <span className="g-dot" style={{ background: PRIORITY_COLOR[m.priority] || "var(--warn)" }} />{m.priority}
                  </span>
                  <Pill status={m.status} />
                </span>
              </div>
            ))}
          </div>
        </>
      ) : (
      <div className="g-panel" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
        <div>
          <div className="g-panel-title" style={{ marginBottom: 4 }}>Importação semanal — Pedidos Emergenciais</div>
          <div className="g-muted" style={{ fontSize: 11.5 }}>
            Suba a planilha toda semana neste mesmo padrão de colunas. Itens já existentes (mesmo SAP) são atualizados; itens novos são adicionados — nada é duplicado.
          </div>
        </div>
        <div>
          <input ref={emergFileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImportEmergenciais} />
          <button className="g-btn primary" onClick={() => emergFileRef.current?.click()}>
            <Upload size={14} />Importar Pedidos Emergenciais
          </button>
        </div>
      </div>
      )}

      {matSubTab === "requisicoes" && (
      <div className="g-panel" style={{ padding: 0, overflow: "hidden" }}>
        <div className="tbl-toolbar">
          <label className="tbl-check"><input type="checkbox" checked={groupByPrio} onChange={(e) => setGroupByPrio(e.target.checked)} />Agrupar por prioridade</label>
          <span className="g-muted" style={{ fontSize: 11 }}>{filtered.length} material(is) · clique na seta de uma linha para editar todos os detalhes</span>
        </div>
        <div className="g-table-wrap">
        <table className="g-table tbl-modern">
          <thead>
            <tr>
              <th style={{ width: 34 }}></th>
              <SortTh sortKey="descricao" sort={sort} setSort={setSort} style={{ minWidth: 320 }}>Material</SortTh>
              <SortTh sortKey="quantidade" sort={sort} setSort={setSort}>Qtd</SortTh>
              <SortTh sortKey="dataNecessidade" sort={sort} setSort={setSort} style={{ minWidth: 150 }}>Necessário Até</SortTh>
              <SortTh sortKey="priority" sort={sort} setSort={setSort} style={{ minWidth: 130 }}>Prioridade</SortTh>
              <SortTh sortKey="status" sort={sort} setSort={setSort} style={{ minWidth: 180 }}>Status</SortTh>
              <SortTh sortKey="valor" sort={sort} setSort={setSort}>Valor</SortTh>
              <th style={{ width: 70 }}></th>
            </tr>
          </thead>
          <tbody>
            {matGroups.map((g) => (
              <React.Fragment key={g.key}>
                {groupByPrio && (
                  <tr className="tbl-group" onClick={() => toggleMatGroup(g.key)}>
                    <td colSpan={8}>
                      <span className="tbl-group-arrow">{collapsedMat.has(g.key) ? "▸" : "▾"}</span>
                      Prioridade {g.key}
                      <span className="tbl-group-cnt">{g.rows.length} item(ns)</span>
                      <span className="tbl-group-tot">{g.rows.filter((x) => naoRecebido(x)).length} em aberto</span>
                    </td>
                  </tr>
                )}
                {!(groupByPrio && collapsedMat.has(g.key)) && g.rows.map((m) => {
              const i = materials.indexOf(m);
              const isOpen = expandedRow === i;
              const cor = PRIO_COLOR[m.priority] || "#9499A8";
              const dn = naoRecebido(m) && m.dataNecessidade ? Math.round((new Date(m.dataNecessidade) - new Date(todayISO())) / 86400000) : null;
              return (
                <React.Fragment key={i}>
                  <tr id={`row-${m.id}`} className={"g-row tbl-row" + (newRowId === m.id ? " g-row-flash" : "")} style={{ "--c": cor }}>
                    <td className="tbl-first"></td>
                    <td style={{ minWidth: 320, whiteSpace: "normal", verticalAlign: "top" }}>
                      <ETextArea rows={1} value={m.descricao} onChange={(v) => updMat(i, "descricao", v)} />
                      <div className="tbl-chips">
                        {m.departamento && <span className="tbl-chip">{m.departamento}</span>}
                        {m.tmMaster && <span className="tbl-chip">TM {m.tmMaster}</span>}
                        {m.sap && <span className="tbl-chip">SAP {m.sap}</span>}
                        {m.rc && <span className="tbl-chip">RC {m.rc}</span>}
                        {m.po && <span className="tbl-chip">PO {m.po}</span>}
                      </div>
                    </td>
                    <td><ENum value={m.quantidade} onChange={(v) => updMat(i, "quantidade", v)} /></td>
                    <td style={{ minWidth: 150 }}>
                      <EDate value={m.dataNecessidade} onChange={(v) => updMat(i, "dataNecessidade", v)} />
                      {dn !== null && <div className="tbl-sub" style={{ color: dn < 0 ? "#B91C1C" : dn <= 7 ? "#B45309" : undefined, fontWeight: dn <= 7 ? 700 : 400 }}>{dn < 0 ? `atrasado ${Math.abs(dn)} dia(s)` : dn === 0 ? "é hoje" : `em ${dn} dia(s)`}</div>}
                    </td>
                    <td style={{ minWidth: 130 }}><ESelect value={m.priority} onChange={(v) => updMat(i, "priority", v)} options={PRIORITY} /></td>
                    <td style={{ minWidth: 180 }}><ESelect value={m.status} onChange={(v) => updMat(i, "status", v)} options={MAT_STATUS} /></td>
                    <td><ENum value={m.valor} onChange={(v) => updMat(i, "valor", v)} /></td>
                    <td style={{ whiteSpace: "nowrap" }}>
                      <span className="g-btn ghost" onClick={() => setExpandedRow(isOpen ? null : i)} title="Ver todos os detalhes">
                        {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                      </span>
                      <span className="g-btn ghost danger" onClick={() => remMat(i)}><Trash2 size={13} /></span>
                    </td>
                  </tr>
                  {isOpen && (
                    <tr className="tbl-detail">
                      <td></td>
                      <td colSpan={7}>
                        <div className="tbl-detail-grid">
                          <div className="g-field"><label>TM Master</label><EText value={m.tmMaster} onChange={(v) => updMat(i, "tmMaster", v)} mono /></div>
                          <div className="g-field"><label>Departamento</label><EText value={m.departamento} onChange={(v) => updMat(i, "departamento", v)} /></div>
                          <div className="g-field"><label>SAP</label><EText value={m.sap} onChange={(v) => updMat(i, "sap", v)} mono /></div>
                          <div className="g-field"><label>Reserva</label><EText value={m.reserva} onChange={(v) => updMat(i, "reserva", v)} mono /></div>
                          <div className="g-field"><label>RC</label><EText value={m.rc} onChange={(v) => updMat(i, "rc", v)} mono /></div>
                          <div className="g-field"><label>PO</label><EText value={m.po} onChange={(v) => updMat(i, "po", v)} mono /></div>
                          <div className="g-field"><label>Linha da PO</label><EText value={m.linhaPo} onChange={(v) => updMat(i, "linhaPo", v)} mono /></div>
                          <div className="g-field"><label>Data da Solicitação</label><EDate value={m.dataSolicitacao} onChange={(v) => updMat(i, "dataSolicitacao", v)} /></div>
                          <div className="g-field"><label>ETA</label>
                            {m.eta
                              ? <EDate value={m.eta} onChange={(v) => updMat(i, "eta", v)} />
                              : <span className="g-flex"><span style={{ color: "var(--crit)", fontSize: 11 }}>sem ETA</span>
                                  <input type="date" className="g-edit mono" onChange={(e) => updMat(i, "eta", e.target.value)} /></span>}
                          </div>
                          <div className="g-field"><label>Data de Recebimento</label><EDate value={m.dataRecebimento} onChange={(v) => updMat(i, "dataRecebimento", v)} /></div>
                          <div className="g-field" style={{ minWidth: 220 }}>
                            <label>Vincular a um serviço (opcional)</label>
                            <select className="g-edit" value={m.wp || ""} onChange={(e) => updMat(i, "wp", e.target.value)}>
                              <option value="">— nenhum —</option>
                              {workPackages.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
                            </select>
                          </div>
                          <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Observação</label><ETextArea rows={1} value={m.obs} onChange={(v) => updMat(i, "obs", v)} /></div>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
              </React.Fragment>
            ))}
          </tbody>
        </table>
        </div>
        {filtered.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum material encontrado com esses filtros.</div>}
      </div>
      )}
    </>
  );
}

/* ============================================================
   PAYMENTS — Pago / Pendente / Atrasado
   ============================================================ */
/* ============================================================
   PAYMENTS SECTION — two pages: full Dashboard, and Status view
   ============================================================ */
function MultiSelectStatus({ options, selected, onChange, labelFor }) {
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
  const getLabel = labelFor || ((o) => o);
  const label = selected.length === 0 ? "Todos" : selected.length === 1 ? getLabel(selected[0]) : `${selected.length} selecionados`;
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
              {getLabel(opt)}
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

function PaymentsSection({ paySubTab, setPaySubTab, serviceInvoices, updInv, remInv, addInv, setReportFn, newRowId }) {
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
    <div className="pay-view">
      <div className="g-mode-toggle" style={{ marginBottom: 16, width: "fit-content" }}>
        <button className={paySubTab === "total" ? "active" : ""} onClick={() => setPaySubTab("total")}>Dashboard Total</button>
        <button className={paySubTab === "status" ? "active" : ""} onClick={() => setPaySubTab("status")}>Status dos Pagamentos</button>
        <button className={paySubTab === "dashboard" ? "active" : ""} onClick={() => setPaySubTab("dashboard")}>Dashboard de Valores</button>
      </div>

      {/* filtro compartilhado — presente em todas as páginas da aba Pagamentos */}
      <div className="dsh-period" style={{ alignItems: "flex-end" }}>
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
        <div className="g-alert" style={{ background: "rgba(59,130,246,0.1)", borderColor: "rgba(59,130,246,0.4)", color: "var(--accent)", justifyContent: "space-between", display: "flex", alignItems: "center" }}>
          <span><strong>{selectedIds.size}</strong> serviço(s) selecionado(s) — os KPIs abaixo refletem só a seleção.</span>
          <span className="g-btn ghost" onClick={clearSelection} style={{ color: "var(--accent)" }}><X size={13} />Limpar seleção</span>
        </div>
      )}

      {paySubTab === "total" && <PaymentsTotalView serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} f={f} selectedIds={selectedIds} toggleSelect={toggleSelect} setReportFn={setReportFn} newRowId={newRowId} />}
      {paySubTab === "status" && <PaymentsStatusView serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} f={f} setF={setF} selectedIds={selectedIds} toggleSelect={toggleSelect} setReportFn={setReportFn} newRowId={newRowId} />}
      {paySubTab === "dashboard" && <PaymentsValoresView serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} f={f} selectedIds={selectedIds} toggleSelect={toggleSelect} setReportFn={setReportFn} newRowId={newRowId} />}
    </div>
  );
}

/* ---------- Página 1: Dashboard Total (todas as colunas da planilha + filtros + métricas de prazo) ---------- */
function PaymentsTotalView({ serviceInvoices, updInv, remInv, f, selectedIds, toggleSelect, setReportFn, newRowId }) {
  React.useEffect(() => {
    if (!newRowId) return;
    const el = document.getElementById(`row-${newRowId}`);
    if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
  }, [newRowId]);
  const [sort, setSort] = useState({ key: "date", dir: 1 });
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
  const [groupByEmpresa, setGroupByEmpresa] = useState(true);
  const [collapsedGroups, setCollapsedGroups] = useState(new Set());
  const [expandedId, setExpandedId] = useState(null);
  const toggleGroup = (k) => setCollapsedGroups((p) => { const n = new Set(p); n.has(k) ? n.delete(k) : n.add(k); return n; });
  const groups = useMemo(() => {
    if (!groupByEmpresa) return [{ key: "Todos", rows: sorted, total: 0, atrasados: 0 }];
    const map = new Map();
    sorted.forEach((r) => {
      const k = (r.empresa || "").trim() || "Sem empresa";
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(r);
    });
    return [...map.entries()].sort((x, y) => x[0].localeCompare(y[0], "pt-BR")).map(([key, rows]) => ({
      key, rows,
      total: rows.reduce((t, r) => t + Number(r.valorTotal || 0), 0),
      atrasados: rows.filter((r) => r.statusPagamento !== "Pago" && r.statusPagamento !== "Cancelado" && Number(r.daysOpenTotal || 0) > 60).length,
    }));
  }, [sorted, groupByEmpresa]);

  /* quando há seleção, os KPIs refletem só os serviços selecionados (dentro do filtro atual) */
  const activeRows = selectedIds.size > 0 ? filtered.filter((r) => selectedIds.has(r.id)) : filtered;

  const avg = (arr) => (arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0);
  const execPayVals = activeRows.map(execToPayDays).filter((v) => v !== null);
  /* valores negativos (MD enviada depois da execução, inconsistência) não entram na média */
  const mdExecVals = activeRows.map(mdToExecDays).filter((v) => v !== null && v >= 0);
  const totalDiasAberto = activeRows.reduce((s, r) => s + Number(r.daysOpenTotal || 0), 0);
  const valorTotalSum = activeRows.reduce((s, r) => s + Number(r.valorTotal || 0), 0);
  const emAtraso = activeRows.filter((r) => Number(r.daysOpenTotal || 0) > 60).length;
  const statusBar = useMemo(() => {
    const map = {};
    activeRows.forEach((r) => { const k = r.statusPagamento || "Sem status"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([name, value]) => ({ name, value, color: STATUS_PAGAMENTO_COLOR[name] || "#9499A8" }));
  }, [activeRows]);

  React.useEffect(() => {
    if (!setReportFn) return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      let y = pdfHeader(doc, "Relatório de Pagamentos — Dashboard Total",
        `${activeRows.length} registro(s)${selectedIds.size > 0 ? " (seleção aplicada)" : ""} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [
        { label: "Registros", value: activeRows.length },
        { label: "Valor Total", value: fmt(valorTotalSum) },
        { label: "Serviços em Atraso (+60d)", value: emAtraso },
        { label: "Média Execução → Pagamento", value: `${avg(execPayVals).toFixed(1)} dias` },
        { label: "Média MD → Execução", value: `${avg(mdExecVals).toFixed(1)} dias` },
        { label: "Total de Dias em Aberto", value: `${totalDiasAberto} dias` },
      ]);
      y = pdfSectionTitle(doc, y, "Registros");
      pdfTable(doc, y,
        ["Data", "Manutenção", "Empresa", "RC", "PO/Contrato", "Valor Total", "Status Pagamento", "Data Pagamento"],
        activeRows.map((r) => [fmtDate(r.date), r.assunto, r.empresa, r.rc || "—", r.poContrato || "—", fmt(r.valorTotal), r.statusPagamento, r.dataPagamento ? fmtDate(r.dataPagamento) : "—"])
      );
      pdfSave(doc, "relatorio-pagamentos-total");
    });
  }, [activeRows, valorTotalSum, emAtraso, execPayVals, mdExecVals, totalDiasAberto, selectedIds, setReportFn]);

  return (
    <>
      <div className="pay-hero">
        <div className="dsh-card pay-main">
          <div className="dsh-label">Valor Total</div>
          <div className="dsh-big" style={{ color: "var(--teal)" }}>{fmt(valorTotalSum)}</div>
          <div className="dsh-sub">{activeRows.length} registro(s){selectedIds.size > 0 ? " selecionado(s)" : ""}</div>
        </div>
        <div className="dsh-card pay-mini" style={{ borderTop: "3px solid var(--crit)" }}>
          <div className="dsh-label">Em Atraso (+60d)</div>
          <div className="pay-num" style={{ color: "var(--crit)" }}>{emAtraso}</div>
        </div>
        <div className="dsh-card pay-mini">
          <div className="dsh-label">Execução → Pagamento</div>
          <div className="pay-num">{avg(execPayVals).toFixed(1)}<small> dias</small></div>
        </div>
        <div className="dsh-card pay-mini">
          <div className="dsh-label">MD → Execução</div>
          <div className="pay-num">{avg(mdExecVals).toFixed(1)}<small> dias</small></div>
        </div>
        <div className="dsh-card pay-mini">
          <div className="dsh-label">Dias em Aberto (Total)</div>
          <div className="pay-num">{totalDiasAberto}<small> dias</small></div>
        </div>
      </div>

      {statusBar.length > 0 && (
        <div className="dsh-card" style={{ marginBottom: 14, padding: "14px 18px" }}>
          <div className="dsh-title" style={{ marginBottom: 8 }}>Situação dos Pagamentos</div>
          <div className="pay-statusbar">
            {statusBar.map((x) => <div key={x.name} style={{ flex: x.value, background: x.color }} title={`${x.name}: ${x.value}`} />)}
          </div>
          <div className="dsh-legend" style={{ flexDirection: "row", flexWrap: "wrap", gap: "6px 18px", marginTop: 10 }}>
            {statusBar.map((x) => <div key={x.name}><i style={{ background: x.color }} />{x.name}<b style={{ marginLeft: 6 }}>{x.value}</b></div>)}
          </div>
        </div>
      )}

      <div className="g-panel" style={{ padding: 0, overflow: "hidden" }}>
        <div className="tbl-toolbar">
          <label className="tbl-check"><input type="checkbox" checked={groupByEmpresa} onChange={(e) => setGroupByEmpresa(e.target.checked)} />Agrupar por empresa</label>
          <span className="g-muted" style={{ fontSize: 11 }}>{filtered.length} registro(s) · clique na seta de uma linha para ver e editar todos os detalhes</span>
        </div>
        <div className="g-table-wrap">
        <table className="g-table tbl-modern">
          <thead>
            <tr>
              <th style={{ width: 34 }}></th>
              <SortTh sortKey="assunto" sort={sort} setSort={setSort} style={{ minWidth: 320 }}>Serviço</SortTh>
              <SortTh sortKey="empresa" sort={sort} setSort={setSort} style={{ minWidth: 130 }}>Empresa</SortTh>
              <SortTh sortKey="statusPagamento" sort={sort} setSort={setSort} style={{ minWidth: 210 }}>Status do Pagamento</SortTh>
              <SortTh sortKey="daysOpenTotal" sort={sort} setSort={setSort}>Tempo em Aberto</SortTh>
              <SortTh sortKey="valorTotal" sort={sort} setSort={setSort} style={{ minWidth: 120 }}>Valor Total</SortTh>
              <th style={{ width: 70 }}></th>
            </tr>
          </thead>
          <tbody>
            {groups.map((g) => (
              <React.Fragment key={g.key}>
                {groupByEmpresa && (
                  <tr className="tbl-group" onClick={() => toggleGroup(g.key)}>
                    <td colSpan={7}>
                      <span className="tbl-group-arrow">{collapsedGroups.has(g.key) ? "▸" : "▾"}</span>
                      {g.key}
                      <span className="tbl-group-cnt">{g.rows.length} serviço(s)</span>
                      <span className="tbl-group-tot">{fmt(g.total)}{g.atrasados > 0 ? ` · ${g.atrasados} em atraso` : ""}</span>
                    </td>
                  </tr>
                )}
                {!(groupByEmpresa && collapsedGroups.has(g.key)) && g.rows.map((r) => {
                  const i = serviceInvoices.indexOf(r);
                  const isOpen = expandedId === r.id;
                  const cor = STATUS_PAGAMENTO_COLOR[r.statusPagamento] || "#9499A8";
                  const dias = Number(r.daysOpenTotal || 0);
                  const diasCls = r.statusPagamento === "Pago" ? "d1" : dias > 60 ? "d3" : dias > 45 ? "d2" : "d1";
                  return (
                    <React.Fragment key={r.id}>
                      <tr id={`row-${r.id}`} className={"g-row tbl-row" + (newRowId === r.id ? " g-row-flash" : "")} style={{ "--c": cor, ...(selectedIds.has(r.id) ? { background: "rgba(59,130,246,0.08)" } : {}) }}>
                        <td className="tbl-first"><input type="checkbox" checked={selectedIds.has(r.id)} onChange={() => toggleSelect(r.id)} /></td>
                        <td style={{ minWidth: 320, whiteSpace: "normal", verticalAlign: "top" }}>
                          <ETextArea rows={1} value={r.assunto} onChange={(v) => updInv(i, "assunto", v)} />
                          <div className="tbl-chips">
                            {r.date && <span className="tbl-chip">Exec. {fmtDate(r.date)}</span>}
                            {r.rc && <span className="tbl-chip">RC {r.rc}</span>}
                            {r.poContrato && <span className="tbl-chip">PO {r.poContrato}</span>}
                            {r.medicao && <span className="tbl-chip">Medição {r.medicao}</span>}
                          </div>
                        </td>
                        <td style={{ minWidth: 130 }}><EText value={r.empresa} onChange={(v) => updInv(i, "empresa", v)} /></td>
                        <td style={{ minWidth: 210 }}>
                          <StatusPagamentoSelect value={r.statusPagamento} onChange={(v) => updInv(i, "statusPagamento", v)} />
                          {r.statusPagamento === "Pago" && r.dataPagamento && <div className="tbl-sub">Pago em {fmtDate(r.dataPagamento)}</div>}
                        </td>
                        <td><span className={`tbl-days ${diasCls}`}>{dias} dias</span></td>
                        <td style={{ minWidth: 120 }}><ENum value={r.valorTotal} onChange={(v) => updInv(i, "valorTotal", v)} /></td>
                        <td style={{ whiteSpace: "nowrap" }}>
                          <span className="g-btn ghost" onClick={() => setExpandedId(isOpen ? null : r.id)} title="Ver todos os detalhes">
                            {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                          </span>
                          <span className="g-btn ghost danger" onClick={() => remInv(i)}><Trash2 size={13} /></span>
                        </td>
                      </tr>
                      {isOpen && (
                        <tr className="tbl-detail">
                          <td></td>
                          <td colSpan={6}>
                            <div className="tbl-detail-grid">
                              <div className="g-field"><label>Data de Execução</label><EDate value={r.date} onChange={(v) => updInv(i, "date", v)} /></div>
                              <div className="g-field"><label>MD</label><ESelect value={r.md} onChange={(v) => updInv(i, "md", v)} options={["Sim", "Não"]} /></div>
                              <div className="g-field"><label>Envio MD</label><EDate value={r.mdSentDate} onChange={(v) => updInv(i, "mdSentDate", v)} /></div>
                              <div className="g-field"><label>Dias em Aberto (Total)</label><ENum value={r.daysOpenTotal} onChange={(v) => updInv(i, "daysOpenTotal", v)} /></div>
                              <div className="g-field"><label>RC</label><EText value={r.rc} onChange={(v) => updInv(i, "rc", v)} mono /></div>
                              <div className="g-field"><label>Status do Serviço</label><ESelect value={r.serviceStatus} onChange={(v) => updInv(i, "serviceStatus", v)} options={["Aberto", "Fechado"]} /></div>
                              <div className="g-field"><label>PO / Contrato</label><EText value={r.poContrato} onChange={(v) => updInv(i, "poContrato", v)} mono /></div>
                              <div className="g-field"><label>Medição</label><EText value={r.medicao} onChange={(v) => updInv(i, "medicao", v)} mono /></div>
                              <div className="g-field"><label>Data do Pagamento</label><EDate value={r.dataPagamento} onChange={(v) => updInv(i, "dataPagamento", v)} /></div>
                              <div className="g-field"><label>Execução → Pagamento</label><div style={{ padding: "7px 0", fontWeight: 600 }}>{execToPayDays(r) ?? "—"} {execToPayDays(r) !== null ? "dias" : ""}</div></div>
                              <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Observações</label><ETextArea value={r.obs} onChange={(v) => updInv(i, "obs", v)} /></div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </React.Fragment>
            ))}
          </tbody>
        </table>
        </div>
        {filtered.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum registro encontrado com esses filtros.</div>}
      </div>
    </>
  );
}

/* ---------- Página 2: Status dos Pagamentos (baseada na planilha "Pagamento Pendente") ---------- */
function PaymentsStatusView({ serviceInvoices, updInv, remInv, f, setF, selectedIds, toggleSelect, setReportFn }) {
  const [sort, setSort] = useState({ key: "date", dir: 1 });
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

  React.useEffect(() => {
    if (!setReportFn) return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      let y = pdfHeader(doc, "Relatório de Status dos Pagamentos",
        `${activeRows.length} registro(s)${selectedIds.size > 0 ? " (seleção aplicada)" : ""} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [
        { label: "Todos", value: activeRows.length },
        ...kpiStatuses.map((s) => ({ label: s, value: countOf(s) })),
      ]);
      y = pdfSectionTitle(doc, y, "Registros");
      pdfTable(doc, y,
        ["Data", "Manutenção", "Empresa", "PO/Contrato", "Medição", "Dias em Aberto", "Status Pagamento"],
        activeRows.map((r) => [fmtDate(r.date), r.assunto, r.empresa, r.poContrato || "—", r.medicao || "—", r.daysOpenTotal, r.statusPagamento])
      );
      pdfSave(doc, "relatorio-status-pagamentos");
    });
  }, [activeRows, kpiStatuses, selectedIds, setReportFn]);

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
                <tr className="g-row" key={r.id} style={color ? { borderLeft: `3px solid ${color}`, background: selectedIds.has(r.id) ? "rgba(59,130,246,0.08)" : "rgba(10,14,20,0.015)" } : (selectedIds.has(r.id) ? { background: "rgba(59,130,246,0.08)" } : undefined)}>
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
function PaymentsValoresView({ serviceInvoices, updInv, remInv, f, selectedIds, toggleSelect, setReportFn }) {
  const [situationFilter, setSituationFilter] = useState("Todos");
  const [sort, setSort] = useState({ key: "date", dir: 1 });

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

  React.useEffect(() => {
    if (!setReportFn) return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      let y = pdfHeader(doc, "Relatório — Dashboard de Valores",
        `${activeRows.length} registro(s)${selectedIds.size > 0 ? " (seleção aplicada)" : ""} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, cards.map((c) => ({ label: c.label, value: c.value })));
      y = pdfSectionTitle(doc, y, "Registros");
      pdfTable(doc, y,
        ["Data", "Serviço", "Empresa", "PO", "Valor", "Dias em Atraso", "Status", "Situação"],
        activeRows.map((r) => [fmtDate(r.date), r.assunto, r.empresa, r.poContrato || "—", fmt(r.valorTotal), r.daysOpenTotal, r.statusPagamento, r._situation])
      );
      pdfSave(doc, "relatorio-dashboard-valores");
    });
  }, [activeRows, cards, selectedIds, setReportFn]);

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
                <tr className="g-row" key={r.id} style={selectedIds.has(r.id) ? { background: "rgba(59,130,246,0.06)" } : undefined}>
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
function CostsView({ serviceInvoices, updInv, exchangeRate, setExchangeRate, setReportFn }) {
  const [costSubTab, setCostSubTab] = useState("rateio"); // "rateio" | "dashboard" | "previsao"
  const [expandedRow, setExpandedRow] = useState(null);
  const [categoriaTableCollapsed, setCategoriaTableCollapsed] = useState(false);
  const currentMonth = useMemo(() => `${new Date().getFullYear()}-${pad2(new Date().getMonth() + 1)}`, []);
  const [cf, setCf] = useState({ statuses: [], servico: "", empresa: "", categorias: [], provisionadoMeses: [currentMonth] });
  const hasActiveFilter = cf.statuses.length > 0 || cf.servico || cf.empresa || cf.categorias.length > 0 ||
    cf.provisionadoMeses.length !== 1 || cf.provisionadoMeses[0] !== currentMonth;

  /* esta aba só considera serviços que ainda não estão como "Pago" na aba Pagamentos */
  const naoPagos = useMemo(() => serviceInvoices.filter((r) => r.statusPagamento !== "Pago"), [serviceInvoices]);
  const statusOptions = useMemo(() => STATUS_PAGAMENTO_OPTIONS.filter((s) => s !== "Pago"), []);

  const allocationsOf = (r) => r.allocations || [];
  const allocatedSum = (r) => allocationsOf(r).reduce((s, a) => s + Number(a.valor || 0), 0);

  /* opções de mês disponíveis para o filtro Provisionado, geradas a partir dos dados reais */
  const monthLabel = (ym) => {
    if (!ym) return "Sem previsão";
    const [y, m] = ym.split("-");
    return `${MONTH_NAMES[Number(m) - 1]}/${y}`;
  };
  const provisionadoMesOptions = useMemo(() => {
    const set = new Set(serviceInvoices.map((r) => r.previsaoMes).filter(Boolean));
    set.add(currentMonth);
    return [...set].sort();
  }, [serviceInvoices, currentMonth]);

  /* Único filtro de tempo desta aba: o(s) MÊS(ES) PROVISIONADO(S) de cada serviço (não a data de
     execução). Isso evita que pagamentos de backlog (executados num mês anterior, mas provisionados
     para o mês corrente) fiquem de fora quando o mês selecionado é o atual. Sem nenhum mês
     selecionado, mostra tudo (todos os meses somados). Múltiplos meses podem ser combinados. */
  const filtered = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return naoPagos.filter((r) => {
      const inProvisionado = cf.provisionadoMeses.length === 0 || cf.provisionadoMeses.includes(r.previsaoMes);
      const inStatus = cf.statuses.length === 0 || cf.statuses.includes(r.statusPagamento);
      const inServico = !cf.servico || norm(r.assunto).includes(norm(cf.servico));
      const inEmpresa = !cf.empresa || norm(r.empresa).includes(norm(cf.empresa));
      const inCategoria = cf.categorias.length === 0 || allocationsOf(r).some((a) => cf.categorias.includes(a.category));
      return inProvisionado && inStatus && inServico && inEmpresa && inCategoria;
    });
  }, [naoPagos, cf]);

  const addAllocation = (i, r) => updInv(i, "allocations", [...allocationsOf(r), { category: CATEGORIES[0], valor: 0 }]);
  /* atalho: joga o valor total do serviço inteiro pra CAPEX de uma vez, sem precisar montar o
     rateio manualmente — útil já que CAPEX não tem teto de orçamento (aceita qualquer valor) */
  const marcarComoCapex = (i, r) => updInv(i, "allocations", [{ category: "CAPEX", valor: Number(r.valorTotal || 0) }]);
  const updAllocation = (i, r, ai, field, value) => {
    const next = allocationsOf(r).map((a, idx) => (idx === ai ? { ...a, [field]: value } : a));
    updInv(i, "allocations", next);
  };
  const remAllocation = (i, r, ai) => updInv(i, "allocations", allocationsOf(r).filter((_, idx) => idx !== ai));

  /* resumo por categoria: Orçado é um valor mensal fixo (não acumula entre meses); Realizado é
     calculado a partir do MÊS PROVISIONADO de cada serviço (não da data de execução), então backlogs
     de meses anteriores aparecem certinho no mês em que o pagamento foi de fato provisionado */
  const categoryCosts = useMemo(() => {
    return CATEGORIES.map((cat) => {
      const orcadoUsd = CATEGORY_BUDGET_USD[cat] || 0;
      const orcadoBrl = orcadoUsd * exchangeRate;
      const realizado = filtered.reduce((s, r) => s + allocationsOf(r).filter((a) => a.category === cat).reduce((s2, a) => s2 + Number(a.valor || 0), 0), 0);
      /* CAPEX é uma categoria sem teto de orçamento — pode receber qualquer valor, então não faz
         sentido calcular "disponível" pra ela (não tem limite pra estourar) */
      const ilimitado = cat === "CAPEX";
      return { category: cat, orcadoUsd, orcadoBrl, realizado, disponivel: orcadoBrl - realizado, ilimitado };
    });
  }, [filtered, exchangeRate]);

  /* os totais de Orçado/Realizado/% Consumido consideram só as categorias com teto de orçamento —
     CAPEX fica de fora dessa conta (é ilimitado) e aparece separado, no seu próprio KPI */
  const categoryCostsLimitadas = categoryCosts.filter((c) => !c.ilimitado);
  const capexRow = categoryCosts.find((c) => c.category === "CAPEX");
  const totalOrcadoBrl = categoryCostsLimitadas.reduce((s, c) => s + c.orcadoBrl, 0);
  const totalRealizado = categoryCostsLimitadas.reduce((s, c) => s + c.realizado, 0);
  const totalDisponivel = totalOrcadoBrl - totalRealizado;
  const pctConsumido = totalOrcadoBrl ? Math.round((totalRealizado / totalOrcadoBrl) * 100) : 0;
  const semRateioCompleto = filtered.filter((r) => Math.round(allocatedSum(r)) !== Math.round(Number(r.valorTotal || 0))).length;

  /* Previsão de provisionamento: para qual mês cada serviço foi provisionado (r.previsaoMes, "YYYY-MM").
     O valor total é separado entre CAPEX (sem teto) e as demais categorias (com orçamento), com base
     no que de fato foi rateado em cada uma — não no valor bruto do serviço. */
  const comPrevisao = filtered.filter((r) => r.previsaoMes);
  const semPrevisao = filtered.filter((r) => !r.previsaoMes);
  const capexProvisionado = comPrevisao.reduce((s, r) => s + allocationsOf(r).filter((a) => a.category === "CAPEX").reduce((s2, a) => s2 + Number(a.valor || 0), 0), 0);
  const outrasCategoriasProvisionado = comPrevisao.reduce((s, r) => s + allocationsOf(r).filter((a) => a.category !== "CAPEX").reduce((s2, a) => s2 + Number(a.valor || 0), 0), 0);
  const totalProvisionado = capexProvisionado + outrasCategoriasProvisionado;
  const provisionadoPorMes = useMemo(() => {
    const map = {};
    comPrevisao.forEach((r) => {
      if (!map[r.previsaoMes]) map[r.previsaoMes] = { capex: 0, outras: 0 };
      allocationsOf(r).forEach((a) => {
        if (a.category === "CAPEX") map[r.previsaoMes].capex += Number(a.valor || 0);
        else map[r.previsaoMes].outras += Number(a.valor || 0);
      });
    });
    return Object.keys(map).sort().map((ym) => ({ mes: monthLabel(ym), capex: map[ym].capex, outras: map[ym].outras }));
  }, [comPrevisao]);

  /* análise combinada de custo + pagamento (sub-aba Dashboard) — mesmo mês provisionado/serviço/empresa
     do filtro acima, mas sem excluir os já pagos, pra dar a visão completa da situação de pagamento */
  const allInPeriod = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return serviceInvoices.filter((r) => {
      const inProvisionado = cf.provisionadoMeses.length === 0 || cf.provisionadoMeses.includes(r.previsaoMes);
      const inServico = !cf.servico || norm(r.assunto).includes(norm(cf.servico));
      const inEmpresa = !cf.empresa || norm(r.empresa).includes(norm(cf.empresa));
      return inProvisionado && inServico && inEmpresa;
    });
  }, [serviceInvoices, cf]);
  const pagosPeriodo = allInPeriod.filter((r) => invoiceSituation(r) === "Pago");
  const pendentesPeriodo = allInPeriod.filter((r) => invoiceSituation(r) === "Pendente");
  const atrasadosPeriodo = allInPeriod.filter((r) => invoiceSituation(r) === "Atrasado");
  const sumVal = (arr) => arr.reduce((s, r) => s + Number(r.valorTotal || 0), 0);

  /* registra o gerador de PDF desta página — reflete exatamente o filtro/dados atuais da aba Custos,
     com cada parte do relatório bem separada (uma seção por página) */
  React.useEffect(() => {
    if (!setReportFn || costSubTab !== "rateio") return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      const mesesLabel = cf.provisionadoMeses.length === 0
        ? "Todos os meses"
        : cf.provisionadoMeses.length === 1
          ? `Mês provisionado: ${monthLabel(cf.provisionadoMeses[0])}`
          : `Meses provisionados: ${cf.provisionadoMeses.map(monthLabel).join(", ")}`;
      let y = pdfHeader(doc, "Relatório de Custos — Rateio por Categoria",
        `${mesesLabel} · Câmbio US$→R$ ${exchangeRate} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [
        { label: "Total Realizado (rateado)", value: fmt(totalRealizado) },
        { label: "Total Orçado", value: fmt(totalOrcadoBrl) },
        { label: "Saldo Disponível", value: fmt(totalDisponivel) },
        { label: "% Orçamento Consumido", value: `${pctConsumido}%` },
        { label: "Serviços sem Rateio Completo", value: semRateioCompleto },
        { label: "Registros no período", value: filtered.length },
      ]);

      /* ---- seção 1: custo por categoria — categorias com orçamento e CAPEX sempre separados ---- */
      y = pdfSectionTitle(doc, y, "1. Custo por categoria (OPEX) — Orçado × Realizado × Disponível");
      const categoriasComOrcamento = categoryCosts.filter((c) => !c.ilimitado);
      y = pdfTable(doc, y,
        ["Categoria", "Orçado (US$)", "Orçado (R$)", "Realizado (R$)", "Disponível (R$)", "Ordem (Compra de Serviços)"],
        categoriasComOrcamento.map((c) => [c.category, fmtBudgetUsd(c.orcadoUsd), fmtBudgetBrl(c.orcadoUsd, c.orcadoBrl), fmt(c.realizado), fmt(c.disponivel), adpServicosLabel(c.category)]),
        { columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" }, 4: { halign: "right" } } }
      );
      if (capexRow) {
        y = pdfSectionTitle(doc, y, "1b. CAPEX (sem teto de orçamento)");
        y = pdfTable(doc, y,
          ["Categoria", "Realizado (R$)", "Observação"],
          [["CAPEX", fmt(capexRow.realizado), "Sem orçamento limitado — contabiliza apenas o que foi utilizado"]],
          { columnStyles: { 1: { halign: "right" } } }
        );
      }

      /* ---- seção 2: rateio detalhado por serviço — Categoria, Ordem, Valor, Serviço, Empresa, Status
         e Justificativa Geral, sempre com CAPEX numa tabela separada das demais categorias ---- */
      const allocationRowsOutras = [];
      const allocationRowsCapex = [];
      filtered.forEach((r) => {
        allocationsOf(r).forEach((a) => {
          const row = [fmtDate(r.date), r.assunto, r.empresa, a.category, adpServicosLabel(a.category), fmt(a.valor), r.statusPagamento, r.justificativaGeral || "—"];
          (a.category === "CAPEX" ? allocationRowsCapex : allocationRowsOutras).push(row);
        });
      });
      const rateioColumns = ["Data", "Serviço", "Empresa", "Categoria", "Ordem", "Valor", "Status", "Justificativa Geral"];
      const rateioColumnStyles = {
        0: { cellWidth: 14 }, 1: { cellWidth: 26 }, 2: { cellWidth: 20 }, 3: { cellWidth: 18 },
        4: { cellWidth: 13 }, 5: { cellWidth: 17, halign: "right" }, 6: { cellWidth: 24 }, 7: { cellWidth: 50 },
      };
      if (allocationRowsOutras.length > 0) {
        y = pdfSectionTitle(doc, y, "2. Rateio por Categoria — OPEX");
        y = pdfTable(doc, y, rateioColumns, allocationRowsOutras, { columnStyles: rateioColumnStyles });
      }
      if (allocationRowsCapex.length > 0) {
        y = pdfSectionTitle(doc, y, "2b. Rateio por Categoria — CAPEX");
        y = pdfTable(doc, y, rateioColumns, allocationRowsCapex, { columnStyles: rateioColumnStyles });
      }

      /* ---- seção 2c: tabela invertida — agrupada por CATEGORIA (em vez de por serviço), mostrando
         a Ordem de cada categoria e todos os serviços/valores que a compõem ---- */
      const porCategoria = [...CATEGORIES].map((cat) => {
        const rows = [];
        filtered.forEach((r) => {
          allocationsOf(r).filter((a) => a.category === cat).forEach((a) => {
            rows.push([r.assunto, r.empresa, fmt(a.valor), r.statusPagamento]);
          });
        });
        const subtotal = filtered.reduce((s, r) => s + allocationsOf(r).filter((a) => a.category === cat).reduce((s2, a) => s2 + Number(a.valor || 0), 0), 0);
        return { cat, ordem: adpServicosLabel(cat), rows, subtotal };
      }).filter((g) => g.rows.length > 0);
      if (porCategoria.length > 0) {
        y = pdfSectionTitle(doc, y, "2c. Rateio agrupado por Categoria (Ordem, Serviços e Valores)");
        porCategoria.forEach((g) => {
          y = pdfSectionTitle(doc, y, `${g.cat} — Ordem: ${g.ordem} · ${g.rows.length} serviço(s) · Subtotal: ${fmt(g.subtotal)}`);
          y = pdfTable(doc, y,
            ["Serviço", "Empresa", "Valor", "Status de Pagamento"],
            g.rows,
            { columnStyles: { 0: { cellWidth: 60 }, 1: { cellWidth: 40 }, 2: { cellWidth: 24, halign: "right" }, 3: { cellWidth: 42 } } }
          );
        });
      }

      /* ---- seção 3: situação dos pagamentos no período ---- */
      y = pdfSectionTitle(doc, y, "3. Situação dos pagamentos no período");
      const sumValRep = (arr) => arr.reduce((s, r) => s + Number(r.valorTotal || 0), 0);
      y = pdfKpis(doc, y, [
        { label: "Pago", value: `${pagosPeriodo.length} · ${fmt(sumValRep(pagosPeriodo))}` },
        { label: "Pendente", value: `${pendentesPeriodo.length} · ${fmt(sumValRep(pendentesPeriodo))}` },
        { label: "Atrasado", value: `${atrasadosPeriodo.length} · ${fmt(sumValRep(atrasadosPeriodo))}` },
      ]);

      /* ---- seção 4: serviços agrupados por mês provisionado ---- */
      y = pdfSectionTitle(doc, y, "4. Serviços por mês provisionado");
      y = pdfKpis(doc, y, [
        { label: "Provisionado — OPEX", value: fmt(outrasCategoriasProvisionado) },
        { label: "Provisionado — CAPEX", value: fmt(capexProvisionado) },
        { label: "Serviços com Previsão", value: comPrevisao.length },
        { label: "Serviços sem Previsão", value: semPrevisao.length },
      ]);
      if (comPrevisao.length === 0) {
        doc.setFontSize(9); doc.setTextColor(...PDF_MUTED);
        doc.text("Nenhum serviço com previsão de mês definida no período selecionado.", 14, y);
        doc.setTextColor(0, 0, 0);
        y += 10;
      } else {
        const byMonth = {};
        comPrevisao.forEach((r) => { (byMonth[r.previsaoMes] = byMonth[r.previsaoMes] || []).push(r); });
        Object.keys(byMonth).sort().forEach((ym) => {
          const rows = byMonth[ym];
          const subtotal = rows.reduce((s, r) => s + Number(r.valorTotal || 0), 0);
          const subCapex = rows.reduce((s, r) => s + allocationsOf(r).filter((a) => a.category === "CAPEX").reduce((s2, a) => s2 + Number(a.valor || 0), 0), 0);
          const subOutras = rows.reduce((s, r) => s + allocationsOf(r).filter((a) => a.category !== "CAPEX").reduce((s2, a) => s2 + Number(a.valor || 0), 0), 0);
          y = pdfSectionTitle(doc, y, `${monthLabel(ym)} — ${rows.length} serviço(s) · Outras: ${fmt(subOutras)} · CAPEX: ${fmt(subCapex)} · Total: ${fmt(subtotal)}`);
          y = pdfTable(doc, y,
            ["Serviço", "Empresa", "Valor", "Status de Pagamento"],
            rows.map((r) => [r.assunto, r.empresa, fmt(r.valorTotal), r.statusPagamento]),
            { columnStyles: { 2: { halign: "right" } } }
          );
        });
      }
      if (semPrevisao.length > 0) {
        y = pdfSectionTitle(doc, y, `Sem previsão de mês definida — ${semPrevisao.length} serviço(s)`);
        y = pdfTable(doc, y,
          ["Serviço", "Empresa", "Valor", "Status de Pagamento"],
          semPrevisao.map((r) => [r.assunto, r.empresa, fmt(r.valorTotal), r.statusPagamento]),
          { columnStyles: { 2: { halign: "right" } } }
        );
      }

      /* ---- seção 5: lista completa de serviços (não pagos) ---- */
      y = pdfSectionTitle(doc, y, "5. Serviços (não pagos) — lista completa");
      pdfTable(doc, y,
        ["Data", "Serviço", "Empresa", "Valor", "PO", "Status de Pagamento", "Rateado"],
        filtered.map((r) => [
          fmtDate(r.date), r.assunto, r.empresa, fmt(r.valorTotal), r.poContrato, r.statusPagamento,
          `${fmt(allocatedSum(r))} / ${fmt(r.valorTotal)}`,
        ]),
        { columnStyles: { 3: { halign: "right" } } }
      );
      pdfSave(doc, "relatorio-custos-rateio");
    });
  }, [filtered, categoryCosts, totalRealizado, totalOrcadoBrl, totalDisponivel, pctConsumido, semRateioCompleto,
      pagosPeriodo, pendentesPeriodo, atrasadosPeriodo, comPrevisao, semPrevisao, capexProvisionado, outrasCategoriasProvisionado, exchangeRate, cf, costSubTab, setReportFn]);

  /* relatório próprio do Dashboard Financeiro — reflete o que essa sub-aba mostra de fato:
     KPIs de Custos/Pagamentos/Provisionamento, os dados dos três gráficos, e a tabela de
     serviços por mês de provisionamento (com a divergência execução × previsão) */
  React.useEffect(() => {
    if (!setReportFn || costSubTab !== "dashboard") return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      const mesesLabel = cf.provisionadoMeses.length === 0
        ? "Todos os meses"
        : cf.provisionadoMeses.length === 1
          ? `Mês provisionado: ${monthLabel(cf.provisionadoMeses[0])}`
          : `Meses provisionados: ${cf.provisionadoMeses.map(monthLabel).join(", ")}`;
      let y = pdfHeader(doc, "Relatório — Dashboard Financeiro",
        `${mesesLabel} · Câmbio US$→R$ ${exchangeRate} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);

      y = pdfSectionTitle(doc, y, "Custos");
      y = pdfKpis(doc, y, [
        { label: "Total Realizado (rateado)", value: fmt(totalRealizado) },
        { label: "Total Orçado", value: fmt(totalOrcadoBrl) },
        { label: "Saldo Disponível", value: fmt(totalDisponivel) },
        { label: "% Orçamento Consumido", value: `${pctConsumido}%` },
      ]);

      const sumValRep2 = (arr) => arr.reduce((s, r) => s + Number(r.valorTotal || 0), 0);
      y = pdfSectionTitle(doc, y, "Pagamentos");
      y = pdfKpis(doc, y, [
        { label: "Pago", value: `${pagosPeriodo.length} · ${fmt(sumValRep2(pagosPeriodo))}` },
        { label: "Pendente", value: `${pendentesPeriodo.length} · ${fmt(sumValRep2(pendentesPeriodo))}` },
        { label: "Atrasado", value: `${atrasadosPeriodo.length} · ${fmt(sumValRep2(atrasadosPeriodo))}` },
      ]);

      y = pdfSectionTitle(doc, y, "Provisionamento");
      y = pdfKpis(doc, y, [
        { label: "Provisionado — OPEX", value: fmt(outrasCategoriasProvisionado) },
        { label: "Provisionado — CAPEX", value: fmt(capexProvisionado) },
        { label: "Serviços com Previsão", value: comPrevisao.length },
        { label: "Serviços sem Previsão", value: semPrevisao.length },
      ]);

      /* gráfico 1: custo por categoria (dados) */
      y = pdfSectionTitle(doc, y, "Gráfico — Custo por categoria (Orçado × Realizado)");
      y = pdfTable(doc, y,
        ["Categoria", "Orçado (R$)", "Realizado (R$)"],
        categoryCosts.map((c) => [c.category, c.ilimitado ? "—" : fmt(c.orcadoBrl), fmt(c.realizado)]),
        { columnStyles: { 1: { halign: "right" }, 2: { halign: "right" } } }
      );

      /* gráfico 2: situação dos pagamentos (dados) */
      y = pdfSectionTitle(doc, y, "Gráfico — Situação dos pagamentos no período");
      y = pdfTable(doc, y,
        ["Situação", "Valor"],
        [
          ["Pago", fmt(sumValRep2(pagosPeriodo))],
          ["Pendente", fmt(sumValRep2(pendentesPeriodo))],
          ["Atrasado", fmt(sumValRep2(atrasadosPeriodo))],
        ],
        { columnStyles: { 1: { halign: "right" } } }
      );

      /* gráfico 3: valor provisionado por mês — OPEX × CAPEX (dados) */
      y = pdfSectionTitle(doc, y, "Gráfico — Valor provisionado por mês (OPEX × CAPEX)");
      if (provisionadoPorMes.length === 0) {
        doc.setFontSize(9); doc.setTextColor(...PDF_MUTED);
        doc.text("Nenhum serviço com previsão de mês definida no período selecionado.", 14, y);
        doc.setTextColor(0, 0, 0);
        y += 10;
      } else {
        y = pdfTable(doc, y,
          ["Mês", "OPEX (R$)", "CAPEX (R$)", "Total (R$)"],
          provisionadoPorMes.map((m) => [m.mes, fmt(m.outras), fmt(m.capex), fmt(m.outras + m.capex)]),
          { columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" } } }
        );
      }

      /* detalhamento: serviços por mês de provisionamento, com a divergência execução × previsão */
      y = pdfSectionTitle(doc, y, "Detalhamento — Serviços por mês de provisionamento");
      if (filtered.length === 0) {
        doc.setFontSize(9); doc.setTextColor(...PDF_MUTED);
        doc.text("Nenhum serviço encontrado com esses filtros.", 14, y);
        doc.setTextColor(0, 0, 0);
      } else {
        pdfTable(doc, y,
          ["Serviço", "Empresa", "Valor", "Data (execução)", "Status de Pagamento", "Previsão"],
          filtered.map((r) => {
            const execMonth = (r.date || "").slice(0, 7);
            const divergente = r.previsaoMes && execMonth && r.previsaoMes !== execMonth;
            return [
              r.assunto, r.empresa, fmt(r.valorTotal), fmtDate(r.date), r.statusPagamento,
              `${r.previsaoMes ? monthLabel(r.previsaoMes) : "—"}${divergente ? " ⚠" : ""}`,
            ];
          }),
          { columnStyles: { 2: { halign: "right" } } }
        );
      }
      pdfSave(doc, "relatorio-custos-dashboard-financeiro");
    });
  }, [filtered, categoryCosts, totalRealizado, totalOrcadoBrl, totalDisponivel, pctConsumido,
      pagosPeriodo, pendentesPeriodo, atrasadosPeriodo, comPrevisao, semPrevisao, capexProvisionado,
      outrasCategoriasProvisionado, provisionadoPorMes, exchangeRate, cf, costSubTab, setReportFn]);

  return (
    <>
      <div className="g-mode-toggle" style={{ marginBottom: 16, width: "fit-content" }}>
        <button className={costSubTab === "rateio" ? "active" : ""} onClick={() => setCostSubTab("rateio")}>Rateio por Categoria</button>
        <button className={costSubTab === "dashboard" ? "active" : ""} onClick={() => setCostSubTab("dashboard")}>Dashboard Financeiro</button>
      </div>

      <div className="g-alert" style={{ background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>
        {costSubTab === "rateio"
          ? <>Esta página mostra apenas serviços que ainda <strong>não</strong> estão marcados como "Pago" na aba Pagamentos. Os valores de orçamento por categoria são mensais — ao trocar o período para outro mês, o realizado zera e o orçado volta inteiro.</>
          : <>Visão combinada: custo por categoria (Orçado × Realizado) e situação dos pagamentos (Pago/Pendente/Atrasado), para o mesmo período e filtros abaixo.</>}
      </div>
      <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
        {costSubTab === "rateio" && (
          <>
            <div className="g-field">
              <label>Status de pagamento (múltipla escolha)</label>
              <MultiSelectStatus options={statusOptions} selected={cf.statuses} onChange={(v) => setCf((p) => ({ ...p, statuses: v }))} />
            </div>
            <div className="g-field">
              <label>Categoria (múltipla escolha)</label>
              <MultiSelectStatus options={CATEGORIES} selected={cf.categorias} onChange={(v) => setCf((p) => ({ ...p, categorias: v }))} />
            </div>
          </>
        )}
        <div className="g-field">
          <label>Serviço</label>
          <input type="text" value={cf.servico} onChange={(e) => setCf((p) => ({ ...p, servico: e.target.value }))} placeholder="digitar..." style={{ minWidth: 150 }} />
        </div>
        <div className="g-field">
          <label>Empresa</label>
          <input type="text" value={cf.empresa} onChange={(e) => setCf((p) => ({ ...p, empresa: e.target.value }))} placeholder="digitar..." style={{ minWidth: 130 }} />
        </div>
        <div className="g-field">
          <label>Provisionado (múltiplos meses)</label>
          <MultiSelectStatus options={provisionadoMesOptions} selected={cf.provisionadoMeses} labelFor={monthLabel}
            onChange={(v) => setCf((p) => ({ ...p, provisionadoMeses: v }))} />
        </div>
        <div className="g-field">
          <label>&nbsp;</label>
          <button className="g-btn" onClick={() => setCf((p) => ({ ...p, provisionadoMeses: [] }))} disabled={cf.provisionadoMeses.length === 0} style={{ opacity: cf.provisionadoMeses.length ? 1 : 0.5 }}>
            <X size={13} />Ver todos os meses
          </button>
        </div>
        <div className="g-field">
          <label>&nbsp;</label>
          <button className="g-btn" onClick={() => setCf({ statuses: [], servico: "", empresa: "", categorias: [], provisionadoMeses: [currentMonth] })}
            disabled={!hasActiveFilter} style={{ opacity: hasActiveFilter ? 1 : 0.5 }}>
            <X size={13} />Limpar filtro
          </button>
        </div>
      </div>

      {costSubTab === "dashboard" ? (
        <>
          <div className="g-section-label">Custos</div>
          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
            {bigKpi("Total Realizado (rateado)", fmt(totalRealizado), "var(--teal)", Wallet)}
            {bigKpi("Total Orçado", fmt(totalOrcadoBrl), "var(--text-dim)", Calculator)}
            {bigKpi("Saldo Disponível", fmt(totalDisponivel), totalDisponivel < 0 ? "var(--crit)" : "var(--ok)", Wallet)}
            {bigKpi("% Orçamento Consumido", `${pctConsumido}%`, pctConsumido > 100 ? "var(--crit)" : "var(--warn)", AlertTriangle)}
          </div>
          <div className="g-section-label">Pagamentos</div>
          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(3, 1fr)" }}>
            {bigKpi("Pago", `${pagosPeriodo.length} · ${fmt(sumVal(pagosPeriodo))}`, "var(--ok)", Wallet)}
            {bigKpi("Pendente", `${pendentesPeriodo.length} · ${fmt(sumVal(pendentesPeriodo))}`, "var(--warn)", AlertTriangle)}
            {bigKpi("Atrasado", `${atrasadosPeriodo.length} · ${fmt(sumVal(atrasadosPeriodo))}`, "var(--crit)", AlertTriangle)}
          </div>
          <div className="g-section-label">Provisionamento</div>
          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
            {bigKpi("Provisionado — OPEX", fmt(outrasCategoriasProvisionado), "var(--teal)", Wallet)}
            {bigKpi("Provisionado — CAPEX", fmt(capexProvisionado), "var(--accent)", Wallet)}
            {bigKpi("Serviços com Previsão", comPrevisao.length, "var(--ok)", Calculator)}
            {bigKpi("Serviços sem Previsão", semPrevisao.length, "var(--crit)", AlertTriangle)}
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Valor provisionado por mês — OPEX × CAPEX</span></div>
            {provisionadoPorMes.length === 0 ? (
              <div className="g-muted">Nenhum serviço com previsão de mês definida ainda — preencha a coluna "Previsão" na aba Rateio por Categoria.</div>
            ) : (
              <div style={{ width: "100%", height: 260 }}>
                <ResponsiveContainer>
                  <BarChart data={provisionadoPorMes} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                    <XAxis dataKey="mes" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                    <YAxis tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={40} tickFormatter={(v) => `${Math.round(v / 1000)}k`} />
                    <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} formatter={(v) => fmt(v)} />
                    <Bar dataKey="outras" name="OPEX" radius={[3, 3, 0, 0]} fill="var(--accent)">
                      <LabelList dataKey="outras" position="top" formatter={(v) => v ? fmt(v) : ""} style={{ fill: "var(--text-dim)", fontSize: 9, fontFamily: "var(--sans)" }} />
                    </Bar>
                    <Bar dataKey="capex" name="CAPEX" radius={[3, 3, 0, 0]} fill="var(--teal)">
                      <LabelList dataKey="capex" position="top" formatter={(v) => v ? fmt(v) : ""} style={{ fill: "var(--text-dim)", fontSize: 9, fontFamily: "var(--sans)" }} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Situação dos pagamentos no período</span></div>
            <div style={{ width: "100%", height: 220 }}>
              <ResponsiveContainer>
                <BarChart data={[
                  { situacao: "Pago", valor: sumVal(pagosPeriodo) },
                  { situacao: "Pendente", valor: sumVal(pendentesPeriodo) },
                  { situacao: "Atrasado", valor: sumVal(atrasadosPeriodo) },
                ]} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
                  <XAxis dataKey="situacao" tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                  <YAxis tick={{ fill: "var(--text-faint)", fontSize: 10 }} axisLine={false} tickLine={false} width={40} tickFormatter={(v) => `${Math.round(v / 1000)}k`} />
                  <Tooltip contentStyle={{ background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 }} labelStyle={{ color: "var(--text)" }} formatter={(v) => fmt(v)} />
                  <Bar dataKey="valor" radius={[3, 3, 0, 0]}>
                    <LabelList dataKey="valor" position="top" formatter={(v) => fmt(v)} style={{ fill: "var(--text-dim)", fontSize: 10, fontFamily: "var(--sans)" }} />
                    <Cell fill="var(--ok)" /><Cell fill="var(--warn)" /><Cell fill="var(--crit)" />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Serviços — mês de provisionamento</span></div>
            <div className="g-table-wrap">
            <table className="g-table">
              <thead>
                <tr><th style={{ minWidth: 220 }}>Serviço</th><th>Empresa</th><th>Valor</th><th>Data (execução)</th><th>Status de pagamento</th><th>Previsão</th></tr>
              </thead>
              <tbody>
                {filtered.map((r) => {
                  const i = serviceInvoices.indexOf(r);
                  const execMonth = (r.date || "").slice(0, 7);
                  const divergente = r.previsaoMes && execMonth && r.previsaoMes !== execMonth;
                  return (
                    <tr className="g-row" key={r.id} style={divergente ? { background: "rgba(59,130,246,0.06)" } : undefined}>
                      <td style={{ minWidth: 220, whiteSpace: "normal" }}>{r.assunto}</td>
                      <td style={{ minWidth: 130 }}>{r.empresa}</td>
                      <td style={{ fontFamily: "var(--mono)" }}>{fmt(r.valorTotal)}</td>
                      <td style={{ fontFamily: "var(--mono)" }}>{fmtDate(r.date)}</td>
                      <td><StatusPagamentoSelect value={r.statusPagamento} onChange={(v) => updInv(i, "statusPagamento", v)} /></td>
                      <td style={{ minWidth: 150 }}>
                        <div className="g-flex" style={{ gap: 5 }}>
                          <input type="month" className="g-edit" value={r.previsaoMes || ""} onChange={(e) => updInv(i, "previsaoMes", e.target.value)} />
                          {divergente && <span title={`Executado em ${execMonth} mas provisionado para ${r.previsaoMes} — pagamento ficou para outro mês`} style={{ fontSize: 13 }}>⚠️</span>}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            </div>
            {filtered.length === 0 && <div className="g-muted" style={{ marginTop: 10 }}>Nenhum serviço encontrado com esses filtros.</div>}
          </div>
        </>
      ) : (
      <>
      {/* KPIs de análise */}
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
        {bigKpi("Total Realizado (rateado)", fmt(totalRealizado), "var(--teal)", Wallet)}
        {bigKpi("Total Orçado", fmt(totalOrcadoBrl), "var(--text-dim)", Calculator)}
        {bigKpi("Saldo Disponível", fmt(totalDisponivel), totalDisponivel < 0 ? "var(--crit)" : "var(--ok)", Wallet)}
        {bigKpi("% do Orçamento Consumido", `${pctConsumido}%`, pctConsumido > 100 ? "var(--crit)" : "var(--warn)", AlertTriangle)}
        {bigKpi("Serviços sem Rateio Completo", semRateioCompleto, "var(--crit)", AlertTriangle)}
      </div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(1, 1fr)" }}>
        {bigKpi("CAPEX Utilizado (sem teto de orçamento)", fmt(capexRow?.realizado || 0), "var(--accent)", Wallet)}
      </div>

      {/* Custo por categoria — Orçado (US$/R$) × Realizado × Disponível */}
      <div className="g-panel">
        <div className="g-panel-head">
          <span className="g-flex" style={{ gap: 8 }}>
            <span className="g-btn ghost" onClick={() => setCategoriaTableCollapsed((c) => !c)} title={categoriaTableCollapsed ? "Expandir" : "Minimizar"}>
              {categoriaTableCollapsed ? <ChevronRight size={14} /> : <ChevronDown size={14} />}
            </span>
            <span className="g-panel-title">Custo por categoria — Orçado × Realizado × Disponível</span>
          </span>
          <span className="g-flex" style={{ fontSize: 11 }}>
            <span className="g-muted" style={{ fontFamily: "var(--mono)" }}>Câmbio US$→R$</span>
            <input type="number" step="0.01" value={exchangeRate} onChange={(e) => setExchangeRate(Number(e.target.value))}
              style={{ width: 64, background: "var(--panel-raised)", border: "1px solid var(--border)", color: "var(--text)", fontFamily: "var(--mono)", fontSize: 11, padding: "4px 6px", borderRadius: 3 }} />
          </span>
        </div>
        {!categoriaTableCollapsed && (
        <div className="g-table-wrap">
        <table className="g-table">
          <thead>
            <tr><th>Categoria</th><th>Orçado (US$)</th><th>Orçado (R$)</th><th>Realizado (R$)</th><th>Disponível (R$)</th><th>Ordem (Compra de Serviços)</th></tr>
          </thead>
          <tbody>
            {categoryCosts.map((c) => (
              <tr key={c.category}>
                <td>{c.category}</td>
                <td style={{ fontFamily: "var(--mono)" }}>{fmtBudgetUsd(c.orcadoUsd)}</td>
                <td style={{ fontFamily: "var(--mono)" }}>{fmtBudgetBrl(c.orcadoUsd, c.orcadoBrl)}</td>
                <td style={{ fontFamily: "var(--mono)" }}>{fmt(c.realizado)}</td>
                <td style={{ fontFamily: "var(--mono)", color: c.ilimitado ? "var(--text-faint)" : (c.disponivel < 0 ? "var(--crit)" : "var(--ok)"), fontWeight: c.ilimitado ? 400 : 700 }}>
                  {c.ilimitado ? "—" : fmt(c.disponivel)}
                </td>
                <td style={{ fontFamily: "var(--mono)", fontSize: 10.5 }}>{adpServicosLabel(c.category)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        </div>
        )}
      </div>

      {/* Tabela de serviços com rateio por categoria */}
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Serviços — rateio de custo por categoria</span></div>
        <div className="g-table-wrap">
        <table className="g-table">
          <thead>
            <tr>
              <th></th><th style={{ minWidth: 220 }}>Serviço</th><th>Empresa</th><th>Valor</th><th>PO</th>
              <th style={{ minWidth: 210 }}>Status de pagamento</th><th>Rateado</th><th>Previsão</th>
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
                    <td style={{ minWidth: 130 }}>
                      <input type="month" className="g-edit" value={r.previsaoMes || ""} onChange={(e) => updInv(i, "previsaoMes", e.target.value)} />
                    </td>
                  </tr>
                  {isOpen && (
                    <tr className="g-expand-row">
                      <td></td>
                      <td colSpan={7} style={{ padding: "10px 8px 16px 8px" }}>
                        <div className="g-panel-title" style={{ marginBottom: 8 }}>
                          Rateio por categoria — {fmt(alocado)} alocado de {fmt(total)}
                          {!completo && <span style={{ color: "var(--warn)", marginLeft: 8, fontWeight: 400, fontSize: 11 }}>(ainda não bate com o valor total)</span>}
                        </div>
                        <div className="g-flex" style={{ gap: 10, marginBottom: 4 }}>
                          <div style={{ minWidth: 180, flexShrink: 0 }} className="g-gantt-mini-label">Categoria</div>
                          <div style={{ width: 80, flexShrink: 0 }} className="g-gantt-mini-label">Ordem</div>
                          <div style={{ width: 110, flexShrink: 0 }} className="g-gantt-mini-label">Valor</div>
                        </div>
                        {allocationsOf(r).map((a, ai) => (
                          <div className="g-flex" key={ai} style={{ gap: 10, marginBottom: 6, alignItems: "center" }}>
                            <div style={{ minWidth: 180, flexShrink: 0 }}>
                              <select className="g-edit" value={a.category} onChange={(e) => updAllocation(i, r, ai, "category", e.target.value)}>
                                {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                              </select>
                            </div>
                            <div style={{ width: 80, flexShrink: 0, fontFamily: "var(--mono)", fontSize: 11.5, color: "var(--text-faint)" }}>
                              {adpServicosLabel(a.category)}
                            </div>
                            <input type="number" className="g-edit num" style={{ width: 110, flexShrink: 0 }} value={a.valor}
                              onChange={(e) => updAllocation(i, r, ai, "valor", Number(e.target.value))} />
                            <span className="g-btn ghost danger" onClick={() => remAllocation(i, r, ai)}><Trash2 size={13} /></span>
                          </div>
                        ))}
                        <div className="g-flex" style={{ gap: 8 }}>
                          <button className="g-btn" onClick={() => addAllocation(i, r)}><Plus size={13} />Adicionar categoria</button>
                          <button className="g-btn" onClick={() => marcarComoCapex(i, r)} title="Joga o valor total deste serviço inteiro pra CAPEX, sem precisar montar o rateio manualmente">
                            💰 Marcar 100% como CAPEX (sem ratear)
                          </button>
                        </div>

                        <div className="g-panel-title" style={{ marginTop: 16, marginBottom: 6 }}>Justificativa geral do serviço</div>
                        <textarea className="g-edit-wrap" rows={2} placeholder="Justificativa geral do rateio deste serviço (por que foi dividido dessa forma entre as categorias acima)..."
                          style={{ width: "100%", background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4 }}
                          value={r.justificativaGeral || ""} onChange={(e) => updInv(i, "justificativaGeral", e.target.value)} />
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
      )}
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
        <div className="g-alert" style={{ background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>
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
