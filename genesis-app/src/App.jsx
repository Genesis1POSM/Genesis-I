import React, { useState, useMemo, useRef } from "react";
import {
  LayoutGrid, Ship, Wrench, Package, Wallet, Calculator,
  Plus, Trash2, ChevronDown, ChevronUp, ChevronRight, AlertTriangle,
  Download, Upload, FileText, LogOut, Lock, User, X, Settings, DollarSign, Clock, ClipboardList, Gauge, Bell
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, LabelList, PieChart, Pie, Legend
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
    .g-bell-wrap { position: relative; flex-shrink: 0; }
    .g-bell { position: relative; display: flex; align-items: center; justify-content: center; width: 36px; height: 36px; border-radius: 9px; cursor: pointer; color: #3A4A66; border: 1px solid transparent; background: transparent; }
    .g-bell:hover { background: var(--navbar-alt); }
    .g-bell-badge { position: absolute; top: -4px; right: -8px; min-width: 14px; height: 14px; padding: 0 3px; border-radius: 8px; background: #E11D2E; color: #fff; font-size: 9px; font-weight: 700; display: flex; align-items: center; justify-content: center; border: 2px solid #fff; box-sizing: content-box; line-height: 1; }
    .g-bell-panel { position: absolute; right: 0; top: 44px; width: 560px; max-width: 92vw; background: #fff; border: 1px solid var(--border); border-radius: 12px; box-shadow: 0 18px 50px rgba(18,32,58,.22); z-index: 60; overflow: hidden; }
    .g-bell-head { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; padding: 12px 14px; border-bottom: 1px solid var(--border-soft); }
    .g-bell-list { max-height: 380px; overflow-y: auto; }
    .g-bell-item { display: grid; grid-template-columns: 104px 1fr; gap: 10px; padding: 9px 14px; border-bottom: 1px solid var(--border-soft); font-size: 12.5px; align-items: center; }
    .g-bell-days { white-space: nowrap; font-weight: 700; text-align: center; padding: 4px 6px; border-radius: 6px; font-size: 11.5px; }
    .g-bell-days.d1 { background: #FEE2E2; color: #B91C1C; } .g-bell-days.d2 { background: #FEF3C7; color: #B45309; } .g-bell-days.d3 { background: #E0E7FF; color: #3730A3; }
    .g-ro .g-btn.primary:not(.g-keep), .g-ro .g-btn.ghost.danger:not(.g-keep), .g-ro .g-edit-only { display: none !important; }
    .g-ro .g-table:not(.g-editavel) tbody input, .g-ro .g-table:not(.g-editavel) tbody select, .g-ro .g-table:not(.g-editavel) tbody textarea { pointer-events: none; }
    .g-alert-banner { margin: 10px 22px 0 22px; display: flex; align-items: center; gap: 12px; padding: 10px 14px; border-radius: 10px; background: #FFF7E6; border: 1px solid #F5C77A; color: #8A5300; font-size: 13px; }
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
    .g-mat-scroll { overflow: auto; max-height: calc(100vh - 230px); scrollbar-width: auto; }
    .g-mat-scroll::-webkit-scrollbar { height: 14px; width: 14px; }
    .g-mat-scroll::-webkit-scrollbar-thumb { background: #8a94a6; border-radius: 8px; border: 3px solid transparent; background-clip: content-box; }
    .g-mat-scroll::-webkit-scrollbar-track { background: rgba(128,128,128,.15); }
    .g-mat-scroll thead th { position: sticky; top: 0; z-index: 2; }
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
const LOGO_MARK = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFAAAACACAYAAACcL+6CAAALiElEQVR42u2cfUxV5x3Hv+fCgcOVlwUUcVocyizrBmvXzYyyrV2yaZb9IV1WNSHZHy4uaUGzbEtru6ylfyxlyZIttF1W1zZN5opgVLDqUGCAhYIgChcEfIH2gijceq8CejlwgLM/9F4P9557znnOOy/fhIQL95xcPvf38n2e58cFz/Ow3Vd1Nc9v384HHr/2T56/2M8P2PG1OmBHnToFJCcHH8YzQEs3NtnxpdoT4Ntvh/2oocuWr9SGAPv7K8R+PDYOjHqxYwWgnGpqXhD7MUMD7ZdRuQJQSf0TkTMWaO1bSWFp+XzZOHMm4q973QDLsisAI6q93SX1a/800OdmBlYAEqavMI3tZmdsBZAVsS9C0VFAi83qoH0A9vdXMAqe5rlrLztjH4AR7ItYGtvJztgH4IcfKn5qQ+dKCofbl05lVOgooMdGdsYeAOvrXaSX2MXO2APg6dNET7eTnbEcIMdxRPVPaGc4jlsBSLe18Wqu89wFBm/Sh1ZSuKlJ1WXOWKDtMgpWAB4+rPrSzuvLvQYODxcqtS9idXBwFPBNIHv5AmxufkfL5f5poOc6XMsX4IkTmi53xgLN3cs5hcvKtHXwKKDtynIF2NzM63Eb/zRwxY1Dyw/gyZO63MZqO2MdwOpq3W5lpZ2xBqAG+yJWB3vc1tkZawBqtC9issrOWAPwo490vZ0zFjjTvlxSWObsV20a97qt2Z0xH6DLZUiq+aeBnkF6YOkD1Dn6hGncM2D+Jqv5AEtKDLktHQWc7VjqNTDC6JouKcxOWnJmbC5AhWe/WtLY7DNjcwHKzL5oiT4A4LhJ00fgzANogH0RU9f1SVNXJeYBlBld0xp9AXHzCXDfMi+NHYs9fcO6sWPS1DNj8wDKjK7pEX0BnT4/ucRqoIH2RUye8Xum2RlzABpgXyJFX7DkmmRnzAFYXm6qtUhNijfNzhgP0OfLRnOzqdEHADUdt0wZgTMeoIrRNb3U1DXPL36AhKNrekRfII0/dc0t7ghkWRYs4eiaXprmpvDZ5SnDN1kNBch0dPCMBdE3zU0BAMb9c4aPwBmbwipH1/TUVbexZ8bGpvCBA5akbkBJzij8p/b2Ik3h4eFCBtZrcJQzdHfGOIAGnP2SRJ9QF/v8rsUHUOPomp4y0s4YApDjOM2ja3pFX5IzCp9dnlpcANVO3ku/Keq3qMb9c4aNwBmTwjqNrgW7+cwkaDqBOPqEOtXkKVg8AHUcXQMAJiZBUwTG0RzarjgWSQrrOLomjEC10TfDsYbaGf0BVlaabl+UqrbV57I/QBM3T6Wib4Zjg9EXSONzXdM2T2EDNk/VKoYOXwe1XdN/BE5fgAad/WqpfUJ57k3qPgKnL8CGBthZqfEJqL/g0fXMOFrXbllSAsbi6BOLPACYneNAOzi0XUmyaQor/NgSK2qfUK3XfbqeGesH0ODRNS21b3aOW5DG5zp8lfYDaPLZr+p1usOvq53RB6BJ9kVN9Inp0oBft1WJPgBNtC+kEqZvQCMTc7jYO+qyD8CKCstBydU+odYnRsE14LBRClt09qtG3Ow0uNlpHK6/bROA/f0VZnwIkxrvJyWPP1oXO6MdYFXVCwBgNMRx/xyxcRaLPqGO1d2otB7gJ588WokYCDLJGaWpcYQt65yzONd12+IU9vmyWRH7Yq+PiRWPPgBo/dyp2c5oAygxumY0xND9PpLoE+rCZW2brNoAHj0qvblg4ZpXLvoAIIm+h7Pn71qUwj5f9oSCs1/WwAjUGn0AUHPprqZNVvUA+/pcMQqfqgpiXJxhte+BLXrwqsa5eE0jcOoBEo6uEUPcskXQgaeCNkav2hf7sAQk0fdQd/5GgekAJw4cwAzhNUQQMzOD3+Zkxuta+4QRCADljT6TU3h4uDBYi4yCuHVrUeDbb2U68/WufUK57zCq7Yw6gCGja7pDzMsDHnvs3cDDtBRUbUqjiZZsUrVvWuQ+au2MOoA1NeFdUU+IRUVhP/rFDxyY4mjdal9APD+vyc5QPE84SOXzZU+kpER8t2IIXwAjApbheUrsuc8WDvFq1rxidS8UIkU5cOH9bIqmaYMj0OVySb3nWiORkejuf9+fPui+w+oWfQF4AOAaG1NlZ8gBKvivc9UQS0qAvDwq0vOeehyb9//8K6KpTNp5Q5Wzdi2O1/UWGA/w4egapzNE7NkDvPIKJfe03xWkUj/7XrwkRCXi+flg9AUeV7WQ10EygP39FV7B6JpeEJknnwQ++IBS+jLe3JtAbUqjNXVeigr/01uHx4k3WckAdnaGnf3qArGuLof0nf/HH1Yrvkas8wqjTyjSM2MigCOlpeJ1RwNEpq/vCJKTiT9KNjkR3Uf/vDl/ZGKOuPaJRR8AZKeuwbHGIYNS2OfLdrS0RC7eKiAy1dVAVtZOtXUsLQVVB/etRgCimtoXqvLmEaLdGeUAFXzqmlKIMwCY0lJg+3YKGrUtN5l6Od8Jjz9ade1b+K460Njh5XUHePvjj5UtoRRATNy3D9i3TzO8gF7auYH65dboMIgktS+gtbEOfHpJeTdWtBLhOA43YmJ4hmClEclkJOblAU1NusFbuFJx8VPT0ZK+Tw6gZ9yLjNQEtPzrGUq3CKQHBoKja0qtScRIPHEiBwbp+Fs5OWprHwBM3PeBiabQeu2OYjujLIVDRtfUQkxU2XFJOvOJvzyR7/FHk9c+AImrkoN18Fjt1UrdAI6InH2QQOQAJNbXa+q4JJ35/f2JGOfiiaNv4r4vWAfPufw6NREJ+6IUYkppKfDccxRM0k++n0a9nO8MQiSKPgBMNIWGnjFFm6zyd5YZXZODyO/dq2vHJe3MShqHMPoCGpueVzQCJwvwtszZrxRENjcXqw8eNB1eQH/97SbqGxscshEojD5SOyNrYz6nKMWfvBFqcVZNTVEMY+3oOcdxePrXXbxU9EXcZpvlMVolbfYdcrsvJC9WGIkpQ0NFVsMDAJqmcfZv383v9nypOPqCaTw+K2tnJAF6ysuJR9dmAKTU1y84FLJaaSmoqil+Ct1DY7K1L3RZJ2dnJAFygt1npRDXlZWZ2nFJOvO7L34dwkiUij6ldsYhZV9mQuyLHER6715g927bwQt25l1bqNd3ZKDb86V89D20M+UdtyTtjIPUvkSCOG9xx1WqN1/MonY9vQ7srPLTSCk7ExHgzffei9ydQreN0tOx/uTJHCwSffSnLCojNUERxLWxDhypu0FmY1iWxa24ONm7Mw9hZgwNFdmpaSjRqBc71hX8t3LjKmWjw9cO/1j0zFg0ApmODl7JP0OxADJs1nFJO7PbK7/77L4f+VPgRAF6amsVvYgMm3Zcks5c/lquIoiRzoxFAfoV/OPMquJiW3dcpdq5LZl6fddmuO/PSdbBSGfG4QCHhwvZIemTqZjcXKS+8caih6e0Mwc2WcXsTDjAtrbg6JpYHaTS05Ha2Lhk4AX07+JvSndmxiE6AhcGcODIkYUWJeT3X2tqKiKdYFoMomkax0ueyZdK46P/G5S2MSzL4gsJ+5LV13fEjF1lK1XbOsr/9Pft2JgWq8jOOELtS0R4ZWVY6vCCa+b9j4s2FbeXCzszdiixL8lLpOOSrpnDIDIOXOj1RU7hLzZu5EM7cOLzz+Orx44tG3hC7f5jJ9/QMwYm+tGfvy45fsGZsUPKvjDp6VhTXr4s4UXqzKEjcI8ANje/418mHVdrZ3ZdfTQCFwR4s6ICzsBKBEBWU9OiXOMasWY+9Oq3g8u9UDtD8TwPlmXRGxfHOx/C+051tS6TU0tJZ1t8/PZXW7AxhYbby2GmZhtF0/SDCAzYFz+ADcXFK/BEtC03mXrrVw87s2AEziG0L0m7di2pNa7eOrDnCeo3P1wDsPPBM2NqZmYGI5mZPL1+PVIbG6nl3DSUiOM4/Oildh7AAzvDDw0VdgA87/Vm8zyPlS/5L+84n41nT/LecT77/29cXcXFyQ3CAAAAAElFTkSuQmCC";
const LOGO_FULL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAUAAAACpCAYAAABarORUAAA+XklEQVR42u2dd1wUx/vHn9kr3MEdHKCAID2IKCpYUFBBVKJgCRq7KJaEqBgNikZjiy3WGFtUEmsssdeoWBCJiELsWJAiTaVLL8fd7vz+UPxRduGolu+8X697vZTdnZ2dmf3s80x5BoBAaGJwcHBnvGoVxl/odRu9Ezv+dRHjZ3G4HSkZQlPDJ0VAaHJ8fBRgZQXQsSMYawN++QYg7Chpi4SmhyJFQGhSBAKA8eOnQ1wcgKYYNMUAiWkAlibgiEjpEIgAEj5r97e0FEFc3HDg8d7/jYcAEMByBmMeKSECEUDC54wWnD0LZQKoUAAI+QB3YwAQQjQpHgIRQMLny/nz66CkRAoYv/8TgwFKlKAV9givIAVEIAJI+GyhN23SA4WCAlSxx09NALwrEdCBlBCBCCDhswRj3Abi4lqBQgHlBRABQKkC4PlL0MIYa5KSIhABJHx+7NrlzJPJWrMdUpQCGDQHx6vhMIwUFKGpIHOvCE30qaUA1q8XQ2Ii62EeDyA9B/j7LpI2SSAWIOFzc39p2gi6dJkKxcXsDREB5OYDWJvBBIyxHikxAhFAwufDqFECiI624hRIDMADADkNjofOg4QUGIEIIOHzACEAe/txEB5e7WlCAcCzRAAFBU6k0AhEAAmfif+LAYqL/VQ5jaYBFHJYSAqN0BSQDmdCE+gfVgdLS7lKDZICuBMNWggAMCk6ArEACZ88N25Mh9xcXZW8ZQDIk4PmvWd4Iik4AhFAwqfP/v12QNMCVb1lgQDUQx5Cf1JwBOICEz5t97ew0BDattWHnByVzkcIoKQEIOI5SDDGmgihPFKKBGIBEj5NgoL6gaVl79pcUloKYNQcet95Cq6kAAnEAiR8miAEsHEjhsePa3UZjwLIKQLR6RAg64IJxAIkfKLuL8NIoVWriZCeXjsBRACpWQBiDXDHGKuTkiQQASR8euzYoQlJSc61Fk4AAAaAwTD6YjjokIIkEBeY8OlRXNwRLl+u06UiAcCjWAARH5oBwEtSmARiARI+LYTC34CuW5R7hACUDIClEewkBUkgAkj49Fi/ni4f+r5WbjB+uzb4+n0yEEIgAkj4xMDPnk2EoiLTul6PAKBUroDXbxjt5FTcjZQogQgg4dNh796uQFFqdb2cwQwolKWgKaWaXf0PJpACJTQGZBCE0PDWH8Y8sLbWhIyMOqdBMzQwQMObXAZCH5doYox5ZNtMArEACR8/Dx70ByurYXUdAMEYoFQhBwoQFBWVgJ6OaGhiKriQgiUQC5Dw8ePjg+DlS0FdL6fpUigLhsWnlJCaTan9flxJokQTiAVI+OjdXzXo1m08pKTUOQ0FXfr+3zweBYmpxYBRfgeMMSIlTCACSPh4iYxUh8zM4XW5FKG3ri+DmYoWobIYZBrqi3NzQZsUMIG4wISPl0uXNCEwEMPbmSy1Ej+GYYBhlIDKXYoxBolICHeiS3hdZ+UrSAETiAVI+HiRyTZBXl6dXFWaoYHBNKs45hRgtHSk2p+kgAlEAAkfLfT+/aZ1Wf3BYAylihJWwxFjDFJ1Edx4xLQmJUwgAkj4KMFRUV9BXJwRMEytr1XSimrEkYFieRG8SC1pnp2N7UhJE4gAEj4+zp935olEzetkOdLKagSQBqVSAZqaWoZhj5lhpKAJDQUZBCE0jPWHsUBhYCARpKU1dLpA00rg8RDkFWL4459MMdkyk0AsQMLHRWZme4GT05CGTpZm3lqGCADyCwrBzEjdMzcPtyEFTiACSPh4mDFDDLGxzRsySQy4wqgwAhryikUWv58o0CMFTiACSPhY3F8+yGSeEBlZ+4sFb9f+Yqg6cELTCsDlRpRFQgQJrwuhmGbIqhACEUDCR4MQBILZdVG/YgUAw1Sd/MJgpsqKEACAErkceKD4CQDUSLETiAASPjznzqnD6dPy2l+oeDcKhytYegBvV4VU/hvGAGo8JdyJLtU+dgyIBUggAkj4CMjLmwtpaXWwyASgVAIoaaaySw00o2BxtRlgsALeFFAU0s2ZQgqeQASQ8OH5++8edd37gw0li/jRDA1KWgkACERCHu/u05IRpOAJRAAJHxSMcSt48EAXSksbKj3ALH1/FELv/65QKCE0ihYXFGADUgMEIoCED8eRI6OhefNWDZUczSir9P2VWYBlKJSloNdMo8PtR8XjSQUQ6gNZCUKoj7WG5ObmoFaP4KdlIHjXx8fQVSxCmlYAU04UKaChqAjD/mt5AlILBGIBEj4UX6g5Ok4GubxBEmMYGnClRW4IIUDo/xe/YcAAiIG0NwVACVBfjHEzUg0EIoCEpmfBAi2IjzduEGsSMCgZJZuVWcH9fdtHiIEPpSDiq/f642SmOakIAhFAQhO3HAogM7Mt/Pdf/d1fxB0Oi206DAACAR8g8kUOpOfK2yAyI5BABJDQlGCa5oOe3qa6bn3JZulV/VvFydBvQ+bT70Uzt7AUpGK0iWEw6csmEAEkNCEI0bB7d6NGpaIZuqIAlguMwDAY1AU0XLtfhIFsmE4gAkhoUkJDf1Dm5Gg2RFIKpaLKul+GoauMCFewDuHtBkoxqYwo7EGxK6kQAhFAQtPx999ufISoUgCoj/nF5wMoaAw0gyu4w+X7BDHg93EB31uHtBIAMEjURaJLt5L9SYUQ6tT+SBEQagvGWFhqYiISFhcDAwAMvJ2kUpfGpFQCCPkIeBSq7n4VXOHyo8KIUcCVSKUIYyxECJWS2iEQC5DQuFy4MJ6vr99DXm7zI+W7X31hGBqUdGkF66+iK4wBl+sLlMvlYNBct8eNu2+8ScUQiAVIaORPJgXF69bpiOPiBJVHQJQN0KjKT3ous/4qCmTFkWEBD8ObN/mCTcdKhKRyCMQCJDSu+0vTJsLWraeUZmezu7SqWoI8HoBY/E7kKlqAFf/PVBDDyhunM4wSUt6UQHM9bXeMsYzUEIEIIKHxWLtWzHvxwry6gQ+VRNDQEMDZhbLSLuaJRAAKJQalsrRC/97bf1e0BitbhIiigMIlQCsUAw6de02iwxCIC0xoJBCCwvh4K+ryZaBraDw1usMCQRIMm5jxDYBy1uOCGEB8K4SU5fSuqtjRLEvllDQNQj5AeHQx5CuRhFQSgViAhEYTQDVLyxVlI781WXmclqCmJkBR0T8AEMejUHxmdskFiUTtXcDT/7f2ylt/DKOsYg0qlKWAMQMIAIrkAL3aiDZQZFkcgQggobEEkD53zvC9KNVRBBkNjTyYNu0OQohhMMAQJ/SIB4oCQLz31l/lAAiVJ0q/DY6K3p0NIOaVQOiDN60ZsmM6gQggoTHAd+58xzx8KCkvRbUWQYoCpn//F8jTc0/Zn4a6Ndttb8XclzO8d9ZeVbGrMhqMmXIDIggAGLgRnSlKTsFdSU0RiAASGp7z5/sKBQJxZSOLgZpXgygB3i77GDCggL90qU/l4+P6SL9vZ4YKGCSqYO1hjFmXxFW0CDFghgE9LZn05LUYslkSgQggoYGtP4y1S0+f1irNzWW1+OgaLEE+jweMjc1rsLPri0xMqsTQsrZQe2hrot7XqLn6a0SpVbT+KvX9MeWCJJSNDCMEUFoKsOfKaxnGmLRrAhFAQgMSE9MbNDVdihQKwBwWH5c7zAMALBTmUQ4OE9Dy5eFct/AboxXeyoieoCZEefhd06xq/eH3gyXl3WIEAPnFRaCmJXVKTIRepMIIRAAJDUaKn58EoqOFNbm9bCIo0NYG3sqVB9C+fVdqus9KH9kVt87C2SAUQ6mi4sjvO2f3vUVYMVQ+gIiHQAQSvcV7HhiSGiMQASQ0lPurIdXRGVT88qVKFl+ZOGIAEOnpAcyaFU75+c0HZc1rRBgMMM+r+VHPbigcqYkrRIlhGLrS6HDFwRIBj4Lol9mQVljaE2OsQWqOQASQUG+Kr17VFTDM12xbH2EOEaQBQKShAeDmFgbOzkMQQnmq3g8hlNe+ZcmQjuZUmJIvfb807m3/H1PF/S2fFwqXgqWhzOfe0xSyWRKBCCCh/uTt3y/LOnqU20JkEUEhADBWVo/AwcEdubjUet9MdxezlDYtXrlrqxVEUnzR24ku5Sw+9n1AMIj4ANfvpsFvJ9NJ2yYQASTUn2Y9eqwTKBRAA/d0l/IDIwgAcKtWebx588ajmTPz6nrfmeO65S0a1XK4ugiiipXovcvLtib47d8BKFBCQlYxtDfXmElqjkAEkFBvsk+csELlXFuG47yyPkFpp06gtmrVNhg58kl97/1lT63n3w3S+NPUWAfkyoqhsiojLy0GBmOQiHjwODp/CKk5AhFAQr3A+fnt6du3NcqLnpLDEkQAIDUxARg16ndq6NBFCKF6x0ilMcCovvpb+7alf2/WvDnQDPtiXwYzQFEUKGkliCkMN6NzRKlv5O1JDRKIABLqTPHhwzM1NTX1KisZmyUokskgq3PnADRnzvSGEL/3wopQ6Q+jjaa30c8PkEnVq4bEQggYhoZiedHbvDE0qGlo6AWHpxI3mEAEkFBH4aEoyFiyhC5NSQE2u0tZTgTFACAXCP68cfKkb6O54vcf+1JI/gdDCctZnQhoWgmlipL3AyM8CoG8sBRWHYpVkFokEAEk1AkmPb2bqFWrwYXVzN9TwtuVHoybW5J2evrKEfXbJK5ajh0bQf/7e4eFBnolsQWlABRCb6dFvxscKZsigwAgr0gOSE2tI8bYlNQkgQggodakLFpkKMjI0C8B7oEPBABq/fuDZNWqxQihxEa3ShHKWDHGeIWlEYbsYgVQCEGpUg4KZem7lSFl8wEVoCWVdFm9K6ozqUkCEUBCrcAYIyopyVz+/DkgAGDb/5cHAGo2NjTq02cS6tx5X1Plzamj4T53O+kkg+Z8JiM/HxiGfi9+ZYgFFDxPzIHQhEIzjDEJk0ogAkioBfn5zdTMzNYUlXN/FeVEkAIAvo5OaYm19TTxnDl7mjp7s8fb7GnVXDhZS0NQrKAZFksRoKi4CKyaC9fkp+STVSEETg+GQKjCfQCZrHnzNzgjo0obEcK7vT4mTtyqv3v3DITQB4vDPH7Zw6Xn7mQs1uSzxAxkAJpp8bGGolg39NDAbFKrBGIBElSi9ZEjeyErq4r4YQBAIhGo+/nl6O/evfxDih8AwL5F7Td59dbLKVLwK+6vCQA8CiAum0Yzvuk2g9QogQggQWWyTp2yZFtwqyYWg9aIEUnSb77pjRBK/+AuDEJvvvdo17ufg2FSARZUEUE+xYPLYYnDSI0SiAASVAJjbFocFCRh6IpupQAA+GZm6XmdOg1Fbdve/1jy26oVut/FGIa2NtJMV5bbiBMDgAaF4PydNxTGWExqlkAEkFAjyn/++UYqk5mUH/mlAIDR15eXGhn1NZw58+7HlueZE2zuWjcT9NGWCAvkSvy+c5thFCBrptnqWkTWbFKzBCKAhBqJXbyYr0xNpQD+f90vT0+vVNfff4lJaGjkx5rvv1bYPR7/ZYvVero6IFe83S+YhzCAnOavP/iEbJpOIAJIqMH9ffOmvZ6h4ZDC/Pz3bqSkWTOQTZ9+TcPffxuUlHy8lisN4DvUcO9AB2kIUpcCjTFQFIKMNwXAU5d6Jifnk+AIhArwSREQypO1d68er7DQusz1laqrA3h63pQsWjQMIVT4sedfQ0PjFX71yuN1VunlwIe4O0UXAgYaSjHfOvBWuj6pYQKxAAns8HigCAszzrt1CxAAiHk8wJaW1wVGRh6fgviVgYyMikylYo9WzQUhGIlAIqTgeXwOnLufo0uRma8EAoHV/cVYnDN//ovHADgGACe1avVafubMJ+s2nruea9Xl27AXRp5Xsbr7Rbxkb2wUGQ0mEAisDAfgvejYUfkEAKd27owVV66sBv6n20tCIYBz11+ucpx+D8v6X8RuM0OVAMN5pKYJBEIV5OfOffsQQJloZYULdu3a+Lk8164zSRvbTb6FeW6ByiOXk78lNU0gEKqQtXz53ZcmJjjBy2vj5xRBBWOMfFbf32jm9R9eGvDsHqlpAoFQWSS0YszNY1/16rX7c3w+BAB9vr+9p53Xv08xxgJS4wQCobwAGjy2tPw249kz6ef6jNGZWFPmcckLYywiNU4gEAgEAoFAIBAIBAKBQCAQCAQCgUAgEAgEAoFAIBAIBAKBQCAQCAQCgUAgEAgEAoFAIBAIBAKBQCAQCAQCgUAgEAgEAoFAIBAIBAKBQCAQCAQCgUAgEAgEAoFAIBAIBAKBQCAQCAQCgUAgEAgElUCkCAifExhjCgB4bMd4PJ6CYRhSSAQigITPkxMnTgxLTk7e9fr1a0Do/5s3RVHg4OBw3tPTcxxCiCYlRSAQPjvmz58/pm3bthgAqvzWr19PY4wlpJQI7z+MDZIIRQHGWJSWlibZunWrxMXFRQIA738uLi6SrVu3StLS0iQYY3U+n/9hzF2EAGOslpaWJpk2bZqkefPmFfJ49OhRCcZYIhAIGsIVQxhj0UfoIvJSUlI0MMafpfWvpqaGq6k/BXnlCeXh1/NlMt6zZ0/rgIAA3SVLlixISkqyiomJgYSEhDKxAQCAqKgoOHjwIISHh4O1tXX2oEGDltvZ2cUsXrw4BSH0uNFVnqLgxo0bXVavXm2+dOnSn+Li4lo/evQIioqKAACAx+NBbGws7Ny5Ex4/fgwjR45c6uTk9N/UqVPvI4Sy6nLPiIiIBadOneqFEOpbVg4fA/v27euDEDpRUlLSCgBSGiLN06dPd46KitKm6ab3LNXU1AAAMv39/e+T15nQJAKIMW7n4+Mzrnfv3n0oiuoYHh4O4eHhnOenpaVBWloa3Lp1CwDAAAB+f/XqFdy/f//V5MmTD23duvUPsVgc2xgPGB0dbTlv3rwZixYtmnjnzh3puXPnqpxD0zS8evUKXr16BZcvXwYA+CUpKQmCgoKuBQYGnh44cOAWpVJZq/umpqb2O378uM3mzZt7ff/999c/EutPNnr06NkURUnGjBnTYKr86tWr3y9cuOBQWlra5M8kEAhAIpEEAUBf8joTGpXHjx/r+Pr6Lra1tX2lr6/P2s9S25++vj42NjZOmjNnzs6QkJAW5Tuu6+vqHTx4MMDY2DhRS0urTnnj8XjYysoKjxkz5lBqaqp+bfK2ffv2IKlUir29vWMwxgYfQ/2NGTPGzsDAAHfr1g0vWbKkwfK0bNmyEB6P1yDtoS4/MzOzf8ry8vPPP4+2s7Pj6gMsIX2AhLqIiaGXl9fXvXr1yqUoqlEaMY/Hwz169GC8vLx8MMbN65lfvo+Pzy4zM7MGyRtFUXjAgAHMqVOnvlc1D97e3kECgQBraWnh2bNn+1EU9aHr8IsBAwakAADW0NDAAKDbUGmvWLEiRCwWfxDx4/P52NramgggoXFISkrS8ff3v2Vubt4kDVpfXx9Pmzbt3xs3bnSoa559fX13tmjRosHz5uzsjMPCwlQSwW7dul0VCAQYAHDXrl1LMMbqH7IeFy1atFFbW/v9swQEBLgTAST8r1OtWTJz5kzZiBEjLm/fvr1bfHy8yokihEBNTQ3EYjGoqalBbayftLQ02LNnT8+VK1deXbNmjXVtH2j//v3rLl26NDklpeb+fXV1ddDW1gaZTFbWmV4tcrk8EQAuqmBtCaOiojQUireDjnfu3BFMmTLlzw/18h0+fNj12rVrk3JycgAAgM/ng6Gh4aLPphF/YOua8OnC5xKwn376yfzEiRMnYmNj7VUZANDU1AR7e3ugaTo+MjIyzNDQUKmtrQ15eXmQmZmpZmtr20+pVGrfvHkTyoSBi+LiYrh06VKzN2/eBI8YMcL96NGjD1V5mLt3734xfvz4XrGx1Y+nODg4gEKhCJVKpXG2trZQUlICt27dAqlU2jU3N7f18+fPq1zzxRdfwKBBg75zcnKqcbCmsLDQzdTU1LpMcGiapiIiIsb8+eef1wHgzyauY7Rhw4bhERER0vJ/TE5O1mqoG9A0LaqpThsLpVIJ6enpAvIqExpMAIcNG6b1999/n05ISGhf3dIhhBBoaWmBm5tbzOvXr1f7+vrKhw8f/lggEDx89uzZ+/MEAgFcuXLF+eTJk8Yikailvr7+opMnT2rk5+dXZ0XBvXv3WlAUdWrlypWDFyxYUON0mS1btvSSy+WduY4bGBiAmZnZYS8vr7O+vr5XEEKZ//77b/l7tvb19e3k6Ojof/z4cbuCgoL3lqK9vf2xcePGRSxcuLDGQn306FEHDQ0N7fJ/e/DgAZw7d84nLi7urKWlZVpTVXBgYODsqVOnTi3/N4Zh4MGDBw22JszAwGDrd999N4CmaVVGiRQJCQn58fHxU9g+NAAAkyZNilNXV/9PqVTWaNohhHgKheK/nTt3kreZ0CAvjI6bm9t/NY3qaWho4H79+kXPmTOnT05Ojo6qI6RCoRDi4+MNDh48+JOzs/NrTU3Nau8jEAjw8OHDM0pKSlrX4HZqeHt7nwXuQRbay8vrEMZYqIILq+Xr6zukS5cuOQKBAHfu3DktKyurbS3c8HkdOnSokgdNTU08Y8YM76aqy9jYWL2ePXvGAMugTq9evSIbzMR8O+dTqOJPcP36dfs+ffpw1vny5cv71yI9NYwxn/QBEuoNxths/vz5N8t3lrP92rdvj4cOHbqgvqsJ0tPTW/n5+Z0xMDCoUQSnT59+s7q0oqKizH18fDjTMDU1Lahtfvfv399n0qRJadu3b/evzXXjxo2bZ2RkxJoPJyengri4OLem6BebPn365ubNm1fJA0IIN2/e/AnGmPch2tmmTZs6ubq6ctaVh4eHU13TJgJIqDO+vr5fW1lZVStG3bt3x1u2bJnfgPP1xG5ubqNqmleopaX1ZtmyZaO47rtnzx4zDw8PzutnzJhRgDFWq23+wsPDO9f2GolEMrdsBLjyTyQSYU9Pz1ON2umHEOzdu7dTq1atErnKQyQSxWKMtT9EO9uyZUvn6gTQ09OzBxFAQlPwvo/lzZs3WnFxcbtiYmI4T+7SpQt88803P86cOXNVQy3vQggVX7ly5bC1tfXX6urquVznFRYWat+7d28lwzBGbMdzcnJALpdXKwrvXoRa0bVr1zu1FHR+y5YtW3INCpSUlMCTJ088Nm/ePLmx1kQzDEM9ffr0cnx8vAnXOebm5uYPHz6cTV4BAhFAANi5c+f0S5cucY4M6uvrg7u7+9yJEyeubYyYahEREScnT558WVNTk/W4UqmE0NBQi3Xr1k1ic2XNzMyA61oAgLNnzyLgiBPXwFi4uLiMrO6EmJgY4ZUrV2YrFAqNxsjA9u3bJ/3xxx/a1Y3MKhQK6unTp8QaIhABxBgL/vzzzzFcJ/H5fGjduvW2ZcuWrWusjJSUlMCKFSsW9ejRI4nrnMzMTHj8+PEstmP9+vXDRkZGnKvxk5OTRVOnTl3P4zWuBiYlJWkCgF5N54WEhNjs3r17T0PfH2PM/+OPP2bn5uZW20dRWFgI9+/fx+QVeDswhzEWJyUlicPCwsRJSUlijLFYlQGzetSTsOyeR48eFYeFhb2/Z1PPaxQIBIAxFlV+/tq8K2URoTDG4rCwsPLPo9ZI5ScoX2fl7ieqTTQnPgDAjh075peWlrbmcmttbGxez5gxY39ISEijVoSWltZzPz+/kQ8ePLj1+vVr1nOOHDmiPmLEiF4AEFz+72KxOOnhw4c/CwSC5WyWD03TVFBQ0DR/f39m9erVMxBCjfLyh4aG4uTk5BrPy8vLg7lz5/Y6fvy4xbBhw140UKPQnDVr1q8JCQmtauqiyMvLg6CgoP9pAcQYa8+YMaO9VCr9asGCBd/Fx8dDdnY2iMViMDU1BSMjo5SvvvrKb+XKldnt2rULrW+3D8ZY7dChQ91Wrlwp/vXXXzempKQYJyYmQl5eHqirq0PLli3ByMgo1dPT8+cBAwYkTZo0KRohVKuIPadPn5ZmZWV1zMzMhPJdLCUlJdCqVauSYcOGPUAIycvys3jx4m75+fndlyxZ8mNcXBw/IyMDdHR0wN7eHoYNG+bn4eFxw9vb+1l1Qr5582bHa9euGS1btuz32NhYYWZmJgiFQjA1NQVTU9Oszp07j7t+/Xq0RCJJqWf5iQ8fPtx527Zt4jVr1qzIzc1tW1ZnIpEIDA0NwcTEROnr67sGIXRzw4YNSQih+BrV39bW9jc1NTXWjmNNTU3cr1+/35uwUWq5u7sHAkcHuUQiwT///HMw27X79u2b5OjoWO1giqmpKXZwcNj95MmTkY2R/9mzZ3fq3LmzSsu4pFIpnjRpUiDGWNoQ9968ebODs7OzSvdGCGGpVLqlsS1iNj7wIIiIx+PBjz/+ONXR0fEK23Sl8j8dHR08atQo7O3t/UN94kRu3Lhxsp2d3eF+/fphoVBY4/K+Pn364AEDBtzevHnzktr0FQcEBHTx9fXFTk5OuEePHu9/jo6OeM2aNQqMsTkAwJo1a0Y6OTkdcXBwwO+MAdbZF/b29sn+/v5zKy/lRAjBr7/+6tWlS5fDzs7OmM/ncz6Pnp4e7ty5863ffvttVl3KTiQSwYkTJ74ZOHDgib59+2KZTFbj2n07Ozs8aNCgJ8uWLVuCMTbmTDwuLq7L4MGDM7kSMzAwyIuKivJsyhdkypQp33CNRiOEsKGhYSibaY0xbjZ27NizXCOwUG4k1t7evnj16tXno6OjLRt4Okh7ro8J28/CwqJ45cqVw+o7pQhjjMaOHXsHarGO1s3N7SHG2OJ/RQC3bNmSFxkZ2W/IkCH/1TaakZaWFh4xYsQujHGt+m3/+eefrkOHDo0wNDSkoQ5rnS0sLPDXX38d/uDBg6GqzLxYv359p+7du7OmNXfuXFxQUODx7bff3tLT0ytWNbBJixYtcEBAwJ5yVqb1uHHjIgwMDEq5xJPtZ2dnh1etWrWs/LzNmvjrr780xo8f/4+5uTldnchWZ/CMGDEidvv27WasN5g0aVKfdu3acSbQu3fv5039giQmJn7p5OSUzpWnLl26KF68eDGeQwh07ezsUlQpHKFQiF1dXekFCxbsO3/+/MD6RoJGCMGQIUOG1LaSRowYQdc3WMKSJUsGq6urM7W5r5eXF379+nX//xUB9PHxUfbu3Zupa+gubW1tPGbMGD9VPlYYY4Gjo2MfOzu7grq8uFApUpKTkxP+8ccfp6oigC4uLlztTDlmzBi6JgOB7aemppYXGBg4cOHChV42NjZZdX2mbt264T/++GNETWJOURSEhIR4d+nSJbq+5UdRFLa1tU1s3bp1qyo36tmz59cikYjzwm3btsU1+cgMRYFMJjvG9YXq0KED/uOPP+ZwXR8eHj61b9++tYoo4u7ujidMmLA7Pj6+ztYuxhgFBAQ8rkslzZ49e3Fd7xsfH9/L3d09tbYvdr9+/fDJkyfd/1cEsCF+bdq0UWZkZEhr6qtasWLFGm1tbUVD3rtVq1bY29t7cl0FEOoZdcfDwwNbWFjUO62hQ4ferMmSvnv3rk+vXr0aLP8IIWxhYRE1dOjQCivK+O7u7t43btzgqkjQ1dXd29QvCMMwMGrUKObQoUOQl5dX5XhaWhoEBgZyzsVxcnLaHhYWhrOzsxdHRka2qClSsVKphIsXL4JMJpsYERExcsOGDTHt2rUb6u7u/qKWkaBRdna2MZd1yNWJjhCCS5cu+cfFxV20tLT8r7aiO3jw4H63b9/W5wpJz3XvrKwsiIyMhP9V+Hw+6OjoAEVR73eQy87OhpKSkuq6jKgDBw5sAYAJXOeMHDnSNTAwcC5b2y1DV1cXlEplaX5+fpZMJoO8vDzQ0tLSUyqVvNxc9umw0dHRwOPxdixYsID/yy+/BNR3UIbP54Ourm7ZCC5gjCEjIwPYprkplUq4cOECazoSiQQ0NTWBYRhACEFJSQlkZ2dz3vf69etOfn5+ZgDwhO14ZmamZrdu3fy5gpqIRCLg8XhF7du3z3VzcwMtLS2Qy+UQHBwM4eHhWnK5XL3ynGCMMSQnJ1u3adPmalRUVM/WrVvHl30xalJPww/ROPfs2XNUT0+P03XV1dWdrYI4NB80aND52gZG5fF4uFu3bnm//fbb4tqsHMAYU97e3llsaYrFYtynTx/MFZ1aLBbjkSNHnk9LS1M5UjNCCE6dOtWvU6dOnM/yxRdfYCcnJ9ZjzZo1w3Z2dv+TFqCdnR02Nzc/HBAQcOj+/fu/PH78eN3Zs2cPDR8+/J/quoT4fD5u27ZtEJcLt3///m4WFha5UM0a+s6dO6euWrXq0P3794eXv/bZs2c+U6dOPdSuXbus6vrWXFxc8PXr113rYwE6ODjgLl26nD1w4MChyMjI9ZGRkb/s27fvUNu2bS/o6Oio9J7o6upiV1fXkpEjRx4KDAw89OzZs6URERFbli1bdrh169ZZ1bmkBw8efMSW/7S0NIPRo0f/wxVfUkdHhxk4cODJ4ODgvggh4PF4738IIXj+/HnfsWPHhhkaGnK58njNmjW3yvcd1fSgxh9CAM+cOXO0ho7q2SqKksbatWuXOzs759c2mrW5uTn28vI6p+oABcYYde/enbXi1dXV8fjx4/fq6+tzLk+ztLTEY8aMUXlvCx6PB+PGjbtV3TMsWbIkdPz48dV1wH/5vySAurq6uFOnTmc2bNgwlm2+HcaYOn78+CZra2tlNa7oK4xxV5Zr9b7++us7XINgBgYGeNCgQceePXtW7VrnW7du9fL29n5RXb16enqGsIlwTQIokUiwh4dHxI4dOyZqaGiwPT+/TZs2K2rqJ7S3t8c//PDDtoSEhAFsz3Dnzp0x7du3T+cSwO+//z6Sw3r+0tLSkvWeLVu2xG5ubpNrit2JMdZbv379dS79sLW1zXz69OlQVQXQ6EMIYGBg4NEagiSovIxLIBBAcHCw2dq1awNtbGyKarNHiIaGBraxsbkYGhpa41SV4OBgvp6eXhbXaNTYsWO9AwICTkilUs77ffnll8kYYz1Vnmv37t2bub507zqcMx48eNB72rRpnC/z4sWLzzXmhN+PSQD19PSYX3/99QzGWFzTh+ynn346xRWpyNnZGUdERPhX7rf28/ObbGxsrORqR5MnTz6m6lapGOM2I0aMeMglpsbGxvjChQvrayOAmpqaeMOGDfcwxrLq7v3gwQNbd3f3PK76sbGxeR0UFDQUY0xV93GePXt2AJsIURSFXV1dI9nE19vbO4mj7nCfPn1UnkaDMZZOmDDhNteH6Ntvv10LAECpq1c/+Ni1a9cPEso9JycHuJbcIYSgNnOjFAoFuLq6JsydO7f/6dOnxw4dOnR/z549Vbq2sLAQXrx40X/KlCn7vby8qu24tbKyapebm8vawHV0dMDZ2VnDx8dnpJ2d3Z1q+kdaLl++fFlN+Tp69OgXK1eu7Mk1YbxFixZgYWHxS4cOHUJ0dXWLudIRCoUeAPDR7V/cGPj6+pbMmjVrBEKouIauBfzdd99toCiqiO14aWkpvHnzpkIHHE3TOtnZ2d8nJyezTqlyc3N78ddffw1HCJWo2L3xdMqUKXPbtGmTwmbppaenw7p169rWpotm4sSJ4OfnNxIhlFPDO//43r17u7iOz5o1y79Pnz4nEUKc/fA0TYO3t3eCTCZj2MYW2NrtlStXXC5cuNCSTUzt7e2jT5w4ofJKDIRQvo6OzvgWLVpUOZaamgrR0dFOGGMjisfjPayuj8nf3/+7D9FY4+PjgWsAQiqVgq2tbZ3Stba2PrV///7xEyZMGO7r63vZ1tYWapoMLJfLITk5+SuZTPZ3dcIbExMzWUNDg9W60NDQAAsLCz5CSOnu7r65TZs2nGJ9+vTpkVevXvWsrq/x5s2b83Jycuy4vr5du3Z9tHHjxn8AABUXF/9bTZ7prKys/wX9Ay0tLQAAleY6mZqa3jAxMWEdxWAYpkpUcycnJ2lQUBDrPjatW7eGfv36La/tGvrevXtf+vbbbyPZXD65XA7FxcU9g4ODXVRN7106NVoOcrkc8vLy8rneiwcPHqi0P0a7du1OpKens67+KNuTu9IA00/5+fmIbcDI0dHxqkwmu1ub8vv111+zPT09aY4BoO7btm1rS4WGhm6ubu1hdHT04A/RWP/66y9lWURmNuvmq6++qvPEYaVSCZMnTz6+devWrxYuXNhu4sSJj3R0dKr9Mufm5sLx48cHBQYGenGdExsbawYAiKPxMaampqUAAPPnzz9iZWV1RiQSsYkbREZGyvz8/NyXLFnCWjHHjx93vXTpkjeXcBkbG4O/v/8iPT29GACgz7FthvyO5ORkuHjx4v+EAKoSYbr8R6R9+/YqtzEfH5/5XEsg7ezs0r28vO7UZeP4tm3bTtDT0yvhMBI0li9frqdqaLrajBqLRCLEla5MJhOrmEy6XC4v5rIQK3Pw4EERh9FTMnv27BiMcbva/ADAsKio6C+2BCMjI+Hw4cPALy0tjTc1NQWuTY/27dvHNHVDTUpKavfll1+24Zq+oq2tDQMGDChaunRpve7zzh15DAAdVq5cOSosLOzA+fPnOc3BrKwsWLx4sZdQKDzAlrenT58quSKwJCcnx1hZWV17d9/SmJiYg+np6e63bt0SslmBxcXFPubm5i94PN6ayo3lxIkT+6Kioji/5E5OTnucnZ3PlrlzPB4vk6Io1i6FFy9ewM6dOxkgVMHISPXu7+jo6P4cQgKhoaHbpFLp47rkwdnZOXvUqFG8tWvXVjmWlpYGLi4u3RiGOYAQarJNWRiGUVVJeUhFdcYYtzAyMtJmE8Di4mLR4sWLf6ttDFI+nw9cH6X09HT4+uuvu/Fnz56NNm3ahOPj4xHbFyM7O7vlmTNnpn711Vfbm6qAAwMDu8lkMs5tMRMTEx916dKlQYOKLliw4PC9e/coiURy8NixY6xioVAooKCgwCEwMPCr3r17n6l8PDg4GHPNIXv9+nWyVCqNLtdfeGzChAk6//333w62So+NjYUDBw7MVSqVG8sWryOEYNGiRX5r1qzR4eqy8PDwAC8vr42HDh16/3dHR0cUHR0N6enpVa559eoV9eOPP7YICQmJIZJXsSzZLHQu7t69y2q6GBoawjfffMP/6aef6poVhVQqvQAAX7EdtLCwGA0APwNAyqdc3m/evHHW0dExZ+sbzM/Ph99++63B72lgYGBG9e7d+9+4uLg9XG5wUVGRxoEDBzwwxjpNURAYY9HOnTstwsPDWY+rq6uDp6dnTm2jZKhCx44dD3l7e4/r0KFDdjUVpb1582Zjthfm4cOHFJsFyOfzwcTEpLiyS79o0aLTgwcPjucq+7CwMJ2FCxceKPvyJSQkWF68eHGGQqFgdUE0NTWLbGxsZri7u1eYYOrm5sZpzVAUxbO0tJxKJK9+LmNaWhpne7W2tq5zHvh8Ph0WFnaS63hBQYH01atX/E+9rGNiYqzFYnGTDsbRNE1TCCHlkCFDMlu2bMl6UmFhISQnJw/cunWrU1NkasaMGc0wxvO4Gp9CoSgdMGDAnsa6v4eHxwEnJ6dQtjlSAAAvX76EkJAQtcodxAzD2NrY2LTlMsUtLS2r/N3S0jKte/fuP3G9IEVFRbBr1y6HXbt2dcQYU2vWrFkXFRVlxmad8ng8cHV1jVm3bt3vCKEKPnO7du1AW1ubSwDh9evXPYDQaGJZn8jfGGNITU0trMYdhXdLWT9plEol+hD3pQAA3N3dtxUVFXHGpLt9+zZERESswRhrNXIjUsvMzNzx33/cq8H69OlTPHDgwL2NmY9x48bdfrcAm+v4LKVSWcEKjIuL69GyZUtLrhfA3NycNa1Zs2Zd6t69+4V3IZLYzH+T+/fv/xkcHNzl/PnzQ7i2EtXX14dBgwZ9xzY1wdramjNaNsMw8OLFCxoI9YKrfDHGUNNSzJpccRMTExnXcTU1NRCLxZ98+UkkElzdYGz5FR8N8XvnVQn5AADdu3dPnD9/fuiqVas4QyMdP368jZWV1SwAWNIYBcDn82Hnzp0bQ0NDB3Cdo6OjA66urhuCgoKgMTfivnjx4trc3Fx/ANDmqCxDANCo9JHgXMMpEAigffv2XA08G2M8+ObNm/nPnj0Ts1ngly9f7pidnX0zMTGRNQ0tLS2YPn16xKRJk+5Mnlx1rXybNm1eyuXyDABozuIGwPXr14mC1RNLS0sIDQ2t8vesrCw4e/ZsnT8wSqVSbd26df5nzpzhcuNuqqur533q5dehQ4dr6enpvsASTV1dXb3E0NAwpKSkpEGsRIqiICsrCzDG99/b5gMGDJh/7dq18Vx9b8XFxRAQEDB9xIgRYUePHr3U0AUwfvx42zlz5nyZk5PDeY6FhUWEu7t7wI8//tiolbF06dKy2fwquztPnz7FXALI4/HAyMiIV83XjV68ePFvmzZt+oltEfnz588hOjqa83pDQ8MH9vb2npVd33I8i4+PjwSA3mwCeO/ePRIav544Ozvz9+3bV+XvaWlp8OTJE0uMsTpCqKguRuDDhw9Z3Qd1dXUICQn5i8fj5X0GRXirqKjoBUJIr/L7paWlJY+JifkRIfSwIW+4fPny/98UqXv37qnm5uazdHS4xzpSU1N1nj17dmTy5MkNtq8tn8+HefPmdbp48WJgfn4+pwVqZWUFY8aMOdahQ4c0rnO8vb3HLViw4MGlS5fM65One/fu+WtqanK6+0lJSVeh0qjbv//+izMzM1nPVygUWE9PL6S6fpwlS5Zs9fDwSOWKScjVx9SiRQvlpEmTzgwYMCClGoHNyc3NTWDri8IYA0JIs6kGuT5XIiMjJ7FNHMYYg7a29rDLly/b1CXda9eumZ4+fZr1mLW1NYwaNaqoMTYpa2oQQrSbmxuPrf0rlUqtdevW+TTGfalyGWB27tx5bdSoUa+4FhvTNA2RkZFaISEhF1xdXUeruma1GktKc/369YuOHDkSkpKSYlTNRFFsaWm5w9/ffz2XEDx//rxramrqX2vWrOmwYcOGkOXLl39Z3VpFLoRCIWzbtq0PTdOsFpdYLIagoKATCKEK5l5oaCjiEsC8vLw8hUKxowazPMXGxmaKpqZmem3y27Nnzxc//vjjz9W9BAzDgIWFBS0UCrnctzYvXrwYQ2Ss7vTv3z9m1KhRXCImCggIGFvbNHk8Hhw5ciRAoVCosfUNSqXS6Hnz5n028cxcXV2v0TRdpSFnZGTA5s2bOwUHBzdrNAF817f10MHBYWjHjh2rXRsVGxvLf/LkyaFp06Ydv3PnTuu63PjJkydew4cPP7Zp06Zl8fHx1a6xHT58OL548eLC6l7ymTNn/nzp0iVQKpVw6dIl48OHD1+YOXPm3zUt/GZxxXsfP37crrCQfeCtbdu2sHr1an4lIZeNGTPmK640NTQ0mG7duhXU8DGAhQsXnhkyZEhkTdEuyvWbgK+v7y+qWABffPEF4koXY4yioqL0iYzVnX79+mVLJJIDHBFW4PHjx367du3yq02a8+bN63/06FEbtkEUHR0dcHd3D29ot/BDMnny5AW2trasjbmwsLDrrl27TtQncvrLly/tEhMTp965c0fAKoAAABMmTIho2bJlX0NDw2pf2PT0dNi7d29Pd3f3kDlz5hz97rvvzLKysozZQke9C7ioe+PGDZMVK1YMWLZsWcygQYN2/fPPP19yrUAp+8oNHDhQPnPmzDE8Hi+L4+UVzp49e2N4eHj/SgLL27t37whzc/NH69evP7Fx48b27/KnVvkeGGPZlStXTNauXbvlzJkzJ/Lz8w248vPmzZvQnj177i//d7lcrm9gYMAZVsrGxgZUDam1YMGCb9q3b19Q03kaGhpgZGS0xMXFZZ8q6dra2gKXAObn58PDhw/JapD6uXAFY8aMOd6jRw+abTQzOjoadu3a9eu2bdvm1DQths/ng7+/f9+goKCjhYWFrF6WiYlJ+vTp0+d8TmWIMaaGDBlyiW1UOzs7G86fP+88atSoiwEBASY1BXGplK7WkiVLevj5+V2ZOXPmZqVS+f4GrDVx7NixB97e3p7BwcGHkpKSON3c4uJiKC4u1lu/fv1wExOT4QghnJKSssXQ0PB+u3btQCgUQmpqKiQnJ4tnzZr1fVpams2RI0eAYRiVJpn2799fPmHCBO8ePXoc4SqwgIAAv6CgoJlsgwd5eXmQl5dnPGfOHGMTE5Ohz58/x8XFxcd0dHQumpmZAY/Hg4yMDMEPP/wwJSkpqeO5c+egugjQWlpaimnTph21tLSs4P7evHkTXr16xXkd1xQYDkstYePGjZtTUlJ+evnyJZe7DF26dIn9+eefz3FF6WWzXLlWNuTm5sKtW7eIitUTFxeXMzNnzpx19erVtQBQ5WsTFhaGsrKy1rZq1Uq8bt26Rx4eHufLL2ETiURw48YNj+XLl7c9derUL3Fxcazvp4WFBfz0009/SKXStM/sI6LYvHmzv6mpqVlUVFRbNhE8cuSIc05OTuLIkSM3DR8+/D93d/dLCKFMto9IYmKi55w5c2ReXl5Tg4KCHFJTUwEhBH379h0PAFs5BRAAYP/+/UGDBw/21NDQuPDs2TNZTe5bYmIi7NixAwHADHV1dYiMjASKoqC4uBiysrJqtZQFIQRDhw4FLy8v7yFDhhyp7txbt24xNe3DW5a/7du3IwAYIRQKRyQkJABCCPLz82HTpk0q5Wvs2LHP/f39t1T++9mzZ6E6S9bQUPWg2jRNg0QiWSeRSNoBwCC2c/T19WH06NH7HRwc7quabrNmzTij3uTk5MCVK1fIXMAGYPv27Zt9fHzmb9++ndWLeP78OfD5/KVLly5lFi1adMnHxydbW1sbMjIy4Pbt26IZM2YMePz4sRrXfE8dHR1wdHTca2Njs/xzLL8ZM2ZEbdy48fTmzZvbvnjBPjX50qVLIBKJZsbFxcHs2bPDnJycEjp06ACamppQWFgIL1++hPDwcMGECRO+jomJoRISEipoy86dO79DCG1VaaXPhg0bOo4aNSpTW1u70TaaKf+TyWTY3d39enp6ukoB+zp16iSYOnXqeltb20bLE5/Px99++212ampqN7YF2b17925TXWDStWvX5tZ2683Tp09PsbW1VbCFRvfy8gqpbV9IRkZGH11dXc5Q7XZ2dn/XFCi0ofjA+wKrFD+Pz+fD4sWLU4EjnPy5c+c4A/Lm5OQ4+Pr6vuHabAzKBQbV0tLC2traWFNTE9e0qZWWlhbu2rXrThcXF07DpbqAqHPmzMEYY5VGo7W1tZdz7cY2f/58VxVdT10NDY1YYN/aNrKa6zQ3bdoUqOrWpWKx+H05ymQy/G5SNef5RkZGij///HNBtRZgGbNmzbqXm5s7YMqUKeOePHni++jRo0Z7MVxdXaFbt24bVq1aNVtPT7UB5rt37yru3r3rf/To0byFCxd6vnz50p4t1lhdEQgEMGnSpKzRo0cPNzAwuM12jq6urg5XYFIAAJlMVhYGX2U8PT13TJs2rdvjx4+9Kwk+GBoajq3tnLLmzZsHAUAGALAuWXB0dByVn5//OwCEEjuufshksoi4uLjBPB4vYMeOHZxRjRiGAa65o5UxNjaGDh067Dt//vw39d0M6RNwhfMwxoNLS0vP7tmzp9/Tp0+rPf9dV5zK6aurq/Pv3LmTDcAyCMLR9xX+999/T/fy8uo3cODAM/r6+rg+6xvLIxQKwdzcHE+dOvXfiRMnDt6wYcPsuqQzYsSIZdu3bx+yZMmSa126dMFcUz5UpWz9rouLy6q5c+e69+rVK5jLXe/fv79XdWlFRUUF8Pn8Wg8y/Prrr8v79OmjKPdVhq5du25Zs2bNm7o8U8uWLTkHYjIzM+HJkydN1sgbc+5aNWnXaiUBl9CokndLS8vQTZs2fdWpUyf/du3a1Xk9MJ/Ph65du9Lz5s3be+7cuW9VET+u/NVGOJVKJdQlfqEqecEY11iGCKFSf3//kT169BjctWvXV/Xds5uiKODz+TB06FBF3759v96xY8d+lSzA8sydO/cyxjj4+PHjNhcuXNgRHBxszePxdOLi4motesbGxoAxTp4wYUKajY2Nz7Bhw56pGi6ciz59+iRijD0cHBxsHj169Oe2bdv0FAqFSUJCgsovnFQqBSMjI+jUqdP9iRMn/tW/f/+NbIEMygtgSUnJaAMDA2BzjxFCcPDgwToFwxSLxXEDBw6c0KxZs4NZWVnQuXPnV7///vv2Oq4oABcXF0lQUBBrPouKiuDy5cu8phA/hBCvefPmwFZmGGOQSqV1zoe6urpAX1+/StoMw4CamhpVG0tcXV1dUjkdjDHo6+sDn88XqvCcsTwe79ewsLDkbdu2+YeGhnbIzMwUqmL1GRgYgEAgeOPm5vasZ8+ewyZMmJBRzUqf8u8WT09Pj/X5JRIJqPr8EolErWyry8r1UwsxZwwMDDTkcnkF8aUoCrS0tDRSU1NrKr9cADgXExPz4OzZs9t2797dsaioyDAhIUFlMdfV1QV9fX1o2bLlg4EDB560t7ff2rNnz+zt299G9+PXofEqAOARADglJye7njx5su+qVau0XFxcRufn5+skJydDZmYmpKWlAcMwIBAIQFtbG1q0aAFmZmYgFApT//33390//PADM3369IMIoagGfrnkAPAAALpgjFv/9ttvY9etWyd2dHT8JisrS+vly5eQkpICRUVFwOfzoayxtGzZEmiafpCenn5h69atad27d9988OBBlb5wYrH4F3d3d002q5PP54NSqXwZEBBQp+c5ePDgf1OnTr1/6tSpdi4uLn4IoWd1LZvu3buvFYlEOpUbMMMwoK2tDRRFJTSFAOrp6b3s2LHjL1KpFCqXmVKphNatW9c5H+bm5g8cHR1/MTExqfCilpSUgIGBwWsAUOkjS9M0mJubr3R3d5eUz6NSqQRTU1OQSCQ3VE2na9euRxFCRx8+fDjW19e3jVQq/VIsFndOTk6G7OxsUCgUIBaLQVNTEywsLODp06cnO3bsGDVjxoyr9vb2wbt374aJEyeq+vyvCgoKftHU1KxQtiUlJdCiRYtcUDFuoKen52U+n6+o7L4rlUowNDSMVrE68gcMGLCWpulm5YWUz+cDQiitJte2DCsrq2QAGIQxtluzZs3Xq1evVnNzc/MuKirSS01Nhfz8fCgtLQUejwdSqRS0tbXB1NQUbt++vb9169bJ8+fPz3JwcPjtXX963V2CaiwVKCoq6nj69GntK1euwKNHj+DWrVtA0zRoaGhAmzZtwNHREQYOHAhubm4ZCKFH0ITw+XzIzMzsdODAAVlwcDDcuHED0tPTQUNDA+zt7cHR0RFcXV3B3d09FiGU+LH1iZw8eXLos2fP1i1YsMCS9LB9+mCMjcPCwlpdvXoVnj59CoWFhaCrqwsWFhbg6ekJ7du3v1VXK/9/AR6PB0ql0u706dO6ERER8OLFCygsLAQ1NTUwNTWFVq1aQf/+/cHU1DQEIaSsLq3/AxSSWQ43bLNFAAAAAElFTkSuQmCC";
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
const MAT_STATUS = ["Solicitado", "Em aprovação", "Cotação", "Cotação recebida", "Em aprovação comercial", "PO emitida", "Em fabricação", "Em trânsito", "Recebido", "Entregue a bordo", "Dentro do Prazo", "Fora do Prazo", "Recusado", "Devolvido"];
const PAY_STATUS = ["Orçamento", "Aprovado", "PO emitida", "Serviço executado", "Medição aprovada", "NF recebida", "NF validada", "Pagamento programado", "Pago"];
const PRIORITY = ["Baixa", "Normal", "Média", "Alta", "Crítica", "Importante", "Emergencial", "Sobressalente crítico"];
/* Materiais (nova aba): departamentos, status e colunas da requisição */
const MAT_DEPTS = ["Máquinas", "Guindaste", "Marinha", "SMS", "Convés", "Base Prio", "Elétrica", "OIM", "Subsea", "Almoxarifado"];
/* Registro de Bordo: departamentos próprios e data de início (só vale a partir daqui) */
const BORDO_DEPTS = ["Marinha", "Convés", "Máquinas", "Elétrica", "Guindaste"];
const BORDO_INICIO = "2026-11-18";
const BORDO_IMPORT_MIN = "2025-01-01"; /* a importação por planilha aceita datas retroativas desde aqui */
const DEPTS_TODOS = [...new Set([...MAT_DEPTS, ...BORDO_DEPTS])];
const MAT_STATUS2 = ["Aguardando Entrega", "Cancelado", "Em Processo de Compras", "Recebido", "Recusado", "Stand By"];
const MAT_PROC = ["Aberto", "Fechado"];
const MAT_STATUS_COR = { "Aguardando Entrega": "#3B82F6", "Cancelado": "#6B7280", "Em Processo de Compras": "#F5A623", "Recebido": "#22C55E", "Recusado": "#EF4444", "Stand By": "#9499A8" };
const MAT2_COLS = [
  ["tmMaster", "TM Master"], ["departamento", "Departamento"], ["priority", "Prioridade"], ["sap", "SAP"], ["descricao", "Descrição"],
  ["quantidade", "Quantidade"], ["justificativa", "Justificativa"], ["status", "Status"], ["statusProcesso", "Status do Processo"], ["oe", "OE"],
  ["dataSolicitacao", "Data da Solicitação"], ["reserva", "Reserva"], ["rc", "RC"], ["po", "PO"], ["linhaPo", "Item da PO"],
  ["valor", "Valor Total"], ["eta", "ETA"], ["dataRecebimento", "Data de Recebimento"], ["obs", "Observação"], ["id", "ID"],
];
/* quem pode editar uma requisição: administrador e Almoxarifado em todas; os demais, só as do seu departamento
   (mesma regra aplicada pelo servidor em matmerge.js) */
/* departamentos de Materiais de um usuário: lista nova (matDepts) ou o campo antigo de um só (matDept) */
const deptsDoUsuario = (u) => (Array.isArray(u?.matDepts) ? u.matDepts : u?.matDept ? [u.matDept] : []).map((d) => MAT_DEPTS.find((x) => String(x).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "") === String(d).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "")) || d).filter((d) => MAT_DEPTS.includes(d));
/* liberação do Registro de Bordo é separada da de Materiais (bordoDepts); contas antigas, sem essa lista, herdam as marcações de Materiais que existem no Registro de Bordo */
const deptsBordoUsuario = (u) => (Array.isArray(u?.bordoDepts) ? u.bordoDepts : (Array.isArray(u?.matDepts) ? u.matDepts : u?.matDept ? [u.matDept] : [])).filter((d) => BORDO_DEPTS.includes(d));
/* departamento "canônico": ignora acento, maiúscula/minúscula e espaços (ex.: "MAQUINAS" → "Máquinas") */
const semAcentoD = (x) => String(x ?? "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, " ").trim();
const canonDept = (d) => MAT_DEPTS.find((x) => semAcentoD(x) === semAcentoD(d)) || BORDO_DEPTS.find((x) => semAcentoD(x) === semAcentoD(d)) || d;
const canonMat = (rows) => (rows || []).map((r) => { const c = canonDept(r && r.departamento); return r && c !== r.departamento ? { ...r, departamento: c } : r; });
const podeEditarMat = (u, depto) => {
  if (!u) return false;
  if (isAdminUser(u)) return true;
  const ds = deptsDoUsuario(u);
  return ds.includes("Almoxarifado") || ds.includes(canonDept(depto));
};
/* Registro de Bordo: administrador edita tudo; os demais, só os departamentos que lhes foram atribuídos
   (Almoxarifado aqui é um departamento como os outros). Mesma regra no servidor (matmerge.js). */
const podeEditarBordo = (u, depto) => {
  if (!u) return false;
  if (isAdminUser(u)) return true;
  return deptsBordoUsuario(u).includes(depto);
};
/* junta a lista local com a do servidor, linha a linha (a edição mais recente de cada requisição vence) */
const mesclarListaCliente = (local, localDel, srv, srvDel, pode) => {
  const del = new Set([...(localDel || []), ...(srvDel || [])]);
  const m = new Map((srv || []).map((r) => [r.id, r]));
  (local || []).forEach((r) => { const e = m.get(r.id); if (!e) { if (pode(r.departamento)) m.set(r.id, r); } else if ((r.upd || 0) > (e.upd || 0) && pode(e.departamento, e)) m.set(r.id, r); });
  del.forEach((id) => m.delete(id));
  return [...m.values()];
};
const mesclarMatCliente = (local, localDel, srv, srvDel, u) => mesclarListaCliente(local, localDel, srv, srvDel, (d) => podeEditarMat(u, d));
const mesclarBordoCliente = (local, localDel, srv, srvDel, u) => mesclarListaCliente(local, localDel, srv, srvDel, (d, row) => podeEditarBordo(u, d) || bordoDono(row, u));
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
  ["repeatOf", "Repetição de (ID)"],
  ["planoAcao", "Plano de Ação"], ["precisaMaterial", "Precisa de Material"], ["materialNecessario", "Material Necessário"],
  ["impacto", "Impacto"], ["novaDataPrevista", "Nova Data Prevista"],
];
const PLAN_COLS = [
  ["nome", "Nome"], ["departamento", "Departamento"], ["empresa", "Empresa"], ["descricaoProblema", "Descrição do Problema"],
  ["planoAcao", "Plano de Ação"], ["rc", "RC"], ["obs", "Observação"],
  ["precisaMaterial", "Precisa de Material"], ["materialNecessario", "Material Necessário"], ["poMaterial", "PO"],
  ["impacto", "Impacto"], ["status", "Status"], ["dataExecucao", "Data de Execução"],
  ["jobTipo", "Job Type (vinculado)"], ["jobNo", "Job No (vinculado)"], ["justificativa", "Justificativa"],
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
  ["planoAcao", "Plano de Ação"], ["po", "PO"], ["linkedPlanId", "Adicionado ao Planejamento"],
];
const TM_HISTORY_COLS = [
  ["jobHistoryNumber", "Job History Number"], ["componentCode", "ComponentCode"], ["componentName", "ComponentName"],
  ["dateDone", "DateDone"], ["jobType", "JobType"], ["jobNo", "JobNo"], ["jobName", "JobName"],
  ["doneByName", "DoneByName"], ["serviceReport", "ServiceReport"], ["remarks", "Remarks"], ["postponed", "Postponed"], ["reason", "Reason"],
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
const exportSheets = (fileBase, sheets) => {
  const wb = XLSX.utils.book_new();
  sheets.forEach(([name, rows]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{ "—": "Sem registros" }]), name.slice(0, 31)));
  XLSX.writeFile(wb, `${fileBase}-${todayISO()}.xlsx`);
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
/* ---------- alerta de vencimento (sino) ---------- */
const diasParaVencer = (d) => {
  if (d.dueDate && d.dueDate >= "2020-01-01") {
    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    const [y, m, day] = d.dueDate.split("-").map(Number);
    return Math.round((new Date(y, m - 1, day) - hoje) / 86400000);
  }
  if (d.diffUnit === "D" && d.diffValue !== null && d.diffValue !== undefined) return d.diffValue;
  return null; /* jobs em horas de máquina não têm data */
};
const jobsAVencer = (tmDue, dias, incluirVencidos) => (tmDue || [])
  .map((d) => ({ d, dias: diasParaVencer(d) }))
  .filter((x) => x.dias !== null && x.dias <= dias && (incluirVencidos || x.dias >= 0))
  .sort((a, b) => a.dias - b.dias);

const AlertaVencimentos = ({ itens, dias, setDias, incluirVencidos, setIncluirVencidos, onAbrir }) => {
  const [aberto, setAberto] = useState(false);
  const cls = (n) => (n < 0 ? "d1" : n <= 7 ? "d1" : n <= 14 ? "d2" : "d3");
  const rotulo = (n) => (n < 0 ? `Vencido há ${-n}d` : n === 0 ? "Vence hoje" : `Vence em ${n}d`);
  return (
    <div className="g-bell-wrap">
      <button className="g-bell" onClick={() => setAberto((v) => !v)} title={`Jobs que vencem em até ${dias} dias`} aria-label="Alertas de vencimento">
        <Bell size={18} />
        {itens.length > 0 && <span className="g-bell-badge">{itens.length > 99 ? "99+" : itens.length}</span>}
      </button>
      {aberto && (
        <>
          <div style={{ position: "fixed", inset: 0, zIndex: 55 }} onClick={() => setAberto(false)} />
          <div className="g-bell-panel">
            <div className="g-bell-head">
              <b style={{ fontSize: 14 }}>{itens.length} job(s) {incluirVencidos ? "vencidos ou a vencer" : "a vencer"} em até</b>
              <select value={dias} onChange={(e) => setDias(Number(e.target.value))} style={{ padding: "4px 8px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 12.5 }}>
                {[7, 10, 15, 20, 30, 40].map((n) => <option key={n} value={n}>{n} dias</option>)}
              </select>
              <label style={{ fontSize: 12, display: "flex", alignItems: "center", gap: 5, marginLeft: "auto" }}>
                <input type="checkbox" checked={incluirVencidos} onChange={(e) => setIncluirVencidos(e.target.checked)} />Incluir vencidos
              </label>
            </div>
            <div className="g-bell-list">
              {itens.length === 0 && <div className="g-muted" style={{ padding: 22, textAlign: "center", fontSize: 13 }}>Nenhum job nessa janela.</div>}
              {itens.slice(0, 300).map(({ d, dias: n }) => (
                <div className="g-bell-item" key={d.id || `${d.code}-${d.jobNo}-${d.dueDate}`}>
                  <span className={`g-bell-days ${cls(n)}`}>{rotulo(n)}</span>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{d.jobName || "—"}</div>
                    <div className="g-muted" style={{ fontSize: 11.5 }}>
                      {d.code || "—"} · {d.component || "—"} · {d.jobType || "—"} {d.jobNo || ""} · Due {d.dueDate ? `${d.dueDate.slice(8, 10)}/${d.dueDate.slice(5, 7)}/${d.dueDate.slice(0, 4)}` : (d.dueRaw || "—")}
                    </div>
                  </div>
                </div>
              ))}
            </div>
            <div style={{ padding: "10px 14px", display: "flex", gap: 10, alignItems: "center" }}>
              <span className="g-muted" style={{ fontSize: 11.5 }}>Jobs em horas de máquina não entram (não têm data).</span>
              <button className="g-btn primary g-keep" style={{ marginLeft: "auto" }} onClick={() => { setAberto(false); onAbrir(); }}>Abrir TM Master</button>
            </div>
          </div>
        </>
      )}
    </div>
  );
};

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
  { id: uid("USR"), name: "Lethicia", username: "lethicia", password: "GenesisI", admin: true },
  { id: uid("USR"), name: "Davi", username: "davi", password: "GenesisI" },
  { id: uid("USR"), name: "Mariana", username: "mariana", password: "GenesisI" },
];

/* PERMISSÕES — admin (só a Lethicia) vê Configurações e gerencia usuários.
   role "viewer" = somente leitura; tabs = abas liberadas (sem o campo = todas). */
const PERM_TABS = [
  { key: "dashboard", label: "Dashboard" },
  { key: "manutencao", label: "Manutenção" },
  { key: "tmmaster", label: "TM Master" },
  { key: "materials", label: "Materiais" },
  { key: "payments", label: "Pagamentos" },
  { key: "costs", label: "Custos" },
];
const isAdminUser = (u) => !!u && (u.admin === true || (u.admin === undefined && String(u.username || "").toLowerCase() === "lethicia"));
const isViewerUser = (u) => !!u && !isAdminUser(u) && u.role === "viewer";
/* "Serviços" e "Planejamento" viraram a aba "Manutenção": quem tinha uma delas ganha Manutenção */
const normTabs = (tabs) => (Array.isArray(tabs) ? [...new Set(tabs.map((k) => (k === "services" || k === "planejamento" ? "manutencao" : k)))] : tabs);
const allowedTabsOf = (u) => {
  const all = PERM_TABS.map((t) => t.key);
  if (!u || isAdminUser(u) || !Array.isArray(u.tabs)) return all;
  const ut = normTabs(u.tabs);
  const ok = all.filter((k) => ut.includes(k));
  return ok.length ? ok : all.slice(0, 1);
};

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
const INITIAL_TM_BORDO = [];
const INITIAL_TM_FAILURES = { reports: [], actions: [], jobs: [], jobHistory: [], orders: [], importedAt: "" };

const SaveCtx = React.createContext("ok");
function SaveBadge() {
  const st = React.useContext(SaveCtx);
  if (st === "oculto") return null;
  const m = { ok: ["✔ Salvo no servidor", "#15803d", "rgba(34,197,94,.12)"], saving: ["⏳ Salvando…", "#92400e", "rgba(245,166,35,.15)"], pend: ["⏳ Salvando…", "#92400e", "rgba(245,166,35,.15)"], erro: ["⚠ NÃO SALVO — não feche nem atualize a página; o portal tenta de novo sozinho", "#fff", "#b91c1c"] }[st] || [];
  return <span title="Situação do salvamento no servidor" style={{ display: "inline-flex", alignItems: "center", fontSize: 12, fontWeight: 700, padding: "6px 12px", borderRadius: 999, color: m[1], background: m[2] }}>{m[0]}</span>;
}

export default function Root() {
  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState(null);
  const [avisoSalvar, setAvisoSalvar] = useState("");
  const [saveStatus, setSaveStatus] = useState("ok");
  const seqSaveRef = React.useRef(0);
  React.useEffect(() => {
    if (saveStatus === "ok") { window.onbeforeunload = null; return; }
    window.onbeforeunload = (e) => { e.preventDefault(); e.returnValue = ""; return ""; };
    return () => { window.onbeforeunload = null; };
  }, [saveStatus]);
  const lastSentRef = React.useRef(null); /* JSON por chave do que o servidor já tem */
  const retryRef = React.useRef(null);
  const [saveRetry, setSaveRetry] = useState(0);
  /* trava de segurança: só habilita o auto-save quando o GET /api/state realmente teve sucesso.
     Se a conexão falhar (backend reiniciando, erro de rede, etc.), NUNCA deixamos o app salvar —
     senão o estado local (ainda nos valores INITIAL_* vazios) seria gravado por cima dos dados
     reais do servidor 700ms depois, apagando tudo. Ver efeito de load logo abaixo. */
  const [saveEnabled, setSaveEnabled] = useState(false);
  const [loadAttempt, setLoadAttempt] = useState(0);

  const [users, setUsers] = useState(DEFAULT_USERS);
  const [currentUserRaw, setCurrentUser] = useState(null);
  /* usuários/permissões: salvos por um canal próprio, leve e imediato (não dependem do salvamento grande de todo o sistema) */
  const usersSalvos = useRef(null);
  const [usersSave, setUsersSave] = useState({ estado: "ok", msg: "" });
  const currentUser = currentUserRaw ? (users.find((x) => x.id === currentUserRaw.id) || currentUserRaw) : null;
  const roMode = isViewerUser(currentUser);
  const noop = () => {};

  const [workPackages, setWorkPackages] = useState(INITIAL_WORK_PACKAGES);
  const [materials, setMaterials] = useState(INITIAL_MATERIALS);
  const [matReqs, setMatReqs] = useState([]);
  const [matDeleted, setMatDeleted] = useState([]);
  const [bordoDeleted, setBordoDeleted] = useState([]);
  const [regObrig, setRegObrig] = useState(null);
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
  const [tmFailures, setTmFailures] = useState(INITIAL_TM_FAILURES);
  const [tmBordo, setTmBordo] = useState(INITIAL_TM_BORDO);
  const [portcallReqs, setPortcallReqs] = useState([]);
  const [reparoItens, setReparoItens] = useState([]);
  const [planosBordo, setPlanosBordo] = useState([]);

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
          if (d.users) { setUsers(d.users); usersSalvos.current = JSON.stringify(d.users); }
          if (d.workPackages) setWorkPackages(d.workPackages);
          if (d.materials) setMaterials(d.materials);
          if (Array.isArray(d.matReqs)) setMatReqs(canonMat(d.matReqs));
          if (Array.isArray(d.matDeleted)) setMatDeleted(d.matDeleted);
          if (Array.isArray(d.bordoDeleted)) setBordoDeleted(d.bordoDeleted);
          if (Array.isArray(d.regObrig)) setRegObrig(d.regObrig);
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
          if (Array.isArray(d.tmBordo)) setTmBordo(d.tmBordo);
          if (Array.isArray(d.portcallReqs)) setPortcallReqs(d.portcallReqs);
          if (Array.isArray(d.reparoItens)) setReparoItens(d.reparoItens);
          if (Array.isArray(d.planosBordo)) setPlanosBordo(d.planosBordo);
          if (d.tmFailures) setTmFailures({ ...INITIAL_TM_FAILURES, ...d.tmFailures });
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
    const cargaMat = matReqs;
    const cheio = () => (roMode ? { matReqs, matDeleted, tmBordo, bordoDeleted } : { workPackages, materials, payments, serviceInvoices, portCallMeta, opCategories, exchangeRate, planningItems, docagemItems, tmDue, tmHistory, tmDueSnapshots, tmFailures, tmBordo, portcallReqs, reparoItens, planosBordo, matReqs, matDeleted, bordoDeleted, regObrig: regObrig || undefined });
    /* primeira rodada após o carregamento: só memoriza o que já está no servidor (não reenvia tudo) */
    if (lastSentRef.current === null) { const f0 = cheio(); const b0 = {}; Object.keys(f0).forEach((k) => { b0[k] = JSON.stringify(f0[k] === undefined ? null : f0[k]); }); lastSentRef.current = b0; return; }
    const seq = ++seqSaveRef.current;
    setSaveStatus((x) => (x === "erro" ? x : "pend"));
    const t = setTimeout(() => {
      /* envia SÓ o que mudou (o servidor mescla por chave): payload pequeno, bem mais rápido e confiável */
      const full = cheio(); const strs = {};
      Object.keys(full).forEach((k) => { strs[k] = JSON.stringify(full[k] === undefined ? null : full[k]); });
      const ult = lastSentRef.current || {};
      const mud = new Set(Object.keys(strs).filter((k) => strs[k] !== ult[k]));
      if (mud.has("matReqs") || mud.has("matDeleted")) { mud.add("matReqs"); mud.add("matDeleted"); }
      if (mud.has("tmBordo") || mud.has("bordoDeleted")) { mud.add("tmBordo"); mud.add("bordoDeleted"); }
      if (!mud.size) { if (seq === seqSaveRef.current) setSaveStatus("ok"); return; }
      setSaveStatus((x) => (x === "erro" ? x : "saving"));
      const parcial = {}; mud.forEach((k) => { if (full[k] !== undefined) parcial[k] = full[k]; });
      fetch("/api/state", {
        method: "PUT",
        headers: { "Content-Type": "application/json", ...(currentUserRaw ? { "x-genesis-user": currentUserRaw.username || "", "x-genesis-pass": currentUser?.password || "", "x-genesis-oldpass": currentUserRaw.password || "" } : {}) },
        body: JSON.stringify({ data: parcial }),
      }).then(async (r) => {
        if (!r.ok) { let det = ""; try { const j = await r.json(); det = j.detalhe || j.error || ""; } catch (_) {} throw new Error(r.status === 413 ? "dados grandes demais (413)" : `servidor respondeu ${r.status}${det ? ": " + det : ""}`); }
        return r.json();
      }).then((j) => {
        const nu = { ...(lastSentRef.current || {}) }; mud.forEach((k) => { nu[k] = strs[k]; }); lastSentRef.current = nu;
        setAvisoSalvar((a) => (a && a.startsWith("Não foi possível salvar") ? "" : a));
        if (seq === seqSaveRef.current) setSaveStatus("ok"); else setSaveStatus("pend");
        /* o servidor devolve a lista de Materiais já mesclada e filtrada pela permissão: adota na tela */
        if (j && Array.isArray(j.matReqs)) {
          /* edição que o servidor NÃO aceitou (permissão): avisa em vez de desfazer em silêncio */
          const srvMap = new Map(j.matReqs.map((x) => [x.id, x]));
          const recusadas = cargaMat.filter((x) => { const e = srvMap.get(x.id); return e && (x.upd || 0) > (e.upd || 0) && podeEditarMat(currentUser, x.departamento); });
          if (recusadas.length) setAvisoSalvar(`O servidor não aceitou ${recusadas.length} alteração(ões) de Materiais (ex.: ${recusadas[0].departamento}). Confira se o seu usuário tem o departamento marcado em Configurações e se o server.js e o matmerge.js do Render estão atualizados.`);
          setMatReqs((cur) => { const novo = mesclarMatCliente(cur, [], j.matReqs, j.matDeleted, currentUser); return JSON.stringify(novo) === JSON.stringify(cur) ? cur : novo; });
          setMatDeleted((cur) => ((j.matDeleted || []).length === cur.length ? cur : j.matDeleted));
        }
        if (j && Array.isArray(j.tmBordo)) {
          setTmBordo((cur) => { const novo = mesclarBordoCliente(cur, [], j.tmBordo, j.bordoDeleted, currentUser); return JSON.stringify(novo) === JSON.stringify(cur) ? cur : novo; });
          setBordoDeleted((cur) => ((j.bordoDeleted || []).length === cur.length ? cur : j.bordoDeleted));
        }
      }).catch((e) => { setSaveStatus("erro"); setAvisoSalvar(`Não foi possível salvar no servidor agora (${(e && e.message) || "sem conexão"}). Suas alterações ficam neste navegador e o portal tenta de novo sozinho.`); clearTimeout(retryRef.current); retryRef.current = setTimeout(() => setSaveRetry((n) => n + 1), 8000); });
    }, 700);
    return () => clearTimeout(t);
  }, [loaded, saveEnabled, roMode, currentUserRaw?.id, workPackages, materials, payments, serviceInvoices, portCallMeta, opCategories, exchangeRate, planningItems, docagemItems, tmDue, tmHistory, tmDueSnapshots, tmFailures, tmBordo, portcallReqs, reparoItens, planosBordo, matReqs, matDeleted, bordoDeleted, regObrig, saveRetry]);

  /* atualiza Materiais a cada 20 s com o que os outros departamentos salvaram */
  const salvarUsuariosAgora = React.useCallback(async () => {
    if (!currentUserRaw) return;
    const json = JSON.stringify(users);
    const hdr = { "Content-Type": "application/json", "x-genesis-user": currentUserRaw.username || "", "x-genesis-pass": currentUser?.password || "", "x-genesis-oldpass": currentUserRaw.password || "" };
    setUsersSave({ estado: "salvando", msg: "Salvando usuários…" });
    try {
      let r = await fetch("/api/usuarios", { method: "PUT", headers: hdr, body: JSON.stringify({ users }) });
      if (r.status === 404 || r.status === 405) {
        /* servidor mais antigo, sem a rota de usuários: usa o salvamento geral */
        r = await fetch("/api/state", { method: "PUT", headers: hdr, body: JSON.stringify({ data: { users } }) });
      }
      if (!r.ok) { const j = await r.json().catch(() => ({})); throw new Error(`${j.error || "o servidor recusou"} (código ${r.status})`); }
      /* confirma lendo de volta do servidor — só diz "salvo" se realmente está lá */
      const v = await fetch("/api/usuarios", { cache: "no-store" }).then((x) => x.json()).catch(() => null);
      const lido = v && Array.isArray(v.users) ? v.users : null;
      if (!lido) throw new Error("o servidor não confirmou a gravação (server.js desatualizado no Render? troque também o server.js)");
      if (lido.length !== users.length || JSON.stringify(lido.map((u) => [u.id, u.username, u.password, u.role, u.tabs, u.matDepts, u.bordoDepts])) !== JSON.stringify(users.map((u) => [u.id, u.username, u.password, u.role, u.tabs, u.matDepts, u.bordoDepts]))) {
        throw new Error(`o servidor tem ${lido.length} usuário(s), mas aqui há ${users.length}; a gravação não foi aceita`);
      }
      usersSalvos.current = json;
      setUsersSave({ estado: "ok", msg: `Confirmado no servidor: ${lido.length} usuário(s) — ${new Date().toLocaleTimeString("pt-BR")}` });
    } catch (e) { setUsersSave({ estado: "erro", msg: "NÃO salvou os usuários: " + (e.message || "falha de conexão") }); }
  }, [users, currentUser, currentUserRaw]);
  React.useEffect(() => {
    if (!loaded || !saveEnabled || !currentUserRaw || !isAdminUser(currentUser)) return;
    if (JSON.stringify(users) === usersSalvos.current) return;
    setUsersSave({ estado: "salvando", msg: "Salvando usuários…" });
    const t = setTimeout(salvarUsuariosAgora, 400);
    return () => clearTimeout(t);
  }, [loaded, saveEnabled, currentUserRaw, currentUser, users, salvarUsuariosAgora]);
  /* avisa antes de fechar/atualizar a página se houver alteração de usuários ainda não salva */
  React.useEffect(() => {
    const h = (e) => { if (usersSave.estado !== "ok") { e.preventDefault(); e.returnValue = ""; } };
    window.addEventListener("beforeunload", h); return () => window.removeEventListener("beforeunload", h);
  }, [usersSave.estado]);

  React.useEffect(() => {
    if (!loaded || !saveEnabled || !currentUserRaw) return;
    const t = setInterval(() => {
      fetch("/api/materiais").then((r) => r.json()).then((j) => {
        if (!j || !Array.isArray(j.matReqs)) return;
        setMatReqs((cur) => { const novo = mesclarMatCliente(cur, [], j.matReqs, j.matDeleted, currentUser); return JSON.stringify(novo) === JSON.stringify(cur) ? cur : novo; });
      }).catch(() => {});
      /* permissões: quem não é administrador recebe aqui as mudanças de abas/departamentos feitas em Configurações */
      if (!isAdminUser(currentUser)) fetch("/api/usuarios").then((r) => r.json()).then((j) => {
        if (!j || !Array.isArray(j.users) || !j.users.length) return;
        setUsers((cur) => (JSON.stringify(cur) === JSON.stringify(j.users) ? cur : j.users));
      }).catch(() => {});
      fetch("/api/bordo").then((r) => r.json()).then((j) => {
        if (!j || !Array.isArray(j.tmBordo)) return;
        setTmBordo((cur) => { const novo = mesclarBordoCliente(cur, [], j.tmBordo, j.bordoDeleted, currentUser); return JSON.stringify(novo) === JSON.stringify(cur) ? cur : novo; });
      }).catch(() => {});
    }, 20000);
    return () => clearInterval(t);
  }, [loaded, saveEnabled, currentUserRaw, currentUser]);

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
    <SaveCtx.Provider value={isAdminUser(currentUser) ? saveStatus : "oculto"}>
    <Genesis
      currentUser={currentUser}
      onLogout={() => setCurrentUser(null)}
      readOnly={roMode}
      users={users} setUsers={isAdminUser(currentUser) ? setUsers : noop} usersSave={usersSave} salvarUsuarios={salvarUsuariosAgora}
      workPackages={workPackages} setWorkPackages={roMode ? noop : setWorkPackages}
      materials={materials} setMaterials={roMode ? noop : setMaterials}
      matReqs={matReqs} setMatReqs={setMatReqs} setMatDeleted={setMatDeleted}
      payments={payments} setPayments={roMode ? noop : setPayments}
      serviceInvoices={serviceInvoices} setServiceInvoices={roMode ? noop : setServiceInvoices}
      portCallMeta={portCallMeta} setPortCallMeta={roMode ? noop : setPortCallMeta}
      opCategories={opCategories} setOpCategories={roMode ? noop : setOpCategories}
      exchangeRate={exchangeRate} setExchangeRate={roMode ? noop : setExchangeRate}
      planningItems={planningItems} setPlanningItems={roMode ? noop : setPlanningItems}
      docagemItems={docagemItems} setDocagemItems={roMode ? noop : setDocagemItems}
      tmDue={tmDue} setTmDue={roMode ? noop : setTmDue} tmHistory={tmHistory} setTmHistory={roMode ? noop : setTmHistory}
      tmDueSnapshots={tmDueSnapshots} setTmDueSnapshots={roMode ? noop : setTmDueSnapshots}
      tmFailures={tmFailures} setTmFailures={roMode ? noop : setTmFailures}
      tmBordo={tmBordo} setTmBordo={setTmBordo} setBordoDeleted={setBordoDeleted} regObrig={regObrig} setRegObrig={roMode ? noop : setRegObrig}
      portcallReqs={portcallReqs} setPortcallReqs={roMode ? noop : setPortcallReqs} reparoItens={reparoItens} setReparoItens={roMode ? noop : setReparoItens} planosBordo={planosBordo} setPlanosBordo={roMode ? noop : setPlanosBordo}
      loadError={loadError}
    />
      {avisoSalvar && isAdminUser(currentUser) && (
        <div role="alert" style={{ position: "fixed", left: 16, right: 16, bottom: 16, zIndex: 9999, maxWidth: 720, margin: "0 auto", background: "#7f1d1d", color: "#fff", padding: "12px 16px", borderRadius: 8, boxShadow: "0 6px 24px rgba(0,0,0,.35)", fontSize: 13, lineHeight: 1.45, display: "flex", gap: 12, alignItems: "flex-start" }}>
          <span style={{ flex: 1 }}>{avisoSalvar}</span>
          <span style={{ cursor: "pointer", fontWeight: 700 }} onClick={() => setAvisoSalvar("")}>✕</span>
        </div>
      )}
    </SaveCtx.Provider>
  );
}

/* ============================================================
   MAIN APP
   ============================================================ */
function Genesis({ currentUser, onLogout, readOnly, users, setUsers,
  usersSave, salvarUsuarios, workPackages, setWorkPackages, materials, setMaterials, matReqs, setMatReqs, setMatDeleted, payments, setPayments,
  serviceInvoices, setServiceInvoices, portCallMeta, setPortCallMeta,
  opCategories, setOpCategories, exchangeRate, setExchangeRate,
  planningItems, setPlanningItems, docagemItems, setDocagemItems,
  tmDue, setTmDue, tmHistory, setTmHistory, tmDueSnapshots, setTmDueSnapshots, tmFailures, setTmFailures, tmBordo, setTmBordo, setBordoDeleted, regObrig, setRegObrig, portcallReqs, setPortcallReqs, reparoItens, setReparoItens, planosBordo, setPlanosBordo, loadError }) {
  const isAdmin = isAdminUser(currentUser);
  const allowedTabs = allowedTabsOf(currentUser);
  const [tabState, setTab] = useState(allowedTabs[0]);
  const tab = tabState === "settings" ? (isAdmin ? "settings" : allowedTabs[0]) : (allowedTabs.includes(tabState) ? tabState : allowedTabs[0]);
  const [newRowId, setNewRowId] = useState(null);
  /* usado sempre que uma linha nova é criada (novo serviço, novo material, novo registro de pagamento):
     guarda o id por alguns segundos pra a linha poder ser destacada e "scrollada" até a visão do usuário */
  const flashNewRow = (id) => {
    setNewRowId(id);
    setTimeout(() => setNewRowId((cur) => (cur === id ? null : cur)), 2500);
  };
  const [expandedWp, setExpandedWp] = useState(null);
  const [paySubTab, setPaySubTab] = useState("total"); // "total" | "status" | "dashboard"
  const [manSub, setManSub] = useState("services"); // "services" | "planejamento" | "portcall" | "reparo"
  const [planSubTab, setPlanSubTab] = useState("funil"); // "funil" | "mapeados" | "docagem" | ...
  const [tmSubTab, setTmSubTab] = useState("due"); // "due" | "history" | "metricas"
  const [alertaDias, setAlertaDias] = useState(20);
  const [alertaVencidos, setAlertaVencidos] = useState(false);
  const [alertaBannerFechado, setAlertaBannerFechado] = useState(false);
  const alertaItens = useMemo(() => jobsAVencer(tmDue, alertaDias, alertaVencidos), [tmDue, alertaDias, alertaVencidos]);
  const alertaVinte = useMemo(() => jobsAVencer(tmDue, alertaDias, false), [tmDue, alertaDias]);
  const abrirTmMaster = () => { setReportFn(null); setExportXlsxFn(null); setTmSubTab("due"); setTab("tmmaster"); };
  const [importMsg, setImportMsg] = useState(null);
  const [reportFn, setReportFn] = useState(null);
  /* mesmo princípio do reportFn (PDF), mas pra "Exportar planilha": cada aba pode registrar sua
     própria exportação em Excel, refletindo exatamente o que está filtrado ali. Abas que não
     registram nada continuam usando o exportador global (handleExportXlsx, com tudo do sistema). */
  const [exportXlsxFn, setExportXlsxFn] = useState(null);
  /* ao trocar de aba principal, limpa os exportadores ANTES dos efeitos da aba nova registrarem os seus
     (layout effect roda antes dos efeitos passivos). Isso impede que o relatório/planilha da aba anterior
     "vaze" pra uma aba que não registra nada (ex.: Configurações). */
  React.useLayoutEffect(() => { setReportFn(null); setExportXlsxFn(null); }, [tab]);
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
              ? { ...r, id: old.id, planoAcao: old.planoAcao || "", po: old.po || "", linkedPlanId: old.linkedPlanId || null }
              : { ...r, planoAcao: "", po: "", linkedPlanId: null };
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
          postponed: (row["Postponed"] || "").toString().trim(),
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

  /* Failure Reports: um Excel com 5 abas (Failure Reports, Actions, Jobs, Job History, Orders) — mescla por número */
  const handleImportFailures = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        const p = parseFailureWorkbook(wb);
        if (!p.reports || p.reports.length === 0) { setImportMsg("Não encontrei a aba \"Failure Reports\" na planilha."); setTimeout(() => setImportMsg(null), 6000); return; }
        setTmFailures((prev) => {
          const base = { ...INITIAL_TM_FAILURES, ...prev };
          const r1 = mergeBy(base.reports, p.reports, (x) => x.no);
          const r2 = mergeBy(base.actions, p.actions || [], (x) => x.no);
          const r3 = mergeBy(base.jobs, p.jobs || [], (x) => `${x.jobType}|${x.jobNo}|${x.frNo}`);
          const r4 = mergeBy(base.jobHistory, p.jobHistory || [], (x) => x.number);
          const r5 = mergeBy(base.orders, p.orders || [], (x) => x.no);
          setImportMsg(`Failure Reports importado: ${r1.add} novo(s)/${r1.upd} atualizado(s) · Actions ${r2.add}+${r2.upd} · Jobs ${r3.add}+${r3.upd} · Job History ${r4.add}+${r4.upd} · Orders ${r5.add}+${r5.upd}.`);
          return { reports: r1.list, actions: r2.list, jobs: r3.list, jobHistory: r4.list, orders: r5.list, importedAt: new Date().toLocaleString("pt-BR") };
        });
      } catch (err) {
        setImportMsg("Erro ao ler a planilha de Failure Reports. Confira se as abas seguem o padrão esperado.");
      }
      setTimeout(() => setImportMsg(null), 8000);
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
      dataExecucao: "", linkedServiceId: null, jobTipo: "", jobNo: "", justificativa: "",
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
        dataExecucao: "", linkedServiceId: null, linkedFromTmDueId: dueRow.id, jobTipo: dueRow.jobType || "", jobNo: dueRow.jobNo || "", justificativa: "",
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
    const paraGraduar = planningItems.filter((p) => p.cronoConf && p.cronoIni && p.cronoFim && !p.linkedServiceId);
    if (paraGraduar.length === 0) return;
    const linkByPlanId = {};
    const novosServicos = paraGraduar.map((p) => {
      const serviceId = uid("PC-2026-08");
      linkByPlanId[p.id] = serviceId;
      const start = p.cronoIni;
      const end = p.cronoFim;
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
    const paraGraduar = docagemItems.filter((d) => d.cronoConf && d.cronoIni && d.cronoFim && !d.linkedServiceId);
    if (paraGraduar.length === 0) return;
    const linkByDocId = {};
    const novosServicos = paraGraduar.map((d) => {
      const serviceId = uid("PC-2026-08");
      linkByDocId[d.id] = serviceId;
      const start = d.cronoIni;
      const end = d.cronoFim;
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

  /* Cronograma ↔ Serviços (nos dois sentidos): quem mudou por último vale. Mudou no cronograma → atualiza o Serviço;
     mudou o horário no Serviço (ex.: arrastou no Gantt) → atualiza o cronograma. Serviço concluído/cancelado encerra o item. */
  const funilSync = useRef({});
  React.useEffect(() => {
    const pares = [
      ...planningItems.filter((p) => p.cronoConf && p.linkedServiceId).map((p) => ({ src: "M", id: p.id, wp: p.linkedServiceId, nome: p.nome, empresa: p.empresa || "", plano: p.planoAcao || "", ini: p.cronoIni, fim: p.cronoFim, status: p.status })),
      ...docagemItems.filter((d) => d.cronoConf && d.linkedServiceId).map((d) => ({ src: "D", id: d.id, wp: d.linkedServiceId, nome: d.nome, empresa: d.empresa || "", plano: d.planoAcao || "", ini: d.cronoIni, fim: d.cronoFim, status: d.status })),
    ];
    if (!pares.length) return;
    const wpById = new Map(workPackages.map((w) => [w.id, w]));
    const push = new Map(), pull = new Map(), enc = new Map();
    pares.forEach((x) => {
      const w = wpById.get(x.wp); if (!w) return;
      const sig = `${x.ini}|${x.fim}|${x.nome}|${x.empresa}|${x.plano}`;
      const wsig = `${w.start}|${w.end}`;
      const last = funilSync.current[x.wp];
      if (!last) { funilSync.current[x.wp] = { sig, wsig }; if (wsig !== `${x.ini}|${x.fim}`) push.set(x.wp, x); }
      else if (last.sig !== sig) { push.set(x.wp, x); funilSync.current[x.wp] = { sig, wsig: `${x.ini}|${x.fim}` }; }
      else if (last.wsig !== wsig) { pull.set(x.wp, { x, w }); funilSync.current[x.wp] = { sig: `${w.start}|${w.end}|${x.nome}|${x.empresa}|${x.plano}`, wsig }; }
      if ((w.status === "Concluído" || w.status === "Cancelado") && x.status !== w.status) enc.set(`${x.src}:${x.id}`, w.status);
    });
    if (push.size) setWorkPackages((prev) => prev.map((w) => { const x = push.get(w.id); return x ? { ...w, start: x.ini || w.start, end: x.fim || w.end, name: x.nome, empresa: x.empresa, planoAcao: x.plano } : w; }));
    if (pull.size) {
      setPlanningItems((prev) => prev.map((p) => { const q = p.linkedServiceId && pull.get(p.linkedServiceId); return q ? { ...p, cronoIni: q.w.start, cronoFim: q.w.end } : p; }));
      setDocagemItems((prev) => prev.map((d) => { const q = d.linkedServiceId && pull.get(d.linkedServiceId); return q ? { ...d, cronoIni: q.w.start, cronoFim: q.w.end } : d; }));
    }
    if (enc.size) {
      setPlanningItems((prev) => prev.map((p) => (enc.has(`M:${p.id}`) ? { ...p, status: enc.get(`M:${p.id}`) } : p)));
      setDocagemItems((prev) => prev.map((d) => (enc.has(`D:${d.id}`) ? { ...d, status: enc.get(`D:${d.id}`) } : d)));
    }
  }, [planningItems, docagemItems, workPackages]);

  /* Serviço Concluído → entra sozinho em Pagamentos (status "Serviço executado"), uma única vez por serviço (vínculo wpId).
     Se já existir um pagamento com o mesmo nome de serviço, ele é apenas ligado ao serviço (e avança para "Serviço executado"
     se ainda estiver em Orçamento/Aprovado/PO emitida). */
  React.useEffect(() => {
    const concl = workPackages.filter((w) => w.status === "Concluído");
    if (!concl.length) return;
    setPayments((prev) => {
      const norm = (v) => String(v || "").trim().toLowerCase();
      const ligados = new Set(prev.map((x) => x.wpId).filter(Boolean));
      const porNome = new Map(prev.filter((x) => !x.wpId).map((x) => [norm(x.service), x]));
      const ANTES = ["Orçamento", "Aprovado", "PO emitida"];
      let mudou = false; const ligar = new Map(); const novos = [];
      concl.forEach((w) => {
        if (ligados.has(w.id)) return;
        const ex = porNome.get(norm(w.name));
        if (ex && !ligar.has(ex.id)) { ligar.set(ex.id, w.id); mudou = true; porNome.delete(norm(w.name)); return; }
        novos.push({ id: uid("PAY"), wpId: w.id, service: w.name || "—", po: "", poValue: Number(w.committed) || 0, nf: "", nfValue: 0, issue: "", due: "", status: "Serviço executado", auto: true });
        mudou = true;
      });
      if (!mudou) return prev;
      return [...prev.map((x) => (ligar.has(x.id) ? { ...x, wpId: ligar.get(x.id), status: ANTES.includes(x.status) ? "Serviço executado" : x.status } : x)), ...novos];
    });
  }, [workPackages]);

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

    const requisicoesAbertas = matReqs.filter((m) => m.statusProcesso !== "Fechado").length;
    const materiaisUrgentes = matReqs.filter((m) =>
      ["Alta", "Crítica", "Emergencial", "Sobressalente crítico"].includes(m.priority) && m.statusProcesso !== "Fechado"
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
  }, [workPackages, filteredWorkPackages, materials, matReqs, payments, portCallRange]);

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
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(matReqs, MAT2_COLS)), "Materiais");
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
        /* Procura a aba e a linha de cabeçalho certas: a planilha pode ter várias abas (ex.: uma base
           bruta gigante + a aba "PEDIDOS") e títulos acima do cabeçalho. Escolhe a aba/linha em que
           mais colunas são reconhecidas. */
        let best = { score: -1, name: null, headerRow: 0 };
        wb.SheetNames.forEach((name) => {
          const sh = wb.Sheets[name];
          const grid = XLSX.utils.sheet_to_json(sh, { header: 1, defval: "", range: 0, blankrows: false });
          for (let ri = 0; ri < Math.min(grid.length, 8); ri++) {
            const keys = new Set((grid[ri] || []).map((h) => HEADER_MAP[normHeader(h)]).filter(Boolean));
            const score = keys.size + (keys.has("descricao") ? 1 : 0) + (keys.has("tmMaster") ? 1 : 0);
            if (score > best.score) best = { score, name, headerRow: ri };
          }
        });
        if (!best.name || best.score < 3) {
          const nomes = wb.SheetNames.join(", ");
          setImportMsg(`Não encontrei as colunas de materiais em nenhuma aba (abas: ${nomes}). A planilha precisa ter colunas como "TM Master", "SAP", "Descrição", "Quantidade", "Status".`);
          setTimeout(() => setImportMsg(null), 12000);
          e.target.value = "";
          return;
        }
        const sheet = wb.Sheets[best.name];
        const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
        range.s.r = best.headerRow;
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "", range, blankrows: false })
          .filter((row) => Object.values(row).some((v) => String(v).replace(/\u00a0/g, "").trim() !== ""));
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
          if (s.toLowerCase() === "normal") return "Média";
          const f = PRIORITY.find((p) => p.toLowerCase() === s.toLowerCase());
          return f || (s || "Média");
        };
        const normStatus = (v, ctx) => {
          const s = (v || "").toString().trim();
          if (s) {
            const f = MAT_STATUS.find((m) => m.toLowerCase() === s.toLowerCase());
            return f || s;
          }
          /* sem status na planilha: deduz pelo que já foi preenchido */
          if (ctx.dataRecebimento) return "Recebido";
          if (ctx.po) return "PO emitida";
          return "Solicitado";
        };

        const parsedRows = json.map((row) => {
          const r = {};
          Object.keys(row).forEach((h) => { if (rowKeyMap[h]) r[rowKeyMap[h]] = row[h]; });
          const base = {
            tmMaster: (r.tmMaster || "").toString().trim(),
            departamento: (r.departamento || "").toString(),
            sap: (r.sap || "").toString().trim(),
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
          };
          base.status = normStatus(r.status, base);
          return base;
        }).filter((r) => r.descricao || r.sap || r.tmMaster);

        if (parsedRows.length === 0) {
          const headersFound = json.length > 0 ? Object.keys(json[0]).join(", ") : "(planilha vazia)";
          setImportMsg(`Nenhuma linha reconhecida nessa planilha. Cabeçalhos encontrados: ${headersFound}. Renomeie as colunas para algo como "Descrição"/"Material", "Quantidade", "Status" etc., ou avise quais são esses nomes para eu ajustar a importação.`);
          setTimeout(() => setImportMsg(null), 12000);
          e.target.value = "";
          return;
        }

        {
          const prev = materials;
          /* casa cada linha da planilha com um item já existente (mesmo TM Master + SAP; em caso de
             vários, prefere o de mesma RC/PO), sem nunca juntar duas linhas diferentes da planilha
             num mesmo item — assim pedidos distintos do mesmo material não se sobrescrevem */
          const keyOf = (m) => `${String(m.tmMaster || "").trim()}|${String(m.sap || "").trim()}`;
          const byKey = new Map();
          prev.forEach((m, idx) => {
            const k = keyOf(m);
            if (!byKey.has(k)) byKey.set(k, []);
            byKey.get(k).push(idx);
          });
          const claimed = new Set();
          const next = [...prev];
          let updated = 0, added = 0;
          parsedRows.forEach((row) => {
            const cands = (byKey.get(keyOf(row)) || []).filter((i) => !claimed.has(i));
            const pick = cands.find((i) => String(next[i].rc || "") === row.rc && String(next[i].po || "") === row.po)
              || cands.find((i) => String(next[i].rc || "") === row.rc)
              || cands.find((i) => !next[i].rc || !next[i].po)
              || cands[0];
            if (pick !== undefined && (row.tmMaster || row.sap)) {
              claimed.add(pick);
              next[pick] = { ...next[pick], ...row, id: next[pick].id, wp: next[pick].wp };
              updated++;
            } else {
              next.push({ id: uid("MAT"), wp: "", ...row });
              added++;
            }
          });
          setMaterials(next);
          setImportMsg(`Planilha importada (aba "${best.name.trim()}"): ${added} novo(s), ${updated} atualizado(s), ${parsedRows.length} linha(s) lida(s).`);
        }
      } catch (err) {
        setImportMsg(`Erro ao ler a planilha: ${err && err.message ? err.message : "formato inesperado"}.`);
      }
      setTimeout(() => setImportMsg(null), 9000);
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
          if (json.length && Object.keys(json[0]).includes("Status do Processo")) { setMatReqs(sheetToRows(json, MAT2_COLS).map((r) => ({ ...r, upd: Date.now() }))); imported.push(`${json.length} materiais`); }
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
    { key: "manutencao", label: "Manutenção", icon: Wrench },
    { key: "tmmaster", label: "TM Master", icon: Gauge },
    { key: "materials", label: "Materiais", icon: Package },
    { key: "payments", label: "Pagamentos", icon: Wallet },
    { key: "costs", label: "Custos", icon: Calculator },
    { key: "settings", label: "Configurações", icon: Settings },
  ].filter((n) => (n.key === "settings" ? isAdmin : allowedTabs.includes(n.key)));

  return (
    <div className={`genesis${readOnly ? " g-ro" : ""}`}>
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
            <div key={n.key} className={`g-nav-item ${tab === n.key ? "active" : ""}`} onClick={() => { if (n.key !== tab) { setReportFn(null); setExportXlsxFn(null); setTab(n.key); } }}>
              <n.icon size={14} />{n.label}
            </div>
          ))}
        </div>
        {allowedTabs.includes("tmmaster") && <AlertaVencimentos itens={alertaItens} dias={alertaDias} setDias={setAlertaDias} incluirVencidos={alertaVencidos} setIncluirVencidos={setAlertaVencidos} onAbrir={abrirTmMaster} />}
        {currentUser && <span style={{ fontSize: 12, whiteSpace: "nowrap", color: "var(--navbar-text)" }}>{currentUser.name}{readOnly ? " (somente leitura)" : ""}</span>}
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
          {tab !== "tmmaster" && tab !== "materials" && !(tab === "manutencao" && manSub !== "services") && <button className="g-btn g-edit-only" onClick={handleImportClick} title="Importar planilha (.xlsx) — reconhece abas Servicos, Materiais, Pagamentos, StatusPagamentos e Alocacoes"><Upload size={14} />Importar</button>}
          {tab !== "materials" && <button className="g-btn" onClick={() => exportXlsxFn && exportXlsxFn()} disabled={!exportXlsxFn} style={{ opacity: exportXlsxFn ? 1 : 0.5 }} title="Exportar em planilha (.xlsx) o conteúdo desta página, já filtrado"><Download size={14} />Exportar planilha</button>}
          {tab !== "materials" && <button className="g-btn" onClick={() => reportFn && reportFn()} disabled={!reportFn}
            title="Exportar relatório em PDF, com o conteúdo exato da página aberta"><FileText size={14} />Exportar relatório</button>}
          {tab === "manutencao" && manSub === "services" && <button className="g-btn primary" onClick={addWp}><Plus size={14} />Novo serviço</button>}
          {tab === "manutencao" && manSub === "planejamento" && planSubTab === "mapeados" && <button className="g-btn primary" onClick={addPlanningItem}><Plus size={14} />Novo mapeamento</button>}
          {tab === "payments" && <button className="g-btn primary" onClick={addInv}><Plus size={14} />Novo registro</button>}
        </div>
      </div>

      {allowedTabs.includes("tmmaster") && !alertaBannerFechado && alertaVinte.length > 0 && (
        <div className="g-alert-banner">
          <Bell size={16} />
          <span><b>{alertaVinte.length} job(s)</b> vencem em até {alertaDias} dias. Clique no sino para ver a lista.</span>
          <button className="g-btn" style={{ marginLeft: "auto" }} onClick={abrirTmMaster}>Abrir TM Master</button>
          <button className="g-btn" onClick={() => setAlertaBannerFechado(true)} title="Dispensar aviso"><X size={13} /></button>
        </div>
      )}

      {importMsg && (
        <div style={{ margin: "10px 22px 0 22px" }}>
          <div className="g-alert" style={{ background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>
            {importMsg}
          </div>
        </div>
      )}

      <div className="g-body">
        {tab === "dashboard" && (
          <DashboardView setExportXlsxFn={setExportXlsxFn} kpis={kpis} workPackages={workPackages} disciplineCosts={disciplineCosts}
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

        {tab === "manutencao" && (
          <div className="g-mode-toggle" style={{ marginBottom: 16, width: "fit-content" }}>
            {[["services", "Serviços"], ["planejamento", "Planejamento"], ["reparo", "Reparo e Retorno"], ["resumo", "Resumo por E-mail"]].map(([k, l]) => (
              <button key={k} className={manSub === k ? "active" : ""} onClick={() => { if (k !== manSub) { setReportFn(null); setExportXlsxFn(null); setManSub(k); } }}>{l}</button>
            ))}
          </div>
        )}

        {tab === "manutencao" && manSub === "services" && (
          <ServicesView workPackages={workPackages} updWp={updWp} remWp={remWp} repeatWp={repeatWp}
            expandedWp={expandedWp} setExpandedWp={setExpandedWp} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} newRowId={newRowId} />
        )}

        {tab === "manutencao" && manSub === "resumo" && (
          <ResumoEmailManutencao workPackages={workPackages} planningItems={planningItems} portcallReqs={portcallReqs} reparoItens={reparoItens} tmDue={tmDue} tmHistory={tmHistory}
            allPortCallDates={allPortCallDates} portCallLabel={portCallLabel} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} />
        )}

        {tab === "manutencao" && manSub === "reparo" && (
          <ReparoRetornoView reparoItens={reparoItens} setReparoItens={setReparoItens} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} />
        )}

        {tab === "manutencao" && manSub === "docagem" && (
          <PlanejamentoView soDocagem tmDue={tmDue} tmHistory={tmHistory} tmFailures={tmFailures} reparoItens={reparoItens} portCallMeta={portCallMeta} portcallReqs={portcallReqs} setPortcallReqs={setPortcallReqs} workPackages={workPackages} updWp={updWp} materials={[]} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn}
            allPortCallDates={allPortCallDates} portCallLabel={portCallLabel} addWpOnDate={addWpOnDate}
            planningItems={planningItems} updPlan={updPlan} remPlan={remPlan} addPlanningItem={addPlanningItem}
            docagemItems={docagemItems} updDocagem={updDocagem} remDocagem={remDocagem} addDocagemItem={addDocagemItem}
            handleImportDocagem={handleImportDocagem} planosBordo={planosBordo} setPlanosBordo={setPlanosBordo} planSubTab="docagem" setPlanSubTab={setPlanSubTab}
            newRowId={newRowId} handleImportPlanejamento={handleImportPlanejamento} />
        )}

        {tab === "manutencao" && manSub === "planejamento" && (
          <PlanejamentoView tmDue={tmDue} tmHistory={tmHistory} tmFailures={tmFailures} reparoItens={reparoItens} portCallMeta={portCallMeta} portcallReqs={portcallReqs} setPortcallReqs={setPortcallReqs} workPackages={workPackages} updWp={updWp} materials={[]} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn}
            allPortCallDates={allPortCallDates} portCallLabel={portCallLabel} addWpOnDate={addWpOnDate}
            planningItems={planningItems} updPlan={updPlan} remPlan={remPlan} addPlanningItem={addPlanningItem}
            docagemItems={docagemItems} updDocagem={updDocagem} remDocagem={remDocagem} addDocagemItem={addDocagemItem}
            handleImportDocagem={handleImportDocagem} planosBordo={planosBordo} setPlanosBordo={setPlanosBordo} planSubTab={planSubTab} setPlanSubTab={setPlanSubTab} setPlanningItems={setPlanningItems} setDocagemItems={setDocagemItems}
            newRowId={newRowId} handleImportPlanejamento={handleImportPlanejamento} />
        )}

        {tab === "tmmaster" && (
          <TmMasterView tmDue={tmDue} tmHistory={tmHistory} tmDueSnapshots={tmDueSnapshots} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn}
            handleImportTmDue={handleImportTmDue} handleImportTmHistory={handleImportTmHistory}
            tmFailures={tmFailures} handleImportFailures={handleImportFailures} tmBordo={tmBordo} setTmBordo={setTmBordo} setBordoDeleted={setBordoDeleted} regObrig={regObrig} setRegObrig={setRegObrig} currentUser={currentUser}
            tmSubTab={tmSubTab} setTmSubTab={setTmSubTab} updTmDue={updTmDue} onToggleAddToPlanning={onToggleAddToPlanning} />
        )}

        {tab === "materials" && <MaterialsView setExportXlsxFn={setExportXlsxFn} setReportFn={setReportFn} matReqs={matReqs} setMatReqs={setMatReqs} setMatDeleted={setMatDeleted} currentUser={currentUser} />}

        {tab === "payments" && (
          <PaymentsSection setExportXlsxFn={setExportXlsxFn}
            paySubTab={paySubTab} setPaySubTab={setPaySubTab}
            serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} addInv={addInv}
            setReportFn={setReportFn} newRowId={newRowId}
          />
        )}

        {tab === "costs" && (
          <CostsView setExportXlsxFn={setExportXlsxFn} serviceInvoices={serviceInvoices} updInv={updInv} setReportFn={setReportFn}
            exchangeRate={exchangeRate} setExchangeRate={setExchangeRate} />
        )}

        {tab === "settings" && (
          <SettingsView currentUser={currentUser} users={users} setUsers={setUsers} usersSave={usersSave} salvarUsuarios={salvarUsuarios} />
        )}
      </div>
    </div>
  );
}

/* ============================================================
   DASHBOARD — exact KPI set requested
   ============================================================ */
const DonutChart = ({ data, center, centerLabel, legend }) => {
  const dados = data.length ? data : [{ name: "Sem dados", value: 1, color: "#E7E9F0" }];
  return (
    <div className="dsh-donut">
      <div className="dsh-donut-chart">
        <PieChart width={150} height={150}>
          <Pie data={dados} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={52} outerRadius={74} paddingAngle={dados.length > 1 ? 2 : 0} stroke="none" isAnimationActive={false}>
            {dados.map((d, i) => <Cell key={i} fill={d.color} />)}
          </Pie>
          {data.length > 0 && <Tooltip contentStyle={{ background: "#fff", border: "1px solid var(--border)", borderRadius: 6, fontSize: 11 }} />}
        </PieChart>
        <div className="dsh-donut-center"><b>{center}</b><span>{centerLabel}</span></div>
      </div>
      <div className="dsh-legend">
        {(legend || data).map((d) => (
          <div key={d.name}><i style={{ background: d.color }} />{d.name}<b>{d.value}</b></div>
        ))}
      </div>
    </div>
  );
};

function DashboardView({ setExportXlsxFn, kpis, workPackages, disciplineCosts, serviceInvoices, exchangeRate, setExchangeRate, setReportFn, tmDue = [] }) {

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




  const pctRealizado = totalOrcadoDash > 0 ? Math.min(100, Math.round((totalRealizadoDash / totalOrcadoDash) * 100)) : 0;
  const pctReal = totalOrcadoDash > 0 ? Math.round((totalRealizadoDash / totalOrcadoDash) * 100) : 0;
  const legendServicos = [
    { name: "Concluídos", value: concluidosServicos, color: "#22C55E" },
    { name: "Em andamento", value: emAndamentoServicos, color: "#3B82F6" },
    { name: "Não iniciados", value: naoIniciadosServicos, color: "#9499A8" },
    { name: "Cancelados", value: canceladosServicos, color: "#EF4444" },
  ];
  const donutServicos = legendServicos.filter((d) => d.value > 0);
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

  React.useEffect(() => {
    if (!setExportXlsxFn) return;
    setExportXlsxFn(() => () => {
      exportSheets("genesis-dashboard", [
        ["Financeiro", financeiro.map((k) => ({ Indicador: k.label, Valor: k.value }))],
        ["Serviços", servicos.map((k) => ({ Indicador: k.label, Valor: k.value }))],
        ["Pagamentos", pagamentos.map((k) => ({ Situação: k.label, "Quantidade · Valor": k.value }))],
        ["Aprovação Pendente", topOpenInvoices.map((r) => ({ Data: r.date, Serviço: r.assunto, Empresa: r.empresa, "Valor Total": r.valorTotal, "Dias em Aberto": r.daysOpenTotal, Status: r.statusPagamento }))],
        ["Próximas Manutenções", upcomingMaintenance.map((w) => ({ Manutenção: w.name, Categoria: w.discipline, Empresa: w.empresa, Início: w.start, Status: w.status }))],
        ["Último Port Call", lastPortCallServices.map((w) => ({ "Port Call": lastPortCallLabel, Serviço: w.name, Empresa: w.empresa, Data: w.start, Status: w.status }))],
        ["TM Master", tmPorDepto.map((d) => ({ Departamento: d.departamento, Vencidas: d.Vencidas, "Até 40 dias": d["Até 40 dias"], "No prazo": d["No prazo"], Total: d.total }))],
        ["Materiais Urgentes", kpis.materiaisUrgentes.map((m) => ({ Descrição: m.descricao, Solicitação: m.dataSolicitacao, Prioridade: m.priority, Status: m.status }))],
      ]);
    });
  }, [financeiro, servicos, pagamentos, topOpenInvoices, upcomingMaintenance, lastPortCallServices, tmPorDepto, kpis, setExportXlsxFn]);

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
      y = pdfSectionTitle(doc, y, "TM Master — situação por departamento");
      y = pdfTable(doc, y, ["Departamento", "Vencidas", "Até 40 dias", "No prazo", "Total"],
        tmPorDepto.map((d) => [d.departamento, d.Vencidas, d["Até 40 dias"], d["No prazo"], d.total]),
        { columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" }, 4: { halign: "right" } } });
      if (y > 220) { doc.addPage(); y = 15; }
      y = pdfSectionTitle(doc, y, "Próximas manutenções (Planejamento)");
      y = pdfTable(doc, y, ["Manutenção", "Categoria", "Empresa", "Início", "Status"],
        upcomingMaintenance.map((w) => [w.name, w.discipline || "—", w.empresa || "—", fmtDate((w.start || "").slice(0, 10)), w.status]));
      if (y > 220) { doc.addPage(); y = 15; }
      y = pdfSectionTitle(doc, y, `Último Port Call${lastPortCallLabel ? ` — ${lastPortCallLabel}` : ""}`);
      y = pdfTable(doc, y, ["Serviço", "Empresa", "Data", "Status"],
        lastPortCallServices.map((w) => [w.name, w.empresa || "—", fmtDate((w.start || "").slice(0, 10)), w.status]));
      if (y > 220) { doc.addPage(); y = 15; }
      y = pdfSectionTitle(doc, y, "Pagamentos com Aprovação Pendente há mais tempo");
      y = pdfTable(doc, y,
        ["Data", "Serviço", "Empresa", "Valor", "Dias em Aberto"],
        topOpenInvoices.map((r) => [fmtDate(r.date), r.assunto, r.empresa, fmt(r.valorTotal), r.daysOpenTotal]),
        { columnStyles: { 3: { halign: "right" } } }
      );
      if (y > 220) { doc.addPage(); y = 15; }
      y = pdfSectionTitle(doc, y, "Materiais urgentes");
      pdfTable(doc, y, ["Descrição", "Necessidade", "Prioridade", "Status"],
        kpis.materiaisUrgentes.map((m) => [m.descricao, fmtDate(m.dataSolicitacao), m.priority, m.status]));
      pdfSave(doc, "relatorio-dashboard");
    });
  }, [financeiro, servicos, pagamentos, topOpenInvoices, upcomingMaintenance, lastPortCallServices, lastPortCallLabel, tmPorDepto, kpis, dp, setReportFn]);


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
          <DonutChart data={donutServicos} legend={legendServicos} center={`${taxaConclusaoServicos}%`} centerLabel="concluído" />
          <div className="dsh-chips">
            <div><b>{totalServicos}</b>Total</div>
            <div><b style={{ color: "var(--ok)" }}>{concluidosServicos}</b>Concluídos</div>
            <div><b style={{ color: "var(--accent)" }}>{emAndamentoServicos}</b>Andamento</div>
          </div>
        </div>
        <div className="dsh-card">
          <div className="dsh-title">Pagamentos</div>
          <DonutChart data={donutPagamentos} center={serviceInvoices.length} centerLabel="registros" />
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
                  <div className="dsh-item-sub">{m.departamento} · solicitado {fmtDate(m.dataSolicitacao)}</div>
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

/* ============================================================
   MANUTENÇÃO › Solicitação para o Port Call (e-mail) e Reparo e Retorno
   ============================================================ */
const PC_TIPOS = ["Preventiva", "Preditiva", "Corretiva", "Melhoria / Outra"];

function PortcallRequestView({ portcallReqs, setPortcallReqs, planningItems, tmDue, tmHistory, allPortCallDates, portCallLabel, setReportFn, setExportXlsxFn }) {
  const reqs = portcallReqs || [];
  const [selId, setSelId] = useState(null);
  const [msg, setMsg] = useState("");
  const [para, setPara] = useState(() => { try { return localStorage.getItem("genesis_portcall_para") || ""; } catch (e) { return ""; } });
  const hoje = todayISO();
  const sel = reqs.find((r) => r.id === selId) || reqs[0] || null;
  const inp = { padding: "4px 6px", border: "1px solid var(--border)", borderRadius: 4, fontSize: 12, background: "var(--panel, #fff)", color: "var(--text)", width: "100%" };

  const pcRotulo = (dk) => { try { return dk ? portCallLabel(dk) : "—"; } catch (e) { return dk || "—"; } };
  const novaReq = () => {
    const id = uid("PCR");
    setPortcallReqs((a) => [{ id, portCall: allPortCallDates.find((d) => d >= hoje) || allPortCallDates[allPortCallDates.length - 1] || "", itens: [], criado: hoje, enviado: "" }, ...(a || [])]);
    setSelId(id);
  };
  const updReq = (k, v) => setPortcallReqs((a) => a.map((r) => (r.id === sel.id ? { ...r, [k]: v } : r)));
  const remReq = () => { if (window.confirm("Excluir esta solicitação?")) { setPortcallReqs((a) => a.filter((r) => r.id !== sel.id)); setSelId(null); } };
  const novoItem = (base = {}) => ({ id: uid("PCI"), descricao: "", componente: "", tipo: "Preventiva", justificativa: "", linkado: "nao", jobTipo: "", jobNo: "", oneNo: "", ...base });
  const addItem = (base) => updReq("itens", [...sel.itens, novoItem(base)]);
  const updItem = (id, k, v) => updReq("itens", sel.itens.map((x) => (x.id === id ? { ...x, [k]: v } : x)));
  const remItem = (id) => updReq("itens", sel.itens.filter((x) => x.id !== id));
  const addMapeado = (pid) => {
    const p = planningItems.find((x) => x.id === pid); if (!p) return;
    const temJob = !!(p.jobNo);
    addItem({ descricao: p.nome, justificativa: p.descricaoProblema || p.planoAcao || "", tipo: "Preventiva", linkado: temJob ? "sim" : "nao", jobTipo: p.jobTipo || "", jobNo: p.jobNo || "", planId: p.id });
  };
  const sincronizar = () => {
    let n = 0;
    updReq("itens", sel.itens.map((it) => { const p = planningItems.find((x) => x.id === it.planId); if (!p) return it; n++;
      return { ...it, descricao: p.nome, justificativa: it.justificativa || p.justificativa || p.descricaoProblema || p.planoAcao || "", ...(p.jobNo ? { linkado: "sim", jobTipo: p.jobTipo || "", jobNo: p.jobNo } : {}) }; }));
    setMsg(`${n} item(ns) sincronizado(s) com os Mapeados para Execução (nome, justificativa e job).`); setTimeout(() => setMsg(""), 5000);
  };
  const addTodosMapeados = () => {
    const novos = planningItems.filter((p) => p.status !== "Concluído" && p.status !== "Cancelado" && !sel.itens.some((x) => x.planId === p.id))
      .map((p) => novoItem({ descricao: p.nome, justificativa: p.justificativa || p.descricaoProblema || p.planoAcao || "", linkado: p.jobNo ? "sim" : "nao", jobTipo: p.jobTipo || "", jobNo: p.jobNo || "", planId: p.id }));
    updReq("itens", [...sel.itens, ...novos]); setMsg(`${novos.length} mapeado(s) em aberto adicionado(s).`); setTimeout(() => setMsg(""), 5000);
  };
  const achar = (t, n) => {
    if (!n) return null;
    const k = (x) => `${bordoNorm(x.jobType)}|${bordoNorm(x.jobNo)}`; const alvo = `${bordoNorm(t)}|${bordoNorm(n)}`;
    const d = tmDue.find((x) => k(x) === alvo); if (d) return { onde: "Due", nome: d.jobName };
    const h = tmHistory.find((x) => k(x) === alvo); if (h) return { onde: "History", nome: h.jobName };
    return null;
  };
  const situacao = (it) => {
    if (it.linkado === "sim") {
      if (!it.jobNo) return { cor: "var(--crit)", txt: "Informe o job do plano de manutenção a que está ligado." };
      const a = achar(it.jobTipo, it.jobNo);
      return a ? { cor: "var(--ok)", txt: `Ligado ao job ${it.jobTipo} ${it.jobNo} (${a.onde}: ${a.nome || "—"}).` } : { cor: "var(--warn)", txt: `Job ${it.jobTipo} ${it.jobNo} não encontrado no Due/History — confira o número.` };
    }
    if (it.tipo === "Corretiva") return it.oneNo
      ? { cor: "var(--ok)", txt: `Sem job no plano; ONE job ${it.oneNo} aberto no TM.` }
      : { cor: "var(--crit)", txt: "Corretiva sem job no plano: ABRIR 1 JOB \"ONE\" no TM e informar o número." };
    return { cor: "var(--warn)", txt: "Sem job no plano de manutenção — avaliar incluir no plano." };
  };
  const pendentes = sel ? sel.itens.filter((it) => (it.linkado === "sim" && !it.jobNo) || (it.linkado !== "sim" && it.tipo === "Corretiva" && !it.oneNo)).length : 0;

  const assunto = sel ? `Solicitação de manutenções para o Port Call ${pcRotulo(sel.portCall)}` : "";
  const linhaCels = (it, i) => [i + 1, it.descricao || "—", it.componente || "—", it.tipo, it.justificativa || "—",
    it.linkado === "sim" ? `Sim — ${`${it.jobTipo} ${it.jobNo}`.trim() || "job não informado"}` : "Não",
    it.linkado !== "sim" && it.tipo === "Corretiva" ? (it.oneNo ? `ONE ${it.oneNo}` : "ABRIR 1 job ONE no TM") : "—"];
  const cols = ["#", "Manutenção a realizar no Port Call", "Componente", "Tipo", "Justificativa", "Ligada a job do plano de manutenção?", "Job ONE (corretiva sem job)"];
  const montarHtml = () => {
    const th = (t) => `<th style="border:1px solid #1F3A5F;background:#1F3A5F;color:#fff;padding:5px 7px;font-size:12px;text-align:left">${escHtml(t)}</th>`;
    const td = (t) => `<td style="border:1px solid #cfd8e6;padding:5px 7px;font-size:12px;vertical-align:top">${escHtml(t)}</td>`;
    return `<div style="font-family:Calibri,Arial,sans-serif;color:#1b2430"><h2 style="margin:0 0 4px">${escHtml(assunto)}</h2>
<p style="margin:0 0 8px;font-size:13px">Prezados, solicitamos a realização das manutenções abaixo durante o Port Call de <b>${escHtml(pcRotulo(sel.portCall))}</b>. Para cada item informamos a justificativa e se ele já está ligado a um job do plano de manutenção (TM). Itens corretivos sem job no plano terão <b>1 job ONE</b> aberto no TM.</p>
<table style="border-collapse:collapse;width:100%"><tr>${cols.map(th).join("")}</tr>${sel.itens.map((it, i) => `<tr>${linhaCels(it, i).map(td).join("")}</tr>`).join("") || `<tr>${td("(nenhum item)")}</tr>`}</table>
<p style="margin:10px 0 0;font-size:12px">Favor confirmar o recebimento e a viabilidade de cada item (equipe, material, janela de execução).</p></div>`;
  };
  const montarTexto = () => [assunto, "", `Prezados, solicitamos a realização das manutenções abaixo no Port Call de ${pcRotulo(sel.portCall)}.`, "", cols.join(" | "), ...sel.itens.map((it, i) => linhaCels(it, i).join(" | ")), "", "Favor confirmar o recebimento e a viabilidade de cada item."].join("\n");
  const copiar = async () => {
    const html = montarHtml(); const txt = montarTexto();
    try {
      if (window.ClipboardItem && navigator.clipboard.write) await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([txt], { type: "text/plain" }) })]);
      else await navigator.clipboard.writeText(txt);
      setMsg("Solicitação copiada com a tabela. Cole no corpo do e-mail (Ctrl+V).");
    } catch (e) { setMsg("Não consegui copiar automaticamente — selecione a prévia abaixo e copie com Ctrl+C."); }
    setTimeout(() => setMsg(""), 7000);
  };
  const abrirEmail = async () => {
    try { localStorage.setItem("genesis_portcall_para", para); } catch (e) { /* ignora */ }
    await copiar();
    updReq("enviado", hoje);
    const corpo = "Prezados,\n\nSegue a solicitação de manutenções para o Port Call.\n\n[Cole aqui a solicitação com Ctrl+V — a tabela já está copiada]\n";
    window.location.href = `mailto:${para.split(/[;,]/).map((x) => x.trim()).filter(Boolean).join(",")}?subject=${encodeURIComponent(assunto)}&body=${encodeURIComponent(corpo)}`;
  };

  React.useEffect(() => {
    if (!setExportXlsxFn || !setReportFn) return;
    const lin = () => reqs.flatMap((r) => r.itens.map((it) => ({ "Port Call": pcRotulo(r.portCall), "Criada em": r.criado, "Enviada em": r.enviado || "", Manutenção: it.descricao, Componente: it.componente, Tipo: it.tipo, Justificativa: it.justificativa, "Ligada a job do plano?": it.linkado === "sim" ? "Sim" : "Não", "Job Type": it.jobTipo, "Job No": it.jobNo, "Job ONE aberto no TM": it.oneNo })));
    setExportXlsxFn(() => () => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(lin().length ? lin() : [{ "—": "Sem solicitações" }]), "Solicitações Portcall"); XLSX.writeFile(wb, `genesis-solicitacoes-portcall-${todayISO()}.xlsx`); });
    setReportFn(() => () => {
      const doc = new jsPDF({ orientation: "landscape" });
      let y = pdfHeader(doc, "Manutenção — Solicitação para o Port Call", sel ? `${assunto} · Gerado em ${new Date().toLocaleDateString("pt-BR")}` : "Sem solicitação");
      if (sel) { y = pdfSectionTitle(doc, y, "Itens solicitados"); pdfTable(doc, y, cols, sel.itens.map((it, i) => linhaCels(it, i))); }
      pdfSave(doc, "solicitacao-portcall");
    });
  }, [reqs, sel, portCallLabel]);

  const mapDisp = planningItems.filter((p) => !(sel?.itens || []).some((x) => x.planId === p.id));
  return (
    <>
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Solicitação por e-mail das manutenções para o Port Call</span></div>
        <div className="g-muted" style={{ fontSize: 11.5, lineHeight: 1.5, marginBottom: 10 }}>
          Monte a lista do que precisa ser feito no Port Call, com a <b>justificativa</b> e se o item já está <b>ligado a um job do plano de manutenção</b>. Se for <b>corretiva</b> e <b>não</b> estiver ligada a nenhum job, abra <b>1 job "ONE"</b> no TM e informe o número aqui. Depois copie a tabela e abra o e-mail.
        </div>
        <div className="g-flex" style={{ gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          {reqs.length > 0 && <select value={sel?.id || ""} onChange={(e) => setSelId(e.target.value)} style={{ padding: "6px 8px", border: "1px solid var(--border)", borderRadius: 6 }}>
            {reqs.map((r) => <option key={r.id} value={r.id}>{pcRotulo(r.portCall)} — {r.itens.length} item(ns){r.enviado ? " · enviada" : ""}</option>)}
          </select>}
          <button className="g-btn primary" onClick={novaReq}><Plus size={14} />Nova solicitação</button>
          {sel && <button className="g-btn ghost danger" onClick={remReq}><Trash2 size={13} />Excluir</button>}
        </div>
      </div>
      {msg && <div className="g-alert" style={{ marginBottom: 12, background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>{msg}</div>}
      {!sel && <div className="g-panel"><div className="g-muted">Nenhuma solicitação ainda — clique em “Nova solicitação”.</div></div>}
      {sel && (
        <>
          <div className="g-panel">
            <div className="g-flex" style={{ gap: 14, flexWrap: "wrap", alignItems: "flex-end" }}>
              <div className="g-field"><label>Port Call</label>
                <select value={sel.portCall} onChange={(e) => updReq("portCall", e.target.value)} style={{ padding: "6px 8px", border: "1px solid var(--border)", borderRadius: 6, minWidth: 200 }}>
                  {sel.portCall && !allPortCallDates.includes(sel.portCall) && <option value={sel.portCall}>{sel.portCall}</option>}
                  {allPortCallDates.map((d) => <option key={d} value={d}>{pcRotulo(d)}</option>)}
                </select></div>
              <div className="g-field"><label>Ou data do Port Call</label><input type="date" value={sel.portCall} onChange={(e) => updReq("portCall", e.target.value)} /></div>
              <div className="g-field" style={{ flex: 1, minWidth: 260 }}><label>Destinatários (separe por vírgula)</label><input type="text" value={para} onChange={(e) => setPara(e.target.value)} placeholder="email1@..., email2@..." style={inp} /></div>
              <button className="g-btn g-keep" onClick={copiar}>Copiar tabela</button>
              <button className="g-btn primary g-keep" onClick={abrirEmail} disabled={!sel.itens.length}>Abrir e-mail</button>
            </div>
            {pendentes > 0 && <div className="g-alert" style={{ marginTop: 10 }}><AlertTriangle size={14} style={{ marginTop: 1 }} />{pendentes} item(ns) com pendência: informe o job do plano ou abra 1 job ONE no TM (corretiva sem job).</div>}
          </div>
          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Itens da solicitação ({sel.itens.length})</span>
              <div className="g-flex" style={{ gap: 8 }}>
                {mapDisp.length > 0 && <select value="" onChange={(e) => { if (e.target.value) addMapeado(e.target.value); }} style={{ padding: "5px 8px", border: "1px solid var(--border)", borderRadius: 6, maxWidth: 260 }}>
                  <option value="">+ Adicionar dos Mapeados…</option>{mapDisp.map((p) => <option key={p.id} value={p.id}>{p.nome}</option>)}
                </select>}
                <button className="g-btn" onClick={addTodosMapeados} title="Adiciona todos os mapeados não concluídos que ainda não estão na solicitação">+ Todos os mapeados em aberto</button>
                <button className="g-btn" onClick={sincronizar} title="Atualiza nome, justificativa e job com o que está em Mapeados para Execução">🔄 Sincronizar com Mapeados</button>
                <button className="g-btn" onClick={() => addItem()}><Plus size={14} />Item manual</button>
              </div>
            </div>
            <div className="g-table-wrap">
              <table className="g-table tbl-modern">
                <thead><tr><th>#</th><th style={{ minWidth: 220 }}>Manutenção</th><th style={{ minWidth: 130 }}>Componente</th><th>Tipo</th><th style={{ minWidth: 220 }}>Justificativa</th><th>Ligada a job do plano?</th><th style={{ minWidth: 150 }}>Job (tipo / nº)</th><th style={{ minWidth: 110 }}>Job ONE aberto no TM</th><th style={{ minWidth: 220 }}>Situação</th><th></th></tr></thead>
                <tbody>
                  {sel.itens.length === 0 && <tr><td colSpan={10} className="g-muted" style={{ textAlign: "center", padding: 16 }}>Adicione itens dos Mapeados para Execução ou manualmente.</td></tr>}
                  {sel.itens.map((it, i) => { const s = situacao(it); return (
                    <tr className="g-row tbl-row" key={it.id} style={{ "--c": s.cor }}>
                      <td>{i + 1}</td>
                      <td style={{ verticalAlign: "top" }}><ETextArea rows={1} value={it.descricao} onChange={(v) => updItem(it.id, "descricao", v)} />{it.planId && <span className="tbl-chip" style={{ color: "var(--ok)" }}>🔗 Mapeado</span>}</td>
                      <td style={{ verticalAlign: "top" }}><EText value={it.componente} onChange={(v) => updItem(it.id, "componente", v)} /></td>
                      <td><select value={it.tipo} onChange={(e) => updItem(it.id, "tipo", e.target.value)} style={inp}>{PC_TIPOS.map((t) => <option key={t}>{t}</option>)}</select></td>
                      <td style={{ verticalAlign: "top" }}><ETextArea rows={1} value={it.justificativa} onChange={(v) => updItem(it.id, "justificativa", v)} /></td>
                      <td><select value={it.linkado} onChange={(e) => updItem(it.id, "linkado", e.target.value)} style={inp}><option value="sim">Sim</option><option value="nao">Não</option></select></td>
                      <td>{it.linkado === "sim" ? <div className="g-flex" style={{ gap: 4 }}><input value={it.jobTipo} onChange={(e) => updItem(it.id, "jobTipo", e.target.value.toUpperCase())} placeholder="SRV" style={{ ...inp, width: 56 }} /><input value={it.jobNo} onChange={(e) => updItem(it.id, "jobNo", e.target.value)} placeholder="nº" style={{ ...inp, width: 70 }} /></div> : <span className="g-muted">—</span>}</td>
                      <td>{it.linkado !== "sim" && it.tipo === "Corretiva" ? <input value={it.oneNo} onChange={(e) => updItem(it.id, "oneNo", e.target.value)} placeholder="nº do ONE" style={inp} /> : <span className="g-muted">—</span>}</td>
                      <td style={{ whiteSpace: "normal", color: s.cor, fontSize: 11.5, fontWeight: 600 }}>{s.txt}</td>
                      <td><span className="g-btn ghost danger" onClick={() => remItem(it.id)}><Trash2 size={13} /></span></td>
                    </tr>); })}
                </tbody>
              </table>
            </div>
          </div>
          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Prévia do e-mail</span></div>
            <div style={{ background: "#fff", color: "#1b2430", padding: 12, borderRadius: 6, overflowX: "auto" }} dangerouslySetInnerHTML={{ __html: montarHtml() }} />
          </div>
        </>
      )}
    </>
  );
}

const RR_VAZIO = { item: "", sap: "", qty: "", desc: "", sn: "", nfGnsi: "", nfAcu: "", fornecedor: "", po: "", pad: "", bklOut: "", dataEnvio: "", lomRec: "", dataRecebido: "", status: "ABERTO", obs: "" };
const isIsoD = (v) => /^\d{4}-\d{2}-\d{2}$/.test(String(v || ""));
const DateOuTexto = ({ value, onChange }) => (isIsoD(value) || !value
  ? <input type="date" value={value || ""} onChange={(e) => onChange(e.target.value)} style={{ padding: "3px 4px", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11.5, background: "var(--panel, #fff)", color: "var(--text)" }} />
  : <input type="text" value={value} onChange={(e) => onChange(e.target.value)} title="Valor em texto (ex.: TBA). Apague para escolher uma data." style={{ padding: "3px 4px", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11.5, width: 90, background: "var(--panel, #fff)", color: "var(--text)" }} />);

function ReparoRetornoView({ reparoItens, setReparoItens, setReportFn, setExportXlsxFn }) {
  const itens = reparoItens || [];
  const hoje = todayISO();
  const fileRef = useRef(null);
  const [msg, setMsg] = useState("");
  const [f, setF] = useState({ status: "ABERTO", fornecedor: "", busca: "" });
  const [rrOpen, setRrOpen] = useState(null);
  const dias = (r) => (r.status !== "RECEBIDO" && isIsoD(r.dataEnvio) ? frDias(r.dataEnvio, hoje) : null);
  const rows = useMemo(() => itens.map((r) => ({ ...r, _d: dias(r) })), [itens]);
  const forn = useMemo(() => [...new Set(itens.map((r) => (r.fornecedor || "").trim()).filter(Boolean))].sort(), [itens]);
  const filtered = useMemo(() => {
    const b = f.busca.trim().toLowerCase();
    return rows.filter((r) => (f.status === "Todos" || r.status === f.status) && (!f.fornecedor || r.fornecedor === f.fornecedor)
      && (!b || `${r.item} ${r.sap} ${r.desc} ${r.sn} ${r.nfGnsi} ${r.nfAcu} ${r.fornecedor} ${r.po} ${r.pad} ${r.obs}`.toLowerCase().includes(b)))
      .sort((a, b2) => (b2._d ?? -1) - (a._d ?? -1));
  }, [rows, f]);
  const abertos = rows.filter((r) => r.status !== "RECEBIDO");
  const kpi = { total: rows.length, abertos: abertos.length, receb: rows.length - abertos.length, v90: abertos.filter((r) => r._d !== null && r._d > 90).length, semData: abertos.filter((r) => r._d === null).length, maior: abertos.reduce((m, r) => Math.max(m, r._d ?? 0), 0) };
  const porForn = useMemo(() => { const m = {}; abertos.forEach((r) => { const k = r.fornecedor || "—"; m[k] = m[k] || { n: 0, maior: 0 }; m[k].n++; m[k].maior = Math.max(m[k].maior, r._d ?? 0); }); return Object.entries(m).sort((a, b) => b[1].n - a[1].n); }, [rows]);

  const upd = (id, k, v) => setReparoItens((a) => a.map((r) => (r.id === id ? { ...r, [k]: v, ...(k === "dataRecebido" && isIsoD(v) ? { status: "RECEBIDO" } : {}) } : r)));
  const novo = () => setReparoItens((a) => [{ id: uid("RR"), ...RR_VAZIO, item: String(a.length + 1) }, ...a]);
  const rem = (id) => { if (window.confirm("Excluir este item?")) setReparoItens((a) => a.filter((r) => r.id !== id)); };

  const importar = (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        const nome = wb.SheetNames[0];
        const m = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, defval: "", raw: true });
        const nrm = (c) => String(c).trim().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
        const hi = m.findIndex((row) => row.some((c) => nrm(c) === "item") && row.some((c) => nrm(c).includes("description")));
        if (hi < 0) throw new Error("cabecalho");
        const cab = m[hi].map(nrm);
        const col = (fn) => cab.findIndex(fn);
        const ix = { item: col((c) => c === "item"), sap: col((c) => c === "sap"), qty: col((c) => c === "qty"), desc: col((c) => c.includes("description")), sn: col((c) => c === "s/n"), nfGnsi: col((c) => c.includes("gnsi")), nfAcu: col((c) => c.includes("acu")), fornecedor: col((c) => c.includes("fornecedor")), po: col((c) => c.includes("req/po") || c === "po"), pad: col((c) => c.startsWith("pad")), bklOut: col((c) => c.includes("bkl")), dataEnvio: col((c) => c.includes("sent")), lomRec: col((c) => c.includes("lom")), dataRecebido: col((c) => c.includes("received")), status: col((c) => c === "status"), obs: col((c) => c.includes("remark")) };
        const val = (v) => {
          if (v instanceof Date) return `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}`;
          return String(v === undefined || v === null ? "" : v).replace(/\s*\n\s*/g, " / ").trim();
        };
        const lidos = m.slice(hi + 1).map((row) => { const o = {}; Object.keys(ix).forEach((k) => { o[k] = ix[k] >= 0 ? val(row[ix[k]]) : ""; }); o.status = /receb/i.test(o.status) ? "RECEBIDO" : "ABERTO"; return o; }).filter((o) => o.desc || o.sn || o.sap);
        const chave = (o) => `${bordoNorm(o.item)}|${bordoNorm(o.desc)}|${bordoNorm(o.sn)}`;
        let add = 0, atu = 0;
        setReparoItens((prev) => {
          const arr = [...prev];
          lidos.forEach((o) => { const j = arr.findIndex((x) => chave(x) === chave(o)); if (j >= 0) { arr[j] = { ...arr[j], ...o }; atu++; } else { arr.push({ id: uid("RR"), ...RR_VAZIO, ...o }); add++; } });
          return arr;
        });
        setMsg(`Reparo e Retorno importado da aba "${nome}": ${add} item(ns) novo(s), ${atu} atualizado(s).`);
      } catch (err) { setMsg("Não consegui ler a planilha. Ela precisa ter as colunas ITEM, SAP, QTY, ITEM DESCRIPTION, S/N, ... STATUS, REMARKS."); }
      setTimeout(() => setMsg(""), 8000);
    };
    reader.readAsArrayBuffer(file); e.target.value = "";
  };

  React.useEffect(() => {
    if (!setExportXlsxFn || !setReportFn) return;
    const lin = () => filtered.map((r) => ({ ITEM: r.item, SAP: r.sap, QTY: r.qty, "ITEM DESCRIPTION": r.desc, "S/N": r.sn, "NF_e GNSI #": r.nfGnsi, "NF_e AÇU #": r.nfAcu, "FORNECEDOR DO SERVIÇO": r.fornecedor, "REQ/PO #": r.po, "PAD #": r.pad, "BKL OUT": r.bklOut, "DATE SENT": r.dataEnvio, "LOM REC": r.lomRec, "DATE RECEIVED": r.dataRecebido, STATUS: r.status, "DIAS EM REPARO": r._d === null ? "" : r._d, REMARKS: r.obs }));
    setExportXlsxFn(() => () => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(lin().length ? lin() : [{ "—": "Sem itens" }]), "Reparo e Retorno"); XLSX.writeFile(wb, `genesis-reparo-e-retorno-${todayISO()}.xlsx`); });
    setReportFn(() => () => {
      const doc = new jsPDF({ orientation: "landscape" });
      let y = pdfHeader(doc, "Manutenção — Reparo e Retorno", `${filtered.length} item(ns) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [{ label: "Itens", value: kpi.total }, { label: "Em aberto", value: kpi.abertos }, { label: "Recebidos", value: kpi.receb }, { label: "Abertos há +90 dias", value: kpi.v90 }, { label: "Maior tempo (dias)", value: kpi.maior }]);
      y = pdfSectionTitle(doc, y, "Itens");
      pdfTable(doc, y, ["Item", "SAP", "Qtd", "Descrição", "S/N", "Fornecedor", "REQ/PO", "Enviado", "Recebido", "Status", "Dias", "Remarks"],
        filtered.map((r) => [r.item, r.sap, r.qty, r.desc, r.sn, r.fornecedor, r.po, r.dataEnvio, r.dataRecebido, r.status, r._d ?? "—", r.obs]));
      pdfSave(doc, "relatorio-reparo-e-retorno");
    });
  }, [filtered, kpi]);

  const inp = { padding: "3px 5px", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11.5, background: "var(--panel, #fff)", color: "var(--text)", width: "100%" };
  const T = (r, k, w) => <td style={{ minWidth: w || 90, verticalAlign: "top" }}><input value={r[k] || ""} onChange={(e) => upd(r.id, k, e.target.value)} style={inp} /></td>;
  return (
    <>
      <div className="g-panel" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
        <div style={{ maxWidth: 820 }}>
          <div className="g-panel-title" style={{ marginBottom: 4 }}>Reparo e Retorno — itens enviados para reparo em terra</div>
          <div className="g-muted" style={{ fontSize: 11.5, lineHeight: 1.5 }}>Controle dos itens enviados para reparo (NF, fornecedor, PO, datas de envio e recebimento). Importe a planilha “Reparo e Retorno”; reimportar atualiza os itens existentes (mesmo ITEM + descrição + S/N) sem duplicar. Informar a data de recebimento marca o item como RECEBIDO. O tempo em reparo conta a partir de “Date sent”.</div>
        </div>
        <div className="g-flex" style={{ gap: 8 }}>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={importar} />
          <button className="g-btn primary" onClick={() => fileRef.current?.click()}><Upload size={14} />Importar planilha</button>
          <button className="g-btn primary" onClick={novo}><Plus size={14} />Novo item</button>
        </div>
      </div>
      {msg && <div className="g-alert" style={{ marginBottom: 12, background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>{msg}</div>}
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
        {bigKpi("Total de itens", kpi.total, "var(--teal)", Package)}
        {bigKpi("Em aberto", kpi.abertos, "var(--warn)", Clock)}
        {bigKpi("Recebidos", kpi.receb, "var(--ok)", Package)}
        {bigKpi("Abertos há mais de 90 dias", kpi.v90, kpi.v90 ? "var(--crit)" : "var(--ok)", AlertTriangle)}
        {bigKpi("Maior tempo em reparo (dias)", kpi.maior, kpi.maior > 90 ? "var(--crit)" : "var(--teal)", Clock)}
      </div>
      {porForn.length > 0 && (
        <div className="g-panel">
          <div className="g-panel-head"><span className="g-panel-title">Em aberto por fornecedor</span></div>
          <div className="tbl-chips">{porForn.map(([k, v]) => <span key={k} className="tbl-chip" style={{ cursor: "pointer" }} onClick={() => setF((p) => ({ ...p, fornecedor: k === "—" ? "" : k, status: "ABERTO" }))}>{k}: <b>{v.n}</b>{v.maior ? ` · até ${v.maior} d` : ""}</span>)}</div>
        </div>
      )}
      <div className="g-panel">
        <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", marginBottom: 10, alignItems: "center" }}>
          <div className="g-mode-toggle">{["ABERTO", "RECEBIDO", "Todos"].map((s) => <button key={s} className={f.status === s ? "active" : ""} onClick={() => setF((p) => ({ ...p, status: s }))}>{s}</button>)}</div>
          <select value={f.fornecedor} onChange={(e) => setF((p) => ({ ...p, fornecedor: e.target.value }))} style={{ padding: "6px 8px", border: "1px solid var(--border)", borderRadius: 6 }}><option value="">Todos os fornecedores</option>{forn.map((x) => <option key={x}>{x}</option>)}</select>
          <input value={f.busca} onChange={(e) => setF((p) => ({ ...p, busca: e.target.value }))} placeholder="buscar item, S/N, NF, PO..." style={{ padding: "6px 10px", border: "1px solid var(--border)", borderRadius: 6, minWidth: 240 }} />
          <span className="g-muted" style={{ fontSize: 11.5 }}>{filtered.length} de {rows.length} · ordenado pelos mais antigos em reparo</span>
        </div>
        <div className="g-table-wrap">
          <table className="g-table tbl-modern" style={{ minWidth: 0, width: "100%" }}>
            <thead><tr><th style={{ width: 30 }}></th><th style={{ minWidth: 280 }}>Item</th><th style={{ minWidth: 130 }}>Fornecedor</th><th style={{ minWidth: 190 }}>Documentos</th><th style={{ minWidth: 170 }}>Datas</th><th style={{ minWidth: 140 }}>Status</th><th style={{ minWidth: 200 }}>Remarks</th><th style={{ width: 40 }}></th></tr></thead>
            <tbody>
              {filtered.length === 0 && <tr><td colSpan={8} className="g-muted" style={{ textAlign: "center", padding: 18 }}>{rows.length ? "Nenhum item com esse filtro." : "Nenhum item ainda — importe a planilha “Reparo e Retorno” ou use “Novo item”."}</td></tr>}
              {filtered.map((r) => {
                const aberto = rrOpen === r.id; const recv = r.status === "RECEBIDO";
                const cor = recv ? "#22C55E" : r._d === null ? "#94A3B8" : r._d > 180 ? "#B91C1C" : r._d > 90 ? "#EF4444" : r._d > 30 ? "#F5A623" : "#22C55E";
                const doc = (lbl, v) => { const vazio = !v || /^(tba|n\/a)$/i.test(String(v).trim()); return <div key={lbl} style={{ display: "flex", gap: 6, fontSize: 11.5, lineHeight: 1.55 }}><span style={{ color: "var(--text-faint)", minWidth: 58 }}>{lbl}</span><span style={{ fontFamily: "var(--mono)", color: vazio ? "var(--text-faint)" : "var(--text)", fontWeight: vazio ? 400 : 600 }}>{v || "—"}</span></div>; };
                return (
                  <React.Fragment key={r.id}>
                    <tr className="g-row tbl-row" style={{ "--c": cor }}>
                      <td className="tbl-first"><span className="g-btn ghost" onClick={() => setRrOpen(aberto ? null : r.id)} title="Editar todos os campos">{aberto ? <ChevronDown size={14} /> : <ChevronRight size={14} />}</span></td>
                      <td style={{ whiteSpace: "normal", verticalAlign: "top" }}>
                        <div style={{ fontWeight: 700, fontSize: 13, lineHeight: 1.35 }}>{r.desc || "—"}</div>
                        <div className="tbl-chips">{r.item && <span className="tbl-chip">#{r.item}</span>}{r.sap && <span className="tbl-chip">SAP {r.sap}</span>}{r.qty && <span className="tbl-chip">Qtd {r.qty}</span>}{r.sn && <span className="tbl-chip">S/N {r.sn}</span>}</div>
                      </td>
                      <td style={{ whiteSpace: "normal", verticalAlign: "top", fontWeight: 600 }}>{r.fornecedor || <span className="g-muted">—</span>}</td>
                      <td style={{ verticalAlign: "top" }}>{doc("REQ/PO", r.po)}{doc("NF GNSI", r.nfGnsi)}{doc("NF AÇU", r.nfAcu)}{doc("PAD", r.pad)}</td>
                      <td style={{ verticalAlign: "top" }}>
                        <div style={{ fontSize: 10, color: "var(--text-faint)" }}>Enviado</div><DateOuTexto value={r.dataEnvio} onChange={(v) => upd(r.id, "dataEnvio", v)} />
                        <div style={{ fontSize: 10, color: "var(--text-faint)", marginTop: 4 }}>Recebido</div><DateOuTexto value={r.dataRecebido} onChange={(v) => upd(r.id, "dataRecebido", v)} />
                      </td>
                      <td style={{ verticalAlign: "top" }}>
                        <select value={r.status} onChange={(e) => upd(r.id, "status", e.target.value)} style={{ padding: "4px 8px", borderRadius: 99, border: `1px solid ${cor}`, background: "var(--panel-raised)", color: recv ? "var(--ok)" : "var(--warn)", fontWeight: 700, fontSize: 11.5 }}><option>ABERTO</option><option>RECEBIDO</option></select>
                        <div style={{ marginTop: 6 }}>
                          {recv ? <span style={{ fontSize: 11, color: "var(--ok)", fontWeight: 600 }}>✔ retornou</span>
                            : r._d === null ? <span style={{ fontSize: 11, color: "var(--text-faint)" }}>sem data de envio</span>
                            : <span style={{ fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 99, color: cor, background: `${cor}1f` }}>{r._d} dias em reparo</span>}
                        </div>
                      </td>
                      <td style={{ minWidth: 200, verticalAlign: "top" }}><ETextArea rows={2} value={r.obs} onChange={(v) => upd(r.id, "obs", v)} /></td>
                      <td style={{ verticalAlign: "top" }}><span className="g-btn ghost danger" onClick={() => rem(r.id)}><Trash2 size={13} /></span></td>
                    </tr>
                    {aberto && (
                      <tr className="tbl-detail"><td></td><td colSpan={7}>
                        <div className="tbl-detail-grid">
                          {[["item", "Item"], ["sap", "SAP"], ["qty", "Qtd"], ["sn", "S/N"], ["fornecedor", "Fornecedor do serviço"], ["po", "REQ/PO"], ["nfGnsi", "NF_e GNSI"], ["nfAcu", "NF_e AÇU"], ["pad", "PAD"], ["bklOut", "BKL OUT"], ["lomRec", "LOM REC"]].map(([k, l]) => (
                            <div className="g-field" key={k}><label>{l}</label><EText value={r[k] || ""} onChange={(v) => upd(r.id, k, v)} /></div>
                          ))}
                          <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Descrição do item</label><ETextArea rows={1} value={r.desc} onChange={(v) => upd(r.id, "desc", v)} /></div>
                        </div>
                      </td></tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
}

/* ============================================================
   MANUTENÇÃO › Análise de Escopo do Port Call
   Reúne tudo que pode ser feito no porto (Due, Mapeados, Failure Reports sem job,
   Reparo e Retorno), pontua, aponta pendências e sugere o que fazer.
   ============================================================ */
const ESC_CLASSES = ["Executar neste Port Call", "Resolver pendência antes", "Planejar próximo Port Call", "Acompanhar"];
const ESC_COR = { "Executar neste Port Call": "#22C55E", "Resolver pendência antes": "#EF4444", "Planejar próximo Port Call": "#F5A623", "Acompanhar": "#94A3B8" };

function AnaliseEscopoView({ tmDue, tmHistory, planningItems, tmFailures, reparoItens, allPortCallDates, portCallMeta, portCallLabel, portcallReqs, setPortcallReqs, irSolicitacao, setReportFn, setExportXlsxFn }) {
  const hoje = todayISO();
  const futuros = useMemo(() => allPortCallDates.filter((d) => d >= hoje), [allPortCallDates]);
  const [pcSel, setPcSel] = useState("");
  const pc = pcSel || futuros[0] || "";
  const idxPc = futuros.indexOf(pc);
  const proximo = idxPc >= 0 ? futuros[idxPc + 1] : null;
  const horizonte = proximo ? Math.max(1, -frDias(proximo, pc)) : 45; // dias entre este Port Call e o seguinte (padrão 45)
  const diasAtePc = pc ? -frDias(pc, hoje) : 0; // dias de hoje até o Port Call
  const limiteDias = diasAtePc + horizonte; // o que vence até o Port Call seguinte
  const duracao = pc ? (portCallMeta?.[pc]?.duration || 24) : 24;
  const [cls, setCls] = useState("Todos");
  const [orig, setOrig] = useState("Todas");
  const [msg, setMsg] = useState("");

  const lista = useMemo(() => {
    const F = tmFailures || FR_EMPTY;
    const out = [];
    const jaNoPlano = new Set(planningItems.filter((p) => p.jobNo).map((p) => `${bordoNorm(p.jobTipo)}|${bordoNorm(p.jobNo)}`));
    const isVenc = (d) => (d.diffValue !== null && d.diffValue !== undefined ? d.diffValue < 0 : !!d.dueDate && d.dueDate < hoje);
    const dias = (d) => (d.diffValue !== null && d.diffValue !== undefined && d.diffUnit === "D" ? d.diffValue : (d.dueDate ? -frDias(d.dueDate, hoje) : null)); // <0 atrasado
    const frDeJob = new Map(); (F.jobs || []).forEach((j) => { if (j.frNo) frDeJob.set(`${bordoNorm(j.jobType)}|${bordoNorm(j.jobNo)}`, j.frNo); });
    // 1) Due
    tmDue.forEach((d) => {
      const dd = dias(d); if (dd === null) return;
      const venc = isVenc(d); if (!venc && dd > limiteDias) return;
      const k = `${bordoNorm(d.jobType)}|${bordoNorm(d.jobNo)}`; const fr = frDeJob.get(k);
      const crit = !!critPrefixOf(d.code); const one = String(d.jobType || "").toUpperCase() === "ONE";
      const mot = []; let sc = venc ? 40 + Math.min(Math.abs(dd), 60) / 2 : 30 + Math.max(0, 15 - dd / 4);
      if (venc) mot.push(`vencido há ${Math.abs(dd)} d`); else mot.push(`vence em ${dd} d (antes do próximo Port Call)`);
      if (d.pri === "High") { sc += 20; mot.push("prioridade alta"); }
      if (crit) { sc += 15; mot.push("componente crítico (PMS)"); }
      if (fr) { sc += 10; mot.push(`ligado ao ${fr}`); }
      if (one) { sc += 10; mot.push("corretiva"); }
      const pend = []; if (!d.planoAcao) pend.push("sem plano de ação"); if (!d.po) pend.push("PO não informada");
      out.push({ id: `D${d.id || k}`, origem: "Due", item: d.jobName || d.component || "—", ref: `${d.jobType || ""} ${d.jobNo || ""}`.trim(), dept: d.department || "—", score: Math.round(sc), motivos: mot, pend, bloqueio: false,
        sug: venc ? "Programar a execução neste Port Call e dar baixa no TM; se não for possível, justificar e repactuar a data." : "Executar neste Port Call para não vencer antes do próximo.", tipo: one ? "Corretiva" : "Preventiva", jobTipo: d.jobType || "", jobNo: d.jobNo || "", linkado: true, componente: `${d.code || ""} ${d.component || ""}`.trim() });
    });
    // 2) Mapeados
    const pesoImp = { "Baixo": 10, "Médio": 25, "Alto": 40, "Crítico": 55 };
    planningItems.filter((p) => p.status !== "Concluído" && p.status !== "Cancelado").forEach((p) => {
      let sc = pesoImp[p.impacto || "Baixo"] || 10; const mot = [`impacto ${p.impacto || "Baixo"}`]; const pend = []; let bloq = false;
      if (p.dataExecucao && pc) { const dd = -frDias(p.dataExecucao, pc); if (Math.abs(dd) <= Math.max(1, Math.ceil(duracao / 24))) { sc += 20; mot.push("data de execução coincide com este Port Call"); } }
      if (p.status === "Em Andamento") { sc += 5; mot.push("já em andamento"); }
      if (p.precisaMaterial && !p.poMaterial) { pend.push("precisa de material e ainda sem PO"); bloq = true; }
      if (p.precisaMaterial && p.poMaterial) mot.push(`material com PO ${p.poMaterial}`);
      if (!p.jobNo) pend.push("sem job vinculado no TM"); else if (jaNoPlano.has(`${bordoNorm(p.jobTipo)}|${bordoNorm(p.jobNo)}`) && out.some((o) => o.origem === "Due" && o.ref === `${p.jobTipo} ${p.jobNo}`.trim())) { sc -= 5; mot.push(`mesmo job do Due (${p.jobTipo} ${p.jobNo})`); }
      out.push({ id: `M${p.id}`, origem: "Mapeado", item: p.nome, ref: p.jobNo ? `${p.jobTipo} ${p.jobNo}`.trim() : "", dept: p.departamento || "—", score: Math.round(sc), motivos: mot, pend, bloqueio: bloq,
        sug: bloq ? "Providenciar a PO/material antes do Port Call; sem isso não dá para executar." : (p.jobNo ? "Incluir no escopo do Port Call." : "Incluir no escopo e vincular/abrir o job no TM."), tipo: "Preventiva", jobTipo: p.jobTipo || "", jobNo: p.jobNo || "", linkado: !!p.jobNo, componente: "", planId: p.id, just: p.descricaoProblema || p.planoAcao || "" });
    });
    // 3) Failure Reports em aberto sem job
    frGaps(F, hoje).filter((g) => g.noJob).forEach((g) => {
      let sc = 45 + Math.min((g.dias || 0) / 10, 25); const mot = [`FR aberto há ${g.dias ?? "?"} d`]; const pend = ["sem job no TM"];
      if (g.noAct) pend.push("sem action"); if (g.noOrd) pend.push("sem order de material/serviço");
      if (g.n === 3) { sc += 10; mot.push("sem nenhuma tratativa"); }
      out.push({ id: `F${g.r.no}`, origem: "Failure Report", item: g.r.name || g.r.no, ref: g.r.no, dept: g.r.department || "—", score: Math.round(sc), motivos: mot, pend, bloqueio: g.noOrd && g.noAct,
        sug: "Abrir 1 job ONE no TM, definir a action e avaliar material; incluir no Port Call se a correção puder ser feita em porto.", tipo: "Corretiva", jobTipo: "", jobNo: "", linkado: false, componente: "" });
    });
    // 4) Reparo e Retorno em aberto
    (reparoItens || []).filter((r) => r.status !== "RECEBIDO").forEach((r) => {
      const dd = isIsoD(r.dataEnvio) ? frDias(r.dataEnvio, hoje) : null; let sc = 20; const mot = []; const pend = [];
      if (dd !== null) { mot.push(`em reparo há ${dd} d`); if (dd > 90) sc += 25; if (dd > 180) sc += 10; } else pend.push("sem data de envio");
      if (!r.po || /^(tba|n\/a)$/i.test(r.po)) pend.push("PO não informada");
      out.push({ id: `R${r.id}`, origem: "Reparo e Retorno", item: r.desc || r.sap, ref: r.sn || r.sap, dept: "—", score: Math.round(sc), motivos: mot, pend, bloqueio: false,
        sug: dd !== null && dd > 90 ? `Cobrar ${r.fornecedor || "o fornecedor"} e confirmar retorno até o Port Call; programar a reinstalação.` : "Acompanhar o retorno e programar a reinstalação para o Port Call.", tipo: "Corretiva", jobTipo: "", jobNo: "", linkado: false, componente: r.sap || "" });
    });
    out.forEach((o) => { o.classe = o.bloqueio ? "Resolver pendência antes" : o.score >= 60 ? "Executar neste Port Call" : o.score >= 35 ? "Planejar próximo Port Call" : "Acompanhar"; });
    return out.sort((a, b) => b.score - a.score);
  }, [tmDue, planningItems, tmFailures, reparoItens, pc, limiteDias]);

  const vis = lista.filter((x) => (cls === "Todos" || x.classe === cls) && (orig === "Todas" || x.origem === orig));
  const cont = (c) => lista.filter((x) => x.classe === c).length;
  const porDept = useMemo(() => { const m = {}; lista.filter((x) => x.classe === "Executar neste Port Call").forEach((x) => { m[x.dept] = (m[x.dept] || 0) + 1; }); return Object.entries(m).sort((a, b) => b[1] - a[1]); }, [lista]);
  const exec = lista.filter((x) => x.classe === "Executar neste Port Call");
  const sugestoes = useMemo(() => {
    const S = [];
    const vencCrit = lista.filter((x) => x.origem === "Due" && x.motivos.some((m) => m.startsWith("vencido")) && x.motivos.includes("componente crítico (PMS)")).length;
    if (vencCrit) S.push(`${vencCrit} job(s) vencido(s) de componente crítico: prioridade máxima no escopo deste Port Call.`);
    const semMat = lista.filter((x) => x.bloqueio && x.origem === "Mapeado").length;
    if (semMat) S.push(`${semMat} item(ns) mapeado(s) precisam de material sem PO: peça a PO agora, senão não serão executados no Port Call.`);
    const frSem = lista.filter((x) => x.origem === "Failure Report").length;
    if (frSem) S.push(`${frSem} Failure Report(s) em aberto sem job no TM: abra os jobs ONE e avalie levar a correção para o Port Call.`);
    const rrAnt = lista.filter((x) => x.origem === "Reparo e Retorno" && x.motivos.some((m) => /há (\d+) d/.test(m) && Number(m.match(/há (\d+)/)[1]) > 90)).length;
    if (rrAnt) S.push(`${rrAnt} item(ns) em Reparo e Retorno há mais de 90 dias: cobre os fornecedores e alinhe o retorno ao Port Call.`);
    const semJob = lista.filter((x) => x.origem === "Mapeado" && !x.linkado).length;
    if (semJob) S.push(`${semJob} item(ns) mapeado(s) sem job vinculado: vincule a um job do plano ou abra um ONE (se corretiva).`);
    const horas = exec.length;
    if (pc && horas > Math.max(8, duracao / 2)) S.push(`O Port Call tem ${duracao}h e há ${horas} itens recomendados: avalie a capacidade da equipe e a janela; considere repactuar os de menor score.`);
    if (!S.length) S.push("Sem alertas relevantes: o escopo recomendado está equilibrado.");
    return S;
  }, [lista, exec.length, duracao, pc]);

  const itemDe = (x) => ({ id: uid("PCI"), escId: x.id, descricao: x.item, componente: x.componente || "", tipo: x.tipo, justificativa: x.motivos.join("; "), linkado: x.linkado ? "sim" : "nao", jobTipo: x.jobTipo, jobNo: x.jobNo, oneNo: "", planId: x.planId });
  const jaNoPc = (x) => (portcallReqs || []).some((q) => q.portCall === pc && (q.itens || []).some((i) => i.escId === x.id || (x.planId && i.planId === x.planId) || (x.jobNo && i.jobNo === x.jobNo && i.jobTipo === x.jobTipo)));
  const addAoPortCall = (x) => {
    if (!pc) { setMsg("Escolha antes o Port Call (ou a data) analisado."); setTimeout(() => setMsg(""), 4000); return; }
    if (jaNoPc(x)) { setMsg(`“${x.item}” já está na solicitação do Port Call de ${portCallLabel(pc)}.`); setTimeout(() => setMsg(""), 4000); return; }
    setPortcallReqs((a) => {
      const lst = a || []; const i = lst.findIndex((q) => q.portCall === pc && !q.enviado);
      if (i >= 0) return lst.map((q, k) => (k === i ? { ...q, itens: [...(q.itens || []), itemDe(x)] } : q));
      return [{ id: uid("PCR"), portCall: pc, itens: [itemDe(x)], criado: hoje, enviado: "" }, ...lst];
    });
    setMsg(`“${x.item}” adicionado à solicitação do Port Call de ${portCallLabel(pc)}.`); setTimeout(() => setMsg(""), 4000);
  };
  const criarSolicitacao = () => {
    const itens = exec.slice(0, 40).map((x) => ({ id: uid("PCI"), descricao: x.item, componente: x.componente || "", tipo: x.tipo, justificativa: x.motivos.join("; "), linkado: x.linkado ? "sim" : "nao", jobTipo: x.jobTipo, jobNo: x.jobNo, oneNo: "", planId: x.planId }));
    setPortcallReqs((a) => [{ id: uid("PCR"), portCall: pc, itens, criado: hoje, enviado: "" }, ...(a || [])]);
    setMsg(`Solicitação criada com ${itens.length} item(ns) recomendado(s). Abrindo a aba Solicitação Portcall...`);
    setTimeout(() => { setMsg(""); irSolicitacao(); }, 900);
  };

  React.useEffect(() => {
    if (!setExportXlsxFn || !setReportFn) return;
    const lin = () => vis.map((x, i) => ({ "#": i + 1, Origem: x.origem, Item: x.item, Referência: x.ref, Departamento: x.dept, Score: x.score, Classificação: x.classe, Motivos: x.motivos.join("; "), Pendências: x.pend.join("; "), Sugestão: x.sug }));
    setExportXlsxFn(() => () => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(lin().length ? lin() : [{ "—": "Sem itens" }]), "Análise de Escopo"); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sugestoes.map((t) => ({ Sugestão: t }))), "Sugestões"); XLSX.writeFile(wb, `genesis-analise-escopo-${todayISO()}.xlsx`); });
    setReportFn(() => () => {
      const doc = new jsPDF({ orientation: "landscape" });
      let y = pdfHeader(doc, "Manutenção — Análise de Escopo do Port Call", `${pc ? portCallLabel(pc) : "sem Port Call futuro"} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, ESC_CLASSES.map((c) => ({ label: c, value: cont(c) })));
      y = pdfSectionTitle(doc, y, "Sugestões");
      y = pdfTable(doc, y, ["Sugestão"], sugestoes.map((t) => [t]));
      y = pdfSectionTitle(doc, y, "Escopo analisado");
      pdfTable(doc, y, ["#", "Origem", "Item", "Ref.", "Dept.", "Score", "Classificação", "Motivos", "Pendências", "Sugestão"], vis.map((x, i) => [i + 1, x.origem, x.item, x.ref || "—", x.dept, x.score, x.classe, x.motivos.join("; "), x.pend.join("; ") || "—", x.sug]));
      pdfSave(doc, "analise-escopo-portcall");
    });
  }, [vis, sugestoes, pc]);

  return (
    <>
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Análise de escopo — o que pode ser feito no Port Call</span></div>
        <div className="g-muted" style={{ fontSize: 11.5, lineHeight: 1.5, marginBottom: 10 }}>
          O portal junta tudo que é candidato a ser executado em porto: <b>jobs do Due</b> vencidos ou que vencem antes do próximo Port Call, <b>mapeados para execução</b>, <b>Failure Reports em aberto sem job</b> e <b>itens em Reparo e Retorno</b>. Cada item recebe uma pontuação (vencimento, prioridade, criticidade PMS, Failure Report, impacto, material) e uma classificação com sugestão. Itens que dependem de docagem não entram.
        </div>
        <div className="g-flex" style={{ gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div className="g-field"><label>Port Call analisado</label>
            <select value={pc} onChange={(e) => setPcSel(e.target.value)} style={{ padding: "6px 8px", border: "1px solid var(--border)", borderRadius: 6, minWidth: 220 }}>
              {futuros.length === 0 && <option value="">Sem Port Call futuro cadastrado</option>}
              {futuros.map((d) => <option key={d} value={d}>{portCallLabel(d)}</option>)}
            </select></div>
          <div className="g-field"><label>Ou informe a data</label><input type="date" value={pc} onChange={(e) => setPcSel(e.target.value)} /></div>
          <div className="g-muted" style={{ fontSize: 12 }}>{pc ? `em ${diasAtePc} dia(s) · duração ${duracao} h · considera o que vence em até ${limiteDias} dias${proximo ? " (até o Port Call seguinte)" : " (sem Port Call seguinte: horizonte de 45 dias)"}` : ""}</div>
          <button className="g-btn primary g-keep" onClick={criarSolicitacao} disabled={!exec.length || !pc} style={{ marginLeft: "auto" }}>Criar solicitação com os recomendados ({Math.min(exec.length, 40)})</button>
        </div>
      </div>
      {msg && <div className="g-alert" style={{ marginBottom: 12, background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>{msg}</div>}
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
        {ESC_CLASSES.map((c) => <div key={c} style={{ cursor: "pointer" }} onClick={() => setCls(cls === c ? "Todos" : c)}>{bigKpi(c, cont(c), ESC_COR[c], ClipboardList)}</div>)}
      </div>
      <div className="g-panel" style={{ borderLeft: "3px solid var(--teal)" }}>
        <div className="g-panel-head"><span className="g-panel-title">Sugestões</span></div>
        <ul style={{ margin: 0, paddingLeft: 18, fontSize: 13, lineHeight: 1.7 }}>{sugestoes.map((t, i) => <li key={i}>{t}</li>)}</ul>
        {porDept.length > 0 && <div className="tbl-chips" style={{ marginTop: 10 }}><span className="g-muted" style={{ fontSize: 11.5 }}>Carga recomendada por departamento:</span>{porDept.map(([k, v]) => <span key={k} className="tbl-chip">{k}: <b>{v}</b></span>)}</div>}
      </div>
      <div className="g-panel">
        <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", marginBottom: 10, alignItems: "center" }}>
          <div className="g-mode-toggle">{["Todos", ...ESC_CLASSES].map((c) => <button key={c} className={cls === c ? "active" : ""} onClick={() => setCls(c)}>{c}</button>)}</div>
          <select value={orig} onChange={(e) => setOrig(e.target.value)} style={{ padding: "6px 8px", border: "1px solid var(--border)", borderRadius: 6 }}>{["Todas", "Due", "Mapeado", "Failure Report", "Reparo e Retorno"].map((o) => <option key={o}>{o}</option>)}</select>
          <span className="g-muted" style={{ fontSize: 11.5 }}>{vis.length} de {lista.length} item(ns), do maior para o menor score</span>
        </div>
        <div className="g-table-wrap" style={{ maxHeight: 560, overflowY: "auto" }}>
          <table className="g-table tbl-modern">
            <thead><tr><th>#</th><th>Origem</th><th style={{ minWidth: 240 }}>Item</th><th>Ref.</th><th>Dept.</th><th style={{ textAlign: "right" }}>Score</th><th>Classificação</th><th style={{ minWidth: 220 }}>Por quê</th><th style={{ minWidth: 180 }}>Pendências</th><th style={{ minWidth: 260 }}>Sugestão</th><th style={{ minWidth: 150 }}>Port Call</th></tr></thead>
            <tbody>
              {vis.length === 0 && <tr><td colSpan={11} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nenhum item nessa seleção.</td></tr>}
              {vis.map((x, i) => (
                <tr className="g-row tbl-row" key={x.id} style={{ "--c": ESC_COR[x.classe] }}>
                  <td>{i + 1}</td><td><span className="tbl-chip">{x.origem}</span></td>
                  <td style={{ whiteSpace: "normal" }}><b>{x.item}</b></td>
                  <td style={{ fontFamily: "var(--mono)", whiteSpace: "nowrap" }}>{x.ref || "—"}</td><td>{x.dept}</td>
                  <td style={{ textAlign: "right", fontWeight: 700 }}>{x.score}</td>
                  <td style={{ color: ESC_COR[x.classe], fontWeight: 700, whiteSpace: "normal" }}>{x.classe}</td>
                  <td style={{ whiteSpace: "normal", fontSize: 11.5 }}>{x.motivos.join(" · ") || "—"}</td>
                  <td style={{ whiteSpace: "normal", fontSize: 11.5, color: x.pend.length ? "var(--crit)" : "var(--ok)" }}>{x.pend.join(" · ") || "✔ sem pendência"}</td>
                  <td style={{ whiteSpace: "normal", fontSize: 11.5 }}>{x.sug}</td>
                  <td>{jaNoPc(x) ? <span className="tbl-chip" style={{ color: "var(--ok)" }}>✔ no Port Call {pc ? portCallLabel(pc) : ""}</span>
                    : <button className="g-btn primary" disabled={!pc} onClick={() => addAoPortCall(x)} title="A análise está correta: adiciona este item à solicitação do Port Call da data informada acima"><Plus size={13} />Adicionar ao Port Call</button>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
}

/* ============================================================
   MANUTENÇÃO › Resumo por e-mail (Mapeados + Solicitação Port Call + Serviços + Reparo e Retorno)
   ============================================================ */
function ResumoEmailManutencao({ workPackages, planningItems, portcallReqs, reparoItens, tmDue, tmHistory, allPortCallDates, portCallLabel, setReportFn, setExportXlsxFn }) {
  const hoje = todayISO();
  const futuros = useMemo(() => allPortCallDates.filter((d) => d >= hoje), [allPortCallDates]);
  const [pc, setPc] = useState("");
  const pcAtual = pc || futuros[0] || "";
  const [para, setPara] = useState(() => { try { return localStorage.getItem("genesis_resumo_man_para") || ""; } catch (e) { return ""; } });
  const [msg, setMsg] = useState("");
  const [inc, setInc] = useState({ geral: true, mapeados: true, manuais: true, servicos: true, reparo: true });
  const rotulo = (dk) => { try { return dk ? portCallLabel(dk) : "—"; } catch (e) { return dk || "—"; } };
  const fmt = (iso) => (iso && isIsoD(iso) ? fmtDataCompleta(iso) : (iso || "—"));
  const assunto = `Resumo de Manutenção — Port Call ${rotulo(pcAtual)}`;

  const secoes = useMemo(() => {
    const reqs = portcallReqs || [];
    const reqDoPc = reqs.filter((r) => r.portCall === pcAtual);
    const solicDe = (planId) => { const r = reqDoPc.find((x) => x.itens.some((i) => i.planId === planId)) || reqs.find((x) => x.itens.some((i) => i.planId === planId)); return r || null; };
    const abertos = planningItems.filter((p) => p.status !== "Concluído" && p.status !== "Cancelado");
    const achar = (t, n) => { if (!n) return null; const k = (x) => `${bordoNorm(x.jobType)}|${bordoNorm(x.jobNo)}`; const al = `${bordoNorm(t)}|${bordoNorm(n)}`; return tmDue.find((x) => k(x) === al) || tmHistory.find((x) => k(x) === al) || null; };
    const semSolic = abertos.filter((p) => !solicDe(p.id));
    const S = [];
    const itensPc = reqDoPc.flatMap((r) => r.itens);
    S.push({ id: "geral", titulo: "Situação geral da Manutenção", explicacao: "Visão resumida do que está mapeado, do que já foi solicitado para o Port Call e do que está parado em reparo.",
      cols: ["Indicador", "Quantidade"], rows: [
        { key: "a", cells: ["Mapeados para execução (não concluídos)", abertos.length], auto: `${abertos.filter((p) => p.status === "Em Andamento").length} já em andamento` },
        { key: "b", cells: ["Mapeados já na solicitação do Port Call", abertos.length - semSolic.length], auto: semSolic.length ? `${semSolic.length} ainda fora da solicitação` : "Todos incluídos" },
        { key: "c", cells: ["Mapeados que precisam de material sem PO", abertos.filter((p) => p.precisaMaterial && !p.poMaterial).length], auto: "Sem PO não dá para executar — bordo/terra devem informar a PO" },
        { key: "d", cells: ["Mapeados sem job vinculado no TM", abertos.filter((p) => !p.jobNo).length], auto: "Vincular a um job do plano ou abrir 1 job ONE (corretiva)" },
        { key: "e", cells: [`Itens na solicitação do Port Call ${rotulo(pcAtual)}`, itensPc.length], auto: `${itensPc.filter((i) => i.linkado !== "sim" && i.tipo === "Corretiva" && !i.oneNo).length} corretiva(s) sem job ONE aberto` },
        { key: "f", cells: ["Itens em Reparo e Retorno em aberto", (reparoItens || []).filter((r) => r.status !== "RECEBIDO").length], auto: `${(reparoItens || []).filter((r) => r.status !== "RECEBIDO" && isIsoD(r.dataEnvio) && frDias(r.dataEnvio, hoje) > 90).length} há mais de 90 dias` },
      ] });
    S.push({ id: "mapeados", titulo: `Mapeados para execução e solicitação ao Port Call (${abertos.length})`,
      explicacao: "Tudo que está mapeado para execução, já cruzado com a solicitação do Port Call: a coluna “Na solicitação?” indica se o item foi pedido para o Port Call. Responder na coluna “Resposta”: confirmação de que é viável, PO do material e nº do job.",
      cols: ["Item", "Dept.", "Impacto", "Status", "Data de execução", "Material / PO", "Job vinculado", "Na solicitação?"], blankCols: ["Resposta"],
      rows: abertos.map((p) => {
        const r = solicDe(p.id); const it = r ? r.itens.find((i) => i.planId === p.id) : null; const j = achar(p.jobTipo, p.jobNo);
        const obs = []; if (p.precisaMaterial && !p.poMaterial) obs.push("INFORMAR A PO do material"); if (!p.jobNo) obs.push("sem job vinculado no TM"); else if (!j) obs.push("job não encontrado no TM");
        if (r && r.portCall !== pcAtual) obs.push(`solicitado para outro Port Call (${rotulo(r.portCall)})`); if (!r) obs.push("não está na solicitação do Port Call");
        if (it?.justificativa) obs.push(`Justificativa: ${it.justificativa}`); else if (p.justificativa) obs.push(`Justificativa: ${p.justificativa}`);
        return { key: p.id, cells: [p.nome, p.departamento || "—", p.impacto || "—", p.status, fmt(p.dataExecucao), p.precisaMaterial ? (p.poMaterial ? `PO ${p.poMaterial}` : "precisa — sem PO") : "não precisa", p.jobNo ? `${p.jobTipo || ""} ${p.jobNo}`.trim() : "—", r ? `Sim — ${rotulo(r.portCall)}` : "Não"], auto: obs.join(" · ") || "Sem pendência" };
      }) });
    const manuais = itensPc.filter((i) => !i.planId);
    S.push({ id: "manuais", titulo: `Outros itens solicitados para o Port Call (${manuais.length})`, explicacao: "Itens da solicitação que não vieram dos Mapeados. Corretivas sem job no plano devem ter 1 job ONE aberto no TM.",
      cols: ["Manutenção", "Componente", "Tipo", "Ligada a job do plano?", "Job ONE"],
      rows: manuais.map((i) => ({ key: i.id, cells: [i.descricao || "—", i.componente || "—", i.tipo, i.linkado === "sim" ? `Sim — ${`${i.jobTipo} ${i.jobNo}`.trim() || "job não informado"}` : "Não", i.linkado !== "sim" && i.tipo === "Corretiva" ? (i.oneNo ? `ONE ${i.oneNo}` : "ABRIR no TM") : "—"], auto: i.justificativa ? `Justificativa: ${i.justificativa}` : "Sem justificativa" })) });
    const svc = workPackages.filter((w) => (w.start || "").slice(0, 10) === pcAtual);
    S.push({ id: "servicos", titulo: `Serviços programados para o Port Call (${svc.length})`, explicacao: "Serviços cadastrados em Manutenção › Serviços com início na data do Port Call.",
      cols: ["Manutenção", "Categoria", "Empresa", "RC", "Status", "Progresso"],
      rows: svc.map((w) => ({ key: w.id, cells: [w.name, w.group || "—", w.empresa || "—", w.rc || "—", w.status, `${w.progress ?? 0}%`], auto: w.obs || (w.status === "Concluído" ? "Concluído" : "Acompanhar") })) });
    const rr = (reparoItens || []).filter((r) => r.status !== "RECEBIDO").map((r) => ({ r, d: isIsoD(r.dataEnvio) ? frDias(r.dataEnvio, hoje) : null })).sort((a, b) => (b.d ?? -1) - (a.d ?? -1));
    S.push({ id: "reparo", titulo: `Reparo e Retorno em aberto (${rr.length})`, explicacao: "Itens enviados para reparo que ainda não retornaram, do mais antigo para o mais recente. Para cada item, confirmar previsão de retorno.",
      cols: ["Item", "SAP", "S/N", "Fornecedor", "REQ/PO", "Enviado", "Dias em reparo"], blankCols: ["Previsão de retorno"],
      rows: rr.map(({ r, d }) => ({ key: r.id, cells: [r.desc || "—", r.sap || "—", r.sn || "—", r.fornecedor || "—", r.po || "—", fmt(r.dataEnvio), d ?? "—"], auto: d !== null && d > 90 ? `Em reparo há mais de 90 dias — cobrar ${r.fornecedor || "fornecedor"}${r.obs ? ` · ${r.obs}` : ""}` : (r.obs || "Acompanhar retorno") })) });
    return S.filter((x) => inc[x.id]);
  }, [planningItems, portcallReqs, reparoItens, workPackages, tmDue, tmHistory, pcAtual, inc]);
  secoes.forEach((x, i) => { x._t = `${i + 1}. ${x.titulo}`; });

  const montarHtml = () => {
    const th = (t) => `<th style="border:1px solid #1F3A5F;background:#1F3A5F;color:#fff;padding:5px 7px;font-size:12px;text-align:left">${escHtml(t)}</th>`;
    const td = (t) => `<td style="border:1px solid #cfd8e6;padding:5px 7px;font-size:12px;vertical-align:top">${escHtml(t)}</td>`;
    return `<div style="font-family:Calibri,Arial,sans-serif;color:#1b2430"><h2 style="margin:0 0 4px">${escHtml(assunto)}</h2><p style="margin:0 0 6px;font-size:13px">Prezados, segue o resumo de manutenção para o Port Call de <b>${escHtml(rotulo(pcAtual))}</b>, com os itens mapeados já cruzados com a solicitação ao Port Call.</p>`
      + secoes.map((s) => `<h3 style="margin:16px 0 2px;font-size:15px;color:#1F3A5F">${escHtml(s._t)}</h3><p style="margin:0 0 6px;font-size:12px;color:#51607a">${escHtml(s.explicacao)}</p><table style="border-collapse:collapse;width:100%"><tr>${[...s.cols, "Observação", ...(s.blankCols || [])].map(th).join("")}</tr>${s.rows.length ? s.rows.map((r) => `<tr>${[...r.cells, r.auto, ...(s.blankCols || []).map(() => "")].map(td).join("")}</tr>`).join("") : `<tr>${td("(nenhum item)")}</tr>`}</table>`).join("") + "</div>";
  };
  const montarTexto = () => [assunto, ""].concat(secoes.flatMap((s) => [s._t, s.explicacao, [...s.cols, "Observação", ...(s.blankCols || [])].join(" | "), ...(s.rows.length ? s.rows.map((r) => [...r.cells, r.auto, ...(s.blankCols || []).map(() => "______")].join(" | ")) : ["(nenhum item)"]), ""])).join("\n");
  const copiar = async () => {
    const html = montarHtml(); const txt = montarTexto();
    try {
      if (window.ClipboardItem && navigator.clipboard.write) await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([txt], { type: "text/plain" }) })]);
      else await navigator.clipboard.writeText(txt);
      setMsg("Resumo copiado com as tabelas. Cole no corpo do e-mail (Ctrl+V).");
    } catch (e) { setMsg("Não consegui copiar automaticamente — selecione a prévia abaixo e copie com Ctrl+C."); }
    setTimeout(() => setMsg(""), 7000);
  };
  const abrirEmail = async () => {
    try { localStorage.setItem("genesis_resumo_man_para", para); } catch (e) { /* ignora */ }
    await copiar();
    const corpo = "Prezados,\n\nSegue o resumo de manutenção.\n\n[Cole aqui o resumo com Ctrl+V — as tabelas já estão copiadas]\n";
    window.location.href = `mailto:${para.split(/[;,]/).map((x) => x.trim()).filter(Boolean).join(",")}?subject=${encodeURIComponent(assunto)}&body=${encodeURIComponent(corpo)}`;
  };

  React.useEffect(() => {
    if (!setExportXlsxFn || !setReportFn) return;
    setExportXlsxFn(() => () => { const wb = XLSX.utils.book_new(); secoes.forEach((s, i) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(s.rows.length ? s.rows.map((r) => Object.fromEntries([...s.cols.map((c, k) => [c, r.cells[k]]), ["Observação", r.auto]])) : [{ "—": "Sem itens" }]), `${i + 1}-${s.id}`.slice(0, 31))); XLSX.writeFile(wb, `genesis-resumo-manutencao-${todayISO()}.xlsx`); });
    setReportFn(() => () => {
      const doc = new jsPDF({ orientation: "landscape" });
      let y = pdfHeader(doc, "Manutenção — Resumo", `${assunto} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      secoes.forEach((s) => { y = pdfSectionTitle(doc, y, s._t); y = pdfTable(doc, y, [...s.cols, "Observação"], s.rows.map((r) => [...r.cells, r.auto])); });
      pdfSave(doc, "resumo-manutencao");
    });
  }, [secoes, assunto]);

  const nomes = { geral: "Situação geral", mapeados: "Mapeados + Solicitação Port Call", manuais: "Outros itens solicitados", servicos: "Serviços do Port Call", reparo: "Reparo e Retorno" };
  return (
    <>
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Resumo por e-mail — Manutenção</span></div>
        <div className="g-muted" style={{ fontSize: 11.5, lineHeight: 1.5, marginBottom: 10 }}>
          Um único e-mail com <b>Mapeados para Execução já cruzados com a Solicitação ao Port Call</b>, os serviços do Port Call e o Reparo e Retorno. Cada tabela tem a coluna <b>Observação</b> (pendências) e, onde precisa de resposta, uma coluna em branco. O e-mail não carrega tabelas pelo link, então o portal <b>copia o resumo</b> e abre o e-mail: é só colar (Ctrl+V).
        </div>
        <div className="g-flex" style={{ gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div className="g-field"><label>Port Call</label>
            <select value={pcAtual} onChange={(e) => setPc(e.target.value)} style={{ padding: "6px 8px", border: "1px solid var(--border)", borderRadius: 6, minWidth: 200 }}>
              {pcAtual && !allPortCallDates.includes(pcAtual) && <option value={pcAtual}>{pcAtual}</option>}
              {allPortCallDates.map((d) => <option key={d} value={d}>{rotulo(d)}</option>)}
            </select></div>
          <div className="g-field"><label>Ou data</label><input type="date" value={pcAtual} onChange={(e) => setPc(e.target.value)} /></div>
          <div className="g-field" style={{ flex: 1, minWidth: 260 }}><label>Destinatários (separe por vírgula)</label><input type="text" value={para} onChange={(e) => setPara(e.target.value)} placeholder="email1@..., email2@..." style={{ width: "100%", padding: "6px 8px", border: "1px solid var(--border)", borderRadius: 6 }} /></div>
          <button className="g-btn g-keep" onClick={copiar}>Copiar resumo</button>
          <button className="g-btn primary g-keep" onClick={abrirEmail}>Abrir e-mail</button>
        </div>
        <div className="tbl-chips" style={{ marginTop: 10 }}>
          {Object.keys(nomes).map((k) => <label key={k} className="tbl-check" style={{ marginRight: 12 }}><input type="checkbox" checked={inc[k]} onChange={(e) => setInc((p) => ({ ...p, [k]: e.target.checked }))} />{nomes[k]}</label>)}
        </div>
      </div>
      {msg && <div className="g-alert" style={{ marginBottom: 12, background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>{msg}</div>}
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Prévia do e-mail</span></div>
        <div style={{ background: "#fff", color: "#1b2430", padding: 12, borderRadius: 6, overflowX: "auto" }} dangerouslySetInnerHTML={{ __html: montarHtml() }} />
      </div>
    </>
  );
}

/* ============================================================
   MANUTENÇÃO › Plano de Ação com Bordo
   Lista geral: Mapeados para Execução + Itens DNV (Docagem), cada um com ação, responsável, prazo e resposta do bordo.
   ============================================================ */
const AB_STATUS = ["Pendente", "Enviado ao bordo", "Respondido pelo bordo", "Em andamento", "Concluído"];
const AB_RESP = ["Bordo", "Terra", "Fornecedor", "Bordo + Terra"];
const AB_COR = { "Pendente": "#EF4444", "Enviado ao bordo": "#F5A623", "Respondido pelo bordo": "#3B82F6", "Em andamento": "#8B5CF6", "Concluído": "#22C55E" };

/* atualiza o registro do plano de ação com bordo (status, resposta, responsável, prazo) de um item;
   usado tanto na aba Plano de Ação quanto dentro dos Mapeados/DNV, para ficarem sempre iguais */
function updPlanoBordoKey(setPlanosBordo, key, campo, valor) {
  setPlanosBordo((arr) => {
    const a = arr || []; const i = a.findIndex((x) => x.key === key);
    const ext = campo === "resposta" && valor && !(a[i]?.status && a[i].status !== "Pendente" && a[i].status !== "Enviado ao bordo") ? { status: "Respondido pelo bordo" } : {};
    const novo = { ...(i >= 0 ? a[i] : { key }), [campo]: valor, ...ext, atual: todayISO() };
    return i >= 0 ? a.map((x, k) => (k === i ? novo : x)) : [...a, novo];
  });
}

function PlanoAcaoBordoView({ planningItems, docagemItems, planosBordo, setPlanosBordo, updPlan, remPlan, updDocagem, remDocagem, setReportFn, setExportXlsxFn }) {
  const hoje = todayISO();
  const [f, setF] = useState({ origem: "Todas", status: "Abertos", resp: "Todos", busca: "" });
  const [para, setPara] = useState(() => { try { return localStorage.getItem("genesis_acao_bordo_para") || ""; } catch (e) { return ""; } });
  const [msg, setMsg] = useState("");
  const planos = planosBordo || [];

  const base = useMemo(() => {
    const out = [];
    planningItems.filter((p) => p.status !== "Concluído" && p.status !== "Cancelado").forEach((p) => out.push({ key: `M:${p.id}`, src: "M", id: p.id, origem: "Mapeado", nome: p.nome, localRaw: p.departamento || "", local: p.departamento || "—", ref: p.jobNo ? `${p.jobTipo || ""} ${p.jobNo}`.trim() : "", acaoBase: p.planoAcao || "", material: p.precisaMaterial ? (p.poMaterial ? `PO ${p.poMaterial}` : "precisa — sem PO") : "não precisa", prazoBase: p.dataExecucao || "", empresa: p.empresa || "" }));
    docagemItems.filter((d) => !d.dataConclusao && !/conclu|cancel/i.test(d.status || "")).forEach((d) => out.push({ key: `D:${d.id}`, src: "D", id: d.id, origem: "DNV", nome: d.nome, localRaw: d.localizacao || "", local: d.localizacao || "—", ref: d.tipoPeriodo || "", acaoBase: d.planoAcao || "", material: d.necessitaMaterial ? (d.poRelacionada ? `PO ${d.poRelacionada}` : "precisa — sem PO") : "não precisa", prazoBase: d.previsaoExecucao || "", empresa: d.empresa || "" }));
    return out;
  }, [planningItems, docagemItems]);
  const rows = useMemo(() => base.map((b) => {
    const p = planos.find((x) => x.key === b.key) || {};
    const acao = b.acaoBase;
    const prazo = p.prazo !== undefined ? p.prazo : (isIsoD(b.prazoBase) ? b.prazoBase : "");
    const status = p.status || "Pendente";
    return { ...b, acao, prazo, status, resp: p.resp || "Bordo", resposta: p.resposta || "", atual: p.atual || "", atrasado: !!prazo && prazo < hoje && status !== "Concluído" };
  }), [base, planos]);

  const upd = (key, campo, valor) => updPlanoBordoKey(setPlanosBordo, key, campo, valor);
  /* texto do item (nome, setor/local, ação) é gravado no próprio item dos Mapeados / DNV: o que muda aqui muda lá */
  const updItem = (r, campo, valor) => {
    if (r.src === "M") { const i = planningItems.findIndex((x) => x.id === r.id); if (i >= 0) updPlan(i, campo === "local" ? "departamento" : campo === "acao" ? "planoAcao" : "nome", valor); }
    else { const i = docagemItems.findIndex((x) => x.id === r.id); if (i >= 0) updDocagem(i, campo === "local" ? "localizacao" : campo === "acao" ? "planoAcao" : "nome", valor); }
  };
  const excluirItem = (r) => {
    if (!window.confirm(`Excluir "${r.nome}"? O item também será excluído de ${r.src === "M" ? "Mapeados para Execução" : "Itens DNV (Docagem)"}.`)) return;
    if (r.src === "M") { const i = planningItems.findIndex((x) => x.id === r.id); if (i >= 0) remPlan(i); }
    else { const i = docagemItems.findIndex((x) => x.id === r.id); if (i >= 0) remDocagem(i); }
    setPlanosBordo((arr) => (arr || []).filter((x) => x.key !== r.key));
  };
  /* migração única: ação digitada antes só no plano passa para o item; e limpa registros de itens que já foram excluídos */
  React.useEffect(() => {
    (planosBordo || []).forEach((x) => {
      if (x.acao === undefined || !x.acao) return;
      if (x.key.startsWith("M:")) { const i = planningItems.findIndex((p) => `M:${p.id}` === x.key); if (i >= 0 && !planningItems[i].planoAcao) updPlan(i, "planoAcao", x.acao); }
      else { const i = docagemItems.findIndex((d) => `D:${d.id}` === x.key); if (i >= 0 && !docagemItems[i].planoAcao) updDocagem(i, "planoAcao", x.acao); }
    });
    const vivos = new Set([...planningItems.map((p) => `M:${p.id}`), ...docagemItems.map((d) => `D:${d.id}`)]);
    if ((planosBordo || []).some((x) => x.acao !== undefined || !vivos.has(x.key))) setPlanosBordo((arr) => (arr || []).filter((x) => vivos.has(x.key)).map(({ acao, ...rest }) => rest));
  }, [planosBordo, planningItems, docagemItems]);

  const filtered = rows.filter((r) => (f.origem === "Todas" || r.origem === f.origem) && (f.status === "Todos" || (f.status === "Abertos" ? r.status !== "Concluído" : r.status === f.status)) && (f.resp === "Todos" || r.resp === f.resp)
    && (!f.busca.trim() || `${r.nome} ${r.local} ${r.ref} ${r.acao} ${r.resposta}`.toLowerCase().includes(f.busca.trim().toLowerCase())))
    .sort((a, b) => (b.atrasado - a.atrasado) || ((a.status === "Pendente" ? 0 : 1) - (b.status === "Pendente" ? 0 : 1)));
  const kpi = { total: rows.length, semAcao: rows.filter((r) => !r.acao.trim()).length, aguardando: rows.filter((r) => r.status === "Enviado ao bordo").length, resp: rows.filter((r) => r.status === "Respondido pelo bordo").length, atras: rows.filter((r) => r.atrasado).length, conc: planos.filter((x) => x.status === "Concluído" && base.some((b) => b.key === x.key)).length };

  const paraBordo = rows.filter((r) => /Bordo/.test(r.resp) && r.status !== "Concluído");
  const assunto = `Plano de ação com o bordo — ${paraBordo.length} item(ns) pendente(s)`;
  const cols = ["Origem", "Item", "Setor / Local", "Ref.", "Ação solicitada ao bordo", "Material / PO", "Prazo", "Status"];
  const linha = (r) => [r.origem, r.nome, r.local, r.ref || "—", r.acao || "(definir ação)", r.material, r.prazo ? fmtDataCompleta(r.prazo) : "—", r.status];
  const montarHtml = () => {
    const th = (t) => `<th style="border:1px solid #1F3A5F;background:#1F3A5F;color:#fff;padding:5px 7px;font-size:12px;text-align:left">${escHtml(t)}</th>`;
    const td = (t) => `<td style="border:1px solid #cfd8e6;padding:5px 7px;font-size:12px;vertical-align:top">${escHtml(t)}</td>`;
    return `<div style="font-family:Calibri,Arial,sans-serif;color:#1b2430"><h2 style="margin:0 0 4px">${escHtml(assunto)}</h2><p style="margin:0 0 8px;font-size:13px">Prezados, abaixo o plano de ação dos itens mapeados para execução e dos itens da DNV (docagem). Para <b>cada item</b>, favor responder na coluna “Resposta do bordo”: situação atual, o que é possível fazer a bordo, material necessário (nº da PO) e data prevista.</p>
<table style="border-collapse:collapse;width:100%"><tr>${[...cols, "Resposta do bordo", "Data prevista"].map(th).join("")}</tr>${paraBordo.length ? paraBordo.map((r) => `<tr>${[...linha(r), "", ""].map(td).join("")}</tr>`).join("") : `<tr>${td("(nenhum item)")}</tr>`}</table></div>`;
  };
  const montarTexto = () => [assunto, "", [...cols, "Resposta do bordo", "Data prevista"].join(" | "), ...paraBordo.map((r) => [...linha(r), "______", "______"].join(" | "))].join("\n");
  const copiar = async () => {
    const html = montarHtml(); const txt = montarTexto();
    try {
      if (window.ClipboardItem && navigator.clipboard.write) await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([txt], { type: "text/plain" }) })]);
      else await navigator.clipboard.writeText(txt);
      setMsg("Plano copiado com a tabela. Cole no corpo do e-mail (Ctrl+V).");
    } catch (e) { setMsg("Não consegui copiar automaticamente — copie a prévia com Ctrl+C."); }
    setTimeout(() => setMsg(""), 7000);
  };
  const abrirEmail = async () => {
    try { localStorage.setItem("genesis_acao_bordo_para", para); } catch (e) { /* ignora */ }
    await copiar();
    setPlanosBordo((arr) => { const a = [...(arr || [])]; paraBordo.forEach((r) => { if (r.status === "Pendente") { const i = a.findIndex((x) => x.key === r.key); const n = { ...(i >= 0 ? a[i] : { key: r.key }), status: "Enviado ao bordo", atual: hoje }; if (i >= 0) a[i] = n; else a.push(n); } }); return a; });
    const corpo = "Prezados,\n\nSegue o plano de ação com o bordo.\n\n[Cole aqui o plano com Ctrl+V — a tabela já está copiada]\n";
    window.location.href = `mailto:${para.split(/[;,]/).map((x) => x.trim()).filter(Boolean).join(",")}?subject=${encodeURIComponent(assunto)}&body=${encodeURIComponent(corpo)}`;
  };

  React.useEffect(() => {
    if (!setExportXlsxFn || !setReportFn) return;
    setExportXlsxFn(() => () => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filtered.length ? filtered.map((r) => ({ Origem: r.origem, Item: r.nome, "Setor / Local": r.local, Referência: r.ref, Ação: r.acao, Responsável: r.resp, "Material / PO": r.material, Prazo: r.prazo, Status: r.status, "Resposta do bordo": r.resposta, "Atualizado em": r.atual })) : [{ "—": "Sem itens" }]), "Plano de Ação"); XLSX.writeFile(wb, `genesis-plano-acao-bordo-${todayISO()}.xlsx`); });
    setReportFn(() => () => {
      const doc = new jsPDF({ orientation: "landscape" });
      let y = pdfHeader(doc, "Manutenção — Plano de Ação com Bordo", `${filtered.length} item(ns) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [{ label: "Itens", value: kpi.total }, { label: "Sem ação definida", value: kpi.semAcao }, { label: "Aguardando bordo", value: kpi.aguardando }, { label: "Respondidos", value: kpi.resp }, { label: "Prazo vencido", value: kpi.atras }]);
      pdfTable(doc, y, ["Origem", "Item", "Local", "Ação", "Resp.", "Prazo", "Status", "Resposta do bordo"], filtered.map((r) => [r.origem, r.nome, r.local, r.acao, r.resp, r.prazo || "—", r.status, r.resposta]));
      pdfSave(doc, "plano-acao-bordo");
    });
  }, [filtered]);

  const sel = { padding: "4px 6px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 11.5, background: "var(--panel, #fff)", color: "var(--text)" };
  return (
    <>
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Plano de ação com o bordo — Mapeados + Itens DNV</span></div>
        <div className="g-muted" style={{ fontSize: 11.5, lineHeight: 1.5, marginBottom: 10 }}>
          Lista geral de tudo que está em <b>Mapeados para Execução</b> e nos <b>Itens DNV (Docagem)</b>, ainda não concluído. Para cada item defina a <b>ação</b>, o <b>responsável</b> e o <b>prazo</b>; o botão de e-mail envia ao bordo a tabela dos itens sob responsabilidade do bordo e marca como “Enviado ao bordo”. Quando o bordo responder, registre a <b>resposta</b> (o status vira “Respondido pelo bordo”) e acompanhe até Concluído. Item, setor e ação são <b>os mesmos dos Mapeados / DNV</b>: o que você editar ou excluir aqui muda lá, e vice-versa. A ação vem do “Plano de Ação” já preenchido no item.
        </div>
        <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <input type="text" value={para} onChange={(e) => setPara(e.target.value)} placeholder="E-mail do bordo (vários separados por vírgula)" style={{ flex: 1, minWidth: 280, padding: "7px 10px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 13 }} />
          <button className="g-btn g-keep" onClick={copiar}>Copiar tabela</button>
          <button className="g-btn primary g-keep" onClick={abrirEmail} disabled={!paraBordo.length}>Enviar ao bordo ({paraBordo.length})</button>
        </div>
      </div>
      {msg && <div className="g-alert" style={{ marginBottom: 12, background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>{msg}</div>}
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
        {bigKpi("Itens no plano", kpi.total, "var(--teal)", ClipboardList)}
        {bigKpi("Sem ação definida", kpi.semAcao, kpi.semAcao ? "var(--crit)" : "var(--ok)", AlertTriangle)}
        {bigKpi("Aguardando o bordo", kpi.aguardando, "var(--warn)", Clock)}
        {bigKpi("Respondidos pelo bordo", kpi.resp, "var(--teal)", ClipboardList)}
        {bigKpi("Prazo vencido", kpi.atras, kpi.atras ? "var(--crit)" : "var(--ok)", AlertTriangle)}
      </div>
      <div className="g-panel">
        <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", marginBottom: 10, alignItems: "center" }}>
          <div className="g-mode-toggle">{["Todas", "Mapeado", "DNV"].map((o) => <button key={o} className={f.origem === o ? "active" : ""} onClick={() => setF((p) => ({ ...p, origem: o }))}>{o}</button>)}</div>
          <select value={f.status} onChange={(e) => setF((p) => ({ ...p, status: e.target.value }))} style={sel}>{["Abertos", "Todos", ...AB_STATUS].map((x) => <option key={x}>{x}</option>)}</select>
          <select value={f.resp} onChange={(e) => setF((p) => ({ ...p, resp: e.target.value }))} style={sel}>{["Todos", ...AB_RESP].map((x) => <option key={x}>{x}</option>)}</select>
          <input value={f.busca} onChange={(e) => setF((p) => ({ ...p, busca: e.target.value }))} placeholder="buscar item, ação, resposta..." style={{ ...sel, minWidth: 220, padding: "6px 10px" }} />
          <span className="g-muted" style={{ fontSize: 11.5 }}>{filtered.length} de {rows.length}</span>
        </div>
        <div className="g-table-wrap">
          <table className="g-table tbl-modern" style={{ minWidth: 0, width: "100%" }}>
            <thead><tr><th>Origem</th><th style={{ minWidth: 220 }}>Item</th><th style={{ minWidth: 240 }}>Ação</th><th>Responsável</th><th>Prazo</th><th>Status</th><th style={{ minWidth: 240 }}>Resposta do bordo</th><th></th></tr></thead>
            <tbody>
              {filtered.length === 0 && <tr><td colSpan={8} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nenhum item nesse filtro.</td></tr>}
              {filtered.map((r) => (
                <tr className="g-row tbl-row" key={r.key} style={{ "--c": r.atrasado ? "#EF4444" : AB_COR[r.status] }}>
                  <td style={{ verticalAlign: "top" }}><span className="tbl-chip">{r.origem}</span></td>
                  <td style={{ whiteSpace: "normal", verticalAlign: "top" }}><ETextArea rows={1} value={r.nome} onChange={(v) => updItem(r, "nome", v)} /><div className="tbl-chips"><span className="tbl-chip" title={r.origem === "DNV" ? "Localização (igual na DNV)" : "Departamento (igual nos Mapeados)"}><EText value={r.localRaw} onChange={(v) => updItem(r, "local", v)} /></span>{r.ref && <span className="tbl-chip">{r.ref}</span>}<span className="tbl-chip" style={{ color: /sem PO/.test(r.material) ? "var(--crit)" : undefined }}>{r.material}</span></div></td>
                  <td style={{ verticalAlign: "top" }}><ETextArea rows={2} value={r.acao} onChange={(v) => updItem(r, "acao", v)} />{!r.acao.trim() && <div style={{ fontSize: 10.5, color: "var(--crit)" }}>definir ação</div>}</td>
                  <td style={{ verticalAlign: "top" }}><select value={r.resp} onChange={(e) => upd(r.key, "resp", e.target.value)} style={sel}>{AB_RESP.map((x) => <option key={x}>{x}</option>)}</select></td>
                  <td style={{ verticalAlign: "top" }}><DateOuTexto value={r.prazo} onChange={(v) => upd(r.key, "prazo", v)} />{r.atrasado && <div style={{ fontSize: 10.5, color: "var(--crit)", fontWeight: 700 }}>vencido</div>}</td>
                  <td style={{ verticalAlign: "top" }}><select value={r.status} onChange={(e) => upd(r.key, "status", e.target.value)} style={{ ...sel, borderColor: AB_COR[r.status], color: AB_COR[r.status], fontWeight: 700 }}>{AB_STATUS.map((x) => <option key={x}>{x}</option>)}</select>{r.atual && <div style={{ fontSize: 10, color: "var(--text-faint)" }}>atualizado {fmtDataCompleta(r.atual)}</div>}</td>
                  <td style={{ verticalAlign: "top" }}><ETextArea rows={2} value={r.resposta} onChange={(v) => upd(r.key, "resposta", v)} /></td>
                  <td style={{ verticalAlign: "top" }}><span className="g-btn ghost danger" title={`Excluir também de ${r.src === "M" ? "Mapeados" : "DNV"}`} onClick={() => excluirItem(r)}><Trash2 size={13} /></span></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
}

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
      ]);
      y = pdfSectionTitle(doc, y, "Serviços");
      pdfTable(doc, y,
        ["Data", "Categoria", "Manutenção", "Empresa", "RC", "Status", "Progresso"],
        filtered.map((w) => [fmtDate((w.start || "").slice(0, 10)), w.group || "—", w.name, w.empresa || "—", w.rc || "—", w.status, `${w.progress}%`])
      );
      pdfSave(doc, "relatorio-servicos");
    });
  }, [filtered, concluidos, emAndamento, naoIniciados, cancelados, taxaConclusao, setReportFn]);

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
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(6, 1fr)" }}>
        {bigKpi("Total de Serviços", filtered.length, "var(--teal)", Wrench)}
        {bigKpi("Concluídos", concluidos, "var(--ok)", Wrench)}
        {bigKpi("Em Andamento", emAndamento, "var(--teal)", Wrench)}
        {bigKpi("Não Iniciados", naoIniciados, "var(--text-dim)", Wrench)}
        {bigKpi("Cancelados", cancelados, "var(--text-dim)", X)}
        {bigKpi("Taxa de Conclusão", `${taxaConclusao}%`, "var(--warn)", LayoutGrid)}
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
              <th style={{ width: 90 }}></th>
            </tr>
          </thead>
          <tbody>
            {wpGroups.map((g) => (
              <React.Fragment key={g.key}>
                {groupByPc && (
                  <tr className="tbl-group" onClick={() => toggleWpGroup(g.key)}>
                    <td colSpan={6}>
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
                      <td colSpan={5} style={{ padding: "10px 8px 16px 8px" }}>
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
/* ============================================================
   FUNIL DE MANUTENÇÃO — 3 etapas sobre a mesma lista (Mapeados + Itens DNV)
   1 Mapeamento total  →  2 Análise do porto  →  3 Cronograma (confirmar cria/atualiza o Serviço)
   ============================================================ */
const FUNIL_ORIGENS = ["Bordo", "DNV", "TM Master", "Inspeção", "Outro"];
const FUNIL_DECISOES = ["Faz no porto", "Precisa de parada", "Só docagem", "Aguardando material", "Descartar"];
const FUNIL_DEC_COR = { "Faz no porto": "var(--ok)", "Precisa de parada": "var(--warn)", "Só docagem": "var(--teal)", "Aguardando material": "var(--warn)", "Descartar": "var(--text-faint)" };
const funilAberto = (r) => r.status !== "Concluído" && r.status !== "Cancelado";
const funilIni = (pc) => `${pc}T08:00`;
const funilSomaH = (iniStr, h) => { const d = new Date(iniStr); d.setMinutes(d.getMinutes() + Math.round((Number(h) || 9) * 60)); const p = (n) => String(n).padStart(2, "0"); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`; };

function FunilManutencaoView({ stage, setStage, planningItems, setPlanningItems, docagemItems, setDocagemItems, allPortCallDates, portCallLabel, portCallMeta, setReportFn, setExportXlsxFn }) {
  const hoje = todayISO();
  const rows = useMemo(() => [
    ...planningItems.map((p) => ({ key: "M:" + p.id, src: "M", id: p.id, origem: p.origem || "Bordo", nome: p.nome, area: p.departamento || "", empresa: p.empresa || "", plano: p.planoAcao || "", impacto: p.impacto || "Baixo", status: p.status || "A Executar", material: !!p.precisaMaterial, po: p.poMaterial || "", decisao: p.decisao || "", duracaoH: p.duracaoH || "", analise: p.analise || "", pc: p.cronoPC || "", ordem: p.cronoOrdem || 0, ini: p.cronoIni || "", fim: p.cronoFim || "", conf: !!p.cronoConf, linked: p.linkedServiceId })),
    ...docagemItems.map((d) => ({ key: "D:" + d.id, src: "D", id: d.id, origem: "DNV", nome: d.nome, area: d.localizacao || "", empresa: d.empresa || "", plano: d.planoAcao || "", impacto: d.impacto || "—", status: d.status || "A Executar", material: !!d.necessitaMaterial, po: d.poRelacionada || "", decisao: d.decisao || "", duracaoH: d.duracaoH || "", analise: d.analise || "", pc: d.cronoPC || "", ordem: d.cronoOrdem || 0, ini: d.cronoIni || "", fim: d.cronoFim || "", conf: !!d.cronoConf, linked: d.linkedServiceId })),
  ], [planningItems, docagemItems]);
  const MAPA_M = { area: "departamento", plano: "planoAcao", material: "precisaMaterial", po: "poMaterial", ini: "cronoIni", fim: "cronoFim", pc: "cronoPC", ordem: "cronoOrdem", conf: "cronoConf", decisao: "decisao", duracaoH: "duracaoH", analise: "analise", origem: "origem" };
  const MAPA_D = { area: "localizacao", plano: "planoAcao", material: "necessitaMaterial", po: "poRelacionada", ini: "cronoIni", fim: "cronoFim", pc: "cronoPC", ordem: "cronoOrdem", conf: "cronoConf", decisao: "decisao", duracaoH: "duracaoH", analise: "analise" };
  const set = (r, patch) => {
    const mapa = r.src === "M" ? MAPA_M : MAPA_D;
    const real = Object.fromEntries(Object.entries(patch).map(([k, v]) => [mapa[k] || k, v]));
    (r.src === "M" ? setPlanningItems : setDocagemItems)((arr) => arr.map((x) => (x.id === r.id ? { ...x, ...real } : x)));
  };
  const apagar = (r) => { if (!window.confirm(`Excluir "${r.nome}"?`)) return; (r.src === "M" ? setPlanningItems : setDocagemItems)((arr) => arr.filter((x) => x.id !== r.id)); };
  const inp = { padding: "4px 6px", border: "1px solid var(--border)", borderRadius: 4, fontSize: 12, background: "var(--panel, #fff)", color: "var(--text)", width: "100%" };
  const abertos = rows.filter(funilAberto);
  const norm = (v) => String(v || "").toLowerCase();

  /* ---------- filtros comuns ---------- */
  const [f, setF] = useState({ busca: "", origem: "Todas", dec: "Todas", area: "" });
  const lista = abertos.filter((r) => (f.origem === "Todas" || r.origem === f.origem) && (f.dec === "Todas" || (f.dec === "Sem decisão" ? !r.decisao : r.decisao === f.dec)) && (!f.busca || norm(`${r.nome} ${r.area} ${r.empresa} ${r.plano}`).includes(norm(f.busca))));
  const filtros = (comDec) => (
    <div className="g-filterbar" style={{ padding: "10px 16px", marginBottom: 12, borderRadius: 6 }}>
      <div className="g-field"><label>Buscar</label><input type="text" value={f.busca} onChange={(e) => setF((p) => ({ ...p, busca: e.target.value }))} placeholder="nome, área, empresa…" /></div>
      <div className="g-field"><label>Origem</label><select value={f.origem} onChange={(e) => setF((p) => ({ ...p, origem: e.target.value }))}>{["Todas", ...FUNIL_ORIGENS].map((o) => <option key={o}>{o}</option>)}</select></div>
      {comDec && <div className="g-field"><label>Decisão</label><select value={f.dec} onChange={(e) => setF((p) => ({ ...p, dec: e.target.value }))}>{["Todas", "Sem decisão", ...FUNIL_DECISOES].map((o) => <option key={o}>{o}</option>)}</select></div>}
    </div>
  );

  /* ---------- cabeçalho com as 3 etapas ---------- */
  const nSem = abertos.filter((r) => !r.decisao).length;
  const nPorto = abertos.filter((r) => r.decisao === "Faz no porto").length;
  const nCrono = abertos.filter((r) => r.pc).length;
  const nConf = abertos.filter((r) => r.conf).length;
  const passo = (n, k, titulo, sub) => (
    <div onClick={() => { setReportFn && setReportFn(null); setExportXlsxFn && setExportXlsxFn(null); setStage(k); }} style={{ flex: 1, minWidth: 200, cursor: "pointer", padding: "10px 14px", borderRadius: 8, border: `2px solid ${stage === k ? "var(--teal)" : "var(--border)"}`, background: stage === k ? "rgba(43,108,176,0.08)" : "transparent" }}>
      <div style={{ fontWeight: 700, fontSize: 13 }}><span style={{ display: "inline-block", width: 20, height: 20, lineHeight: "20px", textAlign: "center", borderRadius: 10, background: stage === k ? "var(--teal)" : "var(--border)", color: "#fff", marginRight: 6, fontSize: 12 }}>{n}</span>{titulo}</div>
      <div className="g-muted" style={{ fontSize: 11.5, marginTop: 3 }}>{sub}</div>
    </div>
  );
  const cab = (
    <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
      {passo(1, "mapa", "Mapeamento total", `${abertos.length} itens em aberto (Mapeados + DNV)`)}
      {passo(2, "analise", "Análise do porto", `${nSem} sem decisão · ${nPorto} fazem no porto`)}
      {passo(3, "crono", "Cronograma", `${nCrono} no cronograma · ${nConf} confirmados em Serviços`)}
    </div>
  );

  /* ================= ETAPA 1 ================= */
  const [nv, setNv] = useState({ nome: "", origem: "Bordo", area: "", empresa: "", plano: "", impacto: "Médio" });
  const [msg, setMsg] = useState("");
  const adicionar = () => {
    if (!nv.nome.trim()) { setMsg("Escreva o nome da manutenção."); setTimeout(() => setMsg(""), 3000); return; }
    if (nv.origem === "DNV") {
      setDocagemItems((r) => [...r, { id: uid("DOC"), nome: nv.nome.trim(), localizacao: nv.area, tipoPeriodo: "", empresa: nv.empresa, planoAcao: nv.plano, necessitaMaterial: false, poRelacionada: "", previsaoExecucao: "", dataConclusao: "", status: "A Executar", linkedServiceId: null }]);
    } else {
      setPlanningItems((r) => [...r, { id: uid("PLAN"), nome: nv.nome.trim(), departamento: nv.area, empresa: nv.empresa, descricaoProblema: "", planoAcao: nv.plano, rc: "", obs: "", status: "A Executar", precisaMaterial: false, materialNecessario: "", poMaterial: "", impacto: nv.impacto, dataExecucao: "", linkedServiceId: null, jobTipo: "", jobNo: "", justificativa: "", origem: nv.origem }]);
    }
    setNv((p) => ({ ...p, nome: "", plano: "" })); setMsg("Adicionado à lista."); setTimeout(() => setMsg(""), 2500);
  };
  const enter = (e) => { if (e.key === "Enter") adicionar(); };
  const porOrigem = FUNIL_ORIGENS.map((o) => ({ o, n: abertos.filter((r) => r.origem === o).length })).filter((x) => x.n);

  const etapa1 = (
    <>
      <div className="g-panel" style={{ borderLeft: "4px solid var(--teal)" }}>
        <div className="g-panel-head"><span className="g-panel-title">Lançamento rápido — digite e aperte Enter</span></div>
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr 2fr 1fr auto", gap: 8, alignItems: "end" }}>
          <div className="g-field"><label>Manutenção *</label><input value={nv.nome} onChange={(e) => setNv((p) => ({ ...p, nome: e.target.value }))} onKeyDown={enter} placeholder="ex.: Troca do selo da bomba 2" /></div>
          <div className="g-field"><label>Origem</label><select value={nv.origem} onChange={(e) => setNv((p) => ({ ...p, origem: e.target.value }))}>{FUNIL_ORIGENS.map((o) => <option key={o}>{o}</option>)}</select></div>
          <div className="g-field"><label>Área / localização</label><input value={nv.area} onChange={(e) => setNv((p) => ({ ...p, area: e.target.value }))} onKeyDown={enter} /></div>
          <div className="g-field"><label>Empresa</label><input value={nv.empresa} onChange={(e) => setNv((p) => ({ ...p, empresa: e.target.value }))} onKeyDown={enter} /></div>
          <div className="g-field"><label>Plano de ação</label><input value={nv.plano} onChange={(e) => setNv((p) => ({ ...p, plano: e.target.value }))} onKeyDown={enter} /></div>
          <div className="g-field"><label>Impacto</label><select value={nv.impacto} onChange={(e) => setNv((p) => ({ ...p, impacto: e.target.value }))} disabled={nv.origem === "DNV"}>{IMPACT_LEVELS.map((o) => <option key={o}>{o}</option>)}</select></div>
          <button className="g-btn primary g-keep" onClick={adicionar}><Plus size={14} />Adicionar</button>
        </div>
        {msg && <div className="g-muted" style={{ fontSize: 12, marginTop: 6 }}>{msg}</div>}
        <div className="g-muted" style={{ fontSize: 11, marginTop: 6 }}>Todos os itens mapeados (inclusive os de Mapeados para Execução e DNV já cadastrados) aparecem aqui automaticamente — esta é a lista única de tudo o que está mapeado.</div>
      </div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))" }}>
        {bigKpi("Mapeado em aberto", abertos.length, "var(--teal)", ClipboardList)}
        {porOrigem.map((x) => bigKpi(x.o, x.n, "var(--teal)", Wrench))}
        {bigKpi("Sem plano de ação", abertos.filter((r) => !String(r.plano).trim()).length, abertos.some((r) => !String(r.plano).trim()) ? "var(--warn)" : "var(--ok)", AlertTriangle)}
        {bigKpi("Sem empresa", abertos.filter((r) => !String(r.empresa).trim()).length, "var(--text-dim)", AlertTriangle)}
      </div>
      {filtros(false)}
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Mapeamento total ({lista.length})</span></div>
        <div className="g-table-wrap" style={{ maxHeight: 620, overflow: "auto" }}>
          <table className="g-table tbl-modern g-editavel">
            <thead><tr><th style={{ minWidth: 220 }}>Manutenção</th><th>Origem</th><th style={{ minWidth: 130 }}>Área / localização</th><th style={{ minWidth: 120 }}>Empresa</th><th style={{ minWidth: 240 }}>Plano de ação</th><th>Impacto</th><th>Status</th><th></th></tr></thead>
            <tbody>
              {lista.length === 0 && <tr><td colSpan={8} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nada mapeado com esse filtro.</td></tr>}
              {lista.map((r) => (
                <tr className="g-row tbl-row" key={r.key}>
                  <td><TextoAuto style={inp} value={r.nome} onChange={(v) => set(r, { nome: v })} /></td>
                  <td>{r.src === "M" ? <select style={inp} value={r.origem} onChange={(e) => set(r, { origem: e.target.value })}>{FUNIL_ORIGENS.filter((o) => o !== "DNV").map((o) => <option key={o}>{o}</option>)}</select> : <span className="tbl-chip">DNV</span>}</td>
                  <td><TextoAuto style={inp} value={r.area} onChange={(v) => set(r, { area: v })} /></td>
                  <td><TextoAuto style={inp} value={r.empresa} onChange={(v) => set(r, { empresa: v })} /></td>
                  <td><TextoAuto style={{ ...inp, borderColor: String(r.plano).trim() ? undefined : "var(--warn)" }} value={r.plano} onChange={(v) => set(r, { plano: v })} placeholder="o que será feito" /></td>
                  <td>{r.src === "M" ? <select style={inp} value={r.impacto} onChange={(e) => set(r, { impacto: e.target.value })}>{IMPACT_LEVELS.map((o) => <option key={o}>{o}</option>)}</select> : <span className="g-muted">—</span>}</td>
                  <td><select style={inp} value={r.status} onChange={(e) => set(r, { status: e.target.value })}>{PLAN_STATUS.map((o) => <option key={o}>{o}</option>)}</select></td>
                  <td><span className="g-btn ghost danger" title="Excluir" onClick={() => apagar(r)}><Trash2 size={13} /></span></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );

  /* ================= ETAPA 2 ================= */
  const etapa2 = (
    <>
      <div className="g-alert" style={{ background: "rgba(37,104,160,0.08)", borderColor: "rgba(37,104,160,0.35)", color: "var(--accent)" }}>
        Para cada item mapeado, decida o que pode ser feito <strong>no porto</strong>. Os marcados como <strong>“Faz no porto”</strong> seguem para o Cronograma. Informe a duração estimada e se depende de material — isso evita agendar o que não dá para executar.
      </div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))" }}>
        {bigKpi("Sem decisão", nSem, nSem ? "var(--warn)" : "var(--ok)", AlertTriangle)}
        {FUNIL_DECISOES.map((d) => bigKpi(d, abertos.filter((r) => r.decisao === d).length, FUNIL_DEC_COR[d], Wrench))}
      </div>
      {filtros(true)}
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Análise — o que pode ser feito no porto ({lista.length})</span></div>
        <div className="g-table-wrap" style={{ maxHeight: 620, overflow: "auto" }}>
          <table className="g-table tbl-modern">
            <thead><tr><th style={{ minWidth: 200 }}>Manutenção</th><th>Origem</th><th style={{ minWidth: 200 }}>Plano de ação</th><th style={{ minWidth: 140 }}>Empresa</th><th style={{ minWidth: 150 }}>Decisão</th><th style={{ minWidth: 80 }}>Duração (h)</th><th>Material?</th><th style={{ minWidth: 90 }}>PO</th><th style={{ minWidth: 200 }}>Análise / observação</th></tr></thead>
            <tbody>
              {lista.length === 0 && <tr><td colSpan={9} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nada para analisar com esse filtro.</td></tr>}
              {lista.map((r) => {
                const semPo = r.material && !String(r.po).trim();
                return (
                  <tr className="g-row tbl-row" key={r.key} style={{ "--c": r.decisao ? FUNIL_DEC_COR[r.decisao] : "var(--warn)" }}>
                    <td><b>{r.nome}</b><div className="g-muted" style={{ fontSize: 11 }}>{r.area || "—"}</div></td>
                    <td><span className="tbl-chip">{r.origem}</span></td>
                    <td style={{ whiteSpace: "normal", fontSize: 12 }}>{r.plano || <span className="g-muted">sem plano</span>}</td>
                    <td><input style={inp} value={r.empresa} onChange={(e) => set(r, { empresa: e.target.value })} /></td>
                    <td><select style={{ ...inp, borderColor: r.decisao ? undefined : "var(--warn)", color: r.decisao ? FUNIL_DEC_COR[r.decisao] : undefined, fontWeight: 600 }} value={r.decisao} onChange={(e) => set(r, { decisao: e.target.value })}><option value="">— decidir —</option>{FUNIL_DECISOES.map((o) => <option key={o}>{o}</option>)}</select></td>
                    <td><input style={inp} type="number" min="0" step="0.5" value={r.duracaoH} onChange={(e) => set(r, { duracaoH: e.target.value })} /></td>
                    <td style={{ textAlign: "center" }}><input type="checkbox" checked={r.material} onChange={(e) => set(r, { material: e.target.checked })} /></td>
                    <td><input style={{ ...inp, borderColor: semPo ? "var(--crit)" : undefined }} value={r.po} onChange={(e) => set(r, { po: e.target.value })} placeholder={r.material ? "PO" : "—"} /></td>
                    <td><input style={inp} value={r.analise} onChange={(e) => set(r, { analise: e.target.value })} />{semPo && r.decisao === "Faz no porto" && <div style={{ color: "var(--crit)", fontSize: 10.5 }}>precisa de material e ainda não tem PO</div>}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );

  /* ================= ETAPA 3 ================= */
  const [pc, setPc] = useState("");
  const pcs = allPortCallDates;
  const pcAtual = pc || pcs.find((d) => d >= hoje) || pcs[pcs.length - 1] || "";
  const candidatos = abertos.filter((r) => r.decisao === "Faz no porto" && !r.pc);
  const noPc = abertos.filter((r) => r.pc === pcAtual).sort((a, b) => (a.ordem - b.ordem) || String(a.ini).localeCompare(String(b.ini)));
  const meta = portCallMeta?.[pcAtual];
  const janelaFim = pcAtual ? funilSomaH(funilIni(pcAtual).slice(0, 10) + "T00:00", meta?.duration || 24) : "";
  const proxOrdem = () => Math.max(0, ...rows.filter((r) => r.pc === pcAtual).map((r) => r.ordem)) + 1;
  const fimMax = () => rows.filter((r) => r.pc === pcAtual && r.fim).map((r) => r.fim).sort().pop() || "";
  const agendar = (r, de) => {
    if (!pcAtual) return;
    const ini = de || fimMax() || funilIni(pcAtual);
    set(r, { pc: pcAtual, ordem: proxOrdem(), ini, fim: funilSomaH(ini, r.duracaoH), conf: false });
  };
  const agendarTodos = () => {
    let o = proxOrdem(); let cur = fimMax() || funilIni(pcAtual);
    candidatos.forEach((r) => { const ini = cur; const fim = funilSomaH(ini, r.duracaoH); set(r, { pc: pcAtual, ordem: o++, ini, fim, conf: false }); cur = fim; });
  };
  const tirar = (r) => { if (r.conf && r.linked && !window.confirm("Este item já virou Serviço. Tirar do cronograma NÃO apaga o Serviço — apague-o em Serviços se necessário. Continuar?")) return; set(r, { pc: "", conf: false }); };
  const mover = (r, dir) => {
    const i = noPc.findIndex((x) => x.key === r.key); const j = i + dir; if (j < 0 || j >= noPc.length) return;
    const a = noPc[i], b = noPc[j]; const oa = a.ordem || i + 1, ob = b.ordem || j + 1;
    set(a, { ordem: ob === oa ? ob + dir : ob }); set(b, { ordem: oa });
  };
  const confirmar = (r) => { if (!r.ini || !r.fim) return; set(r, { conf: true }); };
  const confirmarTodos = () => noPc.filter((r) => r.ini && r.fim && !r.conf).forEach((r) => set(r, { conf: true }));
  const fmtH = (s) => (s ? `${fmtDataCompleta(s.slice(0, 10))} ${s.slice(11, 16)}` : "—");
  const pendConf = noPc.filter((r) => !r.conf).length;
  const horasTotal = noPc.reduce((s, r) => s + (r.ini && r.fim ? Math.max(0, (new Date(r.fim) - new Date(r.ini)) / 3600000) : 0), 0);

  const etapa3 = (
    <>
      <div className="g-alert" style={{ background: "rgba(37,104,160,0.08)", borderColor: "rgba(37,104,160,0.35)", color: "var(--accent)" }}>
        Escolha o Port Call, traga os itens aprovados na análise, ajuste a ordem e os horários e clique em <strong>Confirmar</strong>. Cada item confirmado <strong>vira um Serviço automaticamente</strong> (aba Serviços e Gantt) e mudanças aqui passam a valer lá. Se você mudar o horário no próprio Serviço, o cronograma acompanha; se o Serviço for concluído ou cancelado, o item é encerrado aqui também.
      </div>
      <div className="g-panel" style={{ display: "flex", gap: 14, alignItems: "end", flexWrap: "wrap" }}>
        <div className="g-field" style={{ minWidth: 280 }}><label>Port Call do cronograma</label>
          <select value={pcAtual} onChange={(e) => setPc(e.target.value)}>{pcs.length === 0 && <option value="">Nenhum Port Call cadastrado (crie em Port Call)</option>}{pcs.map((d) => <option key={d} value={d}>{portCallLabel(d)}</option>)}</select>
        </div>
        <div className="g-muted" style={{ fontSize: 12 }}>Janela: <b>{meta?.duration || 24}h</b> · Itens: <b>{noPc.length}</b> · Horas somadas: <b>{horasTotal.toFixed(1)}h</b> · Pendentes de confirmação: <b>{pendConf}</b></div>
        <div style={{ marginLeft: "auto" }}><button className="g-btn primary g-keep" disabled={!pendConf} onClick={confirmarTodos}>Confirmar todos ({pendConf})</button></div>
      </div>

      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Candidatos aprovados na análise, ainda sem Port Call ({candidatos.length})</span>{candidatos.length > 0 && <button className="g-btn g-keep" onClick={agendarTodos}>Levar todos para {pcAtual ? portCallLabel(pcAtual).split(" · ")[0] : "o Port Call"}</button>}</div>
        {candidatos.length === 0 ? <div className="g-muted" style={{ fontSize: 12 }}>Nenhum candidato. Na etapa 2, marque “Faz no porto” nos itens que podem ser executados.</div> : (
          <div className="g-table-wrap"><table className="g-table"><tbody>
            {candidatos.map((r) => <tr className="g-row" key={r.key}><td><b>{r.nome}</b><div className="g-muted" style={{ fontSize: 11 }}>{r.origem} · {r.area || "—"} · {r.empresa || "sem empresa"}{r.duracaoH ? ` · ${r.duracaoH}h` : ""}</div></td><td style={{ width: 170 }}><button className="g-btn g-keep" onClick={() => agendar(r)} disabled={!pcAtual}>Levar ao Port Call</button></td></tr>)}
          </tbody></table></div>
        )}
      </div>

      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Cronograma — {pcAtual ? portCallLabel(pcAtual) : "—"}</span></div>
        <div className="g-table-wrap">
          <table className="g-table tbl-modern">
            <thead><tr><th style={{ width: 70 }}>Ordem</th><th style={{ minWidth: 220 }}>Manutenção</th><th style={{ minWidth: 120 }}>Empresa</th><th style={{ minWidth: 190 }}>Início</th><th style={{ minWidth: 190 }}>Fim</th><th>Situação</th><th></th></tr></thead>
            <tbody>
              {noPc.length === 0 && <tr><td colSpan={7} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nenhum item neste Port Call ainda.</td></tr>}
              {noPc.map((r, i) => {
                const fora = janelaFim && r.fim && r.fim > janelaFim;
                const invalido = r.ini && r.fim && r.fim <= r.ini;
                return (
                  <tr className="g-row tbl-row" key={r.key} style={{ "--c": r.conf ? "#22C55E" : "#F59E0B" }}>
                    <td style={{ whiteSpace: "nowrap" }}><span className="g-btn ghost" title="Subir" onClick={() => mover(r, -1)}>▲</span><span className="g-btn ghost" title="Descer" onClick={() => mover(r, 1)}>▼</span></td>
                    <td><b>{i + 1}. {r.nome}</b><div className="g-muted" style={{ fontSize: 11 }}>{r.origem} · {r.area || "—"}</div></td>
                    <td><input style={inp} value={r.empresa} onChange={(e) => set(r, { empresa: e.target.value })} /></td>
                    <td><input style={inp} type="datetime-local" value={r.ini} onChange={(e) => set(r, { ini: e.target.value, conf: r.conf && !!e.target.value })} /></td>
                    <td><input style={{ ...inp, borderColor: fora || invalido ? "var(--crit)" : undefined }} type="datetime-local" value={r.fim} onChange={(e) => set(r, { fim: e.target.value })} />{fora && <div style={{ color: "var(--crit)", fontSize: 10.5 }}>passa da janela do Port Call</div>}{invalido && <div style={{ color: "var(--crit)", fontSize: 10.5 }}>fim antes do início</div>}</td>
                    <td style={{ whiteSpace: "normal" }}>{r.conf ? <span style={{ color: "var(--ok)", fontWeight: 700 }}>✓ Em Serviços{r.linked ? "" : " (criando…)"}</span> : <button className="g-btn primary g-keep" disabled={!r.ini || !r.fim || invalido} onClick={() => confirmar(r)}>Confirmar</button>}</td>
                    <td><span className="g-btn ghost danger" title="Tirar do cronograma" onClick={() => tirar(r)}><Trash2 size={13} /></span></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );

  return (
    <>
      {cab}
      {stage === "mapa" ? etapa1 : stage === "analise" ? etapa2 : etapa3}
    </>
  );
}

function PlanejamentoView({ workPackages, updWp, materials, setReportFn, setExportXlsxFn, allPortCallDates, portCallLabel, addWpOnDate,
  planningItems, updPlan, remPlan, addPlanningItem, newRowId, handleImportPlanejamento,
  docagemItems, updDocagem, remDocagem, addDocagemItem, handleImportDocagem, planosBordo, setPlanosBordo, planSubTab: planSubTabProp, setPlanSubTab, setPlanningItems, setDocagemItems, tmDue, tmHistory, tmFailures, reparoItens, portCallMeta, portcallReqs, setPortcallReqs, soDocagem }) {
  const planSubTab = soDocagem ? "docagem" : planSubTabProp;
  const [funilStage, setFunilStage] = useState("mapa");
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
  /* vínculo Mapeados ↔ Solicitação Portcall */
  const emSolic = useMemo(() => { const m = {}; (portcallReqs || []).forEach((r) => r.itens.forEach((it) => { if (it.planId) (m[it.planId] = m[it.planId] || []).push(r); })); return m; }, [portcallReqs]);
  const solicitar = (p) => {
    if (emSolic[p.id]) { setPlanSubTab("portcall"); return; }
    const item = { id: uid("PCI"), descricao: p.nome, componente: "", tipo: "Preventiva", justificativa: p.justificativa || p.descricaoProblema || p.planoAcao || "", linkado: p.jobNo ? "sim" : "nao", jobTipo: p.jobTipo || "", jobNo: p.jobNo || "", oneNo: "", planId: p.id };
    setPortcallReqs((a) => {
      const arr = a || []; const alvo = arr.find((r) => !r.enviado);
      if (alvo) return arr.map((r) => (r === alvo ? { ...r, itens: [...r.itens, item] } : r));
      return [{ id: uid("PCR"), portCall: allPortCallDates.find((d) => d >= todayISO()) || allPortCallDates[allPortCallDates.length - 1] || "", itens: [item], criado: todayISO(), enviado: "" }, ...arr];
    });
    setPlanSubTab("portcall");
  };
  const mapGroups = useMemo(() => groupRows(filteredMapeados, groupMap, (p) => p.departamento, "Sem departamento"), [filteredMapeados, groupMap]);
  const docGroups = useMemo(() => groupRows(filteredDocagem, groupDoc, (d) => d.localizacao, "Sem localização"), [filteredDocagem, groupDoc]);

  React.useEffect(() => {
    if (!setExportXlsxFn || planSubTab === "portcall" || planSubTab === "escopo" || planSubTab === "acao" || planSubTab === "funil") return;
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
    if (!setReportFn || planSubTab === "portcall" || planSubTab === "escopo" || planSubTab === "acao" || planSubTab === "funil") return;
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
      <div className="g-mode-toggle" style={{ marginBottom: 16, width: "fit-content", display: soDocagem ? "none" : undefined }}>
        <button className={planSubTab === "funil" ? "active" : ""} onClick={() => { setReportFn(null); setExportXlsxFn(null); setPlanSubTab("funil"); }}>Mapear → Analisar → Programar</button>
        <button className={planSubTab === "escopo" ? "active" : ""} onClick={() => setPlanSubTab("escopo")}>Análise de Escopo</button>
        <button className={planSubTab === "portcall" ? "active" : ""} onClick={() => setPlanSubTab("portcall")}>Solicitação Portcall</button>
      </div>

      {planSubTab === "funil" ? (
        <FunilManutencaoView stage={funilStage} setStage={setFunilStage} planningItems={planningItems} setPlanningItems={setPlanningItems} docagemItems={docagemItems} setDocagemItems={setDocagemItems}
          allPortCallDates={allPortCallDates} portCallLabel={portCallLabel} portCallMeta={portCallMeta} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} />
      ) : planSubTab === "acao" ? (
        <PlanoAcaoBordoView planningItems={planningItems} docagemItems={docagemItems} planosBordo={planosBordo} setPlanosBordo={setPlanosBordo} updPlan={updPlan} remPlan={remPlan} updDocagem={updDocagem} remDocagem={remDocagem} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} />
      ) : planSubTab === "escopo" ? (
        <AnaliseEscopoView tmDue={tmDue} tmHistory={tmHistory} planningItems={planningItems} tmFailures={tmFailures} reparoItens={reparoItens}
          allPortCallDates={allPortCallDates} portCallMeta={portCallMeta} portCallLabel={portCallLabel} portcallReqs={portcallReqs} setPortcallReqs={setPortcallReqs}
          irSolicitacao={() => { setReportFn(null); setExportXlsxFn(null); setPlanSubTab("portcall"); }} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} />
      ) : planSubTab === "portcall" ? (
        <PortcallRequestView portcallReqs={portcallReqs} setPortcallReqs={setPortcallReqs} planningItems={planningItems} tmDue={tmDue} tmHistory={tmHistory}
          allPortCallDates={allPortCallDates} portCallLabel={portCallLabel} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} />
      ) : planSubTab === "mapeados" ? (
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
            O item só vira Serviço (aba Serviços e Port Call) quando for <strong>confirmado no Cronograma</strong>,
            na aba “Mapear → Analisar → Programar”. A Data de Execução abaixo é apenas uma previsão informativa.
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
                  <th style={{ minWidth: 170 }}>Job vinculado (TM)</th>
                  <th style={{ minWidth: 150 }}>Status</th>
                  <th style={{ width: 100 }}></th>
                </tr>
              </thead>
              <tbody>
                {mapGroups.map((g) => (
                  <React.Fragment key={g.key}>
                    {groupMap && (
                      <tr className="tbl-group" onClick={() => setCollapsedMap((p) => { const n = new Set(p); n.has(g.key) ? n.delete(g.key) : n.add(g.key); return n; })}>
                        <td colSpan={8}>
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
                          {(planosBordo || []).find((x) => x.key === `M:${p.id}`)?.status && <span className="tbl-chip" title="Status no Plano de Ação com Bordo (mesmo item)">📋 Bordo: {(planosBordo || []).find((x) => x.key === `M:${p.id}`).status}</span>}
                          {emSolic[p.id] && <span className="tbl-chip" style={{ color: "var(--ok)", cursor: "pointer" }} onClick={() => setPlanSubTab("portcall")} title="Já está em uma solicitação ao Port Call">📨 Solicitação Port Call</span>}
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
                      <td style={{ minWidth: 170, verticalAlign: "top" }}>
                        {(() => {
                          const ach = p.jobNo ? ([...(tmDue || []), ...(tmHistory || [])].find((x) => bordoNorm(x.jobType) === bordoNorm(p.jobTipo) && bordoNorm(x.jobNo) === bordoNorm(p.jobNo))) : null;
                          const ini = { fontSize: 11.5, background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 6, padding: "4px 6px", color: "var(--text)" };
                          return (<>
                            <div className="g-flex" style={{ gap: 4 }}>
                              <input type="text" placeholder="Tipo" value={p.jobTipo || ""} onChange={(e) => updPlan(i, "jobTipo", e.target.value.toUpperCase())} style={{ ...ini, width: 58 }} />
                              <input type="text" placeholder="nº do job" value={p.jobNo || ""} onChange={(e) => updPlan(i, "jobNo", e.target.value)} style={{ ...ini, width: 82 }} />
                            </div>
                            {p.jobNo && <div style={{ marginTop: 3, fontSize: 10, color: ach ? "var(--ok)" : "var(--warn)" }} title={ach?.jobName || ""}>{ach ? `✔ no TM: ${(ach.jobName || "").slice(0, 28)}` : "⚠ não achado no TM"}</div>}
                          </>);
                        })()}
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
                        <span className="g-btn ghost g-keep" onClick={() => solicitar(p)} title={emSolic[p.id] ? "Já está na solicitação — abrir" : "Incluir na solicitação ao Port Call"}>📨</span>
                        <span className="g-btn ghost" onClick={() => setExpandedPlanRow(isOpen ? null : p.id)} title="Ver todos os detalhes">
                          {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                        </span>
                        <span className="g-btn ghost danger" onClick={() => remPlan(i)}><Trash2 size={13} /></span>
                      </td>
                    </tr>
                    {isOpen && (
                      <tr className="tbl-detail">
                        <td></td>
                        <td colSpan={7}>
                          <div className="tbl-detail-grid">
                            <div className="g-field"><label>Departamento</label><EText value={p.departamento} onChange={(v) => updPlan(i, "departamento", v)} /></div>
                            <div className="g-field"><label>Empresa</label><EText value={p.empresa} onChange={(v) => updPlan(i, "empresa", v)} /></div>
                            <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Descrição do Problema</label><ETextArea rows={1} value={p.descricaoProblema} onChange={(v) => updPlan(i, "descricaoProblema", v)} /></div>
                            <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Plano de Ação</label><ETextArea rows={1} value={p.planoAcao} onChange={(v) => updPlan(i, "planoAcao", v)} /></div>
                            <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Resposta do bordo (Plano de Ação com Bordo, mesmo item)</label><ETextArea rows={1} value={(planosBordo || []).find((x) => x.key === `M:${p.id}`)?.resposta || ""} onChange={(v) => updPlanoBordoKey(setPlanosBordo, `M:${p.id}`, "resposta", v)} /></div>
                            <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Justificativa (usada na solicitação ao Port Call)</label><ETextArea rows={1} value={p.justificativa || ""} onChange={(v) => updPlan(i, "justificativa", v)} /></div>
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
              <button className="g-btn g-edit-only" onClick={() => docagemFileRef.current?.click()}>
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
                            <div className="g-field" style={{ gridColumn: "1 / -1" }}><label>Resposta do bordo (Plano de Ação com Bordo, mesmo item)</label><ETextArea rows={1} value={(planosBordo || []).find((x) => x.key === `D:${d.id}`)?.resposta || ""} onChange={(v) => updPlanoBordoKey(setPlanosBordo, `D:${d.id}`, "resposta", v)} /></div>
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
const fmtDataCompleta = (d) => { if (!d) return "—"; const [y, m, day] = String(d).split("-"); return `${day}/${m}/${y}`; };

const PlanosAtencao = ({ postergados, vencidos, totalRegistrosPost }) => {
  const [abrePost, setAbrePost] = useState(true);
  const [abreVenc, setAbreVenc] = useState(true);
  const [qPost, setQPost] = useState("");
  const [qVenc, setQVenc] = useState("");
  const norm = (v) => String(v || "").toLowerCase();
  const postVis = postergados.filter((r) => !qPost || norm(`${r.componentCode} ${r.componentName} ${r.jobName} ${r.jobNo} ${r.jobType}`).includes(norm(qPost)));
  const vencVis = vencidos.filter((r) => !qVenc || norm(`${r.code} ${r.component} ${r.jobName} ${r.jobNo} ${r.jobType}`).includes(norm(qVenc)));
  const caixa = { maxHeight: 420, overflowY: "auto" };
  const busca = (v, set, n) => (
    <input value={v} onChange={(e) => set(e.target.value)} placeholder={`Buscar entre ${n}...`}
      style={{ padding: "6px 10px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 12.5, minWidth: 220 }} />
  );
  return (
    <>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(2, 1fr)" }}>
        {bigKpi("Planos Postergados", postergados.length, "var(--warn)", AlertTriangle)}
        {bigKpi("Planos Vencidos", vencidos.length, "var(--crit)", Clock)}
      </div>

      <div className="g-panel" style={{ marginBottom: 14 }}>
        <div className="g-panel-head" style={{ gap: 12, flexWrap: "wrap" }}>
          <span className="g-panel-title">Planos Postergados — Detalhamento</span>
          <span className="g-muted" style={{ fontSize: 12 }}>{postergados.length} plano(s)</span>
          <div style={{ flex: 1 }} />
          {abrePost && busca(qPost, setQPost, postergados.length)}
          <button className="g-btn" onClick={() => setAbrePost((v) => !v)}>{abrePost ? "Recolher" : "Ver detalhes"}</button>
        </div>
        {abrePost && (
          <div className="g-table-wrap" style={caixa}>
            <table className="g-table tbl-modern">
              <thead><tr>
                <th>Código</th><th>Componente</th><th>Job No</th><th>Plano (Job Name)</th><th>Tipo</th>
                <th>Última Postergação</th>
              </tr></thead>
              <tbody>
                {postVis.length === 0 && <tr><td colSpan={6} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nenhum plano postergado.</td></tr>}
                {postVis.map((r) => (
                  <tr className="g-row tbl-row" key={r.key} style={{ "--c": "#F5A623" }}>
                    <td style={{ fontFamily: "var(--mono)" }}>{r.componentCode || "—"}</td>
                    <td>{r.componentName || "—"}</td>
                    <td style={{ fontFamily: "var(--mono)" }}>{r.jobNo || "—"}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 220 }}>{r.jobName || "—"}</td>
                    <td>{r.jobType || "—"}</td>
                    <td style={{ fontFamily: "var(--mono)" }}>{fmtDataCompleta(r.ultimaData)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div className="g-panel" style={{ marginBottom: 14 }}>
        <div className="g-panel-head" style={{ gap: 12, flexWrap: "wrap" }}>
          <span className="g-panel-title">Planos Vencidos — Detalhamento</span>
          <span className="g-muted" style={{ fontSize: 12 }}>{vencidos.length} plano(s) · pela coluna Diff (ou Due date quando não há Diff)</span>
          <div style={{ flex: 1 }} />
          {abreVenc && busca(qVenc, setQVenc, vencidos.length)}
          <button className="g-btn" onClick={() => setAbreVenc((v) => !v)}>{abreVenc ? "Recolher" : "Ver detalhes"}</button>
        </div>
        {abreVenc && (
          <div className="g-table-wrap" style={caixa}>
            <table className="g-table tbl-modern">
              <thead><tr>
                <th>Código</th><th>Componente</th><th>Job No</th><th>Plano (Job Name)</th>
                <th>Tipo</th><th>Prioridade</th><th>Due</th><th>Atraso</th>
              </tr></thead>
              <tbody>
                {vencVis.length === 0 && <tr><td colSpan={8} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nenhum plano vencido.</td></tr>}
                {vencVis.map((r) => (
                  <tr className="g-row tbl-row" key={r.id} style={{ "--c": "#EF4444" }}>
                    <td style={{ fontFamily: "var(--mono)" }}>{r.code || "—"}</td>
                    <td>{r.component || "—"}</td>
                    <td style={{ fontFamily: "var(--mono)" }}>{r.jobNo || "—"}</td>
                    <td style={{ whiteSpace: "normal", minWidth: 220 }}>{r.jobName || "—"}</td>
                    <td>{r.jobType || "—"}</td>
                    <td>{r.pri || "—"}</td>
                    <td style={{ fontFamily: "var(--mono)" }}>{r.dueDate ? fmtDataCompleta(r.dueDate) : (r.dueRaw || "—")}</td>
                    <td><span className="tbl-days d3">{r._atraso}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </>
  );
};

/* ---------- Visão Geral das Métricas: componentes das novas análises ---------- */
/* grupos críticos = 3 primeiros dígitos do código do componente (planilha PMS Change Request 011/2026 - Critical) */
const CRITICAL_GROUPS = [
  { cod: "331", nome: "Offshore Crane", crit: "Operação" },
  { cod: "403", nome: "Steering Gear", crit: "Segurança" },
  { cod: "404", nome: "Side Thrusters", crit: "Operação" },
  { cod: "408", nome: "K-Pos (Sistema de Posicionamento)", crit: "Operação" },
  { cod: "411", nome: "Radares", crit: "Segurança" },
  { cod: "412", nome: "DP Demodulators", crit: "Operação / Segurança" },
  { cod: "413", nome: "Gyro Compasses", crit: "Segurança" },
  { cod: "414", nome: "Navigational Echo Sounder", crit: "Segurança" },
  { cod: "415", nome: "Loading Computer", crit: "Segurança" },
  { cod: "417", nome: "Wind Indicators", crit: "Segurança" },
  { cod: "421", nome: "GMDSS MF/HF - DSC Radio", crit: "Segurança" },
  { cod: "422", nome: "GMDSS Portable VHF", crit: "Segurança" },
  { cod: "427", nome: "Navigation Light Switchboards", crit: "Segurança" },
  { cod: "433", nome: "Windlass / Mooring Winch", crit: "Operação" },
  { cod: "434", nome: "Capstans", crit: "Operação" },
  { cod: "501", nome: "FRC / MOB Boats", crit: "Segurança" },
  { cod: "502", nome: "Inflatable Liferafts", crit: "Segurança" },
  { cod: "503", nome: "Life Buoys", crit: "Segurança" },
  { cod: "505", nome: "Smoke Diving Equipment", crit: "Segurança" },
  { cod: "506", nome: "Portable Gas Detector", crit: "Segurança" },
  { cod: "513", nome: "Watertight Doors", crit: "Segurança" },
  { cod: "571", nome: "AC Accommodation", crit: "Operação" },
  { cod: "573", nome: "Air Condition Units", crit: "Operação" },
  { cod: "574", nome: "Ventilation / Air Cond.", crit: "Operação" },
  { cod: "582", nome: "Sewage Treatment Unit", crit: "Meio Ambiente" },
  { cod: "601", nome: "Motores Principais", crit: "Operação" },
  { cod: "625", nome: "Main Generators", crit: "Operação" },
  { cod: "632", nome: "Shaft Brake", crit: "Operação" },
  { cod: "634", nome: "Pitch Propeller", crit: "Operação" },
  { cod: "635", nome: "Azimuth Thruster", crit: "Operação" },
  { cod: "637", nome: "Main Reduction Gear", crit: "Operação" },
  { cod: "665", nome: "Emergency Engine", crit: "Segurança" },
  { cod: "701", nome: "FO Transfer", crit: "Operação" },
  { cod: "703", nome: "Diesel Oil Filters", crit: "Operação" },
  { cod: "711", nome: "LO Transfer / Filters", crit: "Operação" },
  { cod: "713", nome: "Lub Oil System", crit: "Operação" },
  { cod: "721", nome: "Sea Water Cooling", crit: "Operação" },
  { cod: "722", nome: "Fresh Water Cooling", crit: "Operação" },
  { cod: "731", nome: "Starting Air System", crit: "Operação" },
  { cod: "801", nome: "Anti Heeling Pump", crit: "Segurança / Meio Ambiente" },
  { cod: "803", nome: "Bilge / Oily Water Separator", crit: "Segurança / Meio Ambiente" },
  { cod: "811", nome: "Fire Alarm Central", crit: "Segurança" },
  { cod: "812", nome: "Emergency Shutdown", crit: "Segurança" },
  { cod: "813", nome: "Fire / Wash Down Pump", crit: "Segurança" },
  { cod: "815", nome: "CO2 Plant", crit: "Segurança" },
  { cod: "816", nome: "Deep Fat Fire-Extinguishing", crit: "Segurança" },
  { cod: "865", nome: "Transformers", crit: "Operação" },
  { cod: "866", nome: "Batteries", crit: "Segurança" },
  { cod: "871", nome: "Main Switchboards", crit: "Operação" },
  { cod: "872", nome: "Emergency Switchboards", crit: "Segurança" },
  { cod: "875", nome: "Section & Distribution Boards", crit: "Operação" },
  { cod: "891", nome: "Emergency Lights - Engine Rooms", crit: "Segurança" },
  { cod: "892", nome: "Emergency Lights - Accommodation", crit: "Segurança" },
  { cod: "893", nome: "Emergency Lights - Outside Decks", crit: "Segurança" },
];
const CRITICAL_PREFIXES = CRITICAL_GROUPS.map((g) => g.cod);
const critOf = (cod) => CRITICAL_GROUPS.find((x) => x.cod === cod)?.crit || "";
const critLabel = (cod) => { const g = CRITICAL_GROUPS.find((x) => x.cod === cod); return g ? `${g.cod} - ${g.nome}` : cod; };
const critPrefixOf = (code) => {
  const c = String(code || "").trim();
  for (const p of CRITICAL_PREFIXES) { if (c.startsWith(p) && !/\d/.test(c.charAt(p.length))) return p; }
  return null;
};
/* jobs preditivos: tipo + número do job. Vários componentes do mesmo job feitos juntos contam como 1 */
const PREDITIVOS = [
  { grupo: "Análise de Vibração", tipo: "CHK", no: "7" },
  { grupo: "Análise de Vibração", tipo: "INS", no: "2054" },
  { grupo: "Megagem", tipo: "ANN", no: "28" },
  { grupo: "Megagem", tipo: "CHK", no: "511" },
  { grupo: "Megagem", tipo: "INS", no: "2706" },
  { grupo: "Megagem", tipo: "CHK", no: "101" },
  { grupo: "Análise de Óleo", tipo: "FSA", no: "1" },
  { grupo: "Análise de Óleo", tipo: "FSA", no: "2" },
];
const AX = { fill: "var(--text-faint)", fontSize: 10 };
const TT = { background: "var(--panel-raised)", border: "1px solid var(--border)", borderRadius: 4, fontSize: 11 };

const MesBarChart = ({ title, subtitle, data, bars, height = 240, headRight, nota }) => (
  <div className="g-panel">
    <div className="g-panel-head" style={{ gap: 12, flexWrap: "wrap" }}>
      <span className="g-panel-title">{title}</span>
      <div style={{ flex: 1 }} />
      {headRight}
    </div>
    {subtitle && <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 8 }}>{subtitle}</div>}
    <div style={{ width: "100%", height }}>
      <ResponsiveContainer>
        <BarChart data={data} margin={{ left: 0, right: 8, top: 20, bottom: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
          <XAxis dataKey={data[0] && data[0].mes !== undefined ? "mes" : "nome"} tick={AX} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
          <YAxis allowDecimals={false} tick={AX} axisLine={false} tickLine={false} width={42} />
          <Tooltip contentStyle={TT} labelStyle={{ color: "var(--text)" }} />
          <Legend wrapperStyle={{ fontSize: 11 }} />
          {bars.map((b) => (
            <Bar key={b.key} dataKey={b.key} name={b.name} fill={b.color} radius={[3, 3, 0, 0]} isAnimationActive={false}>
              <LabelList dataKey={b.key} position="top" style={{ fill: "var(--text-dim)", fontSize: 10 }} />
            </Bar>
          ))}
        </BarChart>
      </ResponsiveContainer>
    </div>
    {data.length === 0 && <div className="g-muted" style={{ fontSize: 11.5 }}>Sem dados com data no filtro atual.</div>}
    {nota && <div className="g-muted" style={{ fontSize: 10.5, marginTop: 4 }}>{nota}</div>}
  </div>
);

const PostergadosAnalise = ({ analise, porTipo }) => (
  <div className="g-grid-2">
    <MesBarChart title="Postergados Por Mês" data={analise.rows}
      headRight={<span className="g-muted" style={{ fontSize: 12 }}>Total: <b>{analise.total}</b> · realizados: <b>{analise.realizados}</b> · ainda em aberto: <b>{analise.emAberto}</b> · vencidos: <b>{analise.vencidos}</b></span>}
      subtitle="Postergações registradas em cada mês e, dessas: quantas já foram realizadas depois da última postergação do plano, quantas ainda estão em aberto no Due e quantas dessas já estão vencidas."
      bars={[{ key: "total", name: "Postergados", color: "#F5A623" }, { key: "realizados", name: "Realizados após a postergação", color: "#22C55E" }, { key: "emAberto", name: "Ainda em aberto", color: "#3B82F6" }, { key: "vencidos", name: "Vencidos", color: "#EF4444" }]} />
    <MesBarChart title="Postergados Por Job Type" data={porTipo} subtitle="Postergações do período, por tipo de job."
      bars={[{ key: "total", name: "Postergados", color: "#F5A623" }]} />
  </div>
);

const CumprimentoPrazoPanel = ({ c }) => (
  <div className="g-panel">
    <div className="g-panel-head"><span className="g-panel-title">Cumprimento de Prazo — Due Date × Data Assinada</span></div>
    <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 10 }}>
      Considera os {c.total} jobs fechados com Due date e Data assinada preenchidas (postergações ficam de fora, pois não são execução).
      Atraso = Data assinada − Due date; zero ou negativo conta como no prazo.
    </div>
    <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
      {bigKpi("% No Prazo", `${c.pct}%`, c.pct >= 80 ? "var(--ok)" : c.pct >= 60 ? "var(--warn)" : "var(--crit)", ClipboardList)}
      {bigKpi("Feitos no Prazo ou Antes", c.noPrazo, "var(--ok)", ClipboardList)}
      {bigKpi("Feitos Após o Prazo", c.atrasados, "var(--crit)", AlertTriangle)}
      {bigKpi("Atraso Médio (dos atrasados)", `${c.atrasoMedio}d`, "var(--warn)", Clock)}
      {bigKpi("Antecedência Média (dos no prazo)", `${c.antecedenciaMedia}d`, "var(--teal)", Clock)}
    </div>
    <div className="g-grid-2">
      <MesBarChart title="Distribuição Por Faixa de Atraso" height={220} data={c.faixas}
        bars={[{ key: "total", name: "Jobs", color: "#3B82F6" }]} />
      <MesBarChart title="% No Prazo Por Mês de Assinatura" height={220} data={c.porMes}
        bars={[{ key: "pct", name: "% no prazo", color: "#22C55E" }]}
        nota="Mês em que o job foi assinado. Valores em %." />
    </div>
  </div>
);

const CriticosPanel = ({ kpis, porGrupo, due, post }) => {
  const [vista, setVista] = useState("abertos");
  const [grupo, setGrupo] = useState("Todos");
  const [criticidade, setCriticidade] = useState("Todas");
  const [sit, setSit] = useState("Todos");
  const [q, setQ] = useState("");
  const norm = (v) => String(v || "").toLowerCase();
  const okCrit = (cod) => criticidade === "Todas" || critOf(cod).includes(criticidade);
  const gruposVis = CRITICAL_GROUPS.filter((g) => okCrit(g.cod));
  const porGrupoVis = porGrupo.filter((r) => okCrit(r.grupo));
  const porGrupoGrafico = porGrupoVis.filter((r) => r.vencidos + r.aVencer + r.noPrazo + r.postergados > 0);
  const dueVis = due.filter((r) => okCrit(r._grupo) && (grupo === "Todos" || r._grupo === grupo) && (sit === "Todos" || r._sit === sit)
    && (!q || norm(`${r.code} ${r.component} ${r.jobName} ${r.jobNo} ${r.jobType}`).includes(norm(q))));
  const postVis = post.filter((r) => okCrit(r._grupo) && (grupo === "Todos" || r._grupo === grupo)
    && (!q || norm(`${r.componentCode} ${r.componentName} ${r.jobName} ${r.jobNo} ${r.jobType}`).includes(norm(q))));
  const sitCls = { "Vencido": "d3", "A vencer (até 40d)": "d2", "No prazo": "d1" };
  const corPct = (v) => (v >= 80 ? "var(--ok)" : v >= 60 ? "var(--warn)" : "var(--crit)");
  const selStyle = { padding: "6px 10px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 12.5 };
  const fmtD = (iso, raw) => (iso ? fmtDataCompleta(iso) : (raw || "—"));
  return (
    <div className="g-panel">
      <div className="g-panel-head"><span className="g-panel-title">Análise de Jobs Críticos</span></div>
      <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 10 }}>
        Componentes críticos conforme a planilha PMS Change Request 011/2026 (Critical to Operation / Safety / Environment) — {CRITICAL_GROUPS.length} grupos, identificados pelos 3 primeiros dígitos do código do componente. O gráfico mostra só os grupos com jobs.
      </div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
        {bigKpi("Críticos em Aberto", kpis.abertos, "var(--teal)", ClipboardList)}
        {bigKpi("Vencidos", kpis.vencidos, "var(--crit)", AlertTriangle)}
        {bigKpi("A Vencer (até 40 dias)", kpis.aVencer, "var(--warn)", Clock)}
        {bigKpi("No Prazo (+40 dias)", kpis.noPrazo, "var(--ok)", ClipboardList)}
        {bigKpi("Críticos Postergados", kpis.postergados, "var(--warn)", AlertTriangle)}
      </div>
      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
        {bigKpi("Executados no Período", kpis.executados, "var(--ok)", ClipboardList)}
        {bigKpi("% Executado no Prazo", `${kpis.pctPrazo}%`, corPct(kpis.pctPrazo), ClipboardList)}
        {bigKpi("Taxa de Postergação", `${kpis.taxaPost}%`, kpis.taxaPost <= 5 ? "var(--ok)" : kpis.taxaPost <= 15 ? "var(--warn)" : "var(--crit)", AlertTriangle)}
        {bigKpi("Índice de Eficiência", `${kpis.eficiencia}%`, corPct(kpis.eficiencia), Gauge)}
      </div>
      <div style={{ width: "100%", height: 400 }}>
        <ResponsiveContainer>
          <BarChart data={porGrupoGrafico} margin={{ left: 70, right: 8, top: 20, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
            <XAxis dataKey="nome" tick={AX} axisLine={{ stroke: "var(--border)" }} tickLine={false} interval={0} angle={-40} textAnchor="end" height={120} />
            <YAxis allowDecimals={false} tick={AX} axisLine={false} tickLine={false} width={36} />
            <Tooltip contentStyle={TT} labelStyle={{ color: "var(--text)" }} />
            <Legend wrapperStyle={{ fontSize: 11 }} />
            <Bar dataKey="vencidos" name="Vencidos" fill="#EF4444" radius={[3, 3, 0, 0]} isAnimationActive={false}><LabelList dataKey="vencidos" position="top" style={{ fill: "var(--text-dim)", fontSize: 9 }} /></Bar>
            <Bar dataKey="aVencer" name="A vencer (até 40d)" fill="#F5A623" radius={[3, 3, 0, 0]} isAnimationActive={false}><LabelList dataKey="aVencer" position="top" style={{ fill: "var(--text-dim)", fontSize: 9 }} /></Bar>
            <Bar dataKey="noPrazo" name="No prazo" fill="#22C55E" radius={[3, 3, 0, 0]} isAnimationActive={false}><LabelList dataKey="noPrazo" position="top" style={{ fill: "var(--text-dim)", fontSize: 9 }} /></Bar>
            <Bar dataKey="postergados" name="Postergados" fill="#8B5CF6" radius={[3, 3, 0, 0]} isAnimationActive={false}><LabelList dataKey="postergados" position="top" style={{ fill: "var(--text-dim)", fontSize: 9 }} /></Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div className="g-muted" style={{ fontSize: 10.5, margin: "4px 0 10px" }}>Por grupo de sistema. Vencidos / a vencer / no prazo: jobs em aberto. Postergados: postergações registradas no período.</div>

      <div className="g-table-wrap" style={{ marginBottom: 12 }}>
        <table className="g-table tbl-modern">
          <thead><tr>
            <th>Grupo</th><th>Criticidade</th><th style={{ textAlign: "right" }}>Em Atraso</th><th style={{ textAlign: "right" }}>A Vencer</th><th style={{ textAlign: "right" }}>No Prazo</th>
            <th style={{ textAlign: "right" }}>Executados</th><th style={{ textAlign: "right" }}>Postergados</th><th style={{ textAlign: "right" }}>% No Prazo</th>
            <th style={{ textAlign: "right" }}>Taxa de Postergação</th><th style={{ textAlign: "right" }}>Eficiência</th>
          </tr></thead>
          <tbody>
            {porGrupoVis.map((r) => (
              <tr className="g-row tbl-row" key={r.grupo} style={{ "--c": r.vencidos > 0 ? "#EF4444" : r.aVencer > 0 ? "#F5A623" : "#22C55E" }}>
                <td><b>{r.nome}</b></td><td>{critOf(r.grupo)}</td>
                <td style={{ textAlign: "right" }}>{r.vencidos}</td><td style={{ textAlign: "right" }}>{r.aVencer}</td><td style={{ textAlign: "right" }}>{r.noPrazo}</td>
                <td style={{ textAlign: "right" }}>{r.executados}</td><td style={{ textAlign: "right" }}>{r.postergados}</td>
                <td style={{ textAlign: "right" }}>{r.pctPrazo === null ? "—" : `${r.pctPrazo}%`}</td>
                <td style={{ textAlign: "right" }}>{r.taxaPost === null ? "—" : `${r.taxaPost}%`}</td>
                <td style={{ textAlign: "right", fontWeight: 600 }}>{r.eficiencia === null ? "—" : `${r.eficiencia}%`}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="g-muted" style={{ fontSize: 10.5, marginBottom: 12, lineHeight: 1.5 }}>
        <b>% No Prazo</b>: executados com Data assinada ≤ Due date ÷ executados com as duas datas. <b>Taxa de Postergação</b>: postergados ÷ (executados + postergados).
        <b> Eficiência</b>: executados no prazo ÷ (executados + postergados + em atraso).
      </div>

      <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 8 }}>
        <div className="g-mode-toggle">
          <button className={vista === "abertos" ? "active" : ""} onClick={() => setVista("abertos")}>Em aberto ({due.length})</button>
          <button className={vista === "postergados" ? "active" : ""} onClick={() => setVista("postergados")}>Postergados ({post.length})</button>
        </div>
        <select value={criticidade} onChange={(e) => { setCriticidade(e.target.value); setGrupo("Todos"); }} style={selStyle}>
          <option value="Todas">Todas as criticidades</option>
          <option value="Operação">Critical to Operation</option>
          <option value="Segurança">Critical to Safety</option>
          <option value="Meio Ambiente">Critical to Environment</option>
        </select>
        <select value={grupo} onChange={(e) => setGrupo(e.target.value)} style={selStyle}>
          <option value="Todos">Todos os grupos</option>
          {gruposVis.map((g) => <option key={g.cod} value={g.cod}>{critLabel(g.cod)}</option>)}
        </select>
        {vista === "abertos" && (
          <select value={sit} onChange={(e) => setSit(e.target.value)} style={selStyle}>
            <option value="Todos">Todas as situações</option>
            <option>Vencido</option><option>A vencer (até 40d)</option><option>No prazo</option>
          </select>
        )}
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar..." style={{ ...selStyle, minWidth: 200 }} />
      </div>
      <div className="g-table-wrap" style={{ maxHeight: 420, overflowY: "auto" }}>
        {vista === "abertos" ? (
          <table className="g-table tbl-modern">
            <thead><tr><th>Grupo</th><th>Criticidade</th><th>Código</th><th>Componente</th><th>Job No</th><th>Job Type</th><th>Plano (Job Name)</th><th>Due</th><th>Situação</th></tr></thead>
            <tbody>
              {dueVis.length === 0 && <tr><td colSpan={9} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nenhum job crítico com esse filtro.</td></tr>}
              {dueVis.slice(0, 600).map((r) => (
                <tr className="g-row tbl-row" key={r.id} style={{ "--c": r._sit === "Vencido" ? "#EF4444" : r._sit === "No prazo" ? "#22C55E" : "#F5A623" }}>
                  <td><b>{critLabel(r._grupo)}</b></td><td>{critOf(r._grupo)}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{r.code || "—"}</td>
                  <td>{r.component || "—"}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{r.jobNo || "—"}</td>
                  <td>{r.jobType || "—"}</td>
                  <td style={{ whiteSpace: "normal", minWidth: 200 }}>{r.jobName || "—"}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{fmtD(r.dueDate, r.dueRaw)}</td>
                  <td><span className={`tbl-days ${sitCls[r._sit]}`}>{r._sit === "Vencido" ? `Vencido${r._atraso ? ` há ${r._atraso}` : ""}` : r._sit}</span></td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <table className="g-table tbl-modern">
            <thead><tr><th>Grupo</th><th>Criticidade</th><th>Código</th><th>Componente</th><th>Job No</th><th>Job Type</th><th>Plano (Job Name)</th><th>1ª Postergação</th><th>Última Postergação</th><th>Próximo Due Date</th></tr></thead>
            <tbody>
              {postVis.length === 0 && <tr><td colSpan={10} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nenhum job crítico postergado.</td></tr>}
              {postVis.slice(0, 600).map((r) => (
                <tr className="g-row tbl-row" key={r.key} style={{ "--c": "#8B5CF6" }}>
                  <td><b>{critLabel(r._grupo)}</b></td><td>{critOf(r._grupo)}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{r.componentCode || "—"}</td>
                  <td>{r.componentName || "—"}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{r.jobNo || "—"}</td>
                  <td>{r.jobType || "—"}</td>
                  <td style={{ whiteSpace: "normal", minWidth: 200 }}>{r.jobName || "—"}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{fmtD(r.primeiraData)}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{fmtD(r.ultimaData)}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{fmtD(r.proximoDue, r.proximoDueRaw)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
};

const PreditivosPanel = ({ rows, kpis }) => (
  <div className="g-panel">
    <div className="g-panel-head"><span className="g-panel-title">Jobs Preditivos</span></div>
    <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 10 }}>
      Cada análise reúne seus jobs e conta como um só: vários componentes ou jobs da mesma análise realizados na mesma data valem 1.
    </div>
    <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
      {bigKpi("Em Atraso", kpis.vencidos, "var(--crit)", AlertTriangle)}
      {bigKpi("A Vencer (até 40 dias)", kpis.aVencer, "var(--warn)", Clock)}
      {bigKpi("Postergados", kpis.postergados, "var(--warn)", AlertTriangle)}
      {bigKpi("% Executado no Prazo", `${kpis.pctPrazo}%`, kpis.pctPrazo >= 80 ? "var(--ok)" : kpis.pctPrazo >= 60 ? "var(--warn)" : "var(--crit)", ClipboardList)}
      {bigKpi("Índice de Eficiência", `${kpis.eficiencia}%`, kpis.eficiencia >= 80 ? "var(--ok)" : kpis.eficiencia >= 60 ? "var(--warn)" : "var(--crit)", Gauge)}
    </div>
    <div style={{ width: "100%", height: 260 }}>
      <ResponsiveContainer>
        <BarChart data={rows} margin={{ left: 0, right: 8, top: 10, bottom: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="var(--border-soft)" vertical={false} />
          <XAxis dataKey="grupo" tick={AX} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
          <YAxis allowDecimals={false} tick={AX} axisLine={false} tickLine={false} width={30} />
          <Tooltip contentStyle={TT} labelStyle={{ color: "var(--text)" }} />
          <Legend wrapperStyle={{ fontSize: 11 }} />
          <Bar dataKey="executados" name="Executados" stackId="a" fill="#22C55E" isAnimationActive={false} />
          <Bar dataKey="postergados" name="Postergados" stackId="a" fill="#8B5CF6" isAnimationActive={false} />
          <Bar dataKey="vencidos" name="Em atraso" stackId="a" fill="#EF4444" isAnimationActive={false} />
          <Bar dataKey="aVencer" name="A vencer (até 40d)" stackId="a" fill="#F5A623" radius={[3, 3, 0, 0]} isAnimationActive={false} />
        </BarChart>
      </ResponsiveContainer>
    </div>
    <div className="g-table-wrap" style={{ marginTop: 10 }}>
      <table className="g-table tbl-modern">
        <thead><tr>
          <th>Análise</th><th style={{ textAlign: "right" }}>Em Atraso</th><th style={{ textAlign: "right" }}>A Vencer</th>
          <th style={{ textAlign: "right" }}>Executados</th><th style={{ textAlign: "right" }}>Postergados</th><th style={{ textAlign: "right" }}>% No Prazo</th>
          <th style={{ textAlign: "right" }}>Taxa de Postergação</th><th style={{ textAlign: "right" }}>Eficiência</th>
        </tr></thead>
        <tbody>
          {rows.map((r) => (
            <tr className="g-row tbl-row" key={r.grupo} style={{ "--c": r.vencidos > 0 ? "#EF4444" : r.aVencer > 0 ? "#F5A623" : "#22C55E" }}>
              <td><b>{r.grupo}</b></td>
              <td style={{ textAlign: "right" }}>{r.vencidos}</td>
              <td style={{ textAlign: "right" }}>{r.aVencer}</td>
              <td style={{ textAlign: "right" }}>{r.executados}</td>
              <td style={{ textAlign: "right" }}>{r.postergados}</td>
              <td style={{ textAlign: "right" }}>{r.pctPrazo === null ? "—" : `${r.pctPrazo}%`}</td>
              <td style={{ textAlign: "right" }}>{r.taxaPost === null ? "—" : `${r.taxaPost}%`}</td>
              <td style={{ textAlign: "right", fontWeight: 600 }}>{r.eficiencia === null ? "—" : `${r.eficiencia}%`}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
    <div className="g-muted" style={{ fontSize: 10.5, marginTop: 6, lineHeight: 1.5 }}>
      <b>% No Prazo</b>: executados com Data assinada ≤ Due date ÷ executados com as duas datas. <b>Taxa de Postergação</b>: postergados ÷ (executados + postergados).
      <b> Eficiência</b>: executados no prazo ÷ (executados + postergados + em atraso) — quanto do que venceu foi realmente feito a tempo.
    </div>
  </div>
);

/* ============================================================
   TM MASTER — FAILURE REPORTS (abas do Excel: Failure Reports, Actions, Jobs, Job History, Orders)
   Todas as abas se ligam ao Failure Report pelo número FAIL-xxx:
   Actions.HSEQ Source Number · Jobs.HSEQ Source · Job History.HSEQ Source · Orders.Failure Report No
   ============================================================ */
const FR_EMPTY = { reports: [], actions: [], jobs: [], jobHistory: [], orders: [], importedAt: "" };
const frIsoDate = (v) => {
  if (!v) return "";
  if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;
  const m = String(v).trim().match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  return m ? `${m[3]}-${m[2]}-${m[1]}` : "";
};
const frStr = (v) => (v === null || v === undefined ? "" : String(v).trim());
const frActionConcluida = (a) => /^complete/i.test(a.doneStatus) || !!a.doneDate;
const frOrderEncerrada = (o) => /received|cancel|closed/i.test(o.status);
const frJobVencido = (j, hoje) => !!j.nextDue && j.nextDue < hoje;
const frDias = (iso, hoje) => (iso ? Math.floor((new Date(hoje + "T00:00:00") - new Date(iso + "T00:00:00")) / 86400000) : null);
const FR_FAIXAS = [["0-7 dias", 0, 7], ["8-30 dias", 8, 30], ["31-60 dias", 31, 60], ["61-90 dias", 61, 90], ["+90 dias", 91, 1e9]];
/* Failure Reports em aberto e o que falta em cada um: action, job (aberto ou executado) e order */
const frGaps = (F, hoje) => {
  const by = (arr) => { const m = new Map(); arr.forEach((x) => { if (x.frNo) m.set(x.frNo, (m.get(x.frNo) || 0) + 1); }); return m; };
  const a = by(F.actions), j = by(F.jobs), jh = by(F.jobHistory), o = by(F.orders);
  return F.reports.filter((r) => !/complet|closed/i.test(r.status)).map((r) => {
    const noAct = !a.get(r.no), noJob = !j.get(r.no) && !jh.get(r.no), noOrd = !o.get(r.no);
    const falta = [noAct && "action", noJob && "job", noOrd && "order"].filter(Boolean);
    return { r, noAct, noJob, noOrd, n: falta.length, falta, dias: frDias(r.failureDate || r.submittedDate, hoje) };
  });
};
const frMesKey = (iso) => (iso ? iso.slice(0, 7) : "");
const frMesLabel = (k) => { const [y, m] = k.split("-"); return `${["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][Number(m) - 1]}/${y.slice(2)}`; };

/* converte as abas do Excel em registros normalizados */
function parseFailureWorkbook(wb) {
  const sheet = (name) => {
    const n = wb.SheetNames.find((s) => s.trim().toLowerCase() === name);
    return n ? XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: "", cellDates: true }) : null;
  };
  const out = {};
  const rep = sheet("failure reports");
  if (rep) out.reports = rep.map((r) => ({
    no: frStr(r["Failure Report No."]), name: frStr(r["Failure Report Name"]), unit: frStr(r["Unit"]),
    failureDate: frIsoDate(r["Time of failure"]), failureRaw: frStr(r["Time of failure"] instanceof Date ? "" : r["Time of failure"]),
    discoveredDuring: frStr(r["Discovered During"]), activity: frStr(r["Operation Activity"]), mechanism: frStr(r["Failure Mechanism"]),
    mode: frStr(r["Failure Mode"]), cause: frStr(r["Cause"]), createdBy: frStr(r["Created By"]), oopTime: frStr(r["OOP Time"]),
    department: frStr(r["Responsible Department"]), crew: frStr(r["Responsible Crew"]),
    compCode: frStr(r["Component Code"]), compName: frStr(r["Component Name"]), parent: frStr(r["Parent Component"]),
    criticality: frStr(r["Component Criticality"]).replace(/\*/g, ""), system: frStr(r["Component System"]), maker: frStr(r["Component Maker"]),
    jobsToDo: frStr(r["Jobs To Do"]), jobsDone: frStr(r["Jobs Done"]), sparePart: frStr(r["Spare Part Name"]),
    message: frStr(r["Message"]), status: frStr(r["Status"]), effect: frStr(r["Failure Effect"]), severity: frStr(r["Severity"]),
    probability: frStr(r["Probability"]), risk: frStr(r["Risk"]), nonProductive: frStr(r["Non Productive Time"]),
    timeToRepair: frStr(r["Time to Repair"]), manHours: frStr(r["Man Hours to Repair"]),
    submittedDate: frIsoDate(r["Submitted Date"]), closedDate: frIsoDate(r["Closed Date"]), tags: frStr(r["Failure report tags"]),
  })).filter((r) => r.no);
  const act = sheet("actions");
  if (act) out.actions = act.map((r) => ({
    no: frStr(r["Action No."]), name: frStr(r["Action Name"]), frNo: frStr(r["HSEQ Source Number"]), frName: frStr(r["HSEQ Source Name"]),
    dueStatus: frStr(r["Due Status"]), deadline: frIsoDate(r["Deadline (Due date)"]), doneStatus: frStr(r["Done Status"]),
    category: frStr(r["Action Category"]), type: frStr(r["Action Type"]), priority: frStr(r["Action Priority"]),
    department: frStr(r["Responsible Department"]), crewType: frStr(r["Responsible Crew Type"]), crew: frStr(r["Responsible Crew"]),
    task: frStr(r["Task Description"]), doneDate: frIsoDate(r["Done Date"]), doneBy: frStr(r["Done By"]),
    actionTaken: frStr(r["Action Taken"]), created: frIsoDate(r["Date created"]),
  })).filter((r) => r.no);
  const jobs = sheet("jobs");
  if (jobs) out.jobs = jobs.map((r) => ({
    jobType: frStr(r["Job Type"]), jobNo: frStr(r["JobNo"]), name: frStr(r["Name"]), priority: frStr(r["Priority"]), department: frStr(r["Department"]),
    interval: frStr(r["Interval"]), doneStatus: frStr(r["JobDoneStatus"]), nextDue: frIsoDate(r["Next Due"]), diff: frStr(r["Diff"]),
    created: frIsoDate(r["Created Date"]), relatedAction: frStr(r["Related Action"]), frNo: frStr(r["HSEQ Source"]) || frStr(r["Source No. & Name"]).split(" ")[0],
  })).filter((r) => r.jobNo);
  const jh = sheet("job history");
  if (jh) out.jobHistory = jh.map((r) => ({
    number: frStr(r["Job History Number"]), jobType: frStr(r["JobType"]), jobNo: frStr(r["JobNo"]), name: frStr(r["JobName"]),
    compCode: frStr(r["ComponentCode"]), compName: frStr(r["ComponentName"]), dateDone: frIsoDate(r["DateDone"]), dateSigned: frIsoDate(r["Date signed"]),
    doneBy: frStr(r["DoneByName"]), signedBy: frStr(r["Signed by"]), remarks: frStr(r["Remarks"]), reason: frStr(r["Reason"]), symptom: frStr(r["Symptom"]),
    manHours: frStr(r["ManHours"]), conditionBefore: frStr(r["Condition before"]), conditionAfter: frStr(r["Condition after"]),
    relatedAction: frStr(r["Related Action"]), frNo: frStr(r["HSEQ Source"]),
  })).filter((r) => r.number);
  const ord = sheet("orders");
  if (ord) out.orders = ord.map((r) => ({
    no: frStr(r["Order No."]), name: frStr(r["Name"]), status: frStr(r["Status"]), frNo: frStr(r["Failure Report No"]),
    supplier: frStr(r["Supplier company"]), price: r["Total Price"], currency: frStr(r["Order currency"]), department: frStr(r["Department"]),
    orderType: frStr(r["Order type"]), delivery: frIsoDate(r["Delivery date"]), approved: frIsoDate(r["Date last approved"]),
    account: frStr(r["Account name"]), asap: frStr(r["Asap"]), critical: frStr(r["Critical"]),
  })).filter((r) => r.no);
  return out;
}
const mergeBy = (prev, rows, keyFn) => {
  const map = new Map(prev.map((x) => [keyFn(x), x]));
  let add = 0, upd = 0;
  rows.forEach((r) => { const k = keyFn(r); if (map.has(k)) { upd++; map.set(k, { ...map.get(k), ...r }); } else { add++; map.set(k, r); } });
  return { list: [...map.values()], add, upd };
};

const FailureReportsView = ({ tmFailures, handleImportFailures, setReportFn, setExportXlsxFn, active }) => {
  const fileRef = useRef(null);
  const hoje = todayISO();
  const F = tmFailures || FR_EMPTY;
  const [f, setF] = useState({ busca: "", status: "Todos", dept: "Todos", crit: "Todas", mes: [], pend: false });
  const [open, setOpen] = useState(() => new Set());
  const toggleOpen = (no) => setOpen((p) => { const n = new Set(p); if (n.has(no)) n.delete(no); else n.add(no); return n; });

  /* índice: Failure Report No -> registros ligados nas outras abas */
  const links = useMemo(() => {
    const m = new Map();
    const get = (no) => { if (!m.has(no)) m.set(no, { actions: [], jobs: [], jobHistory: [], orders: [] }); return m.get(no); };
    F.actions.forEach((x) => x.frNo && get(x.frNo).actions.push(x));
    F.jobs.forEach((x) => x.frNo && get(x.frNo).jobs.push(x));
    F.jobHistory.forEach((x) => x.frNo && get(x.frNo).jobHistory.push(x));
    F.orders.forEach((x) => x.frNo && get(x.frNo).orders.push(x));
    return m;
  }, [F]);
  const L = (no) => links.get(no) || { actions: [], jobs: [], jobHistory: [], orders: [] };
  const pendencias = (r) => { const l = L(r.no); return l.actions.filter((a) => !frActionConcluida(a)).length + l.jobs.length + l.orders.filter((o) => !frOrderEncerrada(o)).length; };

  const orfaos = useMemo(() => {
    const nos = new Set(F.reports.map((r) => r.no));
    const c = (arr) => arr.filter((x) => x.frNo && !nos.has(x.frNo)).length;
    return { actions: c(F.actions), jobs: c(F.jobs), jobHistory: c(F.jobHistory), orders: c(F.orders) };
  }, [F]);

  const opts = (k) => ["Todos", ...[...new Set(F.reports.map((r) => r[k]).filter(Boolean))].sort()];
  const meses = ["Todos", ...[...new Set(F.reports.map((r) => frMesKey(r.failureDate)).filter(Boolean))].sort().reverse()];
  const norm = (v) => String(v || "").toLowerCase();
  const filtered = useMemo(() => F.reports.filter((r) =>
    (f.status === "Todos" || r.status === f.status) && (f.dept === "Todos" || r.department === f.dept)
    && (f.crit === "Todas" || r.criticality === f.crit) && (f.mes.length === 0 || f.mes.includes(frMesKey(r.failureDate)))
    && (!f.pend || pendencias(r) > 0)
    && (!f.busca || norm(`${r.no} ${r.name} ${r.compCode} ${r.compName} ${r.cause} ${r.mechanism} ${r.createdBy}`).includes(norm(f.busca)))
  ).sort((a, b) => (b.failureDate || "").localeCompare(a.failureDate || "") || b.no.localeCompare(a.no)), [F, f, links]);

  const nos = new Set(filtered.map((r) => r.no));
  const nameOf = (no) => F.reports.find((r) => r.no === no)?.name || "";
  const vis = (arr) => arr.filter((x) => nos.has(x.frNo));
  const kpi = useMemo(() => ({
    total: filtered.length,
    abertos: filtered.filter((r) => !/complet|closed/i.test(r.status)).length,
    actAbertas: vis(F.actions).filter((a) => !frActionConcluida(a)).length,
    actVencidas: vis(F.actions).filter((a) => !frActionConcluida(a) && a.deadline && a.deadline < hoje).length,
    jobsAbertos: vis(F.jobs).length, jobsVencidos: vis(F.jobs).filter((j) => frJobVencido(j, hoje)).length,
    ordersAbertas: vis(F.orders).filter((o) => !frOrderEncerrada(o)).length,
  }), [filtered, F]);

  /* PENDÊNCIAS: tudo que está em aberto nos Failure Reports filtrados, com há quantos dias está parado */
  const [pendTipo, setPendTipo] = useState("Action");
  const pend = useMemo(() => {
    const out = [];
    vis(F.actions).filter((a) => !frActionConcluida(a)).forEach((a) => {
      const desde = a.created || a.deadline; const atraso = a.deadline && a.deadline < hoje ? frDias(a.deadline, hoje) : 0;
      out.push({ tipo: "Action", frNo: a.frNo, frName: nameOf(a.frNo), ref: a.no, desc: a.name, resp: a.crew || a.crewType || a.department || "—", dept: a.department || "—",
        desde, dias: frDias(desde, hoje), prazo: a.deadline, atraso, situacao: a.dueStatus || a.doneStatus || "Em aberto", prioridade: a.priority });
    });
    vis(F.jobs).forEach((j) => {
      const desde = j.created || j.nextDue; const atraso = j.nextDue && j.nextDue < hoje ? frDias(j.nextDue, hoje) : 0;
      out.push({ tipo: "Job", frNo: j.frNo, frName: nameOf(j.frNo), ref: `${j.jobType} ${j.jobNo}`, desc: j.name, resp: j.department || "—", dept: j.department || "—",
        desde, dias: frDias(desde, hoje), prazo: j.nextDue, atraso, situacao: atraso ? "Vencido" : "Em aberto", prioridade: j.priority });
    });
    vis(F.orders).filter((o) => !frOrderEncerrada(o)).forEach((o) => {
      const desde = o.approved || o.delivery; const atraso = o.delivery && o.delivery < hoje ? frDias(o.delivery, hoje) : 0;
      out.push({ tipo: "Order", frNo: o.frNo, frName: nameOf(o.frNo), ref: o.no, desc: o.name, resp: o.supplier || o.department || "—", dept: o.department || "—",
        desde, dias: frDias(desde, hoje), prazo: o.delivery, atraso, situacao: o.status || "Em aberto", prioridade: "" });
    });
    filtered.filter((r) => !/complet|closed/i.test(r.status)).forEach((r) => {
      const desde = r.failureDate || r.submittedDate; out.push({ tipo: "Failure Report", frNo: r.no, frName: r.name, ref: r.no, desc: r.name, resp: r.createdBy || "—", dept: r.department || "—",
        desde, dias: frDias(desde, hoje), prazo: "", atraso: 0, situacao: r.status || "Em aberto", prioridade: "" });
    });
    return out.sort((a, b) => (b.dias ?? -1) - (a.dias ?? -1));
  }, [filtered, F]);
  const pendVis = pend.filter((p) => pendTipo === "Todos" || p.tipo === pendTipo);
  const pendResumo = useMemo(() => {
    const tipos = ["Action", "Job", "Order", "Failure Report"];
    const porTipo = tipos.map((t) => { const l = pend.filter((p) => p.tipo === t); const ds = l.map((p) => p.dias).filter((d) => d !== null);
      return { tipo: t, total: l.length, vencidos: l.filter((p) => p.atraso > 0).length, media: ds.length ? Math.round(ds.reduce((a, b) => a + b, 0) / ds.length) : 0, maisAntigo: ds.length ? Math.max(...ds) : 0 }; });
    const resp = new Map(); pend.filter((p) => p.tipo === "Action").forEach((p) => { const k = p.resp; const o = resp.get(k) || { resp: k, total: 0, vencidos: 0, maisAntigo: 0 }; o.total++; if (p.atraso > 0) o.vencidos++; o.maisAntigo = Math.max(o.maisAntigo, p.dias || 0); resp.set(k, o); });
    const faixas = FR_FAIXAS.map(([nome, a, b]) => ({ nome, total: pend.filter((p) => p.dias !== null && p.dias >= a && p.dias <= b).length }));
    const acts = pend.filter((p) => p.tipo === "Action");
    const frsComPend = new Set(pend.filter((p) => p.tipo !== "Failure Report").map((p) => p.frNo)).size;
    return { porTipo, porResp: [...resp.values()].sort((a, b) => b.maisAntigo - a.maisAntigo), faixas, acts, frsComPend,
      maisAntiga: pend[0] || null, vencidasTotal: pend.filter((p) => p.atraso > 0).length };
  }, [pend]);
  const diasTxt = (d) => (d === null || d === undefined ? "—" : `${d} d`);

  /* ANÁLISE: Failure Reports em aberto sem action, sem job ou sem order */
  const [gapTipo, setGapTipo] = useState("Com lacuna");
  const gaps = useMemo(() => frGaps({ ...F, reports: filtered }, hoje).sort((a, b) => b.n - a.n || (b.dias ?? -1) - (a.dias ?? -1)), [filtered, F]);
  const gapsVis = gaps.filter((g) => (gapTipo === "Com lacuna" ? g.n > 0 : gapTipo === "Sem action" ? g.noAct : gapTipo === "Sem job" ? g.noJob : gapTipo === "Sem order" ? g.noOrd : g.n === 3));
  const gapKpi = { abertos: gaps.length, semAct: gaps.filter((g) => g.noAct).length, semJob: gaps.filter((g) => g.noJob).length, semOrd: gaps.filter((g) => g.noOrd).length, nada: gaps.filter((g) => g.n === 3).length };
  const abrirFr = (no) => { setOpen((p) => new Set(p).add(no)); setTimeout(() => document.getElementById(`fr-row-${no}`)?.scrollIntoView({ behavior: "smooth", block: "start" }), 80); };
  const porMes = useMemo(() => {
    const m = new Map(); filtered.forEach((r) => { const k = frMesKey(r.failureDate); if (k) m.set(k, (m.get(k) || 0) + 1); });
    return [...m.entries()].sort().map(([k, v]) => ({ mes: frMesLabel(k), total: v }));
  }, [filtered]);
  const contar = (key) => { const m = new Map(); filtered.forEach((r) => { const k = r[key] || "—"; m.set(k, (m.get(k) || 0) + 1); }); return [...m.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8).map(([nome, total]) => ({ nome, total })); };
  const porSistema = useMemo(() => contar("system"), [filtered]);
  const porStatus = useMemo(() => contar("status"), [filtered]);

  const openReps = filtered.filter((r) => open.has(r.no));

  /* exportações: seguem os filtros e trazem o Failure Report ligado em cada aba */
  const pack = () => {
    const resumo = filtered.map((r) => { const l = L(r.no); return {
      "Failure Report No.": r.no, Nome: r.name, "Data da falha": r.failureDate ? fmtDataCompleta(r.failureDate) : r.failureRaw, Status: r.status, Departamento: r.department,
      "Código do componente": r.compCode, Componente: r.compName, Criticidade: r.criticality, Sistema: r.system, Mecanismo: r.mechanism, Causa: r.cause, Efeito: r.effect,
      "Actions (total)": l.actions.length, "Actions abertas": l.actions.filter((a) => !frActionConcluida(a)).length, "Jobs abertos": l.jobs.length,
      "Job History": l.jobHistory.length, Orders: l.orders.length, "Orders em aberto": l.orders.filter((o) => !frOrderEncerrada(o)).length }; });
    const withFr = (arr, mapFn) => vis(arr).map((x) => ({ "Failure Report No.": x.frNo, "Failure Report": nameOf(x.frNo), ...mapFn(x) }));
    return {
      resumo,
      actions: withFr(F.actions, (a) => ({ "Action No.": a.no, Ação: a.name, Situação: a.dueStatus, Prazo: fmtDataCompleta(a.deadline), "Status": a.doneStatus, Categoria: a.category, Tipo: a.type, Prioridade: a.priority, Departamento: a.department, Responsável: a.crew || a.crewType, Descrição: a.task, "Concluída em": fmtDataCompleta(a.doneDate), "Concluída por": a.doneBy, "Ação tomada": a.actionTaken })),
      jobs: withFr(F.jobs, (j) => ({ "Job Type": j.jobType, "Job No": j.jobNo, Plano: j.name, Prioridade: j.priority, Departamento: j.department, Status: j.doneStatus, "Next Due": fmtDataCompleta(j.nextDue), Diff: j.diff })),
      hist: withFr(F.jobHistory, (h) => ({ "Job History": h.number, "Job Type": h.jobType, "Job No": h.jobNo, Plano: h.name, Componente: `${h.compCode} ${h.compName}`.trim(), "Data feita": fmtDataCompleta(h.dateDone), "Assinado em": fmtDataCompleta(h.dateSigned), "Feito por": h.doneBy, Motivo: h.reason, Sintoma: h.symptom, Observações: h.remarks })),
      orders: withFr(F.orders, (o) => ({ "Order No.": o.no, Pedido: o.name, Status: o.status, Fornecedor: o.supplier, "Valor total": o.price, Moeda: o.currency, Entrega: fmtDataCompleta(o.delivery), Departamento: o.department })),
    };
  };

  React.useEffect(() => {
    if (!active || !setExportXlsxFn || !setReportFn) return;
    setExportXlsxFn(() => () => {
      const p = pack(); const wb = XLSX.utils.book_new();
      const add = (n, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ "—": "Sem registros" }]), n);
      add("Failure Reports", p.resumo); add("Pendências (mais antigas)", pend.map((x) => ({ Tipo: x.tipo, "Failure Report No.": x.frNo, "Failure Report": x.frName, Referência: x.ref, Descrição: x.desc, Responsável: x.resp, Departamento: x.dept, "Aberto desde": fmtDataCompleta(x.desde), "Dias parado": x.dias, Prazo: fmtDataCompleta(x.prazo), "Dias de atraso": x.atraso || 0, Situação: x.situacao, Prioridade: x.prioridade }))); add("FR sem tratativa", gaps.filter((g) => g.n > 0).map((g) => ({ "Failure Report No.": g.r.no, Nome: g.r.name, "Data da falha": g.r.failureDate ? fmtDataCompleta(g.r.failureDate) : "", "Dias aberto": g.dias, Status: g.r.status, Departamento: g.r.department, "Tem action": g.noAct ? "NÃO" : "sim", "Tem job": g.noJob ? "NÃO" : "sim", "Tem order": g.noOrd ? "NÃO" : "sim", "O que falta": g.falta.join(", ") }))); add("Actions", p.actions); add("Jobs", p.jobs); add("Job History", p.hist); add("Orders", p.orders);
      XLSX.writeFile(wb, `genesis-tm-master-failure-reports-${todayISO()}.xlsx`);
    });
    setReportFn(() => () => {
      const doc = new jsPDF({ orientation: "landscape" });
      let y = pdfHeader(doc, "TM Master — Failure Reports", `${filtered.length} failure report(s) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [
        { label: "Failure Reports", value: kpi.total }, { label: "Em Aberto", value: kpi.abertos }, { label: "Actions Abertas", value: kpi.actAbertas },
        { label: "Actions Vencidas", value: kpi.actVencidas }, { label: "Jobs Abertos", value: kpi.jobsAbertos }, { label: "Orders em Aberto", value: kpi.ordersAbertas },
      ]);
      y = pdfSectionTitle(doc, y, "Failure Reports e itens ligados");
      y = pdfTable(doc, y, ["Nº", "Nome", "Data", "Status", "Dept.", "Componente", "Criticidade", "Actions (abertas/total)", "Jobs abertos", "Job History", "Orders (abertas/total)"],
        filtered.map((r) => { const l = L(r.no); return [r.no, r.name, r.failureDate ? fmtDataCompleta(r.failureDate) : "—", r.status, r.department, `${r.compCode} ${r.compName}`.trim(), r.criticality || "—",
          `${l.actions.filter((a) => !frActionConcluida(a)).length}/${l.actions.length}`, l.jobs.length, l.jobHistory.length, `${l.orders.filter((o) => !frOrderEncerrada(o)).length}/${l.orders.length}`]; }));
      openReps.forEach((selRep) => {
        y = pdfSectionTitle(doc, y, `Failure Report completo — ${selRep.no} · ${selRep.name}`);
        const l = L(selRep.no);
        y = pdfTable(doc, y, ["Data", "Status", "Componente", "Causa", "Mecanismo", "Efeito", "Severidade", "Risco", "Criado por"], [[selRep.failureDate ? fmtDataCompleta(selRep.failureDate) : "—", selRep.status || "—", `${selRep.compCode} ${selRep.compName}`.trim(), selRep.cause || "—", selRep.mechanism || "—", selRep.effect || "—", selRep.severity || "—", selRep.risk || "—", selRep.createdBy || "—"]]);
        if (l.actions.length) y = pdfTable(doc, y, ["Action", "Ação", "Prazo", "Status", "Prioridade", "Responsável", "Ação tomada"], l.actions.map((a) => [a.no, a.name, fmtDataCompleta(a.deadline), a.doneStatus, a.priority, a.crew || a.crewType, a.actionTaken]));
        if (l.jobs.length) y = pdfTable(doc, y, ["Job", "Plano", "Status", "Next Due", "Diff"], l.jobs.map((j) => [`${j.jobType} ${j.jobNo}`, j.name, j.doneStatus || "—", fmtDataCompleta(j.nextDue), j.diff]));
        if (l.jobHistory.length) y = pdfTable(doc, y, ["Histórico", "Job", "Componente", "Data feita", "Feito por", "Observações"], l.jobHistory.map((h) => [h.number, `${h.jobType} ${h.jobNo}`, `${h.compCode} ${h.compName}`.trim(), fmtDataCompleta(h.dateDone), h.doneBy, h.remarks]));
        if (l.orders.length) y = pdfTable(doc, y, ["Order", "Pedido", "Status", "Fornecedor", "Entrega"], l.orders.map((o) => [o.no, o.name, o.status, o.supplier, fmtDataCompleta(o.delivery)]));
      });
      y = pdfSectionTitle(doc, y, `Failure Reports em aberto sem tratativa — ${gaps.filter((g) => g.n > 0).length} de ${gaps.length}`);
      y = pdfTable(doc, y, ["Failure Report", "Nome", "Data", "Dias aberto", "Status", "Action", "Job", "Order", "O que falta"],
        gaps.filter((g) => g.n > 0).map((g) => [g.r.no, g.r.name, g.r.failureDate ? fmtDataCompleta(g.r.failureDate) : "—", g.dias ?? "—", g.r.status, g.noAct ? "NÃO" : "sim", g.noJob ? "NÃO" : "sim", g.noOrd ? "NÃO" : "sim", g.falta.join(", ")]));
      y = pdfSectionTitle(doc, y, `Resumo de pendências — ${pend.length} item(ns) em aberto, do mais antigo ao mais recente`);
      y = pdfTable(doc, y, ["Tipo", "Em aberto", "Vencidos", "Média parado (dias)", "Mais antigo (dias)"], pendResumo.porTipo.map((t) => [t.tipo, t.total, t.vencidos, t.media, t.maisAntigo]));
      y = pdfTable(doc, y, ["Tipo", "Failure Report", "Referência", "Descrição", "Responsável", "Aberto desde", "Dias parado", "Prazo", "Atraso (dias)", "Situação"],
        pend.slice(0, 80).map((x) => [x.tipo, x.frNo, x.ref, x.desc, x.resp, fmtDataCompleta(x.desde), x.dias ?? "—", fmtDataCompleta(x.prazo), x.atraso || "—", x.situacao]));
      pdfSave(doc, "relatorio-tm-master-failure-reports");
    });
  }, [active, filtered, kpi, openReps, F, pend, pendResumo, gaps]);

  const selStyle = { padding: "6px 10px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 12.5 };
  const stCls = (s) => (/complet|closed/i.test(s) ? "d1" : /draft/i.test(s) ? "d2" : "d3");
  /* visão única do Failure Report: dados da falha + linha do tempo + Actions, Jobs, Job History e Orders, tudo junto */
  const Detalhe = ({ r }) => {
    const l = L(r.no);
    const ev = [];
    if (r.failureDate) ev.push({ d: r.failureDate, t: "Falha", x: `${r.name}${r.compName ? ` — ${r.compCode} ${r.compName}` : ""}`, c: "#EF4444" });
    l.actions.forEach((a) => { if (a.created) ev.push({ d: a.created, t: "Action criada", x: `${a.no} — ${a.name}`, c: "#3B82F6" });
      if (a.doneDate) ev.push({ d: a.doneDate, t: "Action concluída", x: `${a.no} — ${a.name}${a.doneBy ? ` (${a.doneBy})` : ""}`, c: "#22C55E" });
      else if (a.deadline) ev.push({ d: a.deadline, t: a.deadline < hoje ? "Action vencida (prazo)" : "Action: prazo", x: `${a.no} — ${a.name}`, c: a.deadline < hoje ? "#EF4444" : "#F5A623" }); });
    l.jobs.forEach((j) => { if (j.created) ev.push({ d: j.created, t: "Job criado", x: `${j.jobType} ${j.jobNo} — ${j.name}`, c: "#3B82F6" });
      if (j.nextDue) ev.push({ d: j.nextDue, t: j.nextDue < hoje ? "Job vencido (due)" : "Job: due", x: `${j.jobType} ${j.jobNo} — ${j.name}`, c: j.nextDue < hoje ? "#EF4444" : "#F5A623" }); });
    l.jobHistory.forEach((h) => { if (h.dateDone) ev.push({ d: h.dateDone, t: "Job executado", x: `${h.jobType} ${h.jobNo} — ${h.name || h.compName}${h.doneBy ? ` (${h.doneBy})` : ""}`, c: "#22C55E" }); });
    l.orders.forEach((o) => { if (o.approved) ev.push({ d: o.approved, t: "Order aprovada", x: `${o.no} — ${o.name}`, c: "#8B5CF6" });
      if (o.delivery) ev.push({ d: o.delivery, t: frOrderEncerrada(o) ? "Order entregue/encerrada" : (o.delivery < hoje ? "Order: entrega atrasada" : "Order: entrega prevista"), x: `${o.no} — ${o.name}`, c: frOrderEncerrada(o) ? "#22C55E" : (o.delivery < hoje ? "#EF4444" : "#8B5CF6") }); });
    ev.sort((a, b) => a.d.localeCompare(b.d));
    const sec = (titulo, n, corpo) => (<div style={{ marginBottom: 14 }}><div style={{ fontSize: 12, fontWeight: 700, margin: "0 0 6px", color: "var(--text)" }}>{titulo} <span className="g-muted">({n})</span></div>{corpo}</div>);
    return (
      <div style={{ padding: "14px 16px", background: "var(--panel-raised, rgba(0,0,0,0.02))", borderLeft: "3px solid var(--accent, #3B82F6)" }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: "8px 18px", marginBottom: 14, fontSize: 12.5 }}>
          {[["Data da falha", r.failureDate ? fmtDataCompleta(r.failureDate) : "—"], ["Componente", `${r.compCode} ${r.compName}`.trim() || "—"], ["Descoberto durante", r.discoveredDuring], ["Atividade", r.activity],
            ["Mecanismo", r.mechanism], ["Causa", r.cause], ["Efeito", r.effect], ["Severidade / Risco", [r.severity, r.risk].filter(Boolean).join(" / ")], ["Criado por", r.createdBy], ["Fechado em", r.closedDate ? fmtDataCompleta(r.closedDate) : ""]].map(([k, v]) => (
            <div key={k} style={{ minWidth: 0, overflowWrap: "anywhere" }}><div className="g-muted" style={{ fontSize: 10.5, textTransform: "uppercase", letterSpacing: 0.4 }}>{k}</div><div>{v || "—"}</div></div>))}
        </div>
        {sec("Linha do tempo (tudo junto, em ordem de data)", ev.length, ev.length === 0 ? <div className="g-muted" style={{ fontSize: 12 }}>Sem datas registradas.</div> : (
          <div style={{ borderLeft: "2px solid var(--border)", marginLeft: 6, paddingLeft: 14, display: "flex", flexDirection: "column", gap: 5 }}>
            {ev.map((e, i) => (<div key={i} style={{ fontSize: 12.5, position: "relative" }}>
              <span style={{ position: "absolute", left: -20, top: 5, width: 9, height: 9, borderRadius: "50%", background: e.c }} />
              <span style={{ fontFamily: "var(--mono)", marginRight: 8 }}>{fmtDataCompleta(e.d)}</span><b style={{ color: e.c, marginRight: 8 }}>{e.t}</b>{e.x}</div>))}
          </div>))}
        {sec("Actions", l.actions.length, <Tbl cols={["Action", "Ação", "Prazo", "Status", "Prioridade", "Responsável", "Concluída em", "Ação tomada"]} empty="Nenhuma action ligada."
          rows={l.actions.map((a) => [a.no, a.name, fmtDataCompleta(a.deadline), a.doneStatus || a.dueStatus, a.priority, a.crew || a.crewType || a.department, fmtDataCompleta(a.doneDate), a.actionTaken])} />)}
        {sec("Jobs (em aberto)", l.jobs.length, <Tbl cols={["Job", "Plano", "Prioridade", "Status", "Next Due", "Diff"]} empty="Nenhum job em aberto ligado."
          rows={l.jobs.map((j) => [`${j.jobType} ${j.jobNo}`, j.name, j.priority, j.doneStatus, fmtDataCompleta(j.nextDue), j.diff])} />)}
        {sec("Job History (executados)", l.jobHistory.length, <Tbl cols={["Histórico", "Job", "Componente", "Data feita", "Feito por", "Motivo", "Observações"]} empty="Nenhum histórico ligado."
          rows={l.jobHistory.map((h) => [h.number, `${h.jobType} ${h.jobNo}`, `${h.compCode} ${h.compName}`.trim(), fmtDataCompleta(h.dateDone), h.doneBy, h.reason, h.remarks])} />)}
        {sec("Orders", l.orders.length, <Tbl cols={["Order", "Pedido", "Status", "Fornecedor", "Valor", "Entrega"]} empty="Nenhuma order ligada."
          rows={l.orders.map((o) => [o.no, o.name, o.status, o.supplier, `${o.price ?? ""} ${o.currency}`.trim(), fmtDataCompleta(o.delivery)])} />)}
      </div>
    );
  };

  const vazio = F.reports.length === 0;
  const Tbl = ({ cols, rows, empty }) => (
    <div className="g-table-wrap" style={{ maxHeight: 360, overflowY: "auto" }}>
      <table className="g-table tbl-modern">
        <thead><tr>{cols.map((c) => <th key={c}>{c}</th>)}</tr></thead>
        <tbody>
          {rows.length === 0 && <tr><td colSpan={cols.length} className="g-muted" style={{ textAlign: "center", padding: 16 }}>{empty}</td></tr>}
          {rows.map((r, i) => <tr className="g-row tbl-row" key={i}>{r.map((c, j) => <td key={j} style={j === 0 ? { fontFamily: "var(--mono)" } : { whiteSpace: "normal" }}>{c === "" || c === null || c === undefined ? "—" : c}</td>)}</tr>)}
        </tbody>
      </table>
    </div>
  );

  return (
    <>
      <div className="g-panel" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
        <div>
          <div className="g-panel-title" style={{ marginBottom: 4 }}>Importar Failure Reports</div>
          <div className="g-muted" style={{ fontSize: 11.5, maxWidth: 780 }}>
            Suba o Excel com as abas <strong>Failure Reports, Actions, Jobs, Job History e Orders</strong>. Cada aba é ligada ao Failure Report pelo número <strong>FAIL-xxx</strong>
            (Actions → HSEQ Source Number · Jobs e Job History → HSEQ Source · Orders → Failure Report No). A importação mescla por número, então nada é duplicado.
            {F.importedAt && <> Última importação: <strong>{F.importedAt}</strong>.</>}
          </div>
        </div>
        <div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImportFailures} />
          <button className="g-btn primary" onClick={() => fileRef.current?.click()}><Upload size={14} />Importar planilha Failure Reports</button>
        </div>
      </div>

      {vazio ? (
        <div className="g-panel"><div className="g-muted" style={{ textAlign: "center", padding: 24 }}>Nenhum Failure Report importado ainda.</div></div>
      ) : (
        <>
          {(orfaos.actions + orfaos.jobs + orfaos.jobHistory + orfaos.orders) > 0 && (
            <div className="g-alert" style={{ marginBottom: 12 }}>
              <AlertTriangle size={14} style={{ marginTop: 1 }} />
              Itens ligados a um Failure Report que não está na lista importada: {orfaos.actions} action(s), {orfaos.jobs} job(s), {orfaos.jobHistory} histórico(s), {orfaos.orders} order(s). Eles aparecem quando o Failure Report correspondente for importado.
            </div>
          )}
          <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
            <div className="g-field"><label>Buscar</label>
              <input type="text" value={f.busca} onChange={(e) => setF((p) => ({ ...p, busca: e.target.value }))} placeholder="nº, nome, componente, causa..." style={{ minWidth: 220 }} /></div>
            <div className="g-field"><label>Status</label>
              <select value={f.status} onChange={(e) => setF((p) => ({ ...p, status: e.target.value }))}>{opts("status").map((o) => <option key={o}>{o}</option>)}</select></div>
            <div className="g-field"><label>Departamento</label>
              <select value={f.dept} onChange={(e) => setF((p) => ({ ...p, dept: e.target.value }))}>{opts("department").map((o) => <option key={o}>{o}</option>)}</select></div>
            <div className="g-field"><label>Criticidade</label>
              <select value={f.crit} onChange={(e) => setF((p) => ({ ...p, crit: e.target.value }))}>
                <option value="Todas">Todas</option>{opts("criticality").filter((o) => o !== "Todos").map((o) => <option key={o}>{o}</option>)}</select></div>
            <div className="g-field"><label>Mês da falha</label>
              <FiltroMeses meses={meses.filter((m) => m !== "Todos")} sel={f.mes} onChange={(v) => setF((p) => ({ ...p, mes: v }))} label={frMesLabel} /></div>
            <div className="g-field"><label>Pendências</label>
              <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, cursor: "pointer", padding: "6px 0" }}>
                <input type="checkbox" checked={f.pend} onChange={(e) => setF((p) => ({ ...p, pend: e.target.checked }))} />Só com itens em aberto</label></div>
          </div>

          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(6, 1fr)" }}>
            {bigKpi("Failure Reports", kpi.total, "var(--teal)", ClipboardList)}
            {bigKpi("Em Aberto", kpi.abertos, "var(--warn)", AlertTriangle)}
            {bigKpi("Actions Abertas", kpi.actAbertas, "var(--warn)", ClipboardList)}
            {bigKpi("Actions Vencidas", kpi.actVencidas, "var(--crit)", AlertTriangle)}
            {bigKpi("Jobs Abertos", `${kpi.jobsAbertos}${kpi.jobsVencidos ? ` (${kpi.jobsVencidos} venc.)` : ""}`, "var(--teal)", Wrench)}
            {bigKpi("Orders em Aberto", kpi.ordersAbertas, "var(--teal)", Package)}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(340px, 1fr))", gap: 14, marginBottom: 14 }}>
            <MesBarChart title="Failure Reports por Mês" height={220} data={porMes} bars={[{ key: "total", name: "Failure Reports", color: "#EF4444" }]} nota="Mês da data da falha." />
            <MesBarChart title="Por Sistema (top 8)" height={220} data={porSistema} bars={[{ key: "total", name: "Failure Reports", color: "#3B82F6" }]} />
            <MesBarChart title="Por Status" height={220} data={porStatus} bars={[{ key: "total", name: "Failure Reports", color: "#8B5CF6" }]} />
          </div>

          <div className="g-panel" style={{ borderTop: "3px solid var(--crit)" }}>
            <div className="g-panel-head"><span className="g-panel-title">Análise — Failure Reports em aberto sem tratativa (sem action, job ou order)</span></div>
            <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 10, lineHeight: 1.5 }}>
              Failure Report em aberto (status diferente de Completed) que ainda não tem <b>action</b> (o que fazer), <b>job</b> (no TM Master, aberto ou já executado) ou <b>order</b> (material/serviço). Quanto mais lacunas e mais dias aberto, mais urgente definir a tratativa.
            </div>
            <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(5, 1fr)" }}>
              {bigKpi("Failure Reports em aberto", gapKpi.abertos, "var(--teal)", ClipboardList)}
              {bigKpi("Sem Action", gapKpi.semAct, gapKpi.semAct ? "var(--warn)" : "var(--ok)", AlertTriangle)}
              {bigKpi("Sem Job no TM", gapKpi.semJob, gapKpi.semJob ? "var(--warn)" : "var(--ok)", Wrench)}
              {bigKpi("Sem Order", gapKpi.semOrd, gapKpi.semOrd ? "var(--warn)" : "var(--ok)", Package)}
              {bigKpi("Sem nada (sem action, job e order)", gapKpi.nada, gapKpi.nada ? "var(--crit)" : "var(--ok)", AlertTriangle)}
            </div>
            <div className="g-mode-toggle" style={{ marginBottom: 10, width: "fit-content" }}>
              {["Com lacuna", "Sem action", "Sem job", "Sem order", "Sem nada"].map((t) => <button key={t} className={gapTipo === t ? "active" : ""} onClick={() => setGapTipo(t)}>{t} ({t === "Com lacuna" ? gaps.filter((g) => g.n > 0).length : t === "Sem action" ? gapKpi.semAct : t === "Sem job" ? gapKpi.semJob : t === "Sem order" ? gapKpi.semOrd : gapKpi.nada})</button>)}
            </div>
            <div className="g-table-wrap" style={{ maxHeight: 420, overflowY: "auto" }}>
              <table className="g-table tbl-modern">
                <thead><tr><th>Failure Report</th><th>Nome</th><th>Data da falha</th><th style={{ textAlign: "right" }}>Dias aberto</th><th>Status</th><th>Dept.</th><th style={{ textAlign: "center" }}>Action</th><th style={{ textAlign: "center" }}>Job</th><th style={{ textAlign: "center" }}>Order</th><th>O que falta</th></tr></thead>
                <tbody>
                  {gapsVis.length === 0 && <tr><td colSpan={10} className="g-muted" style={{ textAlign: "center", padding: 16 }}>Nenhum Failure Report em aberto nessa condição.</td></tr>}
                  {gapsVis.map((g) => (
                    <tr className="g-row tbl-row" key={g.r.no} onClick={() => abrirFr(g.r.no)} style={{ cursor: "pointer", "--c": g.n === 3 ? "#EF4444" : g.n > 0 ? "#F5A623" : "#22C55E" }}>
                      <td style={{ fontFamily: "var(--mono)", whiteSpace: "nowrap" }}><b>{g.r.no}</b></td>
                      <td style={{ whiteSpace: "normal", minWidth: 220 }}>{g.r.name || "—"}</td>
                      <td style={{ fontFamily: "var(--mono)" }}>{g.r.failureDate ? fmtDataCompleta(g.r.failureDate) : "—"}</td>
                      <td style={{ textAlign: "right", fontWeight: 700 }}>{g.dias === null ? "—" : `${g.dias} d`}</td>
                      <td>{g.r.status}</td><td>{g.r.department || "—"}</td>
                      <td style={{ textAlign: "center", color: g.noAct ? "var(--crit)" : "var(--ok)", fontWeight: 700 }}>{g.noAct ? "✖ não" : "✔"}</td>
                      <td style={{ textAlign: "center", color: g.noJob ? "var(--crit)" : "var(--ok)", fontWeight: 700 }}>{g.noJob ? "✖ não" : "✔"}</td>
                      <td style={{ textAlign: "center", color: g.noOrd ? "var(--crit)" : "var(--ok)", fontWeight: 700 }}>{g.noOrd ? "✖ não" : "✔"}</td>
                      <td style={{ whiteSpace: "normal" }}>{g.n === 0 ? "Completo" : g.n === 3 ? "Sem nenhuma tratativa — definir action, abrir job no TM e avaliar order" : `Sem ${g.falta.join(" e ")}`}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="g-panel">
            <div className="g-panel-head">
              <span className="g-panel-title">Failure Reports ({filtered.length}) — clique numa linha para abrir tudo junto: falha, linha do tempo, actions, jobs, histórico e orders</span>
              <div style={{ flex: 1 }} />
              <button className="g-btn" onClick={() => setOpen(new Set(filtered.map((r) => r.no)))}>Expandir todos</button>
              <button className="g-btn" onClick={() => setOpen(new Set())}>Recolher todos</button>
            </div>
            <div className="g-table-wrap">
              <table className="g-table tbl-modern">
                <thead><tr>
                  <th>Nº</th><th>Nome</th><th>Data</th><th>Componente</th><th>Dept.</th><th>Criticidade</th><th>Status</th>
                  <th style={{ textAlign: "right" }}>Actions</th><th style={{ textAlign: "right" }}>Jobs</th><th style={{ textAlign: "right" }}>Histórico</th><th style={{ textAlign: "right" }}>Orders</th>
                </tr></thead>
                <tbody>
                  {filtered.length === 0 && <tr><td colSpan={11} className="g-muted" style={{ textAlign: "center", padding: 18 }}>Nenhum Failure Report com esse filtro.</td></tr>}
                  {filtered.map((r) => { const l = L(r.no); const ab = l.actions.filter((a) => !frActionConcluida(a)).length; const oa = l.orders.filter((o) => !frOrderEncerrada(o)).length;
                    return (
                      <React.Fragment key={r.no}>
                      <tr id={`fr-row-${r.no}`} className="g-row tbl-row" onClick={() => toggleOpen(r.no)} style={{ cursor: "pointer", "--c": /complet|closed/i.test(r.status) ? "#22C55E" : "#F5A623", background: open.has(r.no) ? "var(--accent-soft, rgba(59,130,246,0.08))" : undefined }}>
                        <td style={{ fontFamily: "var(--mono)", whiteSpace: "nowrap" }}><span style={{ marginRight: 6 }}>{open.has(r.no) ? "▾" : "▸"}</span><b>{r.no}</b></td>
                        <td style={{ whiteSpace: "normal", minWidth: 200 }}>{r.name || "—"}</td>
                        <td style={{ fontFamily: "var(--mono)", whiteSpace: "nowrap" }}>{r.failureDate ? fmtDataCompleta(r.failureDate) : "—"}</td>
                        <td style={{ whiteSpace: "normal" }}>{r.compCode} {r.compName}</td>
                        <td>{r.department || "—"}</td><td>{r.criticality || "—"}</td>
                        <td><span className={`tbl-days ${stCls(r.status)}`}>{r.status || "—"}</span></td>
                        <td style={{ textAlign: "right" }}>{ab > 0 ? <b style={{ color: "var(--warn)" }}>{ab}/{l.actions.length}</b> : `${l.actions.length}`}</td>
                        <td style={{ textAlign: "right" }}>{l.jobs.length}</td><td style={{ textAlign: "right" }}>{l.jobHistory.length}</td>
                        <td style={{ textAlign: "right" }}>{oa > 0 ? <b style={{ color: "var(--warn)" }}>{oa}/{l.orders.length}</b> : `${l.orders.length}`}</td>
                      </tr>
                      {open.has(r.no) && <tr><td colSpan={11} style={{ padding: 0 }}><Detalhe r={r} /></td></tr>}
                      </React.Fragment>
                    ); })}
                </tbody>
              </table>
            </div>
            <div className="g-muted" style={{ fontSize: 10.5, marginTop: 6 }}>Actions: abertas/total (só mostra "abertas/" quando há pendência). Jobs: abertos no Due. Orders: em aberto/total.</div>
          </div>

          <div className="g-panel" style={{ borderTop: "3px solid var(--warn)" }}>
            <div className="g-panel-head"><span className="g-panel-title">Resumo de Pendências — o que está parado há mais tempo</span></div>
            {pend.length === 0 ? (
              <div className="g-muted" style={{ padding: 12 }}>Nenhuma pendência nos Failure Reports do filtro atual.</div>
            ) : (
              <>
                <div style={{ fontSize: 13, lineHeight: 1.6, marginBottom: 12, padding: "10px 14px", background: "var(--panel-raised)", borderRadius: 8, border: "1px solid var(--border-soft)" }}>
                  Há <b>{pendResumo.acts.length}</b> action(s) pendente(s)
                  {pendResumo.acts.length > 0 && <> (<b>{pendResumo.acts.filter((p) => p.atraso > 0).length}</b> com prazo vencido)</>}, <b>{pendResumo.porTipo[1].total}</b> job(s) em aberto e <b>{pendResumo.porTipo[2].total}</b> order(s) em andamento,
                  ligados a <b>{pendResumo.frsComPend}</b> Failure Report(s). {pendResumo.acts[0] && <>A action mais antiga é <b>{pendResumo.acts[0].ref}</b> — “{pendResumo.acts[0].desc}” ({pendResumo.acts[0].frNo}), aberta há <b>{pendResumo.acts[0].dias} dias</b>{pendResumo.acts[0].atraso > 0 && <> e vencida há <b>{pendResumo.acts[0].atraso} dias</b></>}.</>}
                  {" "}{pendResumo.maisAntiga && <>O item pendente mais antigo de todos é {pendResumo.maisAntiga.tipo === "Failure Report" ? "o Failure Report" : `${pendResumo.maisAntiga.tipo.toLowerCase()}`} <b>{pendResumo.maisAntiga.ref}</b>, parado há <b>{pendResumo.maisAntiga.dias} dias</b>.</>}
                </div>
                <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
                  {pendResumo.porTipo.slice(0, 3).map((t) => bigKpi(`${t.tipo}s em aberto`, `${t.total}${t.vencidos ? ` · ${t.vencidos} venc.` : ""}`, t.vencidos ? "var(--crit)" : "var(--teal)", t.tipo === "Action" ? ClipboardList : t.tipo === "Job" ? Wrench : Package))}
                  {bigKpi("Mais antigo (dias)", pendResumo.maisAntiga?.dias ?? 0, "var(--warn)", Clock)}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(340px, 1fr))", gap: 14, marginBottom: 12 }}>
                  <MesBarChart title="Pendências por tempo parado" height={200} data={pendResumo.faixas} bars={[{ key: "total", name: "Pendências", color: "#F5A623" }]} />
                  <div>
                    <div className="g-muted" style={{ fontSize: 11.5, margin: "4px 0 6px", fontWeight: 700 }}>Actions pendentes por responsável</div>
                    <Tbl cols={["Responsável", "Pendentes", "Vencidas", "Mais antiga (dias)"]} empty="Sem actions pendentes."
                      rows={pendResumo.porResp.map((r) => [r.resp, r.total, r.vencidos, r.maisAntigo])} />
                  </div>
                </div>
                <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 8 }}>
                  <div className="g-mode-toggle">
                    {["Todos", "Action", "Job", "Order", "Failure Report"].map((t) => <button key={t} className={pendTipo === t ? "active" : ""} onClick={() => setPendTipo(t)}>{t === "Todos" ? `Todos (${pend.length})` : `${t} (${pend.filter((p) => p.tipo === t).length})`}</button>)}
                  </div>
                  <span className="g-muted" style={{ fontSize: 11 }}>Ordenado do mais antigo para o mais recente. “Dias parado” = hoje − data de criação (Action/Job), aprovação (Order) ou da falha (Failure Report).</span>
                </div>
                <div className="g-table-wrap" style={{ maxHeight: 460, overflowY: "auto" }}>
                  <table className="g-table tbl-modern">
                    <thead><tr><th>Tipo</th><th>Failure Report</th><th>Referência</th><th>Descrição</th><th>Responsável</th><th>Aberto desde</th><th style={{ textAlign: "right" }}>Dias parado</th><th>Prazo</th><th style={{ textAlign: "right" }}>Atraso</th><th>Situação</th></tr></thead>
                    <tbody>
                      {pendVis.slice(0, 300).map((x, i) => (
                        <tr className="g-row tbl-row" key={i} onClick={() => { setOpen((p) => new Set(p).add(x.frNo)); setTimeout(() => document.getElementById(`fr-row-${x.frNo}`)?.scrollIntoView({ behavior: "smooth", block: "start" }), 80); }} style={{ cursor: "pointer", "--c": x.atraso > 0 ? "#EF4444" : (x.dias || 0) > 30 ? "#F5A623" : "#3B82F6" }}>
                          <td><b>{x.tipo}</b></td>
                          <td style={{ fontFamily: "var(--mono)", whiteSpace: "nowrap" }}>{x.frNo}</td>
                          <td style={{ fontFamily: "var(--mono)", whiteSpace: "nowrap" }}>{x.ref}</td>
                          <td style={{ whiteSpace: "normal", minWidth: 220 }}>{x.desc || "—"}</td>
                          <td>{x.resp}</td>
                          <td style={{ fontFamily: "var(--mono)" }}>{x.desde ? fmtDataCompleta(x.desde) : "—"}</td>
                          <td style={{ textAlign: "right", fontWeight: 700 }}>{diasTxt(x.dias)}</td>
                          <td style={{ fontFamily: "var(--mono)" }}>{x.prazo ? fmtDataCompleta(x.prazo) : "—"}</td>
                          <td style={{ textAlign: "right" }}>{x.atraso > 0 ? <span className="tbl-days d3">{x.atraso} d</span> : "—"}</td>
                          <td>{x.situacao}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        </>
      )}
    </>
  );
};

/* ============================================================
   TM MASTER — RESUMO EM TABELAS PARA E-MAIL (cada tópico explicado, com coluna Observação editável).
   O e-mail comum (mailto) não aceita tabelas: por isso o resumo é COPIADO já formatado e colado no e-mail.
   ============================================================ */
const escHtml = (v) => String(v === null || v === undefined ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

const ResumoEmailTm = ({ tmDue, tmHistory, tmFailures }) => {
  const [aberto, setAberto] = useState(false);
  const [para, setPara] = useState(() => { try { return localStorage.getItem("genesis_resumo_para") || ""; } catch (e) { return ""; } });
  const [limite, setLimite] = useState(0);
  const [deptSel, setDeptSel] = useState("Todos");
  const [obs, setObs] = useState({});
  const [msg, setMsg] = useState("");
  const hoje = todayISO();
  const fmt = (iso) => (iso ? fmtDataCompleta(iso) : "—");

  const secoes = useMemo(() => {
    if (!aberto) return [];
    const F = tmFailures || FR_EMPTY;
    const lim = limite === 0 ? 1e9 : limite;
    const isVenc = (d) => (d.diffValue !== null && d.diffValue !== undefined ? d.diffValue < 0 : !!d.dueDate && d.dueDate < hoje);
    const dias = (d) => (d.diffValue !== null && d.diffValue !== undefined && d.diffUnit === "D" ? d.diffValue : (d.dueDate ? -frDias(d.dueDate, hoje) : null)); // negativo = atrasado
    const ate40 = (d) => !isVenc(d) && ((d.diffUnit === "D" && d.diffValue !== null && d.diffValue >= 0 && d.diffValue <= 40) || (d.diffValue === null && d.dueDate && dias(d) >= 0 && dias(d) <= 40));
    const jobTxt = (d) => `${d.jobType || ""} ${d.jobNo || ""}`.trim() || "—";
    const pct = (a, b) => (b ? `${Math.round((a / b) * 100)}%` : "0%");
    const venc = tmDue.filter(isVenc).sort((a, b) => (dias(a) ?? 0) - (dias(b) ?? 0));
    const a40 = tmDue.filter(ate40).sort((a, b) => (dias(a) ?? 0) - (dias(b) ?? 0));
    const crit = tmDue.filter((d) => d.pri === "High" && (isVenc(d) || ate40(d)));
    const one = tmDue.filter((d) => String(d.jobType || "").toUpperCase() === "ONE");
    const critJobs = tmDue.filter((d) => critPrefixOf(d.code));
    const obsJob = (d, vencido) => {
      const p = [];
      if (vencido) p.push(`Vencido há ${Math.abs(dias(d) ?? 0)} dias`); else p.push(`Vence em ${dias(d)} dias`);
      if (d.pri === "High") p.push("prioridade alta");
      if (critPrefixOf(d.code)) p.push("componente crítico (PMS)");
      if (String(d.jobType || "").toUpperCase() === "ONE") p.push("corretiva");
      return p.join(" · ");
    };
    const frMapE = new Map(); (F.jobs || []).concat(F.jobHistory || []).forEach((x) => { if (!x.frNo || !x.jobNo) return; const k = `${bordoNorm(x.jobType)}|${bordoNorm(x.jobNo)}`; if (!frMapE.has(k)) frMapE.set(k, new Set()); frMapE.get(k).add(x.frNo); });
    const frDe = (t, n) => [...(frMapE.get(`${bordoNorm(t)}|${bordoNorm(n)}`) || [])].join(", ");
    const S = [];
    const vencDept = venc.filter((d) => deptSel === "Todos" || (d.department || "—") === deptSel);
    const pedido = (d) => {
      const extra = [`vencido há ${Math.abs(dias(d) ?? 0)} dias`];
      if (d.pri === "High") extra.push("prioridade alta");
      if (critPrefixOf(d.code)) extra.push("componente crítico (PMS)");
      if (String(d.jobType || "").toUpperCase() === "ONE") extra.push("corretiva");
      return `BORDO RESPONDER: (1) foi executado? informe a data e dê baixa no TM; (2) se não, o motivo; (3) nova data prevista; (4) informe o nº da PO deste serviço. [${extra.join(" · ")}]`;
    };
    S.push({ id: "geral", titulo: "Situação geral do Due (planos vencidos)", explicacao: "Fotografia dos planos de manutenção vencidos, ou seja, com prazo (Due/Diff) já passado. Os percentuais usam o total de planos em aberto como base.",
      cols: ["Indicador", "Quantidade"], rows: [
        { key: "t", cells: ["Total de planos em aberto (base)", tmDue.length], auto: "Base de todos os percentuais abaixo" },
        { key: "v", cells: ["Vencidos", venc.length], auto: `${pct(venc.length, tmDue.length)} do total — exigem ação imediata` },
        { key: "c", cells: ["Críticos (prioridade High) vencidos", venc.filter((d) => d.pri === "High").length], auto: "Tratar primeiro" },
        { key: "o", cells: ["Corretivas (Job Type ONE) vencidas", venc.filter((d) => String(d.jobType || "").toUpperCase() === "ONE").length], auto: `de ${one.length} corretivas em aberto` },
        { key: "p", cells: ["Jobs críticos PMS vencidos / em aberto", `${critJobs.filter(isVenc).length} / ${critJobs.length}`], auto: "Componentes críticos para operação, segurança ou meio ambiente" },
      ] });
    S.push({ id: "venc", titulo: `Jobs vencidos que precisam de resposta de bordo (${Math.min(lim, vencDept.length)} de ${vencDept.length}${deptSel !== "Todos" ? ` — ${deptSel}` : ""})`,
      explicacao: "Planos de manutenção cujo prazo já passou, do mais atrasado para o menos atrasado. Para CADA job, bordo deve informar a PO do serviço (coluna “PO”) e responder na coluna “Resposta de bordo”: foi executado (data / baixa feita no TM)? Se não, qual o motivo do atraso e qual a nova data prevista. A coluna “Failure Report” mostra se o job está ligado a alguma falha registrada.",
      cols: ["Departamento", "Job", "Plano", "Componente", "Due", "Atraso (dias)", "Prioridade", "Failure Report"], blankCols: ["PO (nº da PO deste serviço)", "Resposta de bordo"],
      rows: vencDept.slice(0, lim).map((d, i) => ({ key: `${d.id || i}`, cells: [d.department || "—", jobTxt(d), d.jobName || "—", `${d.code || ""} ${d.component || ""}`.trim() || "—", d.dueDate ? fmt(d.dueDate) : (d.dueRaw || "—"), Math.abs(dias(d) ?? 0), d.pri || "—", frDe(d.jobType, d.jobNo) || "—"], auto: pedido(d) })) });
    const dept = {}; tmDue.forEach((d) => { const k = d.department || "—"; dept[k] = dept[k] || { t: 0, v: 0 }; dept[k].t++; if (isVenc(d)) dept[k].v++; });
    S.push({ id: "dept", titulo: "Vencidos por departamento", explicacao: "Mostra onde está concentrado o atraso. O percentual é vencidos ÷ total de planos em aberto do departamento.",
      cols: ["Departamento", "Em aberto", "Vencidos", "% vencido"],
      rows: Object.entries(dept).sort((a, b) => b[1].v - a[1].v).map(([k, v]) => ({ key: k, cells: [k, v.t, v.v, pct(v.v, v.t)], auto: v.v === 0 ? "Sem atraso" : (v.v / v.t > 0.3 ? "Atraso elevado — priorizar" : "Atraso moderado") })) });
    const grp = CRITICAL_GROUPS.map((g) => { const l = tmDue.filter((d) => critPrefixOf(d.code) === g.cod); return { g, t: l.length, v: l.filter(isVenc).length }; }).filter((x) => x.t > 0).sort((a, b) => b.v - a.v);
    S.push({ id: "crit", titulo: "Jobs críticos por grupo de sistema (PMS)", explicacao: "Grupos de componentes críticos (3 primeiros dígitos do código). Mostra só os grupos que têm planos em aberto.",
      cols: ["Grupo", "Criticidade", "Em aberto", "Vencidos"],
      rows: grp.map((x) => ({ key: x.g.cod, cells: [critLabel(x.g.cod), x.g.crit, x.t, x.v], auto: x.v === 0 ? "Em dia" : `${x.v} vencido(s) — ${pct(x.v, x.t)} do grupo` })) });
    const mesAtual = hoje.slice(0, 7);
    const isPost = (h) => (h.remarks || "").toLowerCase().includes("postpon") || String(h.postponed || "").trim().toUpperCase() === "X";
    const postMes = tmHistory.filter((h) => isPost(h) && (h.dateDone || "").slice(0, 7) === mesAtual);
    const execMes = tmHistory.filter((h) => !isPost(h) && (h.dateDone || "").slice(0, 7) === mesAtual);
    S.push({ id: "hist", titulo: "History — mês atual", explicacao: "Jobs fechados neste mês. “Postergado” = job adiado (campo Postponed ou observação com “Postponed”).",
      cols: ["Indicador", "Quantidade"], rows: [
        { key: "e", cells: ["Executados no mês", execMes.length], auto: "Planos concluídos" },
        { key: "p", cells: ["Postergados no mês", postMes.length], auto: postMes.length ? `${pct(postMes.length, postMes.length + execMes.length)} dos jobs do mês foram adiados` : "Nenhum adiamento" },
      ] });
    if (postMes.length) S.push({ id: "post", titulo: `Jobs postergados no mês (${Math.min(lim, postMes.length)} de ${postMes.length})`, explicacao: "Jobs adiados neste mês. A observação vem do campo Remarks do History — edite se quiser.",
      cols: ["Job", "Plano", "Componente", "Data", "Feito por"],
      rows: postMes.slice(0, lim).map((h, i) => ({ key: `${h.jobHistoryNumber || i}`, cells: [`${h.jobType || ""} ${h.jobNo || ""}`.trim() || "—", h.jobName || "—", `${h.componentCode || ""} ${h.componentName || ""}`.trim() || "—", fmt(h.dateDone), h.doneByName || "—"], auto: h.remarks || "Postergado" })) });
    if (F.reports.length) {
      const nos = new Set(F.reports.map((r) => r.no));
      const nm = (no) => F.reports.find((r) => r.no === no)?.name || "";
      const abertos = F.reports.filter((r) => !/complet|closed/i.test(r.status));
      const acts = F.actions.filter((a) => nos.has(a.frNo) && !frActionConcluida(a)).sort((a, b) => (frDias(b.created || b.deadline, hoje) ?? -1) - (frDias(a.created || a.deadline, hoje) ?? -1));
      const actsV = acts.filter((a) => a.deadline && a.deadline < hoje);
      const jobsF = F.jobs.filter((j) => nos.has(j.frNo));
      const ordF = F.orders.filter((o) => nos.has(o.frNo) && !frOrderEncerrada(o));
      const gp = frGaps(F, hoje).filter((g) => g.n > 0).sort((a, b) => b.n - a.n || (b.dias ?? -1) - (a.dias ?? -1));
      S.push({ id: "fr", titulo: "Failure Reports — situação geral", explicacao: "Falhas registradas e tudo que está ligado a elas (actions, jobs e orders). “Em aberto” = status diferente de Completed.",
        cols: ["Indicador", "Quantidade"], rows: [
          { key: "a", cells: ["Failure Reports em aberto / total", `${abertos.length} / ${F.reports.length}`], auto: "Falhas ainda sem encerramento" },
          { key: "b", cells: ["Actions pendentes", acts.length], auto: `${actsV.length} com prazo vencido` },
          { key: "c", cells: ["Jobs ligados em aberto", jobsF.length], auto: `${jobsF.filter((j) => j.nextDue && j.nextDue < hoje).length} vencido(s)` },
          { key: "d", cells: ["Orders em andamento", ordF.length], auto: "Pedidos de peças/serviços ligados às falhas" },
          { key: "e", cells: ["Failure Reports em aberto sem action, job ou order", gp.length], auto: `${gp.filter((g) => g.n === 3).length} sem nenhuma tratativa` },
        ] });
      S.push({ id: "gap", titulo: `Failure Reports em aberto sem tratativa (${Math.min(lim, gp.length)} de ${gp.length})`, explicacao: "Falhas em aberto que ainda não têm action, job no TM ou order. Ação: definir o plano de ação, abrir o job no TM (ONE, se corretiva) e avaliar a necessidade de material.",
        cols: ["Failure Report", "Nome", "Data da falha", "Dias aberto", "Tem action", "Tem job", "Tem order"],
        rows: gp.slice(0, lim).map((g) => ({ key: g.r.no, cells: [g.r.no, g.r.name, fmt(g.r.failureDate), g.dias ?? "—", g.noAct ? "NÃO" : "sim", g.noJob ? "NÃO" : "sim", g.noOrd ? "NÃO" : "sim"], auto: g.n === 3 ? "Sem nenhuma tratativa — definir action, abrir job e avaliar order" : `Falta: ${g.falta.join(" e ")}` })) });
      S.push({ id: "acts", titulo: `Actions pendentes (${Math.min(lim, acts.length)} de ${acts.length}, as paradas há mais tempo)`, explicacao: "Ações corretivas/preventivas abertas, da mais antiga para a mais recente. “Parada há” conta desde a criação da action.",
        cols: ["Action", "Ação", "Failure Report", "Responsável", "Prazo", "Parada há (dias)"],
        rows: acts.slice(0, lim).map((a) => { const at = a.deadline && a.deadline < hoje ? frDias(a.deadline, hoje) : 0; return { key: a.no, cells: [a.no, a.name, `${a.frNo} — ${nm(a.frNo)}`, a.crew || a.crewType || a.department || "—", fmt(a.deadline), frDias(a.created || a.deadline, hoje) ?? "—"], auto: at > 0 ? `Prazo vencido há ${at} dias${a.priority ? ` · prioridade ${a.priority}` : ""}` : `No prazo${a.priority ? ` · prioridade ${a.priority}` : ""}` }; }) });
      if (jobsF.length) S.push({ id: "jf", titulo: "Jobs em aberto ligados a Failure Reports", explicacao: "Planos criados para corrigir as falhas e ainda não executados.",
        cols: ["Job", "Plano", "Failure Report", "Departamento", "Next Due"],
        rows: jobsF.map((j, i) => ({ key: `${j.jobType}${j.jobNo}${i}`, cells: [`${j.jobType} ${j.jobNo}`, j.name, `${j.frNo} — ${nm(j.frNo)}`, j.department || "—", fmt(j.nextDue)], auto: j.nextDue && j.nextDue < hoje ? `Vencido há ${frDias(j.nextDue, hoje)} dias` : (j.nextDue ? `Vence em ${-frDias(j.nextDue, hoje)} dias` : "Sem data") })) });
      if (ordF.length) S.push({ id: "of", titulo: "Orders em andamento ligadas a Failure Reports", explicacao: "Pedidos de material/serviço que ainda não foram recebidos.",
        cols: ["Order", "Pedido", "Failure Report", "Status", "Fornecedor", "Entrega prevista"],
        rows: ordF.map((o) => ({ key: o.no, cells: [o.no, o.name, `${o.frNo} — ${nm(o.frNo)}`, o.status, o.supplier || "—", fmt(o.delivery)], auto: o.delivery && o.delivery < hoje ? `Entrega atrasada há ${frDias(o.delivery, hoje)} dias` : "Aguardando entrega" })) });
    }
    S.forEach((x, i) => { x.titulo = `${i + 1}. ${x.titulo}`; });
    return S;
  }, [aberto, limite, deptSel, tmDue, tmHistory, tmFailures]);

  const deptsVenc = useMemo(() => ([...new Set(tmDue.filter((d) => (d.diffValue !== null && d.diffValue !== undefined ? d.diffValue < 0 : !!d.dueDate && d.dueDate < hoje)).map((d) => d.department || "—"))].sort()), [tmDue]);
  const obsDe = (sid, r) => (obs[`${sid}|${r.key}`] !== undefined ? obs[`${sid}|${r.key}`] : r.auto);
  const assunto = `Genesis I — Resumo TM Master — ${fmt(hoje)}`;

  const montarHtml = () => {
    const th = "background:#1F3A5F;color:#fff;padding:6px 8px;border:1px solid #cbd5e1;text-align:left;font-size:12px";
    const td = "padding:5px 8px;border:1px solid #cbd5e1;font-size:12px;vertical-align:top";
    let h = `<div style="font-family:Arial,Helvetica,sans-serif;color:#1f2937"><h2 style="margin:0 0 4px">${escHtml(assunto)}</h2><p style="margin:0 0 6px;font-size:13px">Prezados, segue o resumo do TM Master. Solicitamos que <b>cada job vencido</b> da tabela “Jobs vencidos” seja respondido por bordo. Para cada job, informe o <b>nº da PO</b> do serviço e a <b>resposta de bordo</b> (respondendo este e-mail ou preenchendo as colunas em amarelo), e dê baixa no TM Master dos jobs já executados.</p><p style="margin:0 0 16px;color:#6b7280;font-size:12px">Situação do TM Master em ${fmt(hoje)}.</p>`;
    secoes.forEach((s) => {
      h += `<h3 style="margin:18px 0 4px;font-size:14px">${escHtml(s.titulo)}</h3><p style="margin:0 0 6px;color:#6b7280;font-size:12px">${escHtml(s.explicacao)}</p>`;
      h += `<table style="border-collapse:collapse;width:100%"><thead><tr>${[...s.cols, "Observação", ...(s.blankCols || [])].map((c) => `<th style="${th}">${escHtml(c)}</th>`).join("")}</tr></thead><tbody>`;
      if (!s.rows.length) h += `<tr><td colspan="${s.cols.length + 1 + (s.blankCols || []).length}" style="${td};color:#6b7280">Nenhum item.</td></tr>`;
      s.rows.forEach((r, i) => { h += `<tr style="background:${i % 2 ? "#f8fafc" : "#fff"}">${r.cells.map((c) => `<td style="${td}">${escHtml(c)}</td>`).join("")}<td style="${td}">${escHtml(obsDe(s.id, r))}</td>${(s.blankCols || []).map(() => `<td style="${td};min-width:170px;background:#fffbe6">&nbsp;</td>`).join("")}</tr>`; });
      h += "</tbody></table>";
    });
    return h + `<p style="margin-top:18px;color:#6b7280;font-size:11px">Enviado a partir do portal Genesis I.</p></div>`;
  };
  const montarTexto = () => {
    const L = [assunto, "", "Prezados, solicitamos que cada job vencido abaixo seja respondido por bordo (execução/baixa no TM, motivo do atraso e nova data prevista).", ""];
    secoes.forEach((s) => {
      L.push(s.titulo, s.explicacao, [...s.cols, "Observação", ...(s.blankCols || [])].join(" | "));
      if (!s.rows.length) L.push("(nenhum item)");
      s.rows.forEach((r) => L.push([...r.cells, obsDe(s.id, r), ...(s.blankCols || []).map(() => "______")].join(" | ")));
      L.push("");
    });
    return L.join("\n");
  };
  const copiar = async () => {
    const html = montarHtml(); const txt = montarTexto();
    try {
      if (window.ClipboardItem && navigator.clipboard.write) await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([txt], { type: "text/plain" }) })]);
      else await navigator.clipboard.writeText(txt);
      setMsg("Resumo copiado com as tabelas. Cole no corpo do e-mail (Ctrl+V).");
    } catch (e) { setMsg("Não consegui copiar automaticamente — selecione a área abaixo e copie com Ctrl+C."); }
    setTimeout(() => setMsg(""), 7000);
    return true;
  };
  const abrirEmail = async () => {
    try { localStorage.setItem("genesis_resumo_para", para); } catch (e) { /* ignora */ }
    await copiar();
    const corpo = "Olá,\n\nSegue o resumo semanal do TM Master.\n\n[Cole aqui o resumo com Ctrl+V — as tabelas já estão copiadas]\n";
    window.location.href = `mailto:${para.split(/[;,]/).map((x) => x.trim()).filter(Boolean).join(",")}?subject=${encodeURIComponent(assunto)}&body=${encodeURIComponent(corpo)}`;
  };

  const inp = { width: "100%", minWidth: 180, padding: "4px 6px", border: "1px solid var(--border)", borderRadius: 4, fontSize: 12, background: "var(--panel, #fff)", color: "var(--text)" };
  return (
    <div style={{ marginBottom: 14 }}>
      <button className="g-btn" onClick={() => setAberto((v) => !v)}><FileText size={14} />{aberto ? "Fechar resumo por e-mail" : "Enviar resumo por e-mail"}</button>
      {aberto && (
        <div className="g-panel" style={{ marginTop: 10 }}>
          <div className="g-panel-head"><span className="g-panel-title">Resumo do TM Master para e-mail — em tabelas</span></div>
          <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 10, lineHeight: 1.5 }}>
            O resumo é completo; na parte do <b>Due</b> entram <b>somente os jobs vencidos</b>. Na tabela de jobs vencidos, a coluna <b>Observação</b> pede a resposta de bordo (executado? motivo? nova data?) e você pode editá-la; as colunas <b>PO</b> e <b>Resposta de bordo</b> vão em branco para eles preencherem. O filtro de departamento vale para a tabela de jobs vencidos.
            Mostra a situação de hoje, com todos os dados (sem os filtros da tela). Como o e-mail comum não recebe tabelas por link, o botão <b>Abrir no e-mail</b> copia o resumo formatado e abre o e-mail com destinatários e assunto — é só colar (Ctrl+V).
          </div>
          <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 10 }}>
            <input type="text" value={para} onChange={(e) => setPara(e.target.value)} placeholder="Para: email1@empresa.com, email2@empresa.com" style={{ flex: 1, minWidth: 280, padding: "7px 10px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 13 }} />
            <label style={{ fontSize: 12, display: "flex", alignItems: "center", gap: 6 }}>Departamento
              <select value={deptSel} onChange={(e) => setDeptSel(e.target.value)} style={{ padding: "5px 8px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 12.5 }}>
                <option>Todos</option>{deptsVenc.map((d) => <option key={d}>{d}</option>)}
              </select></label>
            <label style={{ fontSize: 12, display: "flex", alignItems: "center", gap: 6 }}>Linhas por tabela
              <select value={limite} onChange={(e) => setLimite(Number(e.target.value))} style={{ padding: "5px 8px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 12.5 }}>
                <option value={0}>Todas</option><option value={10}>10</option><option value={20}>20</option><option value={50}>50</option>
              </select></label>
            <button className="g-btn primary g-keep" onClick={abrirEmail}>Abrir no e-mail</button>
            <button className="g-btn" onClick={copiar}>Copiar resumo (com tabelas)</button>
          </div>
          {msg && <div className="g-alert" style={{ marginBottom: 10, background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>{msg}</div>}
          <div className="g-muted" style={{ fontSize: 11.5, marginBottom: 6 }}>Assunto: <b style={{ color: "var(--text)" }}>{assunto}</b></div>
          <div style={{ maxHeight: 620, overflow: "auto", border: "1px solid var(--border)", borderRadius: 6, padding: 12 }}>
            {secoes.map((s) => (
              <div key={s.id} style={{ marginBottom: 18 }}>
                <div style={{ fontSize: 13.5, fontWeight: 700 }}>{s.titulo}</div>
                <div className="g-muted" style={{ fontSize: 11.5, margin: "2px 0 6px" }}>{s.explicacao}</div>
                <div className="g-table-wrap">
                  <table className="g-table tbl-modern">
                    <thead><tr>{[...s.cols, "Observação", ...(s.blankCols || [])].map((c) => <th key={c}>{c}</th>)}</tr></thead>
                    <tbody>
                      {s.rows.length === 0 && <tr><td colSpan={s.cols.length + 1 + (s.blankCols || []).length} className="g-muted" style={{ textAlign: "center", padding: 12 }}>Nenhum item.</td></tr>}
                      {s.rows.map((r) => (
                        <tr className="g-row tbl-row" key={r.key}>
                          {r.cells.map((c, j) => <td key={j} style={{ whiteSpace: "normal" }}>{c === "" || c === null || c === undefined ? "—" : c}</td>)}
                          <td style={{ minWidth: 300 }}><input value={obsDe(s.id, r)} onChange={(e) => setObs((p) => ({ ...p, [`${s.id}|${r.key}`]: e.target.value }))} style={inp} /></td>
                          {(s.blankCols || []).map((bc) => <td key={bc} style={{ minWidth: 150, background: "rgba(245,166,35,0.10)" }} className="g-muted">(bordo preenche)</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

/* ============================================================
   TM MASTER — REGISTRO DE BORDO
   Planilha diária do que o bordo fez (preventiva, preditiva, corretiva). Cada linha diz se houve BAIXA de job ou CRIAÇÃO de job
   no TM Master, e o portal confere sozinho contra o History (baixas) e o Due (jobs abertos) importados.
   ============================================================ */
const BORDO_TIPOS = ["Preventiva", "Preditiva", "Corretiva"];
const BORDO_ACOES = ["Baixa de job", "Criação de job", "Sem job"];
const BORDO_ORIGENS = ["Planilha diária", "E-mail", "Mensagem", "Relatório", "Outro"];
const bordoDate = (v) => {
  if (!v) return "";
  if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;
  const s = String(v).trim(); let m = s.match(/^(\d{2})[./-](\d{2})[./-](\d{4})/); if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[1]}-${m[2]}-${m[3]}` : "";
};
/* data -> AAAA-MM-DD, aceitando ISO com hora, dd/mm/aaaa, dd.mm.aaaa e objetos Date (a conferência não pode falhar por formato) */
const isoDia = (v) => {
  if (!v) return "";
  if (v instanceof Date) return isNaN(v) ? "" : `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;
  const t = String(v).trim(); let m;
  if ((m = t.match(/^(\d{4})-(\d{2})-(\d{2})/))) return `${m[1]}-${m[2]}-${m[3]}`;
  if ((m = t.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})/))) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  return "";
};
const bordoNorm = (v) => String(v === undefined || v === null ? "" : v).trim().toUpperCase();
const bordoPick = (v, list, fallback) => { const t = String(v || "").trim().toLowerCase(); return list.find((x) => x.toLowerCase() === t) || list.find((x) => t && x.toLowerCase().startsWith(t.slice(0, 4))) || fallback; };

/* registro "sem Job" (ação no TM = "Sem job"): exige justificativa e fica EM VALIDAÇÃO até o administrador validar */
/* "dono" do registro: quem lançou (criadoPor); em registros antigos, sem esse dado, vale o nome em "Reportado por" */
const bordoMin = (v) => String(v || "").trim().toLowerCase();
const bordoDono = (r, u) => !!(r && u) && (r.criadoPor ? r.criadoPor === u.username : !!bordoMin(r.por) && (bordoMin(r.por) === bordoMin(u.name) || bordoMin(r.por) === bordoMin(u.username)));
const bordoSemJob = (r) => !!r && !r.semAtv && r.acao === "Sem job";
const BordoRegistroView = ({ tmBordo, setTmBordo, setBordoDeleted, tmDue, tmHistory, tmFailures, regObrig, setRegObrig, currentUser, setReportFn, setExportXlsxFn, active }) => {
  const [sub, setSub] = useState("preench");
  const hoje = todayISO();
  const admin = isAdminUser(currentUser);
  const pode = (d) => podeEditarBordo(currentUser, d);
  /* quem lançou o registro continua podendo editá-lo (além de quem tem o departamento liberado) */
  const podeRow = (r) => pode(r.departamento) || bordoDono(r, currentUser);
  const deptsEditaveis = admin ? BORDO_DEPTS : deptsBordoUsuario(currentUser);
  const podeCriar = deptsEditaveis.length > 0;
  const [msg, setMsg] = useState("");
  const [novo, setNovo] = useState(null);
  const [emailBordo, setEmailBordo] = useState(() => { try { return localStorage.getItem("genesis_email_bordo") || ""; } catch (e) { return ""; } });
  const [f, setF] = useState({ busca: "", tipo: "Todos", status: "Todos", fr: "Todos", dept: "Todos", de: "", ate: "", meus: false });
  const fmt = (iso) => (iso ? fmtDataCompleta(iso) : "—");

  /* índices de conferência contra o TM Master */
  const idx = useMemo(() => {
    const hist = new Map(), due = new Map();
    tmHistory.forEach((h) => { const k = `${bordoNorm(h.jobType)}|${bordoNorm(h.jobNo)}`; if (!hist.has(k)) hist.set(k, []); hist.get(k).push(h); });
    tmDue.forEach((d) => { const k = `${bordoNorm(d.jobType)}|${bordoNorm(d.jobNo)}`; if (!due.has(k)) due.set(k, []); due.get(k).push(d); });
    return { hist, due };
  }, [tmDue, tmHistory]);

  /* Failure Reports: ligação digitada na linha OU detectada automaticamente (jobs / job history ligados ao Failure Report pelo mesmo Job Type + Job No) */
  const frIdx = useMemo(() => {
    const F = tmFailures || FR_EMPTY; const m = new Map();
    F.jobs.concat(F.jobHistory).forEach((x) => { if (!x.frNo || !x.jobNo) return; const k = `${bordoNorm(x.jobType)}|${bordoNorm(x.jobNo)}`; if (!m.has(k)) m.set(k, new Set()); m.get(k).add(x.frNo); });
    return { m, reports: new Map(F.reports.map((r) => [r.no, r])), has: F.reports.length > 0 };
  }, [tmFailures]);
  const frInfo = (r) => {
    const typed = String(r.fr || "").trim();
    const nos = typed ? typed.split(/[,;\s]+/).filter(Boolean) : [...(frIdx.m.get(`${bordoNorm(r.jobTipo)}|${bordoNorm(r.jobNo)}`) || [])];
    const txt = nos.map((n) => { const rep = frIdx.reports.get(n); return rep ? `${n} — ${rep.name} (${rep.status})` : (frIdx.has ? `${n} (não está nos Failure Reports importados)` : n); });
    return { nos, txt, auto: !typed && nos.length > 0 };
  };
  const conferir = (r) => {
    const tem = r.jobTipo && r.jobNo; const k = `${bordoNorm(r.jobTipo)}|${bordoNorm(r.jobNo)}`;
    const dias = (a, b) => Math.round((new Date(isoDia(a) + "T12:00:00") - new Date(isoDia(b) + "T12:00:00")) / 86400000);
    if (r.acao === "Sem job") return { ok: false, cat: "Sem registro no TM", txt: "Atividade sem job no TM — criar o job (ou dar baixa) e informar o nº" };
    if (!tem) return { ok: false, cat: "Falta informar o job", txt: "Informe Job Type e Job No para conferir no TM" };
    const hs = idx.hist.get(k) || []; const ds = idx.due.get(k) || [];
    if (r.acao === "Baixa de job") {
      const dentro = hs.filter((h) => isoDia(h.dateDone) && isoDia(r.data) && dias(h.dateDone, r.data) >= -1 && dias(h.dateDone, r.data) <= 14).sort((a, b) => isoDia(a.dateDone).localeCompare(isoDia(b.dateDone)))[0];
      if (dentro) return { ok: true, cat: "Confirmado no TM", txt: `Baixa confirmada no History em ${fmt(dentro.dateDone)}${dentro.jobHistoryNumber ? ` (${dentro.jobHistoryNumber})` : ""}` };
      if (hs.length) { const datas = [...new Set(hs.map((h) => isoDia(h.dateDone)).filter(Boolean))].sort().reverse(); return { ok: false, cat: "Divergente", txt: `Baixa(s) deste job no History em: ${datas.slice(0, 5).map(fmt).join(", ") || "sem data"}${datas.length > 5 ? "…" : ""} — nenhuma entre 1 dia antes e 14 dias depois da data informada (${fmt(isoDia(r.data))})` }; }
      if (ds.length) return { ok: false, cat: "Não baixado no TM", txt: `Job ainda EM ABERTO no Due${ds[0].dueDate ? ` (vence ${fmt(ds[0].dueDate)})` : ""} — baixa não feita` };
      return { ok: false, cat: "Job não encontrado", txt: "Job não encontrado no Due nem no History importados (nº errado ou planilha do TM desatualizada)" };
    }
    // criação de job
    if (ds.length) return { ok: true, cat: "Confirmado no TM", txt: `Job existe no Due${ds[0].dueDate ? ` — vence ${fmt(ds[0].dueDate)}` : ""}${ds[0].jobName ? ` (${ds[0].jobName})` : ""}` };
    if (hs.length) return { ok: true, cat: "Confirmado no TM", txt: `Job existe no History (último em ${fmt([...hs].sort((a, b) => (b.dateDone || "").localeCompare(a.dateDone || ""))[0].dateDone)})` };
    return { ok: false, cat: "Job não encontrado", txt: "Job criado a bordo ainda não aparece no TM importado — conferir nº ou reimportar o Due" };
  };

  const validarSemJob = (r, c) => {
    if (!bordoSemJob(r)) return c;
    if (!String(r.justificativa || "").trim()) return { ok: false, cat: "Falta justificativa", txt: "Registro sem Job: escreva a justificativa para ele entrar em validação" };
    if (r.valStatus === "Validado") return { ok: true, cat: "Validado (sem job)", txt: `Sem job validado${r.valPor ? ` por ${r.valPor}` : ""}${r.valEm ? ` em ${fmt(r.valEm)}` : ""}${r.valNota ? ` — ${r.valNota}` : ""}` };
    if (r.valStatus === "Recusado") return { ok: false, cat: "Recusado na validação", txt: `Justificativa recusada${r.valPor ? ` por ${r.valPor}` : ""}${r.valNota ? ` — ${r.valNota}` : ""}. Informe o job ou corrija a justificativa.` };
    return { ok: false, cat: "Em validação", txt: "Registro sem Job aguardando validação do administrador" };
  };
  const atrasoEnvio = (r) => (r.recebido && r.data ? Math.round((new Date(r.recebido + "T00:00:00") - new Date(r.data + "T00:00:00")) / 86400000) : null);
  const rows = useMemo(() => tmBordo.filter((r) => (r.data || "") >= BORDO_IMPORT_MIN && BORDO_DEPTS.includes(r.departamento)).map((r) => ({ ...r, _c: r.semAtv ? { ok: true, cat: "Sem atividade", txt: "Departamento informou que não houve atividades neste dia" } : validarSemJob(r, conferir(r)), _f: r.semAtv ? { nos: [], txt: [], auto: false } : frInfo(r), _a: atrasoEnvio(r) })), [tmBordo, idx, frIdx]);
  const norm = (v) => String(v || "").toLowerCase();
  const filtered = rows.filter((r) => (!f.meus || podeRow(r)) && (f.tipo === "Todos" || r.tipo === f.tipo) && (f.status === "Todos" || r._c.cat === f.status) && (f.dept === "Todos" || r.departamento === f.dept)
    && (f.fr === "Todos" || (f.fr === "Com Failure Report" ? r._f.nos.length > 0 : r._f.nos.length === 0))
    && (!f.de || r.data >= f.de) && (!f.ate || r.data <= f.ate)
    && (!f.busca || norm(`${r.descricao} ${r.componente} ${r.jobTipo} ${r.jobNo} ${r.por} ${r.obs} ${r.fr || ""} ${r.departamento || ""}`).includes(norm(f.busca)))).sort((a, b) => (b.data || "").localeCompare(a.data || ""));
  const cats = ["Todos", "Em validação", "Falta justificativa", "Recusado na validação", "Validado (sem job)", "Confirmado no TM", "Não baixado no TM", "Divergente", "Job não encontrado", "Sem registro no TM", "Falta informar o job", "Sem atividade"];
  const iso = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  const diasAtras = (n) => { const d = new Date(hoje + "T00:00:00"); d.setDate(d.getDate() - n); return iso(d); };

  /* controle diário: quais departamentos registraram (ou informaram "sem atividade") em cada um dos últimos 7 dias */
  const obrig = (Array.isArray(regObrig) && regObrig.length ? regObrig : BORDO_DEPTS).filter((d) => BORDO_DEPTS.includes(d));
  const dias7 = Array.from({ length: 7 }, (_, i) => diasAtras(6 - i)).filter((x) => x >= BORDO_INICIO);
  const cont = useMemo(() => { const m = new Map(); tmBordo.forEach((r) => { const k = `${r.departamento}|${r.data}`; m.set(k, (m.get(k) || 0) + 1); }); return m; }, [tmBordo]);
  const feitoEm = (d, dia) => cont.get(`${d}|${dia}`) || 0;
  const pendHoje = obrig.filter((d) => !feitoEm(d, hoje));
  /* pendências em aberto: TODO dia desde o início (até ontem) sem registro continua aparecendo até ser preenchido (ou marcado "sem atividades") */
  const diasDesdeInicio = (() => { const out = []; const d = new Date(BORDO_INICIO + "T00:00:00"); while (iso(d) < hoje && out.length < 400) { out.push(iso(d)); d.setDate(d.getDate() + 1); } return out; })();
  const faltas = obrig.map((d) => ({ d, dias: diasDesdeInicio.filter((x) => !feitoEm(d, x)) }));
  const totalFaltas = faltas.reduce((s, x) => s + x.dias.length, 0);

  const reais = filtered.filter((r) => !r.semAtv);
  const kpi = {
    total: reais.length, ok: reais.filter((r) => r._c.ok).length, pend: reais.filter((r) => !r._c.ok).length,
    corrSemFr: reais.filter((r) => r.tipo === "Corretiva" && r._f.nos.length === 0).length, comFr: reais.filter((r) => r._f.nos.length > 0).length,
    atrasados: reais.filter((r) => r._a !== null && r._a > 1).length,
    baixas: reais.filter((r) => r.acao === "Baixa de job").length, criacoes: reais.filter((r) => r.acao === "Criação de job").length,
  };
  const porTipo = BORDO_TIPOS.map((t) => ({ nome: t, total: reais.filter((r) => r.tipo === t).length }));
  const porDepto = BORDO_DEPTS.map((d) => ({ nome: d, total: reais.filter((r) => r.departamento === d).length })).filter((x) => x.total);
  const porDia = useMemo(() => { const m = new Map(); reais.forEach((r) => { if (r.data) m.set(r.data, (m.get(r.data) || 0) + 1); }); return [...m.entries()].sort().slice(-30).map(([d, n]) => ({ mes: d.slice(8) + "/" + d.slice(5, 7), total: n })); }, [filtered]);

  const NORMA = "NORMA — REGISTRO DIÁRIO DE BORDO\n\n1. Todos os dias, até as 18h (hora de bordo), cada departamento registra no portal (TM Master > Registro de Bordo) TUDO que foi feito no dia: preventiva, preditiva e corretiva.\n2. Cada atividade deve ter o Job Type e o Job No do TM Master e a Ação no TM: baixa de job ou criação de job. Dar a baixa/criar o job no TM ANTES de registrar.\n3. Corretivas devem informar o Failure Report (FAIL-...).\n4. Dia sem atividade também deve ser informado: use o botão “Sem atividades” do departamento.\n5. Atividades comunicadas por mensagem, ligação ou relatório avulso só são consideradas depois de registradas no portal.\n6. O departamento sem registro é cobrado diariamente pelo escritório.";
  const copiarNorma = async () => { try { await navigator.clipboard.writeText(NORMA); setMsg("Norma copiada. Cole no e-mail para o bordo."); } catch (e) { window.prompt("Copie o texto:", NORMA); } setTimeout(() => setMsg(""), 5000); };
  const abrirCobranca = () => {
    try { localStorage.setItem("genesis_email_bordo", emailBordo); } catch (e) { /* ignora */ }
    const hojeTxt = pendHoje.length ? `HOJE (${fmtDataCompleta(hoje)}) ainda sem registro:\n${pendHoje.map((d) => `- ${d}`).join("\n")}\n\n` : "";
    const antes = faltas.filter((x) => x.dias.length).map((x) => `- ${x.d}: ${x.dias.map((d) => fmtDataCompleta(d)).join(", ")}`).join("\n");
    const corpo = (pendHoje.length || antes)
      ? `Prezados,\n\nO REGISTRO DIÁRIO DE BORDO está pendente para os departamentos abaixo.\n\n${hojeTxt}${antes ? `Dias anteriores (últimos 7) sem registro:\n${antes}\n\n` : ""}Favor registrar no portal (TM Master > Registro de Bordo) tudo que foi feito, com o Job Type / Job No do TM, a ação (baixa ou criação de job) e o Failure Report quando houver. Se não houve atividade, use o botão “Sem atividades”.\n\nObrigado.`
      : `Prezados,\n\nTodos os departamentos registraram nos últimos 7 dias. Lembramos que o REGISTRO DIÁRIO DE BORDO deve ser feito todos os dias.\n\nObrigado.`;
    window.location.href = `mailto:${emailBordo.split(/[;,]/).map((x) => x.trim()).filter(Boolean).join(",")}?subject=${encodeURIComponent("COBRANÇA – Registro diário de bordo pendente")}&body=${encodeURIComponent(corpo)}`;
  };

  /* ---------- edição (só as linhas dos departamentos da pessoa; administrador edita todas) ---------- */
  const upd = (id, patch) => setTmBordo((arr) => arr.map((r) => {
    if (!(r.id === id && podeRow(r))) return r;
    const n = { ...r, ...patch, upd: Date.now() };
    /* qualquer mudança do departamento em um registro sem Job devolve para "Em validação"; se o job foi informado, a validação deixa de valer */
    if (bordoSemJob(n)) { if (!("valStatus" in patch)) { n.valStatus = "Em validação"; n.valPor = ""; n.valEm = ""; n.valNota = ""; } }
    else { n.valStatus = ""; n.valPor = ""; n.valEm = ""; n.valNota = ""; n.justificativa = n.justificativa || ""; }
    return n;
  }));
  const validar = (r, status) => {
    if (!admin) return;
    const nota = status === "Recusado" ? (window.prompt("Motivo da recusa (aparece para o departamento):", "") || "") : "";
    if (status === "Recusado" && !nota.trim()) return;
    setTmBordo((arr) => arr.map((x) => (x.id === r.id ? { ...x, valStatus: status, valPor: currentUser?.name || "Administrador", valEm: hoje, valNota: nota, upd: Date.now() } : x)));
  };
  const rem = (r) => {
    if (!window.confirm(`Excluir o registro "${r.descricao || r.jobNo || r.id}"?`)) return;
    setTmBordo((arr) => arr.filter((x) => x.id !== r.id));
    setBordoDeleted((d) => [...new Set([...(d || []), r.id])]);
  };
  const rascunho = () => ({ data: hoje, departamento: deptsEditaveis.length === 1 ? deptsEditaveis[0] : "", tipo: "Preventiva", descricao: "", componente: "", jobTipo: "", jobNo: "", fr: "", acao: "Baixa de job", por: currentUser?.name || "", obs: "", justificativa: "" });
  const salvarNovo = (manter) => {
    if (!novo.departamento) { setMsg("Escolha o departamento."); return; }
    if (!novo.data || novo.data < BORDO_INICIO) { setMsg("O Registro de Bordo vale a partir de " + fmtDataCompleta(BORDO_INICIO) + "."); return; }
    if (!String(novo.descricao).trim()) { setMsg("Descreva a manutenção realizada."); return; }
    if (bordoSemJob(novo) && !String(novo.justificativa || "").trim()) { setMsg("Registro sem Job: a justificativa é obrigatória."); return; }
    const r = { ...novo, ...(bordoSemJob(novo) ? { valStatus: "Em validação" } : { justificativa: "" }), id: uid("BRD"), origem: "Portal", criadoPor: currentUser?.username || "", email: "", recebido: hoje, upd: Date.now() };
    setTmBordo((arr) => [r, ...arr]);
    setNovo(manter ? { ...rascunho(), departamento: novo.departamento, data: novo.data, por: novo.por } : null);
    setMsg("Registro adicionado."); setTimeout(() => setMsg(""), 4000);
  };
  /* ---------- planilha: modelo + importação de vários registros de uma vez ---------- */
  const [imp, setImp] = useState(null);
  const impRef = React.useRef(null);
  const sa = (t) => String(t || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  const baixarModelo = async () => {
    try {
      const qs = new URLSearchParams({ depts: (admin ? BORDO_DEPTS : deptsEditaveis).join("|"), nome: currentUser?.name || "" });
      const r = await fetch("/api/bordo/modelo?" + qs.toString());
      if (!r.ok) throw new Error("indisponível");
      const blob = await r.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = "modelo-registro-de-bordo-OSM.xlsx"; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 2000);
      return;
    } catch (e) { /* sem o modelo profissional: baixa o simples abaixo */ }
    baixarModeloSimples();
  };
  const baixarModeloSimples = () => {
    const dep = deptsEditaveis[0] || "Máquinas";
    const ex = [
      { Data: fmtDataCompleta(hoje), Departamento: dep, Tipo: "Preventiva", "Manutenção realizada": "Troca de filtro do motor de popa BB", Componente: "Motor de popa BB", "Ação no TM": "Baixa de job", "Job Type": "PMS", "Job No": "1234", "Failure Report": "", Observações: "", Justificativa: "" },
      { Data: fmtDataCompleta(hoje), Departamento: dep, Tipo: "Corretiva", "Manutenção realizada": "Reparo de vazamento na linha de óleo", Componente: "Linha de óleo", "Ação no TM": "Criação de job", "Job Type": "COR", "Job No": "5678", "Failure Report": "FR-001", Observações: "", Justificativa: "" },
      { Data: fmtDataCompleta(hoje), Departamento: dep, Tipo: "Preditiva", "Manutenção realizada": "Análise de vibração", Componente: "", "Ação no TM": "Sem job", "Job Type": "", "Job No": "", "Failure Report": "", Observações: "", Justificativa: "Atividade sem job no TM porque ..." },
    ];
    const guia = [["COMO PREENCHER (apague as 3 linhas de exemplo antes de importar)"], [""], ["Data", "dd/mm/aaaa — aceita datas retroativas (a partir de " + fmtDataCompleta(BORDO_IMPORT_MIN) + ") e até hoje"], ["Departamento", (admin ? BORDO_DEPTS : deptsEditaveis).join(" · ")], ["Tipo", BORDO_TIPOS.join(" · ")], ["Ação no TM", BORDO_ACOES.join(" · ")], ["Job Type / Job No", "obrigatórios quando a ação é Baixa ou Criação de job"], ["Justificativa", "obrigatória quando a ação é Sem job (o registro entra em validação)"], ["Várias linhas", "uma linha por atividade; pode ter vários dias e vários jobs na mesma planilha"]];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ex), "Registros"); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(guia), "Instruções");
    XLSX.writeFile(wb, "modelo-registro-de-bordo.xlsx");
  };
  const lerPlanilha = (file) => {
    if (!file) return;
    const rd = new FileReader();
    rd.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames.find((n) => sa(n).startsWith("registro")) || wb.SheetNames[0]];
        const mat = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
        const hi = mat.findIndex((row) => row.some((c) => sa(c).startsWith("data")) && row.some((c) => sa(c).startsWith("depart")));
        if (hi < 0) throw new Error("não achei a linha de cabeçalho (Data, Departamento…). Use o modelo baixado no portal.");
        const lin = mat.slice(hi + 1).map((row) => { const o = {}; mat[hi].forEach((h, j) => { if (String(h).trim()) o[String(h).trim()] = row[j] === undefined ? "" : row[j]; }); o.__n = 0; return o; }).map((o, k) => { o.__n = hi + 2 + k; return o; });
        const col = (o, ...nomes) => { for (const k of Object.keys(o)) if (nomes.some((n) => sa(k).startsWith(sa(n)))) return o[k]; return ""; };
        const dataDe = (v) => { if (typeof v === "number" && v > 20000) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000).toISOString().slice(0, 10); if (v instanceof Date) return isoDia(v); return isoDia(v); };
        const chave = (x) => `${x.departamento}|${x.data}|${sa(x.descricao)}|${bordoNorm(x.jobTipo)}|${bordoNorm(x.jobNo)}`;
        const existentes = new Set(tmBordo.map(chave)); const novas = []; const erros = []; let dup = 0; const vistos = new Set();
        lin.forEach((o, i) => {
          const n = o.__n || i + 2; const vazio = Object.entries(o).every(([k, v]) => k === "__n" || String(v).trim() === ""); if (vazio) return;
          const desc = String(col(o, "manuten", "descri", "atividade")).trim();
          if (/^(exemplo|apague)/i.test(desc)) return;
          const data = dataDe(col(o, "data")); const dep = canonDept(String(col(o, "depart")).trim() || (deptsEditaveis.length === 1 ? deptsEditaveis[0] : ""));
          const tp = BORDO_TIPOS.find((t) => sa(t) === sa(col(o, "tipo"))) ; const ac = (() => { const a = sa(col(o, "acao", "ação")); return BORDO_ACOES.find((x) => sa(x) === a) || (a.startsWith("baixa") ? "Baixa de job" : a.startsWith("cria") ? "Criação de job" : a.startsWith("sem") ? "Sem job" : null); })();
          const reg = { data, departamento: dep, tipo: tp || "", descricao: desc, componente: String(col(o, "componente")).trim(), jobTipo: String(col(o, "job type", "tipo do job", "jobtype")).trim().toUpperCase(), jobNo: String(col(o, "job no", "job n", "numero do job", "jobno")).trim().replace(/\.0+$/, ""), fr: String(col(o, "failure")).trim(), acao: ac || "", por: currentUser?.name || "", obs: String(col(o, "observ")).trim(), justificativa: String(col(o, "justific")).trim() };
          const err = [];
          if (!data) err.push("data inválida ou vazia"); else if (data < BORDO_IMPORT_MIN) err.push("data anterior a " + fmtDataCompleta(BORDO_IMPORT_MIN)); else if (data > hoje) err.push("data no futuro");
          if (!BORDO_DEPTS.includes(dep)) err.push(`departamento “${dep || "vazio"}” inválido`); else if (!pode(dep)) err.push(`você não tem permissão para registrar em ${dep}`);
          if (!tp) err.push("tipo deve ser Preventiva, Preditiva ou Corretiva"); if (!desc) err.push("falta a manutenção realizada"); if (!ac) err.push("ação deve ser Baixa de job, Criação de job ou Sem job");
          if (ac && ac !== "Sem job" && (!reg.jobTipo || !reg.jobNo)) err.push("faltam Job Type e Job No");
          if (ac === "Sem job" && !reg.justificativa) err.push("ação Sem job exige justificativa");
          if (err.length) { erros.push({ n, motivo: err.join("; "), desc: desc || "(sem descrição)" }); return; }
          const k = chave(reg); if (existentes.has(k) || vistos.has(k)) { dup++; return; } vistos.add(k); novas.push(reg);
        });
        setImp({ nome: file.name, novas, erros, dup });
      } catch (e) { setMsg("Não consegui ler a planilha: " + (e?.message || e)); setTimeout(() => setMsg(""), 7000); }
    };
    rd.readAsArrayBuffer(file);
  };
  const confirmarImport = () => {
    if (!imp || !imp.novas.length) { setImp(null); return; }
    const agora = Date.now();
    const regs = imp.novas.map((x, i) => ({ ...x, ...(x.acao === "Sem job" ? { valStatus: "Em validação" } : { justificativa: "" }), id: uid("BRD"), origem: "Planilha", criadoPor: currentUser?.username || "", email: "", recebido: x.data < BORDO_INICIO ? x.data : hoje, retro: x.data < BORDO_INICIO ? true : undefined, upd: agora + i }));
    setTmBordo((arr) => [...regs, ...arr]);
    setMsg(`${regs.length} registro(s) importado(s) da planilha.`); setTimeout(() => setMsg(""), 6000); setImp(null);
  };
  const semAtividade = (d, dia) => {
    if (!pode(d)) return;
    setTmBordo((arr) => [{ id: uid("BRD"), data: dia, departamento: d, tipo: "Preventiva", descricao: "Sem atividades no dia", componente: "", jobTipo: "", jobNo: "", fr: "", acao: "Sem job", origem: "Portal", por: currentUser?.name || "", email: "", obs: "", recebido: hoje, semAtv: true, upd: Date.now() }, ...arr]);
  };

  React.useEffect(() => {
    if (!active || !setExportXlsxFn || !setReportFn) return;
    const lin = () => filtered.map((r) => ({ Data: fmt(r.data), Departamento: r.departamento || "", Tipo: r.semAtv ? "Sem atividade" : r.tipo, Manutenção: r.descricao, Componente: r.componente, "Job Type": r.jobTipo, "Job No": r.jobNo, "Failure Report": r._f.nos.join(", "), "Ação no TM": r.semAtv ? "" : r.acao, "Reportado por": r.por, "Justificativa (sem job)": r.justificativa || "", "Validação": r.valStatus || "", Observação: r.obs, "Lançado em": r.recebido ? fmt(r.recebido) : "", "Dias até o lançamento": r._a ?? "", Conferência: r._c.cat, Detalhe: r._c.txt }));
    const matriz = () => obrig.map((d) => ({ Departamento: d, ...Object.fromEntries(dias7.map((x) => [fmt(x), feitoEm(d, x) ? `${feitoEm(d, x)} registro(s)` : "SEM REGISTRO"])) }));
    setExportXlsxFn(() => () => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(lin().length ? lin() : [{ "—": "Sem registros" }]), "Registro de Bordo"); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(matriz()), "Controle diário"); XLSX.writeFile(wb, `genesis-registro-de-bordo-${todayISO()}.xlsx`); });
    setReportFn(() => () => {
      const doc = new jsPDF({ orientation: "landscape" });
      let y = pdfHeader(doc, "TM Master — Registro de Bordo", `${filtered.length} registro(s) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
      y = pdfKpis(doc, y, [{ label: "Registros", value: kpi.total }, { label: "Confirmados no TM", value: kpi.ok }, { label: "Pendentes no TM", value: kpi.pend }, { label: "Baixas", value: kpi.baixas }, { label: "Jobs criados", value: kpi.criacoes }, { label: "Deptos sem registro hoje", value: pendHoje.length }]);
      y = pdfSectionTitle(doc, y, "Controle diário por departamento (últimos 7 dias)");
      y = pdfTable(doc, y, ["Departamento", ...dias7.map((x) => fmt(x).slice(0, 5))], obrig.map((d) => [d, ...dias7.map((x) => (feitoEm(d, x) ? String(feitoEm(d, x)) : "—"))]));
      y = pdfSectionTitle(doc, y, "Registros e conferência no TM Master");
      pdfTable(doc, y, ["Data", "Depto.", "Tipo", "Manutenção", "Job", "Failure Report", "Ação no TM", "Reportado por", "Conferência", "Detalhe"],
        filtered.map((r) => [fmt(r.data), r.departamento || "—", r.semAtv ? "Sem atividade" : r.tipo, r.descricao, `${r.jobTipo} ${r.jobNo}`.trim() || "—", r._f.nos.join(", ") || "—", r.semAtv ? "—" : r.acao, r.por || "—", r._c.cat, r._c.txt]));
      pdfSave(doc, "relatorio-registro-de-bordo");
    });
  }, [active, filtered, kpi, tmBordo, regObrig]);

  const cor = (c) => (c === "Confirmado no TM" || c === "Sem atividade" || c === "Validado (sem job)" ? "d1" : c === "Divergente" || c === "Falta informar o job" || c === "Em validação" ? "d2" : "d3");
  const inp = { padding: "4px 6px", border: "1px solid var(--border)", borderRadius: 4, fontSize: 12, background: "var(--panel, #fff)", color: "var(--text)", width: "100%" };
  const sel = { padding: "5px 8px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 12, background: "var(--panel, #fff)", color: "var(--text)" };
  const ro = (v) => <span style={{ fontSize: 12 }}>{v === "" || v == null ? "—" : String(v)}</span>;
  const campo = (rot, el, span) => <div className="g-field" style={span ? { gridColumn: "1 / -1" } : undefined}><label>{rot}</label>{el}</div>;
  const nv = (k, tipo = "text", extra = {}) => <input className="g-edit" type={tipo} value={novo[k]} onChange={(e) => setNovo((p) => ({ ...p, [k]: e.target.value }))} {...extra} />;
  const nd = (k, ops) => <select className="g-edit" value={novo[k]} onChange={(e) => setNovo((p) => ({ ...p, [k]: e.target.value }))}>{ops.map((o) => <option key={o} value={o}>{o}</option>)}</select>;
  const th = (t, w) => <th style={{ minWidth: w }}>{t}</th>;

  return (
    <>
      <div className="g-mode-toggle" style={{ marginBottom: 14, width: "fit-content" }}>
        <button className={sub === "preench" ? "active" : ""} onClick={() => setSub("preench")}>Preenchimento</button>
        <button className={sub === "dash" ? "active" : ""} onClick={() => setSub("dash")}>Dashboard</button>
      </div>
      {sub === "preench" && <>
      <div className="g-alert" style={{ marginBottom: 12, background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>
        <Lock size={14} style={{ marginTop: 1 }} />
        {admin ? "Administrador: você edita os registros de todos os departamentos."
          : podeCriar ? <>Você registra e edita só {deptsEditaveis.length > 1 ? "os departamentos" : "o departamento"} <b>{deptsEditaveis.join(", ")}</b>. Os registros dos outros aparecem para consulta.</>
          : "Somente consulta. Peça ao administrador para definir o seu departamento em Configurações para poder registrar."}
      </div>

      <div className="g-panel" style={{ borderLeft: `4px solid ${pendHoje.length ? "var(--crit)" : "var(--ok)"}` }}>
        <div className="g-panel-head" style={{ flexWrap: "wrap", gap: 8 }}>
          <span className="g-panel-title">Controle diário — quem registrou {pendHoje.length ? <span style={{ color: "var(--crit)" }}>· hoje faltam {pendHoje.length} de {obrig.length} departamentos</span> : <span style={{ color: "var(--ok)" }}>· todos os departamentos registraram hoje</span>}</span>
        </div>
        <div className="g-table-wrap">
          <table className="g-table" style={{ minWidth: 0, width: "100%" }}>
            <thead><tr><th>Departamento</th>{dias7.map((d) => <th key={d} style={{ textAlign: "center", ...(d === hoje ? { color: "var(--teal)" } : {}) }}>{fmt(d).slice(0, 5)}{d === hoje ? " (hoje)" : ""}</th>)}<th style={{ textAlign: "center" }}>Pendências em aberto</th></tr></thead>
            <tbody>
              {obrig.map((d) => {
                const f2 = faltas.find((x) => x.d === d);
                return (
                  <tr className="g-row" key={d}>
                    <td><b>{d}</b></td>
                    {dias7.map((dia) => {
                      const n = feitoEm(d, dia);
                      return <td key={dia} style={{ textAlign: "center" }}>
                        {n ? <span style={{ color: "var(--ok)", fontWeight: 700 }}>✓ {n}</span>
                          : <span style={{ color: dia === hoje ? "var(--warn)" : "var(--crit)", fontWeight: 700 }}>{dia === hoje ? "pendente" : "✗ falta"}{pode(d) && <div><span style={{ fontSize: 10, fontWeight: 400, textDecoration: "underline", cursor: "pointer", color: "var(--text-dim)" }} title="Informa que o departamento não teve atividades neste dia" onClick={() => semAtividade(d, dia)}>sem atividades</span></div>}</span>}
                      </td>;
                    })}
                    <td style={{ textAlign: "center", fontWeight: 700, color: f2.dias.length ? "var(--crit)" : "var(--ok)", minWidth: 150 }}>{f2.dias.length}
                      {f2.dias.length > 0 && <div style={{ fontWeight: 400, fontSize: 11, display: "flex", flexWrap: "wrap", gap: 4, justifyContent: "center", marginTop: 3 }}>{f2.dias.map((x) => <span key={x} style={{ background: "rgba(239,68,68,.12)", color: "var(--crit)", borderRadius: 4, padding: "1px 5px" }} title={dias7.includes(x) ? "Em aberto" : "Em aberto desde esta data — continua cobrando até ser preenchido"}>{fmt(x).slice(0, 5)}</span>)}</div>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="g-muted" style={{ fontSize: 11, marginTop: 6, lineHeight: 1.5 }}>
          ✓ = registrou (o número é a quantidade de linhas, incluindo “sem atividades”). O dia de hoje só vira falta quando virar o dia. Uma falta NÃO some com o tempo: fica listada em “Pendências em aberto” até o departamento lançar o registro daquele dia (use a data do dia no formulário) ou marcar “sem atividades”.
        </div>
        {admin && (
          <div style={{ marginTop: 8, fontSize: 12 }}>
            <b>Departamentos que devem registrar todo dia:</b>{" "}
            {BORDO_DEPTS.map((d) => <label key={d} style={{ marginRight: 12, cursor: "pointer", whiteSpace: "nowrap" }}><input type="checkbox" checked={obrig.includes(d)} onChange={(e) => setRegObrig(e.target.checked ? BORDO_DEPTS.filter((x) => obrig.includes(x) || x === d) : BORDO_DEPTS.filter((x) => obrig.includes(x) && x !== d))} /> {d}</label>)}
          </div>
        )}
        <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", alignItems: "center", marginTop: 10 }}>
          <input type="text" value={emailBordo} onChange={(e) => setEmailBordo(e.target.value)} placeholder="E-mail para a cobrança (vários separados por vírgula)" style={{ flex: 1, minWidth: 280, padding: "7px 10px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 13 }} />
          <button className="g-btn primary g-keep" onClick={abrirCobranca}>Abrir e-mail de cobrança ({pendHoje.length} hoje · {totalFaltas} dia(s) em atraso)</button>
          <button className="g-btn" onClick={copiarNorma}>Copiar norma (mensagem padrão)</button>
        </div>
      </div>

      <div className="g-panel" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
        <div style={{ maxWidth: 820 }}>
          <div className="g-panel-title" style={{ marginBottom: 4 }}>Registro de Bordo — preenchido aqui, conferido no TM Master</div>
          <div className="g-muted" style={{ fontSize: 11.5, lineHeight: 1.5 }}>
            Cada departamento registra no próprio portal, no fim do dia, uma linha por atividade (preventiva, preditiva ou corretiva), dizendo se foi <b>baixa de job</b> ou <b>criação de job</b> e o nº do job.
            O portal <b>confere sozinho</b> contra o History e o Due do TM Master. Dia sem atividade: botão “sem atividades” no controle acima. Corretiva: informe o <b>Failure Report</b>; o portal também detecta quando o job já está ligado a um.
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 8 }}>
          <SaveBadge />
          {podeCriar && <div style={{ display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" }}>
            <button className="g-btn g-keep" onClick={baixarModelo}><Download size={14} />Baixar modelo</button>
            <input ref={impRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => { lerPlanilha(e.target.files?.[0]); e.target.value = ""; }} />
            <button className="g-btn g-keep" onClick={() => impRef.current?.click()}><Upload size={14} />Importar planilha</button>
            <button className="g-btn primary g-keep" onClick={() => setNovo(novo ? null : rascunho())}><Plus size={14} />{novo ? "Fechar formulário" : "Novo registro"}</button>
          </div>}
        </div>
      </div>
      {imp && (
        <div className="g-panel" style={{ marginBottom: 12, borderColor: imp.erros.length ? "var(--warn)" : "var(--ok)" }}>
          <div className="g-panel-title" style={{ marginBottom: 6 }}>Importar “{imp.nome}”</div>
          <div style={{ fontSize: 13, marginBottom: 8 }}>
            <b style={{ color: "var(--ok)" }}>{imp.novas.length} linha(s) prontas para importar</b>
            {imp.dup > 0 && <> · {imp.dup} já existiam (ignoradas, sem duplicar)</>}
            {imp.erros.length > 0 && <> · <b style={{ color: "var(--crit)" }}>{imp.erros.length} com erro (não serão importadas)</b></>}
          </div>
          {imp.erros.length > 0 && <div style={{ maxHeight: 160, overflow: "auto", fontSize: 12, background: "rgba(220,38,38,.06)", borderRadius: 6, padding: "6px 10px", marginBottom: 8 }}>
            {imp.erros.slice(0, 50).map((e, i) => <div key={i}>Linha {e.n} — {e.desc}: {e.motivo}</div>)}
            {imp.erros.length > 50 && <div>… e mais {imp.erros.length - 50}</div>}
          </div>}
          <div style={{ display: "flex", gap: 8 }}>
            <button className="g-btn primary g-keep" disabled={!imp.novas.length} onClick={confirmarImport}><Upload size={14} />Importar {imp.novas.length} registro(s)</button>
            <button className="g-btn g-keep" onClick={() => setImp(null)}>Cancelar</button>
          </div>
        </div>
      )}
      {msg && <div className="g-alert" style={{ marginBottom: 12, background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>{msg}</div>}

      {novo && (
        <div className="g-panel" style={{ borderLeft: "4px solid var(--teal)" }}>
          <div className="g-panel-head"><span className="g-panel-title">Novo registro de bordo</span></div>
          <div className="tbl-detail-grid">
            {campo("Data *", nv("data", "date", { min: BORDO_INICIO }))}
            {campo("Departamento *", deptsEditaveis.length === 1 ? <input className="g-edit" value={novo.departamento} disabled /> : <select className="g-edit" value={novo.departamento} onChange={(e) => setNovo((p) => ({ ...p, departamento: e.target.value }))}><option value="">— escolher —</option>{deptsEditaveis.map((o) => <option key={o} value={o}>{o}</option>)}</select>)}
            {campo("Tipo", nd("tipo", BORDO_TIPOS))}
            {campo("Ação no TM", nd("acao", BORDO_ACOES))}
            {campo("Manutenção realizada *", nv("descricao"), true)}
            {campo("Componente", nv("componente"))}
            {campo("Job Type", nv("jobTipo", "text", { onChange: (e) => setNovo((p) => ({ ...p, jobTipo: e.target.value.toUpperCase() })), placeholder: "ex.: ONE" }))}
            {campo("Job No", nv("jobNo"))}
            {campo("Failure Report", nv("fr", "text", { onChange: (e) => setNovo((p) => ({ ...p, fr: e.target.value.toUpperCase() })), placeholder: "FAIL-GEN-2026-xxxx (corretivas)" }))}
            {campo("Reportado por", nv("por"))}
            {bordoSemJob(novo) && campo("Justificativa (sem Job) * — o registro entra em validação", <textarea className="g-edit-wrap" rows={2} value={novo.justificativa || ""} placeholder="Por que não há job? (ex.: serviço sem job cadastrado, aguardando criação…)" onChange={(e) => setNovo((p) => ({ ...p, justificativa: e.target.value }))} style={{ borderColor: String(novo.justificativa || "").trim() ? undefined : "var(--crit)" }} />, true)}
            {campo("Observação", <textarea className="g-edit-wrap" rows={2} value={novo.obs} onChange={(e) => setNovo((p) => ({ ...p, obs: e.target.value }))} />, true)}
          </div>
          <div className="g-flex" style={{ gap: 8, marginTop: 10 }}>
            <button className="g-btn primary g-keep" onClick={() => salvarNovo(false)}>Adicionar registro</button>
            <button className="g-btn" onClick={() => salvarNovo(true)} title="Adiciona e já abre outro em branco, mantendo data, departamento e quem reportou">Adicionar e lançar outro</button>
            <button className="g-btn" onClick={() => setNovo(null)}>Cancelar</button>
          </div>
        </div>
      )}

      </>}
      {sub === "preench" && <>
      <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
        <div className="g-field"><label>Buscar</label><input type="text" value={f.busca} onChange={(e) => setF((p) => ({ ...p, busca: e.target.value }))} placeholder="manutenção, componente, job, quem..." style={{ minWidth: 220 }} /></div>
        <div className="g-field"><label>Departamento</label><select value={f.dept} onChange={(e) => setF((p) => ({ ...p, dept: e.target.value }))}>{["Todos", ...BORDO_DEPTS].map((o) => <option key={o}>{o}</option>)}</select></div>
        <div className="g-field"><label>&nbsp;</label><label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}><input type="checkbox" checked={!!f.meus} onChange={(e) => setF((p) => ({ ...p, meus: e.target.checked }))} />Só os que posso editar</label></div>
        <div className="g-field"><label>Tipo</label><select value={f.tipo} onChange={(e) => setF((p) => ({ ...p, tipo: e.target.value }))}>{["Todos", ...BORDO_TIPOS].map((o) => <option key={o}>{o}</option>)}</select></div>
        <div className="g-field"><label>Conferência no TM</label><select value={f.status} onChange={(e) => setF((p) => ({ ...p, status: e.target.value }))}>{cats.map((o) => <option key={o}>{o}</option>)}</select></div>
        <div className="g-field"><label>Failure Report</label><select value={f.fr} onChange={(e) => setF((p) => ({ ...p, fr: e.target.value }))}>{["Todos", "Com Failure Report", "Sem Failure Report"].map((o) => <option key={o}>{o}</option>)}</select></div>
        <div className="g-field"><label>De</label><input type="date" min={BORDO_IMPORT_MIN} value={f.de} onChange={(e) => setF((p) => ({ ...p, de: e.target.value }))} /></div>
        <div className="g-field"><label>Até</label><input type="date" value={f.ate} onChange={(e) => setF((p) => ({ ...p, ate: e.target.value }))} /></div>
      </div>

      </>}
      {sub === "dash" && <>
      <div className="g-filterbar" style={{ padding: "12px 16px", marginBottom: 14, borderRadius: 6 }}>
        <div className="g-field"><label>Buscar</label><input type="text" value={f.busca} onChange={(e) => setF((p) => ({ ...p, busca: e.target.value }))} placeholder="manutenção, componente, job, quem..." style={{ minWidth: 220 }} /></div>
        <div className="g-field"><label>Departamento</label><select value={f.dept} onChange={(e) => setF((p) => ({ ...p, dept: e.target.value }))}>{["Todos", ...BORDO_DEPTS].map((o) => <option key={o}>{o}</option>)}</select></div>
        <div className="g-field"><label>&nbsp;</label><label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}><input type="checkbox" checked={!!f.meus} onChange={(e) => setF((p) => ({ ...p, meus: e.target.checked }))} />Só os que posso editar</label></div>
        <div className="g-field"><label>Tipo</label><select value={f.tipo} onChange={(e) => setF((p) => ({ ...p, tipo: e.target.value }))}>{["Todos", ...BORDO_TIPOS].map((o) => <option key={o}>{o}</option>)}</select></div>
        <div className="g-field"><label>Conferência no TM</label><select value={f.status} onChange={(e) => setF((p) => ({ ...p, status: e.target.value }))}>{cats.map((o) => <option key={o}>{o}</option>)}</select></div>
        <div className="g-field"><label>Failure Report</label><select value={f.fr} onChange={(e) => setF((p) => ({ ...p, fr: e.target.value }))}>{["Todos", "Com Failure Report", "Sem Failure Report"].map((o) => <option key={o}>{o}</option>)}</select></div>
        <div className="g-field"><label>De</label><input type="date" min={BORDO_IMPORT_MIN} value={f.de} onChange={(e) => setF((p) => ({ ...p, de: e.target.value }))} /></div>
        <div className="g-field"><label>Até</label><input type="date" value={f.ate} onChange={(e) => setF((p) => ({ ...p, ate: e.target.value }))} /></div>
      </div>

      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))" }}>
        {bigKpi("Sem job em validação", rows.filter((r) => r._c.cat === "Em validação" || r._c.cat === "Falta justificativa").length, rows.some((r) => r._c.cat === "Em validação" || r._c.cat === "Falta justificativa") ? "var(--warn)" : "var(--ok)", AlertTriangle)}
        {bigKpi("Departamentos sem registro hoje", pendHoje.length, pendHoje.length ? "var(--crit)" : "var(--ok)", AlertTriangle)}
        {bigKpi("Dias em atraso (em aberto)", totalFaltas, totalFaltas ? "var(--crit)" : "var(--ok)", Clock)}
        {bigKpi("Registros", kpi.total, "var(--teal)", ClipboardList)}
        {bigKpi("Confirmados no TM", kpi.ok, "var(--ok)", ClipboardList)}
        {bigKpi("Pendentes no TM", kpi.pend, kpi.pend ? "var(--crit)" : "var(--ok)", AlertTriangle)}
        {bigKpi("Baixas de job", kpi.baixas, "var(--teal)", Wrench)}
        {bigKpi("Jobs criados", kpi.criacoes, "var(--teal)", Wrench)}
        {bigKpi("Lançados com atraso (>1 dia)", kpi.atrasados, kpi.atrasados ? "var(--warn)" : "var(--ok)", Clock)}
        {bigKpi("Corretivas sem Failure Report", kpi.corrSemFr, kpi.corrSemFr ? "var(--crit)" : "var(--ok)", AlertTriangle)}
      </div>
      {tmBordo.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14, marginBottom: 14 }}>
          <MesBarChart title="Registros por dia (últimos 30 com registro)" height={200} data={porDia} bars={[{ key: "total", name: "Registros", color: "#3B82F6" }]} />
          <MesBarChart title="Por departamento" height={200} data={porDepto} bars={[{ key: "total", name: "Registros", color: "#22C55E" }]} />
          <MesBarChart title="Por tipo de manutenção" height={200} data={porTipo} bars={[{ key: "total", name: "Registros", color: "#8B5CF6" }]} />
        </div>
      )}

      </>}
      {sub === "preench" && (
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Registros ({filtered.length}) — e o que o portal achou no TM Master</span></div>
        <div className="g-table-wrap" style={{ maxHeight: 640, overflow: "auto" }}>
          <table className="g-table tbl-modern g-editavel">
            <thead><tr>{th("Data", 118)}{th("Departamento", 130)}{th("Tipo", 112)}{th("Manutenção", 240)}{th("Componente", 150)}{th("Job Type", 70)}{th("Job No", 70)}{th("Failure Report", 190)}{th("Ação no TM", 130)}{th("Justificativa / validação", 230)}{th("Reportado por", 120)}{th("Observação", 180)}{th("Lançado em", 110)}{th("Conferência no TM", 260)}<th></th></tr></thead>
            <tbody>
              {filtered.length === 0 && <tr><td colSpan={15} className="g-muted" style={{ textAlign: "center", padding: 18 }}>{tmBordo.length ? "Nenhum registro com esse filtro." : "Nenhum registro ainda. Use “Novo registro” para lançar o primeiro."}</td></tr>}
              {filtered.map((r) => {
                const ok = podeRow(r);
                const E = (k, el, ver) => (ok ? el : ro(ver !== undefined ? ver : r[k]));
                return (
                  <tr className="g-row tbl-row" key={r.id} title={ok ? undefined : "Somente leitura: você edita registros dos seus departamentos liberados ou que você mesmo lançou."} style={{ "--c": r._c.ok ? "#22C55E" : "#EF4444" }}>
                    <td>{E("data", <input type="date" min={BORDO_IMPORT_MIN} value={r.data} onChange={(e) => e.target.value >= BORDO_IMPORT_MIN && upd(r.id, { data: e.target.value })} style={inp} />, fmt(r.data))}</td>
                    <td>{ok && !r.semAtv ? <select value={r.departamento || ""} onChange={(e) => upd(r.id, { departamento: e.target.value })} style={inp}>{!r.departamento && <option value="">—</option>}{(r.departamento && !deptsEditaveis.includes(r.departamento) ? [r.departamento, ...deptsEditaveis] : deptsEditaveis).map((o) => <option key={o}>{o}</option>)}</select> : <span className="tbl-chip">{r.departamento || "sem depto."}</span>}</td>
                    <td>{r.semAtv ? <span className="tbl-chip">Sem atividade</span> : E("tipo", <select value={r.tipo} onChange={(e) => upd(r.id, { tipo: e.target.value })} style={inp}>{BORDO_TIPOS.map((o) => <option key={o}>{o}</option>)}</select>)}</td>
                    <td>{E("descricao", <input value={r.descricao} onChange={(e) => upd(r.id, { descricao: e.target.value })} style={inp} />)}</td>
                    <td>{r.semAtv ? ro("") : E("componente", <input value={r.componente} onChange={(e) => upd(r.id, { componente: e.target.value })} style={inp} />)}</td>
                    <td>{r.semAtv ? ro("") : E("jobTipo", <input value={r.jobTipo} onChange={(e) => upd(r.id, { jobTipo: e.target.value.toUpperCase() })} style={inp} />)}</td>
                    <td>{r.semAtv ? ro("") : E("jobNo", <input value={r.jobNo} onChange={(e) => upd(r.id, { jobNo: e.target.value })} style={inp} />)}</td>
                    <td style={{ whiteSpace: "normal" }}>
                      {r.semAtv ? ro("") : E("fr", <input value={r.fr || ""} onChange={(e) => upd(r.id, { fr: e.target.value.toUpperCase() })} placeholder={r._f.auto ? r._f.nos[0] : "FAIL-GEN-2026-xxxx"} style={inp} />)}
                      {r._f.nos.length > 0 && <div style={{ fontSize: 10.5, marginTop: 3, color: "var(--ok)" }}>{r._f.auto ? "ligado automaticamente: " : "ligado: "}{r._f.txt.join(" · ")}</div>}
                      {r._f.nos.length === 0 && !r.semAtv && r.tipo === "Corretiva" && frIdx.has && <div style={{ fontSize: 10.5, marginTop: 3, color: "var(--crit)" }}>Corretiva sem Failure Report ligado</div>}
                    </td>
                    <td>{r.semAtv ? ro("") : E("acao", <select value={r.acao} onChange={(e) => upd(r.id, { acao: e.target.value })} style={inp}>{BORDO_ACOES.map((o) => <option key={o}>{o}</option>)}</select>)}</td>
                    <td style={{ whiteSpace: "normal" }}>
                      {bordoSemJob(r) ? (<>
                        {E("justificativa", <input value={r.justificativa || ""} onChange={(e) => upd(r.id, { justificativa: e.target.value })} placeholder="Justificativa (obrigatória)" style={{ ...inp, borderColor: String(r.justificativa || "").trim() ? undefined : "var(--crit)" }} />)}
                        <div style={{ fontSize: 10.5, marginTop: 3, color: r.valStatus === "Validado" ? "var(--ok)" : r.valStatus === "Recusado" ? "var(--crit)" : "var(--warn)", fontWeight: 700 }}>{r.valStatus || "Em validação"}{r.valPor ? ` · ${r.valPor}` : ""}</div>
                        {admin && String(r.justificativa || "").trim() && <div className="g-flex g-keep" style={{ gap: 4, marginTop: 3 }}>
                          {r.valStatus !== "Validado" && <button className="g-btn primary g-keep" style={{ padding: "2px 8px", fontSize: 11 }} onClick={() => validar(r, "Validado")}>Validar</button>}
                          {r.valStatus !== "Recusado" && <button className="g-btn g-keep" style={{ padding: "2px 8px", fontSize: 11 }} onClick={() => validar(r, "Recusado")}>Recusar</button>}
                        </div>}
                      </>) : <span className="g-muted" style={{ fontSize: 11 }}>—</span>}
                    </td>
                    <td>{E("por", <input value={r.por} onChange={(e) => upd(r.id, { por: e.target.value })} style={inp} />)}</td>
                    <td>{E("obs", <input value={r.obs} onChange={(e) => upd(r.id, { obs: e.target.value })} style={inp} />)}</td>
                    <td style={{ whiteSpace: "normal", fontFamily: "var(--mono)", fontSize: 12 }}>{r.recebido ? fmt(r.recebido) : "—"}{r._a !== null && r._a > 1 && <div style={{ color: "var(--crit)", fontSize: 10.5 }}>lançado +{r._a} dias depois</div>}</td>
                    <td style={{ whiteSpace: "normal" }}><span className={`tbl-days ${cor(r._c.cat)}`}>{r._c.cat}</span><div className="g-muted" style={{ fontSize: 11, marginTop: 3 }}>{r._c.txt}</div></td>
                    <td>{ok ? <span className="g-btn ghost danger g-keep" title="Excluir registro" onClick={() => rem(r)}><Trash2 size={13} /></span> : <Lock size={12} color="var(--text-faint)" title="Só o departamento dono (ou o administrador) edita" />}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="g-muted" style={{ fontSize: 10.5, marginTop: 6, lineHeight: 1.5 }}>
          Conferência: <b>Baixa de job</b> é confirmada quando existe no History um job com o mesmo Job Type e Job No, com data de execução entre 1 dia antes e 14 dias depois da data informada.
          <b> Criação de job</b> é confirmada quando o job aparece no Due ou no History. Para a conferência valer, mantenha o Due e o History do TM Master sempre atualizados.
        </div>
      </div>
      )}
    </>
  );
};

function TmMasterView({ tmDue, tmHistory, tmDueSnapshots, tmFailures, handleImportFailures, tmBordo, setTmBordo, setBordoDeleted, regObrig, setRegObrig, currentUser, setReportFn, setExportXlsxFn, handleImportTmDue, handleImportTmHistory, tmSubTab, setTmSubTab, updTmDue, onToggleAddToPlanning }) {
  const dueFileRef = useRef(null);
  const historyFileRef = useRef(null);

  const isVencida = (r) => r.diffValue !== null && r.diffValue < 0;
  const isAteVencer40 = (r) => r.diffUnit === "D" && r.diffValue !== null && r.diffValue >= 0 && r.diffValue <= 40;
  const isVencidaOuAte40 = (r) => isVencida(r) || isAteVencer40(r);
  const isCritica = (r) => r.pri === "High";
  const isCorretiva = (r) => (r.jobType || "").toUpperCase() === "ONE";
  const isPostergada = (h) => (h.remarks || "").toLowerCase().includes("postpon") || (h.postponed || "").toString().trim().toUpperCase() === "X";

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
  const [histf, setHistf] = useState({ busca: "", jobType: "Todos", department: "Todos", usuario: "Todos", soPostergadas: false, meses: [] });
  const [histPag, setHistPag] = useState(1);
  const [histPorPag, setHistPorPag] = useState(50);
  React.useEffect(() => { setHistPag(1); }, [histf, histPorPag]);
  const histJobTypes = useMemo(() => [...new Set(tmHistory.map((h) => h.jobType).filter(Boolean))].sort(), [tmHistory]);
  const histDepartments = useMemo(() => [...new Set(tmHistory.map((h) => deptOfHistory(h)).filter(Boolean))].sort(), [tmHistory, codeToDept]);
  const histUsuarios = useMemo(() => [...new Set(tmHistory.map((h) => h.doneByName).filter(Boolean))].sort(), [tmHistory]);
  const hasActiveFilterHist = histf.busca || histf.jobType !== "Todos" || histf.department !== "Todos" || histf.usuario !== "Todos" || histf.soPostergadas || (histf.meses || []).length > 0;
  const histMeses = useMemo(() => [...new Set(tmHistory.map((h) => (h.dateDone || "").slice(0, 7)).filter((k) => /^\d{4}-\d{2}$/.test(k)))].sort().reverse(), [tmHistory]);
  const histMesLabel = (k) => `${MONTH_NAMES[Number(k.slice(5, 7)) - 1]} / ${k.slice(0, 4)}`;
  const filteredHistory = useMemo(() => {
    const norm = (s) => (s || "").toString().toLowerCase();
    return tmHistory.filter((h) => {
      const inBusca = !histf.busca || norm(h.jobName).includes(norm(histf.busca)) || norm(h.componentName).includes(norm(histf.busca)) || norm(h.doneByName).includes(norm(histf.busca));
      const inJobType = histf.jobType === "Todos" || h.jobType === histf.jobType;
      const inDep = histf.department === "Todos" || deptOfHistory(h) === histf.department;
      const inUsuario = histf.usuario === "Todos" || h.doneByName === histf.usuario;
      const inPost = !histf.soPostergadas || isPostergada(h);
      const inMes = !(histf.meses || []).length || histf.meses.includes((h.dateDone || "").slice(0, 7));
      return inBusca && inJobType && inDep && inUsuario && inPost && inMes;
    }).sort((a, b) => (b.dateDone || "").localeCompare(a.dateDone || ""));
  }, [tmHistory, histf, codeToDept]);
  const histTotalPag = Math.max(1, Math.ceil(filteredHistory.length / histPorPag));
  const histPagAtual = Math.min(histPag, histTotalPag);
  const histPagina = useMemo(() => filteredHistory.slice((histPagAtual - 1) * histPorPag, histPagAtual * histPorPag), [filteredHistory, histPagAtual, histPorPag]);

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
    const vis = histPagina;
    if (!groupHist) return [{ key: "todos", label: "Todos", rows: vis }];
    const map = new Map();
    vis.forEach((h) => { const k = h.dateDone ? h.dateDone.slice(0, 7) : "sem-data"; if (!map.has(k)) map.set(k, []); map.get(k).push(h); });
    return [...map.entries()].sort((x, y) => (x[0] === "sem-data") - (y[0] === "sem-data") || y[0].localeCompare(x[0])).map(([key, rows]) => ({
      key, rows: rows.sort((p, q) => (q.dateDone || "").localeCompare(p.dateDone || "")),
      label: key === "sem-data" ? "Sem data de fechamento" : `${MONTH_NAMES[Number(key.slice(5, 7)) - 1]} / ${key.slice(0, 4)}`,
    }));
  }, [histPagina, groupHist]);

  /* ---------- filtro de Mês/Ano específico da subaba Métricas — escopa todas as métricas abaixo
     (Due pela data de vencimento, History pela data de fechamento) ---------- */
  const [metricasFiltro, setMetricasFiltro] = useState({ mes: "Todos", ano: "Todos" });
  const [metricasView, setMetricasView] = useState("geral");
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
  /* planos vencidos para o detalhamento: pela coluna Diff (negativa) ou, sem Diff, pela data Due já passada */
  const planosVencidos = useMemo(() => {
    const d0 = new Date(); const hoje = `${d0.getFullYear()}-${String(d0.getMonth() + 1).padStart(2, "0")}-${String(d0.getDate()).padStart(2, "0")}`;
    const semDiff = (d) => d.diffValue === null || d.diffValue === undefined;
    return tmDueForMetrics
      .filter((d) => isVencida(d) || (semDiff(d) && d.dueDate && d.dueDate < hoje))
      .map((d) => {
        let atraso = "—", ord = 0;
        if (!semDiff(d)) { atraso = `${Math.abs(d.diffValue)} ${d.diffUnit === "H" ? "h" : "dias"}`; ord = d.diffUnit === "H" ? d.diffValue / 24 : d.diffValue; }
        else { const n = Math.round((new Date(hoje) - new Date(d.dueDate)) / 86400000); atraso = `${n} dias`; ord = -n; }
        return { ...d, _atraso: atraso, _ord: ord };
      })
      .sort((a, b) => a._ord - b._ord);
  }, [tmDueForMetrics]);
  /* planos postergados: um por plano (componente + job), contando quantas vezes foi postergado no período */
  const planosPostergados = useMemo(() => {
    const map = new Map();
    tmHistoryForMetrics.filter(isPostergada).forEach((h) => {
      const key = `${h.componentCode || ""}|${h.jobNo || ""}|${h.jobName || ""}`;
      const cur = map.get(key);
      if (!cur) map.set(key, { key, componentCode: h.componentCode, componentName: h.componentName, jobNo: h.jobNo, jobName: h.jobName, jobType: h.jobType, vezes: 1, ultimaData: h.dateDone || "", doneByName: h.doneByName });
      else {
        cur.vezes++;
        if ((h.dateDone || "") > cur.ultimaData) { cur.ultimaData = h.dateDone || ""; cur.doneByName = h.doneByName; }
      }
    });
    return [...map.values()].sort((a, b) => b.vezes - a.vezes || b.ultimaData.localeCompare(a.ultimaData));
  }, [tmHistoryForMetrics]);
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

  /* ---------- novas análises da Visão Geral ---------- */
  const hojeLocal = (() => { const d0 = new Date(); return `${d0.getFullYear()}-${String(d0.getMonth() + 1).padStart(2, "0")}-${String(d0.getDate()).padStart(2, "0")}`; })();
  const semDiffJob = (d) => d.diffValue === null || d.diffValue === undefined;
  const isVencidoPlano = (d) => isVencida(d) || (semDiffJob(d) && d.dueDate && d.dueDate < hojeLocal);
  const atrasoDias = (d) => (semDiffJob(d)
    ? (d.dueDate ? Math.max(0, Math.round((new Date(hojeLocal) - new Date(d.dueDate)) / 86400000)) : null)
    : (d.diffUnit === "D" ? Math.max(0, -d.diffValue) : null));
  const jtOf = (x) => String(x.jobType || "").trim().toUpperCase();
  const jobNoOf = (x) => String(x.jobNo === undefined || x.jobNo === null ? "" : x.jobNo).trim();
  const codeOfJob = (x) => String(x.code !== undefined ? x.code : (x.componentCode || "")).trim();
  const ymOf = (iso) => (iso || "").slice(0, 7);
  const ymLabel = (k) => `${MONTH_NAMES[Number(k.slice(5, 7)) - 1].slice(0, 3)}/${k.slice(2, 4)}`;
  const avgOf = (arr) => (arr.length ? Math.round(arr.reduce((a, v) => a + v, 0) / arr.length) : 0);

  /* por mês de vencimento: Due (abertos / vencidos / a vencer) + History (executados / postergados, pelo mês do Due date;
     sem Due date cai no mês em que foi fechado) */
  /* Due date do History só vale como referência se estiver a menos de 180 dias da data de fechamento (a planilha traz placeholders tipo 1900-01-01) */
  const dueValidoHist = (h) => {
    if (!h.dueDate || !h.dateDone) return "";
    const gap = (new Date(h.dueDate) - new Date(h.dateDone)) / 86400000;
    return gap >= -180 && gap <= 180 ? h.dueDate : "";
  };
  const monthlyJobs = (test, comHistory = true, histKey = null) => {
    const map = {};
    const row = (k) => (map[k] = map[k] || { mesKey: k, abertos: 0, vencidos: 0, aVencer: 0, executados: 0, postergados: 0 });
    tmDueForMetrics.filter(test).forEach((d) => {
      if (!d.dueDate || d.dueDate < "2020-01-01") return;
      const r = row(ymOf(d.dueDate)); r.abertos++;
      if (isVencidoPlano(d)) r.vencidos++; else r.aVencer++;
    });
    if (comHistory) tmHistoryForMetrics.filter(test).forEach((h) => {
      const k = histKey ? histKey(h) : ymOf(dueValidoHist(h) || h.dateDone); if (!k) return;
      const r = row(k);
      if (isPostergada(h)) r.postergados++; else r.executados++;
    });
    return Object.values(map).sort((a, b) => a.mesKey.localeCompare(b.mesKey))
      .map((r) => ({ ...r, mes: ymLabel(r.mesKey), previsto: r.abertos + r.executados + r.postergados, realizado: r.executados }));
  };
  const planejamentoMensal = useMemo(() => monthlyJobs(() => true, true, (h) => ymOf(h.dateDone)), [tmDueForMetrics, tmHistoryForMetrics, hojeLocal]);
  const corretivasMensal = useMemo(() => monthlyJobs(isCorretiva, true, (h) => ymOf(h.dateDone)), [tmDueForMetrics, tmHistoryForMetrics, hojeLocal]);
  const aderenciaMensal = useMemo(() => monthlyJobs(() => true), [tmDueForMetrics, tmHistoryForMetrics, hojeLocal]);
  const inspecaoMensal = useMemo(() => monthlyJobs((x) => jtOf(x) === "INS"), [tmDueForMetrics, tmHistoryForMetrics, hojeLocal]);
  const overhaulMensal = useMemo(() => monthlyJobs((x) => jtOf(x) === "OVH"), [tmDueForMetrics, tmHistoryForMetrics, hojeLocal]);
  const aderenciaDue = useMemo(() => {
    const prev = aderenciaMensal.reduce((a, r) => a + r.previsto, 0);
    const aver = aderenciaMensal.reduce((a, r) => a + r.aVencer, 0);
    const real = aderenciaMensal.reduce((a, r) => a + r.realizado, 0);
    const vencidoJa = prev - aver;
    return { pct: vencidoJa > 0 ? Math.round((real / vencidoJa) * 100) : 0 };
  }, [aderenciaMensal]);

  /* postergados: por mês (+ quantos já foram realizados depois da última postergação do mesmo plano) e por job type */
  const postergadosAnalise = useMemo(() => {
    const planKey = (h) => `${codeOfJob(h)}|${jtOf(h)}|${jobNoOf(h)}`;
    const lastPost = new Map(); const lastDone = new Map();
    tmHistory.forEach((h) => {
      if (!h.dateDone) return;
      const k = planKey(h);
      if (isPostergada(h)) { if (!lastPost.has(k) || h.dateDone > lastPost.get(k)) lastPost.set(k, h.dateDone); }
      else if (!lastDone.has(k) || h.dateDone > lastDone.get(k)) lastDone.set(k, h.dateDone);
    });
    /* planos que ainda estão no Due: em aberto, e quais deles já estão vencidos */
    const abertoPlano = new Map();
    tmDue.forEach((d) => { const k = planKey(d); const cur = abertoPlano.get(k) || { aberto: true, vencido: false }; if (isVencidoPlano(d)) cur.vencido = true; abertoPlano.set(k, cur); });
    const map = {};
    tmHistory.forEach((h) => {
      if (!h.dateDone || !isPostergada(h)) return;
      if (metricasFiltro.ano !== "Todos" && h.dateDone.slice(0, 4) !== String(metricasFiltro.ano)) return;
      if (metricasFiltro.mes !== "Todos" && Number(h.dateDone.slice(5, 7)) !== Number(metricasFiltro.mes)) return;
      const k = ymOf(h.dateDone); const pk = planKey(h);
      const r = (map[k] = map[k] || { total: 0, realizados: 0, emAberto: 0, vencidos: 0 });
      r.total++;
      if (lastDone.has(pk) && lastDone.get(pk) > lastPost.get(pk)) r.realizados++;
      const ab = abertoPlano.get(pk);
      if (ab) { r.emAberto++; if (ab.vencido) r.vencidos++; }
    });
    const rows = Object.keys(map).sort().map((k) => ({ mes: ymLabel(k), mesKey: k, ...map[k] }));
    const sm = (f) => rows.reduce((a, r) => a + r[f], 0);
    return { rows, total: sm("total"), realizados: sm("realizados"), emAberto: sm("emAberto"), vencidos: sm("vencidos") };
  }, [tmHistory, tmDue, metricasFiltro.ano, metricasFiltro.mes, hojeLocal]);
  /* backlog de jobs overdue (situação de hoje): em aberto no Due com vencimento passado, por mês de vencimento */
  const backlogOverdueMes = useMemo(() => {
    const m = {};
    tmDue.filter((d) => isVencidoPlano(d) && d.dueDate).forEach((d) => { const k = ymOf(d.dueDate); m[k] = (m[k] || 0) + 1; });
    return Object.keys(m).sort().map((k) => ({ mes: ymLabel(k), mesKey: k, total: m[k] }));
  }, [tmDue, hojeLocal]);
  const postergadosPorTipo = useMemo(() => {
    const map = {};
    tmHistoryForMetrics.filter(isPostergada).forEach((h) => { const k = jtOf(h) || "—"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([nome, total]) => ({ nome, total })).sort((a, b) => b.total - a.total).slice(0, 15);
  }, [tmHistoryForMetrics]);
  const postergadosDoMes = useMemo(() => {
    const key = hojeLocal.slice(0, 7);
    return { key, label: ymLabel(key), total: tmHistory.filter((h) => isPostergada(h) && ymOf(h.dateDone) === key).length };
  }, [tmHistory, hojeLocal]);

  /* cumprimento de prazo: Due date x Data assinada (só execuções, sem postergações) */
  const cumprimentoPrazo = useMemo(() => {
    const rows = tmHistoryForMetrics.filter((h) => !isPostergada(h) && dueValidoHist(h) && h.dateSigned)
      .map((h) => ({ h, atraso: Math.round((new Date(h.dateSigned) - new Date(h.dueDate)) / 86400000) }));
    const noPrazo = rows.filter((r) => r.atraso <= 0); const atrasados = rows.filter((r) => r.atraso > 0);
    const faixaDef = [["No prazo / antecipado", (v) => v <= 0], ["1 a 7 dias", (v) => v >= 1 && v <= 7], ["8 a 30 dias", (v) => v >= 8 && v <= 30], ["31 a 60 dias", (v) => v >= 31 && v <= 60], ["Mais de 60 dias", (v) => v > 60]];
    const faixas = faixaDef.map(([nome, f]) => ({ nome, total: rows.filter((r) => f(r.atraso)).length }));
    const mm = {};
    rows.forEach((r) => { const k = ymOf(r.h.dateSigned); const m = (mm[k] = mm[k] || { t: 0, n: 0 }); m.t++; if (r.atraso <= 0) m.n++; });
    const porMes = Object.keys(mm).sort().map((k) => ({ mes: ymLabel(k), pct: Math.round((mm[k].n / mm[k].t) * 100), total: mm[k].t }));
    return {
      total: rows.length, noPrazo: noPrazo.length, atrasados: atrasados.length,
      pct: rows.length ? Math.round((noPrazo.length / rows.length) * 100) : 0,
      atrasoMedio: avgOf(atrasados.map((r) => r.atraso)), antecedenciaMedia: avgOf(noPrazo.map((r) => -r.atraso)),
      faixas, porMes,
    };
  }, [tmHistoryForMetrics]);

  /* jobs críticos (pelo código do componente) */
  const criticosAnalise = useMemo(() => {
    const planoKey = (x) => `${codeOfJob(x)}|${jtOf(x)}|${jobNoOf(x)}`;
    const proxDue = new Map(); /* próximo Due em aberto de cada plano (o mais antigo) */
    tmDue.forEach((d) => {
      const k = planoKey(d); const cur = proxDue.get(k); const v = d.dueDate || "";
      if (!cur) proxDue.set(k, { dueDate: v, dueRaw: d.dueRaw || "" });
      else if (v && (!cur.dueDate || v < cur.dueDate)) proxDue.set(k, { dueDate: v, dueRaw: d.dueRaw || "" });
    });
    const due = tmDueForMetrics.filter((d) => critPrefixOf(d.code)).map((d) => {
      const venc = isVencidoPlano(d);
      const a = atrasoDias(d);
      return { ...d, _grupo: critPrefixOf(d.code), _sit: venc ? "Vencido" : isAteVencer40(d) ? "A vencer (até 40d)" : "No prazo",
        _atraso: venc ? (a === null ? "" : `${a} dias`) : "", _ord: venc ? -(a || 0) : (semDiffJob(d) ? 1e9 : d.diffValue) };
    }).sort((x, y) => x._ord - y._ord);
    const histCrit = tmHistoryForMetrics.filter((h) => critPrefixOf(h.componentCode));
    const postRegs = histCrit.filter(isPostergada);
    const execs = histCrit.filter((h) => !isPostergada(h));
    const map = new Map();
    postRegs.forEach((h) => {
      const key = planoKey(h);
      const cur = map.get(key);
      if (!cur) {
        const pd = proxDue.get(key);
        map.set(key, { key, _grupo: critPrefixOf(h.componentCode), componentCode: h.componentCode, componentName: h.componentName, jobNo: h.jobNo, jobName: h.jobName, jobType: h.jobType,
          primeiraData: h.dateDone || "", ultimaData: h.dateDone || "", registros: 1, proximoDue: pd ? pd.dueDate : "", proximoDueRaw: pd ? pd.dueRaw : "" });
      } else {
        cur.registros++;
        if ((h.dateDone || "") > cur.ultimaData) cur.ultimaData = h.dateDone || "";
        if (h.dateDone && (!cur.primeiraData || h.dateDone < cur.primeiraData)) cur.primeiraData = h.dateDone;
      }
    });
    const post = [...map.values()].sort((a, b) => b.ultimaData.localeCompare(a.ultimaData));
    const metricas = (dd, ex, ps) => {
      const comDatas = ex.filter((h) => dueValidoHist(h) && h.dateSigned);
      const noPrazoN = comDatas.filter((h) => h.dateSigned <= h.dueDate).length;
      const venc = dd.filter((r) => r._sit === "Vencido").length;
      const base = ex.length + ps.length + venc;
      return {
        vencidos: venc, aVencer: dd.filter((r) => r._sit === "A vencer (até 40d)").length, noPrazo: dd.filter((r) => r._sit === "No prazo").length,
        executados: ex.length, postergados: ps.length,
        pctPrazo: comDatas.length ? Math.round((noPrazoN / comDatas.length) * 100) : null,
        taxaPost: ex.length + ps.length ? Math.round((ps.length / (ex.length + ps.length)) * 100) : null,
        eficiencia: base ? Math.round((noPrazoN / base) * 100) : null, noPrazoN, comDatasN: comDatas.length, baseN: base,
      };
    };
    const porGrupo = CRITICAL_GROUPS.map((g) => ({
      grupo: g.cod, nome: critLabel(g.cod), crit: g.crit,
      ...metricas(due.filter((r) => r._grupo === g.cod), execs.filter((h) => critPrefixOf(h.componentCode) === g.cod), postRegs.filter((h) => critPrefixOf(h.componentCode) === g.cod)),
    }));
    const total = metricas(due, execs, postRegs);
    return {
      due, post, porGrupo,
      kpis: { abertos: due.length, vencidos: total.vencidos, aVencer: total.aVencer, noPrazo: total.noPrazo, postergados: post.length, registrosPost: postRegs.length,
        executados: total.executados, pctPrazo: total.pctPrazo === null ? 0 : total.pctPrazo, taxaPost: total.taxaPost === null ? 0 : total.taxaPost, eficiencia: total.eficiencia === null ? 0 : total.eficiencia },
    };
  }, [tmDue, tmDueForMetrics, tmHistoryForMetrics, hojeLocal]);

  /* jobs preditivos: cada job (tipo + nº) é contado uma vez por data, mesmo com vários componentes */
  const preditivosAnalise = useMemo(() => {
    const uniq = (arr, keyFn) => { const m = new Map(); arr.forEach((x) => { const k = keyFn(x); if (!m.has(k)) m.set(k, x); }); return [...m.values()]; };
    const grupos = [...new Set(PREDITIVOS.map((p) => p.grupo))];
    const rows = grupos.map((grupo) => {
      const jobs = PREDITIVOS.filter((p) => p.grupo === grupo);
      const isJob = (x) => jobs.some((p) => jtOf(x) === p.tipo && jobNoOf(x) === p.no);
      const due = uniq(tmDueForMetrics.filter(isJob), (d) => d.dueDate || d.dueRaw || d.id);
      const hist = uniq(tmHistoryForMetrics.filter(isJob), (h) => `${h.dateDone}|${isPostergada(h) ? "P" : "E"}`);
      const exec = hist.filter((h) => !isPostergada(h)); const post = hist.filter(isPostergada);
      const comDatas = exec.filter((h) => dueValidoHist(h) && h.dateSigned);
      const noPrazo = comDatas.filter((h) => h.dateSigned <= h.dueDate).length;
      const venc = due.filter(isVencidoPlano).length;
      const aVencer = due.filter((d) => !isVencidoPlano(d) && isAteVencer40(d)).length;
      const base = exec.length + post.length + venc;
      return {
        grupo, vencidos: venc, aVencer, executados: exec.length, postergados: post.length,
        noPrazoN: noPrazo, comDatasN: comDatas.length,
        pctPrazo: comDatas.length ? Math.round((noPrazo / comDatas.length) * 100) : null,
        taxaPost: exec.length + post.length ? Math.round((post.length / (exec.length + post.length)) * 100) : null,
        eficiencia: base ? Math.round((noPrazo / base) * 100) : null, baseN: base,
      };
    });
    const sum = (k) => rows.reduce((a, r) => a + r[k], 0);
    const cd = sum("comDatasN"); const base = sum("baseN");
    return { rows, kpis: { vencidos: sum("vencidos"), aVencer: sum("aVencer"), postergados: sum("postergados"),
      pctPrazo: cd ? Math.round((sum("noPrazoN") / cd) * 100) : 0, eficiencia: base ? Math.round((sum("noPrazoN") / base) * 100) : 0 } };
  }, [tmDueForMetrics, tmHistoryForMetrics, hojeLocal]);

  React.useEffect(() => {
    if (!setReportFn || tmSubTab === "falhas" || tmSubTab === "bordo") return;
    setReportFn(() => () => {
      const doc = new jsPDF();
      if (tmSubTab === "metricas" && metricasView === "planos") {
        const periodo = hasActiveFiltroMetricas
          ? `Período: ${metricasFiltro.mes === "Todos" ? "todos os meses" : MONTH_NAMES[Number(metricasFiltro.mes) - 1]} / ${metricasFiltro.ano === "Todos" ? "todos os anos" : metricasFiltro.ano}`
          : "Todos os períodos";
        let y = pdfHeader(doc, "TM Master — Planos Postergados e Vencidos",
          `${periodo} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
        y = pdfKpis(doc, y, [
          { label: "Planos Postergados", value: planosPostergados.length },
          { label: "Planos Vencidos", value: planosVencidos.length },
        ]);
        y = pdfSectionTitle(doc, y, `Planos Postergados (${planosPostergados.length})`);
        y = pdfTable(doc, y, ["Código", "Componente", "Job No", "Plano (Job Name)", "Tipo", "Última Postergação"],
          planosPostergados.map((r) => [r.componentCode || "—", r.componentName || "—", r.jobNo || "—", r.jobName || "—", r.jobType || "—", fmtDataCompleta(r.ultimaData)]));
        y = pdfSectionTitle(doc, y, `Planos Vencidos (${planosVencidos.length})`);
        y = pdfTable(doc, y, ["Código", "Componente", "Job No", "Plano (Job Name)", "Tipo", "Prioridade", "Due", "Atraso"],
          planosVencidos.map((r) => [r.code || "—", r.component || "—", r.jobNo || "—", r.jobName || "—", r.jobType || "—", r.pri || "—", r.dueDate ? fmtDataCompleta(r.dueDate) : (r.dueRaw || "—"), r._atraso]));
        pdfSave(doc, "relatorio-tm-master-planos-postergados-vencidos");
        return;
      }
      if (tmSubTab === "due") {
        const fv = filteredDue;
        let y = pdfHeader(doc, "TM Master — Due (em aberto)",
          `${fv.length} item(ns) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
        y = pdfKpis(doc, y, [
          { label: "Itens no Filtro", value: fv.length },
          { label: "Vencidas", value: fv.filter(isVencida).length },
          { label: "A Vencer em 40 dias", value: fv.filter(isAteVencer40).length },
          { label: "Críticas Vencidas/até 40d", value: fv.filter((d) => isCritica(d) && isVencidaOuAte40(d)).length },
          { label: "Corretivas (ONE)", value: fv.filter(isCorretiva).length },
        ]);
        y = pdfSectionTitle(doc, y, "Itens em aberto");
        pdfTable(doc, y, ["Departamento", "Código", "Componente", "Job No", "Plano (Job Name)", "Tipo", "Pri", "Due", "Situação"],
          fv.map((d) => [d.department || "—", d.code || "—", d.component || "—", d.jobNo || "—", d.jobName || "—", d.jobType || "—", d.pri || "—",
            d.dueDate ? fmtDataCompleta(d.dueDate) : (d.dueRaw || "—"),
            d.diffValue === null || d.diffValue === undefined ? (d.diffRaw || "—")
              : isVencida(d) ? `vencida há ${Math.abs(d.diffValue)} ${d.diffUnit === "H" ? "h" : "dias"}` : `vence em ${d.diffValue} ${d.diffUnit === "H" ? "h" : "dias"}`]));
        pdfSave(doc, "relatorio-tm-master-due");
      }
      if (tmSubTab === "history") {
        const fh = filteredHistory;
        let y = pdfHeader(doc, "TM Master — History (fechados)",
          `${fh.length} item(ns) no filtro atual · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
        y = pdfKpis(doc, y, [
          { label: "Itens no Filtro", value: fh.length },
          { label: "Corretivas Fechadas", value: fh.filter(isCorretiva).length },
          { label: "Postergadas", value: fh.filter(isPostergada).length },
        ]);
        y = pdfSectionTitle(doc, y, "Itens fechados");
        pdfTable(doc, y, ["Data", "Componente", "Job No", "Plano (Job Name)", "Tipo", "Executado Por", "Remarks", "Situação"],
          fh.map((h) => [fmtDataCompleta(h.dateDone), h.componentName || "—", h.jobNo || "—", h.jobName || "—", h.jobType || "—", h.doneByName || "—", h.remarks || "—", isPostergada(h) ? "Postergada" : "Fechada"]));
        pdfSave(doc, "relatorio-tm-master-history");
      }
      if (tmSubTab === "metricas" && metricasView === "geral") {
        const periodo = hasActiveFiltroMetricas
          ? `Período: ${metricasFiltro.mes === "Todos" ? "todos os meses" : MONTH_NAMES[Number(metricasFiltro.mes) - 1]} / ${metricasFiltro.ano === "Todos" ? "todos os anos" : metricasFiltro.ano}`
          : "Todos os períodos";
        let y = pdfHeader(doc, "TM Master — Dashboard", `${periodo} · Gerado em ${new Date().toLocaleDateString("pt-BR")}`);
        y = pdfKpis(doc, y, [
          { label: "Total de Jobs", value: totalDue },
          { label: "Vencidas", value: vencidas.length },
          { label: "A Vencer em 40 dias", value: ateVencer40.length },
          { label: "Críticas Vencidas/até 40d", value: criticasVencidasOu40.length },
          { label: "Corretivas (ONE)", value: corretivasDue.length },
          { label: "Total Fechado", value: totalHistory },
          { label: "Corretivas Fechadas", value: corretivasFechadas.length },
          { label: "Postergadas", value: postergadas.length },
          { label: "Postergados no Mês Atual", value: postergadosDoMes.total },
          { label: "% No Prazo (Due × Assinada)", value: `${cumprimentoPrazo.pct}%` },
        ]);
        const secao = (titulo, cols, rows) => { if (y > 235) { doc.addPage(); y = 15; } y = pdfSectionTitle(doc, y, titulo); y = pdfTable(doc, y, cols, rows); };
        secao("Planejamento mensal", ["Mês", "Abertos", "Vencidos", "A Vencer", "Fechados", "Postergados"], planejamentoMensal.map((r) => [r.mes, r.abertos, r.vencidos, r.aVencer, r.executados, r.postergados]));
        secao("Corretivas (Job Type = One)", ["Mês", "Abertos", "Vencidos", "A Vencer", "Fechados"], corretivasMensal.map((r) => [r.mes, r.abertos, r.vencidos, r.aVencer, r.executados]));
        secao("Postergados por mês", ["Mês", "Postergados", "Realizados após", "Ainda em aberto", "Vencidos"], postergadosAnalise.rows.map((r) => [r.mes, r.total, r.realizados, r.emAberto, r.vencidos]));
        secao("Postergados por Job Type", ["Job Type", "Postergados"], postergadosPorTipo.map((r) => [r.nome, r.total]));
        secao("Aderência no Due", ["Mês", "Previsto", "Realizado", "A Vencer"], aderenciaMensal.map((r) => [r.mes, r.previsto, r.realizado, r.aVencer]));
        secao("Jobs de Inspeção", ["Mês", "Previstos", "Executados", "Postergados"], inspecaoMensal.map((r) => [r.mes, r.previsto, r.realizado, r.postergados]));
        secao("Jobs de Overhaul", ["Mês", "Previstos", "Executados", "Postergados"], overhaulMensal.map((r) => [r.mes, r.previsto, r.realizado, r.postergados]));
        secao(`Jobs críticos — ${criticosAnalise.kpis.abertos} em aberto · ${criticosAnalise.kpis.vencidos} vencidos · ${criticosAnalise.kpis.aVencer} a vencer (40d) · ${criticosAnalise.kpis.postergados} postergados · eficiência ${criticosAnalise.kpis.eficiencia}%`,
          ["Grupo", "Criticidade", "Em atraso", "A vencer", "No prazo", "Executados", "Postergados", "% no prazo", "Taxa postergação", "Eficiência"],
          criticosAnalise.porGrupo.filter((r) => r.vencidos + r.aVencer + r.noPrazo + r.executados + r.postergados > 0).map((r) => [r.nome, r.crit, r.vencidos, r.aVencer, r.noPrazo, r.executados, r.postergados, r.pctPrazo === null ? "—" : `${r.pctPrazo}%`, r.taxaPost === null ? "—" : `${r.taxaPost}%`, r.eficiencia === null ? "—" : `${r.eficiencia}%`]));
        secao(`Jobs preditivos — eficiência ${preditivosAnalise.kpis.eficiencia}%`,
          ["Análise", "Em atraso", "A vencer", "Executados", "Postergados", "% no prazo", "Eficiência"],
          preditivosAnalise.rows.map((r) => [r.grupo, r.vencidos, r.aVencer, r.executados, r.postergados, r.pctPrazo === null ? "—" : `${r.pctPrazo}%`, r.eficiencia === null ? "—" : `${r.eficiencia}%`]));
        secao(`Cumprimento de prazo — ${cumprimentoPrazo.pct}% no prazo · atraso médio ${cumprimentoPrazo.atrasoMedio}d`, ["Faixa", "Jobs"], cumprimentoPrazo.faixas.map((r) => [r.nome, r.total]));
        y = pdfSectionTitle(doc, y, "Due por departamento");
        y = pdfTable(doc, y, ["Departamento", "Quantidade"], duePorDepartamento.map((d) => [d.departamento, d.count]));
        y = pdfSectionTitle(doc, y, "Vencidas por tipo de serviço");
        y = pdfTable(doc, y, ["Tipo", "Quantidade"], vencidasPorJobType.map((d) => [d.tipo, d.count]));
        y = pdfSectionTitle(doc, y, "A vencer (até 40 dias) por tipo de serviço");
        y = pdfTable(doc, y, ["Tipo", "Quantidade"], aVencerPorJobType.map((d) => [d.tipo, d.count]));
        y = pdfSectionTitle(doc, y, "Fechados por departamento");
        y = pdfTable(doc, y, ["Departamento", "Quantidade"], historyPorDepartamento.map((d) => [d.departamento, d.count]));
        y = pdfSectionTitle(doc, y, "Fechados por usuário");
        y = pdfTable(doc, y, ["Usuário", "Quantidade"], historyPorUsuario.map((d) => [d.usuario, d.count]));
        pdfSave(doc, "relatorio-tm-master-metricas");
      }
    });
  }, [tmSubTab, filteredDue, totalDue, vencidas, ateVencer40, criticasVencidasOu40, corretivasDue, duePorDepartamento,
      vencidasPorJobType, aVencerPorJobType, filteredHistory, totalHistory, corretivasFechadas, postergadas, prazoStats,
      historyPorDepartamento, historyPorUsuario, metricasView, planosPostergados, planosVencidos, metricasFiltro, hasActiveFiltroMetricas,
      postergadosDoMes, cumprimentoPrazo, planejamentoMensal, aderenciaMensal, corretivasMensal, postergadosAnalise, postergadosPorTipo, inspecaoMensal, overhaulMensal,
      criticosAnalise, preditivosAnalise, setReportFn]);

  /* "Exportar planilha" no cabeçalho segue a subaba/filtro ativo aqui — Due exporta o filtro de Due,
     History o filtro de History, e Métricas exporta os dois conjuntos já escopados por mês/ano */
  React.useEffect(() => {
    if (!setExportXlsxFn || tmSubTab === "falhas" || tmSubTab === "bordo") return;
    setExportXlsxFn(() => () => {
      const wb = XLSX.utils.book_new();
      if (tmSubTab === "due") {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(filteredDue, TM_DUE_COLS)), "TM Master - Due");
        XLSX.writeFile(wb, `genesis-tm-master-due-${todayISO()}.xlsx`);
      } else if (tmSubTab === "history") {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(filteredHistory, TM_HISTORY_COLS)), "TM Master - History");
        XLSX.writeFile(wb, `genesis-tm-master-history-${todayISO()}.xlsx`);
      } else {
        if (metricasView !== "planos") {
          const add = (nome, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), nome);
          add("Planejamento Mensal", planejamentoMensal.map((r) => ({ "Mês": r.mes, Abertos: r.abertos, Vencidos: r.vencidos, "A Vencer": r.aVencer, Fechados: r.executados, Postergados: r.postergados })));
          add("Aderência no Due", aderenciaMensal.map((r) => ({ "Mês": r.mes, Previsto: r.previsto, Realizado: r.realizado, "A Vencer": r.aVencer })));
          add("Corretivas Mensal", corretivasMensal.map((r) => ({ "Mês": r.mes, Abertos: r.abertos, Vencidos: r.vencidos, "A Vencer": r.aVencer, Fechados: r.executados })));
          add("Postergados por Mês", postergadosAnalise.rows.map((r) => ({ "Mês": r.mes, Postergados: r.total, "Realizados após a postergação": r.realizados, "Ainda em aberto": r.emAberto, Vencidos: r.vencidos })));
          add("Postergados por Job Type", postergadosPorTipo.map((r) => ({ "Job Type": r.nome, Postergados: r.total })));
          add("Jobs de Inspeção", inspecaoMensal.map((r) => ({ "Mês": r.mes, Previstos: r.previsto, Executados: r.realizado, Postergados: r.postergados })));
          add("Jobs de Overhaul", overhaulMensal.map((r) => ({ "Mês": r.mes, Previstos: r.previsto, Executados: r.realizado, Postergados: r.postergados })));
          add("Jobs Críticos por Grupo", criticosAnalise.porGrupo.map((r) => ({ Grupo: r.nome, Criticidade: r.crit, "Em Atraso": r.vencidos, "A Vencer": r.aVencer, "No Prazo": r.noPrazo, Executados: r.executados, Postergados: r.postergados, "% No Prazo": r.pctPrazo, "Taxa de Postergação %": r.taxaPost, "Eficiência %": r.eficiencia })));
          add("Jobs Críticos (abertos)", criticosAnalise.due.map((r) => ({ Grupo: critLabel(r._grupo), Criticidade: critOf(r._grupo), "Código": r.code, Componente: r.component, "Job No": r.jobNo, "Job Type": r.jobType, Plano: r.jobName, Due: r.dueDate || r.dueRaw, "Situação": r._sit, Atraso: r._atraso })));
          add("Jobs Críticos (postergados)", criticosAnalise.post.map((r) => ({ Grupo: critLabel(r._grupo), Criticidade: critOf(r._grupo), "Código": r.componentCode, Componente: r.componentName, "Job No": r.jobNo, "Job Type": r.jobType, Plano: r.jobName, "1ª Postergação": r.primeiraData, "Última Postergação": r.ultimaData, "Próximo Due Date": r.proximoDue || r.proximoDueRaw })));
          add("Jobs Preditivos", preditivosAnalise.rows.map((r) => ({ "Análise": r.grupo, "Em Atraso": r.vencidos, "A Vencer": r.aVencer, Executados: r.executados, Postergados: r.postergados, "% No Prazo": r.pctPrazo, "Taxa de Postergação %": r.taxaPost, "Eficiência %": r.eficiencia })));
          add("Cumprimento de Prazo", cumprimentoPrazo.faixas.map((r) => ({ Faixa: r.nome, Jobs: r.total })));
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(tmDueForMetrics, TM_DUE_COLS)), "Due (filtrado)");
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsToSheet(tmHistoryForMetrics, TM_HISTORY_COLS)), "History (filtrado)");
        }
        if (metricasView === "planos") {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(planosPostergados.map((r) => ({
          "Código": r.componentCode, "Componente": r.componentName, "Job No": r.jobNo, "Plano (Job Name)": r.jobName, "Tipo": r.jobType,
          "Última Postergação": r.ultimaData }))), "Planos Postergados");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(planosVencidos.map((r) => ({
          "Código": r.code, "Componente": r.component, "Job No": r.jobNo, "Plano (Job Name)": r.jobName,
          "Tipo": r.jobType, "Prioridade": r.pri, "Due": r.dueDate || r.dueRaw, "Atraso": r._atraso }))), "Planos Vencidos");
        }
        XLSX.writeFile(wb, `genesis-tm-master-${metricasView === "planos" ? "planos-postergados-vencidos" : "metricas"}-${todayISO()}.xlsx`);
      }
    });
  }, [tmSubTab, filteredDue, filteredHistory, tmDueForMetrics, tmHistoryForMetrics, planosPostergados, planosVencidos, metricasView, planejamentoMensal, aderenciaMensal, corretivasMensal, postergadosAnalise, postergadosPorTipo,
      inspecaoMensal, overhaulMensal, criticosAnalise, preditivosAnalise, cumprimentoPrazo, setExportXlsxFn]);

  return (
    <>
      <div className="g-mode-toggle" style={{ marginBottom: 16, width: "fit-content" }}>
        <button className={tmSubTab === "due" ? "active" : ""} onClick={() => setTmSubTab("due")}>Due (Em Aberto)</button>
        <button className={tmSubTab === "history" ? "active" : ""} onClick={() => setTmSubTab("history")}>History (Fechados)</button>
        <button className={tmSubTab === "falhas" ? "active" : ""} onClick={() => setTmSubTab("falhas")}>Failure Reports</button>
        <button className={tmSubTab === "bordo" ? "active" : ""} onClick={() => setTmSubTab("bordo")}>Registro de Bordo</button>
        <button className={tmSubTab === "metricas" ? "active" : ""} onClick={() => setTmSubTab("metricas")}>Dashboard</button>
      </div>

      <ResumoEmailTm tmDue={tmDue} tmHistory={tmHistory} tmFailures={tmFailures} />

      {tmSubTab === "bordo" && <BordoRegistroView tmBordo={tmBordo || []} setTmBordo={setTmBordo} setBordoDeleted={setBordoDeleted} regObrig={regObrig} setRegObrig={setRegObrig} currentUser={currentUser} tmDue={tmDue} tmHistory={tmHistory} tmFailures={tmFailures} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} active />}

      {tmSubTab === "falhas" && <FailureReportsView tmFailures={tmFailures} handleImportFailures={handleImportFailures} setReportFn={setReportFn} setExportXlsxFn={setExportXlsxFn} active />}

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
                  <th style={{ minWidth: 130 }}>PO</th>
                  <th style={{ minWidth: 130 }}>Planejamento</th>
                </tr>
              </thead>
              <tbody>
                {dueGroups.map((g) => (
                  <React.Fragment key={g.key}>
                    {groupDue && (
                      <tr className="tbl-group" onClick={() => setCollapsedDue((p) => { const n = new Set(p); n.has(g.key) ? n.delete(g.key) : n.add(g.key); return n; })}>
                        <td colSpan={8}>
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
                    <td style={{ minWidth: 130, verticalAlign: "top" }}>
                      <EText value={d.po || ""} onChange={(v) => updTmDue(i, "po", v)} mono />
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
              <label>Mês de fechamento</label>
              <FiltroMeses meses={histMeses} sel={histf.meses || []} onChange={(v) => setHistf((p) => ({ ...p, meses: v }))} label={histMesLabel} />
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
              <button className="g-btn" onClick={() => setHistf({ busca: "", jobType: "Todos", department: "Todos", usuario: "Todos", soPostergadas: false, meses: [] })}
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
            {filteredHistory.length > 0 && (
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginTop: 12 }}>
                <span className="g-muted" style={{ fontSize: 12 }}>
                  Mostrando {(histPagAtual - 1) * histPorPag + 1}–{Math.min(histPagAtual * histPorPag, filteredHistory.length)} de {filteredHistory.length}
                </span>
                <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                  <button className="g-btn g-keep" disabled={histPagAtual <= 1} onClick={() => setHistPag(1)}>«</button>
                  <button className="g-btn g-keep" disabled={histPagAtual <= 1} onClick={() => setHistPag(histPagAtual - 1)}>‹ Anterior</button>
                  <span style={{ fontSize: 12.5, padding: "0 6px" }}>Página <strong>{histPagAtual}</strong> de {histTotalPag}</span>
                  <button className="g-btn g-keep" disabled={histPagAtual >= histTotalPag} onClick={() => setHistPag(histPagAtual + 1)}>Próxima ›</button>
                  <button className="g-btn g-keep" disabled={histPagAtual >= histTotalPag} onClick={() => setHistPag(histTotalPag)}>»</button>
                  <select value={histPorPag} onChange={(e) => setHistPorPag(Number(e.target.value))} style={{ marginLeft: 8 }}>
                    {[25, 50, 100, 200].map((n) => <option key={n} value={n}>{n} por página</option>)}
                  </select>
                </div>
              </div>
            )}
          </div>
        </>
      )}

      {tmSubTab === "metricas" && (
        <>
          <div className="g-mode-toggle" style={{ marginBottom: 14, width: "fit-content" }}>
            <button className={metricasView === "geral" ? "active" : ""} onClick={() => setMetricasView("geral")}>Visão Geral</button>
            <button className={metricasView === "planos" ? "active" : ""} onClick={() => setMetricasView("planos")}>Planos Postergados e Vencidos</button>
          </div>

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

          {metricasView === "planos" && (
            <PlanosAtencao postergados={planosPostergados} vencidos={planosVencidos} totalRegistrosPost={tmHistoryForMetrics.filter(isPostergada).length} />
          )}

          {metricasView === "geral" && (
          <>
          <MesBarChart title="Planejamento Mensal" data={planejamentoMensal}
            subtitle="Abertos, vencidos e a vencer: jobs do Due pelo mês de vencimento. Fechados e Postergados: jobs do History pelo mês em que foram fechados/postergados (respeitam o filtro de mês/ano)."
            bars={[{ key: "abertos", name: "Abertos", color: "#94A3B8" }, { key: "vencidos", name: "Vencidos", color: "#EF4444" }, { key: "aVencer", name: "A Vencer", color: "#F5A623" }, { key: "executados", name: "Fechados", color: "#22C55E" }, { key: "postergados", name: "Postergados", color: "#8B5CF6" }] } />

          <MesBarChart title="Corretivas (Job Type = One)" data={corretivasMensal}
            subtitle="Corretivas (Job type ONE). Abertos, vencidos e a vencer pelo mês de vencimento; fechados pelo mês de fechamento (sem contar postergações)."
            bars={[{ key: "abertos", name: "Abertos", color: "#94A3B8" }, { key: "vencidos", name: "Vencidos", color: "#EF4444" }, { key: "aVencer", name: "A Vencer", color: "#F5A623" }, { key: "executados", name: "Fechados", color: "#22C55E" }] } />

          <PostergadosAnalise analise={postergadosAnalise} porTipo={postergadosPorTipo} />

          <MesBarChart title="Backlog de Jobs em Overdue por Mês" height={240} data={backlogOverdueMes}
            subtitle="Jobs em aberto no Due com vencimento já passado, agrupados pelo mês de vencimento (situação de hoje; não segue o filtro de mês/ano)."
            bars={[{ key: "total", name: "Overdue em aberto", color: "#EF4444" }]} />

          <div className="g-grid-2">
            <MesBarChart title="Jobs de Inspeção" data={inspecaoMensal}
              subtitle="Job Type INS: previstos no mês (pelo Due date), executados e postergados."
              bars={[{ key: "previsto", name: "Previstos", color: "#94A3B8" }, { key: "realizado", name: "Executados", color: "#22C55E" }, { key: "postergados", name: "Postergados", color: "#F5A623" }]} />
            <MesBarChart title="Jobs de Overhaul" data={overhaulMensal}
              subtitle="Job Type OVH: previstos no mês (pelo Due date), executados e postergados."
              bars={[{ key: "previsto", name: "Previstos", color: "#94A3B8" }, { key: "realizado", name: "Executados", color: "#22C55E" }, { key: "postergados", name: "Postergados", color: "#F5A623" }]} />
          </div>

          <CriticosPanel kpis={criticosAnalise.kpis} porGrupo={criticosAnalise.porGrupo} due={criticosAnalise.due} post={criticosAnalise.post} />

          <PreditivosPanel rows={preditivosAnalise.rows} kpis={preditivosAnalise.kpis} />

          <div className="g-section-label">Due — o que está em aberto</div>
          <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(6, 1fr)" }}>
            {bigKpi("Total de Jobs", totalDue, "var(--teal)", ClipboardList)}
            {bigKpi("Vencidas", vencidas.length, "var(--crit)", AlertTriangle)}
            {bigKpi("A Vencer em 40 dias", ateVencer40.length, "var(--warn)", Clock)}
            {bigKpi("Críticas Vencidas/≤40d", criticasVencidasOu40.length, "var(--crit)", AlertTriangle)}
            {bigKpi("Corretivas (ONE)", corretivasDue.length, "var(--text-dim)", Wrench)}
            {bigKpi("Postergados no Mês Atual", postergadosDoMes.total, "var(--warn)", AlertTriangle)}
          </div>

          <div className="g-grid-2">
            <MesBarChart title="Aderência no Due" data={aderenciaMensal}
              headRight={<span className="g-muted" style={{ fontSize: 12 }}>Aderência: <b>{aderenciaDue.pct}%</b></span>}
              subtitle="Por mês de vencimento: o que estava previsto, quanto foi realizado e quanto ainda está a vencer."
              nota="Aderência = realizado ÷ (previsto − a vencer), ou seja, do que já venceu, quanto foi feito. Previsto = abertos + executados + postergados."
              bars={[{ key: "previsto", name: "Previsto", color: "#94A3B8" }, { key: "realizado", name: "Realizado", color: "#22C55E" }, { key: "aVencer", name: "A Vencer", color: "#F5A623" }]} />

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
            {bigKpi("% No Prazo (Due × Assinada)", `${cumprimentoPrazo.pct}%`, cumprimentoPrazo.pct >= 80 ? "var(--ok)" : cumprimentoPrazo.pct >= 60 ? "var(--warn)" : "var(--crit)", Clock)}
          </div>

          <div className="g-panel">
            <div className="g-panel-head"><span className="g-panel-title">Fechados por Mês</span></div>
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

          <CumprimentoPrazoPanel c={cumprimentoPrazo} />

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

          </>
          )}
        </>
      )}
    </>
  );
}

/* ============================================================
   MATERIALS — fully editable, including ID
   ============================================================ */

/* filtro de meses com seleção múltipla (caixinhas) */
function FiltroMeses({ meses, sel, onChange, label }) {
  const [aberto, setAberto] = React.useState(false);
  const ref = React.useRef(null);
  React.useEffect(() => {
    if (!aberto) return;
    const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setAberto(false); };
    document.addEventListener("mousedown", h); return () => document.removeEventListener("mousedown", h);
  }, [aberto]);
  const alterna = (m) => onChange(sel.includes(m) ? sel.filter((x) => x !== m) : [...sel, m]);
  const resumo = sel.length === 0 ? "Todos" : sel.length <= 2 ? [...sel].sort().map(label).join(", ") : `${sel.length} meses`;
  const anos = [...new Set(meses.map((m) => m.slice(0, 4)))];
  return (
    <div ref={ref} style={{ position: "relative", display: "inline-block" }}>
      <button type="button" className="g-keep" onClick={() => setAberto((a) => !a)} style={{ padding: "6px 10px", border: "1px solid var(--border)", borderRadius: 6, background: "var(--panel, #fff)", color: "var(--text)", fontSize: 12.5, cursor: "pointer", minWidth: 150, textAlign: "left", display: "flex", justifyContent: "space-between", gap: 8, ...(sel.length ? { borderColor: "var(--teal)", fontWeight: 600 } : {}) }}>
        <span>{resumo}</span><span style={{ opacity: .6, fontSize: 10 }}>▾</span>
      </button>
      {aberto && (
        <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, zIndex: 50, minWidth: 210, background: "var(--panel, #fff)", border: "1px solid var(--border)", borderRadius: 8, boxShadow: "0 8px 24px rgba(0,0,0,.25)", padding: 8 }}>
          <div style={{ display: "flex", gap: 10, fontSize: 11.5, marginBottom: 6, flexWrap: "wrap" }}>
            <span style={{ cursor: "pointer", textDecoration: "underline" }} onClick={() => onChange([...meses])}>Selecionar todos</span>
            {anos.map((a) => <span key={a} style={{ cursor: "pointer", textDecoration: "underline" }} onClick={() => onChange(meses.filter((m) => m.startsWith(a)))}>Ano {a}</span>)}
            <span style={{ cursor: "pointer", textDecoration: "underline" }} onClick={() => onChange([])}>Limpar</span>
          </div>
          <div style={{ maxHeight: 260, overflowY: "auto" }}>
            {meses.length === 0 && <div className="g-muted" style={{ fontSize: 12 }}>Sem meses</div>}
            {meses.map((m) => (
              <label key={m} style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12.5, padding: "4px 4px", cursor: "pointer" }}>
                <input type="checkbox" checked={sel.includes(m)} onChange={() => alterna(m)} />{label(m)}
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

/* campo de texto que cresce em altura e quebra linha, para mostrar tudo que está escrito */
function TextoAuto({ value, onChange, style, placeholder }) {
  const ref = React.useRef(null);
  const ajusta = () => { const el = ref.current; if (el) { el.style.height = "auto"; el.style.height = el.scrollHeight + 2 + "px"; } };
  React.useLayoutEffect(ajusta, [value]);
  return <textarea ref={ref} rows={1} value={value || ""} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} style={{ ...style, resize: "none", overflow: "hidden", whiteSpace: "pre-wrap", wordBreak: "break-word", lineHeight: 1.35, fontFamily: "inherit", display: "block" }} />;
}

/* filtro em "bloquinho": clica, digita e a lista de opções vai sendo filtrada */
function FiltroBusca({ value, onChange, options, todos, style }) {
  const [aberto, setAberto] = React.useState(false);
  const [q, setQ] = React.useState("");
  const ref = React.useRef(null);
  React.useEffect(() => {
    if (!aberto) return;
    const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setAberto(false); };
    document.addEventListener("mousedown", h); return () => document.removeEventListener("mousedown", h);
  }, [aberto]);
  const ops = [["Todos", todos], ...options.map((o) => [o, o])];
  const norm = (t) => String(t).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const lista = ops.filter(([, l]) => !q || norm(l).includes(norm(q)));
  const atual = (ops.find(([v]) => v === value) || ops[0])[1];
  const escolher = (v) => { onChange(v); setAberto(false); setQ(""); };
  return (
    <div ref={ref} style={{ position: "relative", display: "inline-block" }}>
      <button type="button" className="g-keep" onClick={() => setAberto((a) => !a)} style={{ ...style, cursor: "pointer", minWidth: 150, textAlign: "left", display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center", ...(value !== "Todos" ? { borderColor: "var(--teal)", fontWeight: 600 } : {}) }}>
        <span>{atual}</span><span style={{ opacity: .6, fontSize: 10 }}>▾</span>
      </button>
      {aberto && (
        <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, zIndex: 50, minWidth: 220, background: "var(--panel, #fff)", border: "1px solid var(--border)", borderRadius: 8, boxShadow: "0 8px 24px rgba(0,0,0,.25)", padding: 6 }}>
          <input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="digite para filtrar..." style={{ ...style, width: "100%", boxSizing: "border-box", marginBottom: 6 }}
            onKeyDown={(e) => { if (e.key === "Enter" && lista[0]) escolher(lista[0][0]); if (e.key === "Escape") setAberto(false); }} />
          <div style={{ maxHeight: 240, overflowY: "auto" }}>
            {lista.length === 0 && <div className="g-muted" style={{ fontSize: 12, padding: 6 }}>Nada encontrado</div>}
            {lista.map(([v, l]) => (
              <div key={v} onClick={() => escolher(v)} style={{ padding: "6px 8px", fontSize: 12.5, cursor: "pointer", borderRadius: 5, background: v === value ? "rgba(20,184,166,.15)" : "transparent", fontWeight: v === value ? 600 : 400 }}
                onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(128,128,128,.18)")} onMouseLeave={(e) => (e.currentTarget.style.background = v === value ? "rgba(20,184,166,.15)" : "transparent")}>{l}</div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function MaterialsView({ setExportXlsxFn, setReportFn, matReqs, setMatReqs, setMatDeleted, currentUser }) {
  const admin = isAdminUser(currentUser);
  const meusDeptos = deptsDoUsuario(currentUser);
  const almox = meusDeptos.includes("Almoxarifado");
  const verTodos = admin || almox;
  const pode = (d) => podeEditarMat(currentUser, d);
  const deptsEditaveis = admin || almox ? MAT_DEPTS : meusDeptos;
  const podeCriar = deptsEditaveis.length > 0;

  const [f, setF] = useState({ busca: "", dept: "Todos", status: "Todos", proc: "Todos", prio: "Todos", meu: false });
  const [novo, setNovo] = useState(null);
  const [msg, setMsg] = useState("");
  const [pag, setPag] = useState(0);
  const POR_PAG = 50;
  const rascunho = () => ({ departamento: deptsEditaveis.length === 1 ? deptsEditaveis[0] : "", tmMaster: "", priority: "Média", sap: "", descricao: "", quantidade: 1, justificativa: "", status: "Em Processo de Compras", statusProcesso: "Aberto", oe: "", dataSolicitacao: todayISO(), reserva: "", rc: "", po: "", linhaPo: "", valor: "", eta: "", dataRecebimento: "", obs: "" });

  const upd = (id, patch) => setMatReqs((rows) => rows.map((r) => (r.id === id && pode(r.departamento) ? { ...r, ...patch, upd: Date.now() } : r)));
  const setStatus = (r, v) => {
    const patch = { status: v };
    if (v === "Recebido" && !r.dataRecebimento) patch.dataRecebimento = todayISO();
    if (["Recebido", "Cancelado", "Recusado"].includes(v)) patch.statusProcesso = "Fechado";
    upd(r.id, patch);
  };
  const remover = (r) => {
    if (!window.confirm(`Excluir a requisição "${r.descricao || r.sap || r.id}"?`)) return;
    setMatReqs((rows) => rows.filter((x) => x.id !== r.id));
    setMatDeleted((d) => [...new Set([...(d || []), r.id])]);
  };
  const salvarNovo = () => {
    if (!novo.departamento) { setMsg("Escolha o departamento."); return; }
    if (!String(novo.descricao).trim()) { setMsg("Preencha a descrição."); return; }
    const r = { ...novo, id: uid("MAT"), quantidade: Number(novo.quantidade) || 0, valor: novo.valor === "" ? "" : Number(novo.valor), upd: Date.now() };
    setMatReqs((rows) => [r, ...rows]);
    setNovo(null); setMsg("Requisição adicionada.");
    setTimeout(() => setMsg(""), 4000);
  };

  /* ---------- modelo e importação em massa (Excel) ---------- */
  const fileRef = useRef(null);
  const [impMsg, setImpMsg] = useState(null);
  const nrm = (x) => String(x ?? "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, " ").trim();
  const baixarModelo = () => {
    const wb = XLSX.utils.book_new();
    const cols = MAT2_COLS.filter(([k]) => k !== "id");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([cols.map(([, l]) => l)]), "Materiais");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ["Como preencher"],
      ["Preencha a aba Materiais, uma requisição por linha, e importe no portal. Só Departamento e Descrição são obrigatórios."],
      [`Você só consegue importar linhas do seu departamento (${deptsEditaveis.join(", ") || "nenhum"}). Linhas de outros departamentos são ignoradas.`],
      ["Datas: dd/mm/aaaa. Valor Total e Quantidade: números."],
      [],
      ["Departamento"], ...MAT_DEPTS.map((d) => [d]), [],
      ["Status"], ...MAT_STATUS2.map((d) => [d]), [],
      ["Status do Processo"], ...MAT_PROC.map((d) => [d]), [],
      ["Prioridade"], ...PRIORITY.map((d) => [d]),
    ]), "Instruções");
    XLSX.writeFile(wb, "modelo-materiais.xlsx");
  };
  const aISO = (v) => {
    if (v === "" || v == null) return "";
    if (v instanceof Date) return new Date(v.getTime() + 12 * 3600 * 1000).toISOString().slice(0, 10);
    const t = String(v).trim();
    let m = t.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
    if (m) return `${m[3].length === 2 ? "20" + m[3] : m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
    m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? `${m[1]}-${m[2]}-${m[3]}` : "";
  };
  const aNum = (v) => {
    if (v === "" || v == null) return "";
    if (typeof v === "number") return v;
    let t = String(v).replace(/[^0-9,.-]/g, "");
    if (t.includes(",")) t = t.replace(/\./g, "").replace(",", ".");
    const n = Number(t); return Number.isFinite(n) ? n : "";
  };
  const importar = (e) => {
    const file = e.target.files?.[0]; e.target.value = "";
    if (!file) return;
    setImpMsg({ ok: true, txt: "Lendo a planilha… (arquivos grandes levam alguns segundos)" });
    const reader = new FileReader();
    reader.onload = (evt) => setTimeout(() => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true, dense: true });
        const aba = wb.SheetNames.find((n) => nrm(n) === "materiais") || wb.SheetNames[0];
        const json = XLSX.utils.sheet_to_json(wb.Sheets[aba], { defval: "", blankrows: false });
        const ALIAS = { linhaPo: ["linha da po", "item po"], valor: ["valor", "valor total"], priority: ["priority", "prioridade"], descricao: ["descricao", "material"], dataSolicitacao: ["data da solicitacao", "data solicitacao"], dataRecebimento: ["data de recebimento", "data recebimento"], statusProcesso: ["status do processo", "processo"], tmMaster: ["tm master", "tm"], obs: ["observacao", "observacoes", "obs"] };
        const mapa = {};
        MAT2_COLS.forEach(([k, l]) => { [nrm(l), ...(ALIAS[k] || [])].forEach((a) => { mapa[a] = k; }); });
        const rep = { novas: 0, atual: 0, outroDepto: 0, semDepto: 0, semDesc: 0, valorAjustado: 0 };
        const novasLinhas = [], mudancas = new Map();
        const existentes = [...matReqs];
        const porId = new Map(existentes.map((x) => [x.id, x]));
        const porSap = new Map();
        existentes.forEach((x) => { if (x.sap) { const k = x.departamento + "|" + String(x.sap); (porSap.get(k) || porSap.set(k, []).get(k)).push(x); } });
        const deptDesconhecidos = new Map();
        const porDesc = new Map();
        existentes.forEach((x) => { const k = x.departamento + "|" + nrm(x.tmMaster) + "|" + nrm(x.descricao); (porDesc.get(k) || porDesc.set(k, []).get(k)).push(x); });
        /* casa com requisição já existente: pelo ID (planilha exportada do portal) ou, sem ID, pelo mesmo
           departamento + SAP (+ TM Master, se informado). Se a linha traz uma PO diferente da já registrada,
           é outro pedido e entra como nova. Cada requisição existente é atualizada no máximo uma vez. */
        const usados = new Set();
        const casar = (r) => {
          let ex = r.id ? porId.get(r.id) : null;
          if (ex && usados.has(ex.id)) ex = null;
          if (!ex && r.sap) {
            const c = (porSap.get(r.departamento + "|" + r.sap) || []).filter((x) => !usados.has(x.id) && (!r.tmMaster || String(x.tmMaster || "") === r.tmMaster));
            ex = c.find((x) => r.po && String(x.po || "") === r.po && String(x.linhaPo || "") === (r.linhaPo || "")) || c.find((x) => !String(x.po || "").trim() || !r.po);
          }
          if (!ex && !r.sap && r.descricao) {
            const c = (porDesc.get(r.departamento + "|" + nrm(r.tmMaster) + "|" + nrm(r.descricao)) || []).filter((x) => !usados.has(x.id));
            ex = c.find((x) => r.po && String(x.po || "") === r.po) || c.find((x) => !String(x.po || "").trim() || !r.po) || null;
          }
          if (ex) usados.add(ex.id);
          return ex;
        };
        json.forEach((row) => {
          const r = {};
          Object.entries(row).forEach(([h, v]) => { const k = mapa[nrm(h)]; if (k) r[k] = v; });
          if (Object.values(r).every((v) => v === "" || v == null)) return;
          let depto = MAT_DEPTS.find((d) => nrm(d) === nrm(r.departamento));
          if (!depto) { if (!String(r.departamento || "").trim() && deptsEditaveis.length === 1) depto = deptsEditaveis[0]; else { rep.semDepto++; const nome = String(r.departamento || "").trim() || "(em branco)"; deptDesconhecidos.set(nome, (deptDesconhecidos.get(nome) || 0) + 1); return; } }
          if (!pode(depto)) { rep.outroDepto++; return; }
          const lim = { departamento: depto };
          ["tmMaster", "sap", "descricao", "justificativa", "oe", "reserva", "rc", "po", "linhaPo", "obs", "id"].forEach((k) => { const t = String(r[k] ?? "").trim(); if (t) lim[k] = t; });
          ["dataSolicitacao", "eta", "dataRecebimento"].forEach((k) => { const t = aISO(r[k]); if (t) lim[k] = t; });
          ["quantidade", "valor"].forEach((k) => { const t = aNum(r[k]); if (t !== "") lim[k] = t; });
          const pick = (k, lista, padrao) => { const t = String(r[k] ?? "").trim(); if (!t) return; const SIN = { "em aberto": "Aberto", "aberta": "Aberto", "aberto": "Aberto", "fechada": "Fechado", "encerrado": "Fechado", "standby": "Stand By", "stand by": "Stand By" }; const m = lista.find((o) => nrm(o) === nrm(SIN[nrm(t)] || t)); if (m) lim[k] = m; else rep.valorAjustado++; };
          pick("status", MAT_STATUS2); pick("statusProcesso", MAT_PROC); pick("priority", PRIORITY);
          const ex = casar({ ...lim, sap: lim.sap || "", po: lim.po || "", linhaPo: lim.linhaPo || "", descricao: lim.descricao || "", tmMaster: lim.tmMaster || "" });
          if (ex) { if (pode(ex.departamento)) { mudancas.set(ex.id, { ...(mudancas.get(ex.id) || {}), ...lim, departamento: ex.departamento, id: ex.id }); rep.atual++; } else rep.outroDepto++; return; }
          if (!lim.descricao) { rep.semDesc++; return; }
          novasLinhas.push({ tmMaster: "", priority: "Média", sap: "", justificativa: "", status: "Em Processo de Compras", statusProcesso: "Aberto", oe: "", dataSolicitacao: todayISO(), reserva: "", rc: "", po: "", linhaPo: "", valor: "", eta: "", dataRecebimento: "", obs: "", quantidade: 1, ...lim, id: uid("MAT"), upd: Date.now() });
          rep.novas++;
        });
        const agora = Date.now();
        setMatReqs((rows) => [...novasLinhas, ...rows.map((x) => (mudancas.has(x.id) && pode(x.departamento) ? { ...x, ...mudancas.get(x.id), upd: agora } : x))]);
        setPag(0);
        const ign = [rep.outroDepto && `${rep.outroDepto} de outro departamento`, rep.semDepto && `${rep.semDepto} com departamento que não existe na lista (${[...deptDesconhecidos].map(([n, q]) => `${n}: ${q}`).join(", ")})`, rep.semDesc && `${rep.semDesc} sem descrição`].filter(Boolean);
        setImpMsg({ ok: rep.novas + rep.atual > 0, txt: `Importação concluída: ${rep.novas} nova(s), ${rep.atual} atualizada(s).` + (ign.length ? ` Ignoradas: ${ign.join(", ")}.` : "") + (rep.valorAjustado ? ` ${rep.valorAjustado} valor(es) de lista não reconhecido(s) (status/processo/prioridade) usaram o padrão ou foram mantidos.` : "") });
      } catch (err) { setImpMsg({ ok: false, txt: "Não consegui ler a planilha: " + err.message }); }
    }, 30);
    reader.readAsArrayBuffer(file);
  };

  const filtered = useMemo(() => {
    const q = f.busca.trim().toLowerCase();
    return matReqs.filter((r) =>
      (!f.meu || meusDeptos.includes(r.departamento)) &&
      (f.dept === "Todos" || r.departamento === f.dept) &&
      (f.status === "Todos" || r.status === f.status) &&
      (f.proc === "Todos" || r.statusProcesso === f.proc) &&
      (f.prio === "Todos" || r.priority === f.prio) &&
      (!q || [r.tmMaster, r.sap, r.descricao, r.justificativa, r.oe, r.reserva, r.rc, r.po, r.obs].join(" ").toLowerCase().includes(q))
    ).sort((a, b) => ((a.statusProcesso === "Fechado") - (b.statusProcesso === "Fechado")) || String(b.dataSolicitacao || "").localeCompare(String(a.dataSolicitacao || "")));
  }, [matReqs, f, meusDeptos.join("|")]);

  React.useEffect(() => { setPag(0); }, [f]);
  const totalPag = Math.max(1, Math.ceil(filtered.length / POR_PAG));
  const pagAtual = Math.min(pag, totalPag - 1);
  const visiveis = filtered.slice(pagAtual * POR_PAG, pagAtual * POR_PAG + POR_PAG);
  const navPag = (
    <div className="g-flex" style={{ justifyContent: "space-between", alignItems: "center", gap: 10, padding: "8px 14px", flexWrap: "wrap", fontSize: 12 }}>
      <span className="g-muted">Mostrando {filtered.length ? pagAtual * POR_PAG + 1 : 0}–{Math.min(filtered.length, (pagAtual + 1) * POR_PAG)} de {filtered.length} (a tela mostra {POR_PAG} por vez para não pesar; use a busca e os filtros)</span>
      <span className="g-flex" style={{ gap: 6, alignItems: "center" }}>
        <button className="g-btn g-keep" disabled={pagAtual === 0} onClick={() => setPag(0)}>«</button>
        <button className="g-btn g-keep" disabled={pagAtual === 0} onClick={() => setPag(pagAtual - 1)}>‹ Anterior</button>
        <b>Página {pagAtual + 1} de {totalPag}</b>
        <button className="g-btn g-keep" disabled={pagAtual >= totalPag - 1} onClick={() => setPag(pagAtual + 1)}>Próxima ›</button>
        <button className="g-btn g-keep" disabled={pagAtual >= totalPag - 1} onClick={() => setPag(totalPag - 1)}>»</button>
      </span>
    </div>
  );

  const abertas = matReqs.filter((r) => r.statusProcesso !== "Fechado");
  const kpi = {
    total: matReqs.length, abertas: abertas.length,
    aguard: matReqs.filter((r) => r.status === "Aguardando Entrega").length,
    compras: matReqs.filter((r) => r.status === "Em Processo de Compras").length,
    receb: matReqs.filter((r) => r.status === "Recebido").length,
    semPo: abertas.filter((r) => !String(r.po || "").trim()).length,
    valor: abertas.reduce((s, r) => s + (Number(r.valor) || 0), 0),
  };

  const gerarRelatorio = () => {
    const resumo = [["Total", kpi.total], ["Em aberto", kpi.abertas], ["Aguardando entrega", kpi.aguard], ["Em processo de compras", kpi.compras], ["Recebidos", kpi.receb], ["Em aberto sem PO", kpi.semPo], ["Valor em aberto", kpi.valor], ["Requisições no filtro atual", filtered.length], ["Gerado em", new Date().toLocaleDateString("pt-BR")]].map(([i, v]) => ({ Indicador: i, Valor: v }));
    exportSheets("relatorio-materiais", [["Materiais", rowsToSheet(filtered, MAT2_COLS)], ["Resumo", resumo]]);
  };
  React.useEffect(() => { setExportXlsxFn && setExportXlsxFn(null); setReportFn && setReportFn(null); }, []);

  const sel = { padding: "5px 8px", border: "1px solid var(--border)", borderRadius: 6, fontSize: 12, background: "var(--panel, #fff)", color: "var(--text)" };
  const num = (v) => (v === "" || v == null ? "" : v);
  const campo = (rot, el, span) => <div className="g-field" style={span ? { gridColumn: "1 / -1" } : undefined}><label>{rot}</label>{el}</div>;
  const inp = (k, tipo = "text", extra = {}) => <input className="g-edit" type={tipo} value={num(novo[k])} onChange={(e) => setNovo((p) => ({ ...p, [k]: e.target.value }))} {...extra} />;
  const dd = (k, ops) => <select className="g-edit" value={novo[k]} onChange={(e) => setNovo((p) => ({ ...p, [k]: e.target.value }))}>{ops.map((o) => <option key={o} value={o}>{o}</option>)}</select>;

  const th = (t, w) => <th style={{ minWidth: w }}>{t}</th>;
  /* célula: editável se a pessoa pode editar a linha; senão só mostra o valor */
  const C = (r, k, kind = "text", opts) => {
    const ok = pode(r.departamento);
    const v = r[k];
    if (!ok) return <span style={{ fontSize: 12, ...(k === "status" ? { color: MAT_STATUS_COR[v], fontWeight: 700 } : {}) }}>{kind === "date" ? (v ? fmtDate(v) : "—") : kind === "money" ? (v === "" || v == null ? "—" : Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2 })) : (v === "" || v == null ? "—" : String(v))}</span>;
    if (kind === "select") return <select className="g-edit" value={v || ""} onChange={(e) => (k === "status" ? setStatus(r, e.target.value) : upd(r.id, { [k]: e.target.value }))} style={k === "status" ? { color: MAT_STATUS_COR[v], fontWeight: 700, minWidth: 190 } : undefined}>{!v && <option value="">—</option>}{opts.map((o) => <option key={o} value={o}>{o}</option>)}</select>;
    if (kind === "date") return <EDate value={v} onChange={(x) => upd(r.id, { [k]: x })} />;
    if (kind === "num" || kind === "money") return <input className="g-edit num" type="number" step={kind === "money" ? "0.01" : "1"} value={num(v)} onChange={(e) => upd(r.id, { [k]: e.target.value === "" ? "" : Number(e.target.value) })} />;
    if (kind === "area") return <ETextArea rows={1} value={v} onChange={(x) => upd(r.id, { [k]: x })} />;
    return <EText value={v || ""} onChange={(x) => upd(r.id, { [k]: x })} mono={kind === "mono"} />;
  };

  return (
    <>
      <div className="g-alert" style={{ marginBottom: 12, background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>
        <Lock size={14} style={{ marginTop: 1 }} />
        {admin ? "Administrador: você edita as requisições de todos os departamentos."
          : almox ? "Almoxarifado: você edita as requisições de todos os departamentos."
          : podeCriar ? <>Você edita só as requisições {deptsEditaveis.length > 1 ? "dos departamentos" : "do departamento"} <b>{deptsEditaveis.join(", ")}</b>. As dos outros departamentos aparecem para consulta.</>
          : "Somente consulta. Peça ao administrador para definir o seu departamento em Configurações para poder editar."}
      </div>

      <div className="g-kpi-row" style={{ gridTemplateColumns: "repeat(6, 1fr)" }}>
        {bigKpi("Requisições", kpi.total, "var(--teal)", Package)}
        {bigKpi("Processo aberto", kpi.abertas, "var(--warn)", Clock)}
        {bigKpi("Aguardando entrega", kpi.aguard, "#3B82F6", Package)}
        {bigKpi("Em compras", kpi.compras, "var(--warn)", ClipboardList)}
        {bigKpi("Abertas sem PO", kpi.semPo, kpi.semPo ? "var(--crit)" : "var(--ok)", AlertTriangle)}
        {bigKpi("Valor em aberto", kpi.valor.toLocaleString("pt-BR", { maximumFractionDigits: 0 }), "var(--teal)", Package)}
      </div>

      <div className="g-panel">
        <div className="g-flex" style={{ gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          {podeCriar && <button className="g-btn primary g-keep" onClick={() => setNovo(novo ? null : rascunho())}><Plus size={14} />{novo ? "Fechar formulário" : "Nova requisição"}</button>}
          {podeCriar && <><input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={importar} />
            <button className="g-btn" onClick={() => fileRef.current?.click()} title="Importa várias requisições de uma planilha Excel (só as do seu departamento)"><Upload size={14} />Importar Excel</button>
            <button className="g-btn" onClick={baixarModelo} title="Baixa a planilha modelo com as colunas e as listas permitidas"><Download size={14} />Baixar modelo</button></>}
          <button className="g-btn g-keep" onClick={gerarRelatorio} title="Exporta em planilha (.xlsx) as requisições do filtro atual, com uma aba de resumo"><Download size={14} />Exportar relatório</button>
          <input value={f.busca} onChange={(e) => setF((p) => ({ ...p, busca: e.target.value }))} placeholder="buscar TM Master, SAP, descrição, PO, RC, OE..." style={{ ...sel, minWidth: 280, padding: "7px 10px" }} />
          <FiltroBusca value={f.dept} onChange={(v) => setF((p) => ({ ...p, dept: v }))} options={MAT_DEPTS} todos="Todos os departamentos" style={sel} />
          <FiltroBusca value={f.prio} onChange={(v) => setF((p) => ({ ...p, prio: v }))} options={PRIORITY} todos="Todas as prioridades" style={sel} />
          <FiltroBusca value={f.status} onChange={(v) => setF((p) => ({ ...p, status: v }))} options={MAT_STATUS2} todos="Todos os status" style={sel} />
          <FiltroBusca value={f.proc} onChange={(v) => setF((p) => ({ ...p, proc: v }))} options={MAT_PROC} todos="Qualquer processo" style={sel} />
          {!verTodos && podeCriar && <label style={{ fontSize: 12, display: "flex", gap: 4, alignItems: "center" }}><input type="checkbox" checked={f.meu} onChange={(e) => setF((p) => ({ ...p, meu: e.target.checked }))} />{deptsEditaveis.length > 1 ? "Só os meus departamentos" : "Só o meu departamento"}</label>}
          <span className="g-muted" style={{ fontSize: 11.5 }}>{filtered.length} de {matReqs.length}</span>
        </div>
        {msg && <div style={{ marginTop: 8, fontSize: 12, color: "var(--teal)" }}>{msg}</div>}
        {impMsg && <div className="g-alert" style={{ marginTop: 10, marginBottom: 0, ...(impMsg.ok ? { background: "rgba(34,197,94,0.08)", borderColor: "rgba(34,197,94,0.4)", color: "var(--ok)" } : {}) }}>{impMsg.txt} <span style={{ cursor: "pointer", textDecoration: "underline", marginLeft: 8 }} onClick={() => setImpMsg(null)}>fechar</span></div>}
      </div>

      {novo && (
        <div className="g-panel" style={{ borderLeft: "4px solid var(--teal)" }}>
          <div className="g-panel-head"><span className="g-panel-title">Nova requisição de material</span></div>
          <div className="tbl-detail-grid">
            {campo("Departamento *", deptsEditaveis.length === 1 ? <input className="g-edit" value={novo.departamento} disabled /> : <select className="g-edit" value={novo.departamento} onChange={(e) => setNovo((p) => ({ ...p, departamento: e.target.value }))}><option value="">— escolher —</option>{deptsEditaveis.map((o) => <option key={o} value={o}>{o}</option>)}</select>)}
            {campo("TM Master", inp("tmMaster"))}
            {campo("Prioridade", dd("priority", PRIORITY))}
            {campo("SAP", inp("sap"))}
            {campo("Descrição *", inp("descricao"))}
            {campo("Quantidade", inp("quantidade", "number", { min: 0 }))}
            {campo("Status", dd("status", MAT_STATUS2))}
            {campo("Status do Processo", dd("statusProcesso", MAT_PROC))}
            {campo("OE", inp("oe"))}
            {campo("Data da Solicitação", inp("dataSolicitacao", "date"))}
            {campo("Reserva", inp("reserva"))}
            {campo("RC", inp("rc"))}
            {campo("PO", inp("po"))}
            {campo("Item da PO", inp("linhaPo"))}
            {campo("Valor Total", inp("valor", "number", { step: "0.01", min: 0 }))}
            {campo("ETA", inp("eta", "date"))}
            {campo("Data de Recebimento", inp("dataRecebimento", "date"))}
            {campo("Justificativa", <textarea className="g-edit-wrap" rows={2} value={novo.justificativa} onChange={(e) => setNovo((p) => ({ ...p, justificativa: e.target.value }))} />, true)}
            {campo("Observação", <textarea className="g-edit-wrap" rows={2} value={novo.obs} onChange={(e) => setNovo((p) => ({ ...p, obs: e.target.value }))} />, true)}
          </div>
          <div className="g-flex" style={{ gap: 8, marginTop: 10 }}>
            <button className="g-btn primary g-keep" onClick={salvarNovo}>Adicionar requisição</button>
            <button className="g-btn" onClick={() => setNovo(null)}>Cancelar</button>
          </div>
        </div>
      )}

      <div className="g-panel" style={{ padding: 0, overflow: "hidden" }}>
        {navPag}
        <div className="g-table-wrap g-mat-scroll">
          <table className="g-table tbl-modern g-editavel" style={{ minWidth: 2700 }}>
            <thead><tr>
              {th("TM Master", 110)}{th("Departamento", 130)}{th("Prioridade", 120)}{th("SAP", 100)}{th("Descrição", 240)}{th("Qtd", 70)}{th("Justificativa", 240)}{th("Status", 210)}{th("Status do Processo", 120)}{th("OE", 90)}{th("Data da Solicitação", 130)}{th("Reserva", 100)}{th("RC", 100)}{th("PO", 110)}{th("Item da PO", 80)}{th("Valor Total", 110)}{th("ETA", 130)}{th("Data de Recebimento", 130)}{th("Observação", 260)}<th></th>
            </tr></thead>
            <tbody>
              {filtered.length === 0 && <tr><td colSpan={20} className="g-muted" style={{ textAlign: "center", padding: 18 }}>{matReqs.length ? "Nenhuma requisição nesse filtro." : "Nenhuma requisição ainda. Use “Nova requisição” para cadastrar."}</td></tr>}
              {visiveis.map((r) => {
                const ok = pode(r.departamento);
                return (
                  <tr className="g-row tbl-row" key={r.id} style={{ "--c": MAT_STATUS_COR[r.status] || "#9499A8", opacity: ok || !podeCriar ? 1 : 0.88 }}>
                    <td style={{ verticalAlign: "top" }}>{C(r, "tmMaster", "mono")}</td>
                    <td style={{ verticalAlign: "top" }}><span className="tbl-chip">{r.departamento}</span></td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "priority", "select", PRIORITY)}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "sap", "mono")}</td>
                    <td style={{ verticalAlign: "top", whiteSpace: "normal" }}>{C(r, "descricao", "area")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "quantidade", "num")}</td>
                    <td style={{ verticalAlign: "top", whiteSpace: "normal" }}>{C(r, "justificativa", "area")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "status", "select", MAT_STATUS2)}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "statusProcesso", "select", MAT_PROC)}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "oe", "mono")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "dataSolicitacao", "date")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "reserva", "mono")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "rc", "mono")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "po", "mono")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "linhaPo", "mono")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "valor", "money")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "eta", "date")}</td>
                    <td style={{ verticalAlign: "top" }}>{C(r, "dataRecebimento", "date")}</td>
                    <td style={{ verticalAlign: "top", whiteSpace: "normal" }}>{C(r, "obs", "area")}</td>
                    <td style={{ verticalAlign: "top" }}>{ok ? <span className="g-btn ghost danger g-keep" title="Excluir requisição" onClick={() => remover(r)}><Trash2 size={13} /></span> : <Lock size={12} color="var(--text-faint)" title="Só o departamento dono ou o Almoxarifado edita" />}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        {navPag}
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
  const todosMarcados = options.length > 0 && options.every((o) => selected.includes(o));
  const label = selected.length === 0 || todosMarcados ? "Todos" : selected.length === 1 ? getLabel(selected[0]) : `${selected.length} selecionados`;
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
          <label style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 4px 7px", fontSize: 12, cursor: "pointer", fontWeight: 700, borderBottom: "1px solid var(--border-soft)", marginBottom: 4 }}>
            <input type="checkbox" checked={todosMarcados} onChange={() => onChange(todosMarcados ? [] : [...options])} />
            Selecionar todos
          </label>
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

function PaymentsSection({ setExportXlsxFn, paySubTab, setPaySubTab, serviceInvoices, updInv, remInv, addInv, setReportFn, newRowId }) {
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
  const selectAllIds = (ids) => setSelectedIds(new Set(ids));

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

      {paySubTab === "total" && <PaymentsTotalView setExportXlsxFn={setExportXlsxFn} serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} f={f} selectedIds={selectedIds} toggleSelect={toggleSelect} setReportFn={setReportFn} newRowId={newRowId} />}
      {paySubTab === "status" && <PaymentsStatusView selectAllIds={selectAllIds} setExportXlsxFn={setExportXlsxFn} serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} f={f} setF={setF} selectedIds={selectedIds} toggleSelect={toggleSelect} setReportFn={setReportFn} newRowId={newRowId} />}
      {paySubTab === "dashboard" && <PaymentsValoresView setExportXlsxFn={setExportXlsxFn} serviceInvoices={serviceInvoices} updInv={updInv} remInv={remInv} f={f} selectedIds={selectedIds} toggleSelect={toggleSelect} setReportFn={setReportFn} newRowId={newRowId} />}
    </div>
  );
}

/* ---------- Página 1: Dashboard Total (todas as colunas da planilha + filtros + métricas de prazo) ---------- */
function PaymentsTotalView({ setExportXlsxFn, serviceInvoices, updInv, remInv, f, selectedIds, toggleSelect, setReportFn, newRowId }) {
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
    if (!setExportXlsxFn) return;
    setExportXlsxFn(() => () => {
      const cols = INV_COLS.filter(([k]) => k !== "saldoPo" && k !== "diffDays");
      exportSheets("genesis-pagamentos-total", [["Pagamentos", rowsToSheet(activeRows, cols)]]);
    });
  }, [activeRows, setExportXlsxFn]);

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
function PaymentsStatusView({ selectAllIds, setExportXlsxFn, serviceInvoices, updInv, remInv, f, setF, selectedIds, toggleSelect, setReportFn }) {
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
    if (!setExportXlsxFn) return;
    setExportXlsxFn(() => () => {
      exportSheets("genesis-pagamentos-status", [["Status dos Pagamentos", rowsToSheet(activeRows, INV_COLS)]]);
    });
  }, [activeRows, setExportXlsxFn]);

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
              <th style={{ width: 28 }}>
                <input type="checkbox" title="Selecionar todos os registros listados"
                  checked={sorted.length > 0 && sorted.every((r) => selectedIds.has(r.id))}
                  onChange={(e) => selectAllIds(e.target.checked ? sorted.map((r) => r.id) : [])} />
              </th>
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
function PaymentsValoresView({ setExportXlsxFn, serviceInvoices, updInv, remInv, f, selectedIds, toggleSelect, setReportFn }) {
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
    if (!setExportXlsxFn) return;
    setExportXlsxFn(() => () => {
      const rows = activeRows.map((r) => ({ ...rowsToSheet([r], INV_COLS)[0], "Situação": r._situation }));
      exportSheets("genesis-pagamentos-valores", [["Valores", rows]]);
    });
  }, [activeRows, setExportXlsxFn]);

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
function CostsView({ setExportXlsxFn, serviceInvoices, updInv, exchangeRate, setExchangeRate, setReportFn }) {
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
    if (!setExportXlsxFn) return;
    setExportXlsxFn(() => () => {
      const categorias = categoryCosts.map((c) => ({ Categoria: c.category, "Ordem": adpServicosLabel(c.category), "Orçado (R$)": c.ilimitado ? "Sem teto" : c.orcadoBrl, "Realizado (R$)": c.realizado }));
      if (costSubTab === "dashboard") {
        exportSheets("genesis-custos-dashboard", [["Custo por Categoria", categorias]]);
      } else {
        const rateio = [];
        filtered.forEach((r) => {
          const al = allocationsOf(r);
          if (al.length === 0) rateio.push({ Data: r.date, Serviço: r.assunto, Empresa: r.empresa, Status: r.statusPagamento, "Mês Provisionado": r.previsaoMes, "Valor Total": r.valorTotal, Categoria: "(sem rateio)", Ordem: "", "Valor Rateado": 0, "Justificativa Geral": r.justificativaGeral || "" });
          al.forEach((a) => rateio.push({ Data: r.date, Serviço: r.assunto, Empresa: r.empresa, Status: r.statusPagamento, "Mês Provisionado": r.previsaoMes, "Valor Total": r.valorTotal, Categoria: a.category, Ordem: adpServicosLabel(a.category), "Valor Rateado": a.valor, "Justificativa Geral": r.justificativaGeral || "" }));
        });
        exportSheets("genesis-custos-rateio", [["Rateio", rateio], ["Custo por Categoria", categorias]]);
      }
    });
  }, [costSubTab, filtered, categoryCosts, setExportXlsxFn]);

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
function SettingsView({ currentUser, users, setUsers, usersSave, salvarUsuarios }) {
  const updUser = (i, field, value) => setUsers((r) => r.map((u, idx) => (idx === i ? { ...u, [field]: value } : u)));
  const remUser = (i) => {
    if (isAdminUser(users[i])) { window.alert("O administrador não pode ser removido."); return; }
    if (users.length <= 1) { window.alert("É preciso manter pelo menos um usuário cadastrado."); return; }
    setUsers((r) => r.filter((_, idx) => idx !== i));
  };
  const addUser = () => setUsers((r) => [...r, { id: uid("USR"), name: "Novo usuário", username: `usuario${r.length + 1}`, password: "GenesisI", role: "viewer", tabs: ["dashboard"] }]);
  const setTodasAbas = (i, marcar) => updUser(i, "tabs", marcar ? PERM_TABS.map((t) => t.key) : []);
  const setTodosDeptos = (i, campo, lista, marcar) => updUser(i, campo, marcar ? [...lista] : []);
  const lnk = { fontSize: 11, textDecoration: "underline", cursor: "pointer", color: "var(--teal)", marginRight: 8 };
  const toggleTab = (i, key) => {
    const u = users[i];
    const cur = Array.isArray(u.tabs) ? normTabs(u.tabs) : PERM_TABS.map((t) => t.key);
    const next = cur.includes(key) ? cur.filter((k) => k !== key) : [...cur, key];
    updUser(i, "tabs", next);
  };

  return (
    <>
      <div className="g-panel">
        <div className="g-panel-head"><span className="g-panel-title">Conta atual</span></div>
        <div className="g-muted">Conectado como <strong style={{ color: "var(--text)" }}>{currentUser?.name}</strong> (usuário: {currentUser?.username}) — administrador</div>
      </div>

      <div className="g-panel">
        <div className="g-panel-head">
          <span className="g-panel-title">Usuários, senhas e permissões</span>
          {usersSave && usersSave.msg && <span style={{ fontSize: 12, fontWeight: 700, color: usersSave.estado === "erro" ? "var(--crit)" : usersSave.estado === "salvando" ? "var(--warn)" : "var(--ok)" }}>{usersSave.estado === "ok" ? "✓ " : ""}{usersSave.msg}</span>}
          <button className="g-btn g-keep" onClick={() => salvarUsuarios && salvarUsuarios()} title="Grava agora e confere no servidor">Salvar usuários agora</button>
          <button className="g-btn primary" onClick={addUser}><Plus size={14} />Novo usuário</button>
        </div>
        <div className="g-alert" style={{ background: "rgba(43,108,176,0.08)", borderColor: "rgba(43,108,176,0.35)", color: "var(--teal)" }}>
          <Lock size={14} style={{ marginTop: 1 }} />
          Somente o administrador acessa Configurações. "Visualizador" só consulta as abas em geral (não edita, não importa, não exclui) e vê apenas as abas marcadas. EXCEÇÃO: nas abas Materiais e TM Master › Registro de Bordo, quem tem departamentos marcados nas colunas "aba Materiais" e "aba Registro de Bordo" edita só esses departamentos, cada aba com a sua liberação, mesmo sendo Visualizador (o perfil não importa ali; o que vale é a marcação). Login local ao navegador — não é uma autenticação segura de servidor.
        </div>
        <table className="g-table">
          <thead>
            <tr><th>Nome</th><th>Usuário</th><th>Senha</th><th>Perfil</th><th>Departamentos — aba Materiais</th><th>Departamentos — aba Registro de Bordo</th><th>Abas liberadas</th><th></th></tr>
          </thead>
          <tbody>
            {users.map((u, i) => {
              const adm = isAdminUser(u);
              const tabs = Array.isArray(u.tabs) ? normTabs(u.tabs) : PERM_TABS.map((t) => t.key);
              return (
                <tr className="g-row" key={u.id}>
                  <td style={{ minWidth: 150 }}><EText value={u.name} onChange={(v) => updUser(i, "name", v)} /></td>
                  <td style={{ minWidth: 120 }}>{adm ? <span className="g-mono">{u.username}</span> : <EText value={u.username} onChange={(v) => updUser(i, "username", v)} mono />}</td>
                  <td style={{ minWidth: 120 }}><EText value={u.password} onChange={(v) => updUser(i, "password", v)} mono /></td>
                  <td style={{ minWidth: 150 }}>
                    {adm ? <strong>Administrador</strong> : (
                      <select value={u.role === "viewer" ? "viewer" : "editor"} onChange={(e) => updUser(i, "role", e.target.value)}>
                        <option value="editor">Editor</option>
                        <option value="viewer">Visualizador</option>
                      </select>
                    )}
                  </td>
                  <td style={{ minWidth: 150 }}>
                    {adm ? <span className="g-muted">Todos</span> : (
                      <div style={{ display: "flex", flexDirection: "column", gap: 2 }} title="Departamentos que esta pessoa pode lançar e editar na aba Materiais. Almoxarifado edita todos.">
                        <div><span style={lnk} onClick={() => setTodosDeptos(i, "matDepts", MAT_DEPTS, true)}>Marcar todos</span><span style={lnk} onClick={() => setTodosDeptos(i, "matDepts", MAT_DEPTS, false)}>Limpar</span></div>
                        {MAT_DEPTS.map((d) => (
                          <label key={d} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, cursor: "pointer" }}>
                            <input type="checkbox" checked={deptsDoUsuario(u).includes(d)} onChange={(e) => { const at = deptsDoUsuario(u); updUser(i, "matDepts", e.target.checked ? [...at, d] : at.filter((x) => x !== d)); }} />{d}{d === "Almoxarifado" ? " (edita todos)" : ""}
                          </label>
                        ))}
                        {!deptsDoUsuario(u).length && <span className="g-muted" style={{ fontSize: 11 }}>só consulta</span>}
                      </div>
                    )}
                  </td>
                  <td style={{ minWidth: 150 }}>
                    {adm ? <span className="g-muted">Todos</span> : (
                      <div style={{ display: "flex", flexDirection: "column", gap: 2 }} title="Departamentos que esta pessoa pode lançar e editar na aba TM Master › Registro de Bordo.">
                        <div><span style={lnk} onClick={() => setTodosDeptos(i, "bordoDepts", BORDO_DEPTS, true)}>Marcar todos</span><span style={lnk} onClick={() => setTodosDeptos(i, "bordoDepts", BORDO_DEPTS, false)}>Limpar</span></div>
                        {BORDO_DEPTS.map((d) => (
                          <label key={d} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, cursor: "pointer" }}>
                            <input type="checkbox" checked={deptsBordoUsuario(u).includes(d)} onChange={(e) => { const at = deptsBordoUsuario(u); updUser(i, "bordoDepts", e.target.checked ? [...at, d] : at.filter((x) => x !== d)); }} />{d}
                          </label>
                        ))}
                        {!deptsBordoUsuario(u).length && <span className="g-muted" style={{ fontSize: 11 }}>só consulta</span>}
                      </div>
                    )}
                  </td>
                  <td style={{ minWidth: 330 }}>
                    {adm ? <span className="g-muted">Todas + Configurações</span> : (
                      <div style={{ display: "flex", flexWrap: "wrap", gap: "4px 12px" }}>
                        <div style={{ width: "100%" }}><span style={lnk} onClick={() => setTodasAbas(i, true)}>Marcar todas</span><span style={lnk} onClick={() => setTodasAbas(i, false)}>Limpar</span></div>
                        {PERM_TABS.map((t) => (
                          <label key={t.key} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, cursor: "pointer" }}>
                            <input type="checkbox" checked={tabs.includes(t.key)} onChange={() => toggleTab(i, t.key)} />{t.label}
                          </label>
                        ))}
                      </div>
                    )}
                  </td>
                  <td>{!adm && <span className="g-btn ghost danger" onClick={() => remUser(i)}><Trash2 size={13} /></span>}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </>
  );
}
