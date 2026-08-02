"use client";
import { storage } from "../lib/supabase";

import React, { useState, useEffect, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  Package, Wallet, Plus, X, TrendingUp, TrendingDown, ChevronDown, Trash2,
  Loader2, Check, Users, Settings2, ShoppingCart, Download, Upload,
  Car, ShieldCheck, BarChart3, Handshake, RefreshCw, Wrench, Lock,
  LogOut, Delete, KeyRound, Clock, PlayCircle, Phone, PhoneCall,
  Search, Calendar, AlertTriangle, ArrowRight, Zap, Droplets, Star, Pencil, Save
} from "lucide-react";

/* ═══════════════════════════════════════════════════
   CONSTANTS
═══════════════════════════════════════════════════ */
const STORAGE_KEY = "avtogaz-v2";
const ROLE_LABELS = { admin: "Admin", kassir: "Kassir", usta: "Usta" };

const UNITS = ["dona", "kg", "litr", "metr", "komplekt"];
const SERVICE_TYPES = ["Servis", "Ustanovka", "Detailing", "Moy bo'limi"];
const PAYMENT_TYPES = ["Naqd pul", "Karta (Click/Payme)", "Bank o'tkazma", "Nasiya (qarzga)"];
const SOURCE_TYPES = ["Ta'minotchi", "O'z mahsuloti", "Insider servis"];
const CATEGORIES_DEFAULT = ["Gaz mahsulotlari", "Moy mahsulotlari", "B/U tovarlar", "Boshqa"];

const LEAD_STAGES = [
  { id: "yangi",        label: "Yangi qo'ng'iroq", color: "#2979FF" },
  { id: "tasdiqlangan", label: "Tasdiqlangan",      color: "#FFB300" },
  { id: "keldi",        label: "Keldi",             color: "#00BFA5" },
  { id: "bekor",        label: "Bekor qilindi",     color: "#F44336" },
];

const INCOME_CATEGORIES = ["Xizmat to'lovi", "Erkin savdo", "Hamkordan to'lov", "Rahbardan kirim", "Boshqa kirim"];
const EXPENSE_CATEGORIES = [
  "Ta'minotchiga to'lov", "Usta xizmat haqi", "Rahbarga chiqim",
  "Karta (Click/Payme) chiqim", "Ijara", "Ish haqi", "Kommunal", "Hujjat xarajati", "Boshqa xarajat"
];

const emptyData = () => ({
  settings: {
    usdRate: 12650,
    pins: { admin: "1111", kassir: "2211", usta: "3311" },
    azimKpi: 0,
    categories: [...CATEGORIES_DEFAULT],
  },
  products: [], stockIns: [], stockOuts: [], freeSales: [], cashflow: [],
  serviceCards: [], warrantyClaims: [], partners: [], partnerTx: [],
  ustaLedger: [], leads: [],
  bonusRules: [], bonusAwards: [], nasiyaDebts: [],
  employees: [], employeePayments: [],
  personalDebts: [],
  contractedMasters: [], // kelishilgan (oylik) ustalar — ularning haqi servis foydasiga qo'shiladi
});

/* ═══════════════════════════════════════════════════
   HELPERS
═══════════════════════════════════════════════════ */
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
const todayISO = () => new Date().toISOString().slice(0, 10);
const nowTime = () => new Date().toTimeString().slice(0, 5);
const fmtDate = (iso) => { if (!iso) return "—"; const [y, m, d] = iso.split("-"); return `${d}.${m}.${y}`; };
const num = (v) => Number(v) || 0;
const fmtSum = (n) => Math.round(num(n)).toLocaleString("ru-RU").replace(/,/g, " ") + " so'm";
const fmtUsd = (n) => "$" + (Math.round(num(n) * 100) / 100).toLocaleString("en-US");
const toSum = (a, cur, rate) => (cur === "USD" ? num(a) * rate : num(a));
const toUsd = (a, cur, rate) => (cur === "USD" ? num(a) : rate ? num(a) / rate : 0);

function cardPartsCost(card) { return (card.parts || []).reduce((s, p) => s + num(p.lineTotal), 0); }
// Servis/Moy bo'limi'da lineTotal sotish narxida hisoblanadi (mijozga shuncha yoziladi),
// shuning uchun haqiqiy tan narxni alohida hisoblaymiz — profit shu yerdan chiqishi kerak.
function cardPartsRealCost(card) {
  return (card.parts || []).reduce((s, p) => s + num(p.qty) * num(p.costUnit ?? p.unitCost), 0);
}
function cardUstaFeeSum(card) { return (card.ustaFeeEntries || []).reduce((s, e) => s + num(e.amount), 0); }
function cardStatus(card) { return card.status || "yakunlangan"; }

function supplierDebts(data) {
  const map = {};
  data.stockIns.filter(s => s.sourceType !== "O'z mahsuloti").forEach((s) => {
    const name = s.supplier || "Noma'lum";
    if (!map[name]) map[name] = { name, totalSum: 0, paidSum: 0, sourceType: s.sourceType };
    map[name].totalSum += num(s.totalSum);
    map[name].paidSum += num(s.paidSum);
  });
  return Object.values(map)
    .map((s) => ({ ...s, debtSum: s.totalSum - s.paidSum }))
    .filter((s) => Math.abs(s.debtSum) > 1);
}

function openStockDebts(data) {
  return data.stockIns.filter(
    (s) => s.sourceType !== "O'z mahsuloti" && num(s.totalSum) - num(s.paidSum) > 1
  );
}

function applyFifoPayment(stockIns, supplier, paySum, rate) {
  let remaining = paySum;
  const open = stockIns
    .filter((s) => s.supplier === supplier && num(s.totalSum) - num(s.paidSum) > 0.5)
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  for (const s of open) {
    if (remaining <= 0) break;
    const debt = num(s.totalSum) - num(s.paidSum);
    const pay = Math.min(debt, remaining);
    s.paidSum = num(s.paidSum) + pay;
    remaining -= pay;
  }
}

function partnerBalances(data) {
  return data.partners.map((p) => {
    const tx = data.partnerTx.filter((t) => t.partnerId === p.id);
    const given = tx.filter((t) => t.type === "mahsulot").reduce((s, t) => s + num(t.amountSum), 0);
    const paid = tx.filter((t) => t.type === "tolov").reduce((s, t) => s + num(t.amountSum), 0);
    return { ...p, given, paid, debtSum: given - paid };
  });
}

function ustaNameOptions(data) {
  const names = new Set();
  data.ustaLedger.forEach((e) => e.usta && names.add(e.usta));
  data.serviceCards.forEach((c) => c.usta && names.add(c.usta));
  return Array.from(names).sort();
}

function ustaPendingByName(data) {
  const map = {};
  data.ustaLedger.filter((e) => !e.paid).forEach((e) => {
    if (!map[e.usta]) map[e.usta] = { usta: e.usta, amountSum: 0, count: 0, ids: [] };
    map[e.usta].amountSum += num(e.amountSum);
    map[e.usta].count += 1;
    map[e.usta].ids.push(e.id);
  });
  return Object.values(map).sort((a, b) => b.amountSum - a.amountSum);
}

function ustaPendingGrouped(data) {
  const map = {};
  data.ustaLedger.filter((e) => !e.paid).forEach((e) => {
    const key = e.usta + "|" + e.date;
    if (!map[key]) map[key] = { key, usta: e.usta, date: e.date, amountSum: 0, count: 0, ids: [] };
    map[key].amountSum += num(e.amountSum);
    map[key].count += 1;
    map[key].ids.push(e.id);
  });
  return Object.values(map).sort((a, b) => (b.date + b.usta).localeCompare(a.date + a.usta));
}

function tryDownloadFile(filename, content, mimeType) {
  try {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    return true;
  } catch (e) {
    return false;
  }
}

function doExportExcel(data, rate) {
  const wb = XLSX.utils.book_new();
  const sklad = data.products.map((p) => ({
    "Nomi": p.name, "Kategoriya": p.category || "", "Birlik": p.unit,
    "Kelish narxi": p.costSum, "Sotish (so'm)": p.priceSum, "Sotish (USD)": p.priceUsd || 0,
    "Qoldiq": p.qty, "Qiymati": num(p.qty) * num(p.costSum),
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sklad), "Sklad");

  const kassa = data.cashflow.map((c) => ({
    "Sana": c.date, "Turi": c.type === "kirim" ? "Kirim" : "Chiqim",
    "Turkum": c.category, "Valyuta": c.currency, "Summa": c.amount,
    "Summa (so'm)": c.amountSum, "Izoh": c.note,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(kassa), "Kassa");

  const cards = data.serviceCards.map((c) => ({
    "Sana": c.date, "Holati": cardStatus(c) === "ochiq" ? "Ochiq" : "Yakunlangan",
    "Davlat raqami": c.plate, "Telefon": c.phone, "Mashina": c.carModel,
    "Xizmat turi": c.serviceType, "Usta": c.usta,
    "Tan narx": cardPartsRealCost(c), "Usta haqi": cardUstaFeeSum(c),
    "Hujjat": c.docFee || 0, "Yakuniy": c.finalTotal || 0, "Foyda": c.profitSum || 0,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cards), "Xizmat kartalari");

  const leads = (data.leads || []).map((l) => ({
    "Sana": l.date, "Vaqt": l.time, "Ism": l.name, "Telefon": l.phone,
    "Mashina": l.carModel, "Muammo": l.issue, "Holat": l.stage,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(leads), "Qo'ng'iroqlar");

  const ustaL = data.ustaLedger.map((e) => ({
    "Sana": e.date, "Usta": e.usta, "Summa": e.amountSum,
    "Holati": e.paid ? "To'langan" : "Kutilmoqda",
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ustaL), "Usta hisobi");

  const debts = supplierDebts(data).map((d) => ({
    "Ta'minotchi": d.name, "Jami": d.totalSum, "To'landi": d.paidSum, "Qarz": d.debtSum,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(debts), "Taminotchi qarzi");

  try {
    XLSX.writeFile(wb, `avtogaz-hisobot-${todayISO()}.xlsx`);
    return true;
  } catch (e) {
    console.error("Excel export xatosi:", e);
    return false;
  }
}

/* ═══════════════════════════════════════════════════
   DESIGN TOKENS
═══════════════════════════════════════════════════ */
const T = {
  bg: "#EAEDF5", s1: "#FFFFFF", s2: "#FAFBFD", s3: "#F0F3F9", s4: "#E7EBF3",
  border: "#E3E8F1", border2: "#D2DAE7",
  text: "#12192B", muted: "#828FA3", muted2: "#4F5E76",
  navy: "#0D1526", navy2: "#182236",
  flame: "#E8541F", flameD: "#E8541F14",
  gold: "#CE8A00", goldD: "#CE8A0014",
  teal: "#0C9B80", tealD: "#0C9B8014",
  red: "#DC3545", redD: "#DC354514",
  blue: "#1D6FEB", blueD: "#1D6FEB14",
  purple: "#7150D9", purpleD: "#7150D914",
  shadowSm: "0 1px 2px rgba(13,21,38,.06)",
  shadowMd: "0 8px 24px -6px rgba(13,21,38,.12), 0 2px 8px -2px rgba(13,21,38,.06)",
  shadowLg: "0 24px 60px -14px rgba(13,21,38,.28)",
};

const SERVICE_COLORS = {
  "Servis": T.teal, "Ustanovka": T.flame,
  "Detailing": T.purple, "Moy bo'limi": T.gold,
};

function GlobalStyles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700;800&family=Barlow+Condensed:wght@600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');
      *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
      html{background:${T.bg}}
      body{
        color:${T.text};font-family:'Barlow',sans-serif;font-size:14px;line-height:1.5;
        min-height:100vh;
        background-color:${T.bg};
        background-image:
          radial-gradient(1200px 620px at 4% -14%, ${T.flame}20, transparent 58%),
          radial-gradient(1000px 680px at 106% -8%, ${T.teal}19, transparent 56%),
          radial-gradient(820px 720px at 46% 30%, ${T.purple}0d, transparent 60%),
          radial-gradient(900px 760px at 50% 130%, ${T.gold}12, transparent 60%),
          radial-gradient(#C2CCDE 1.1px, transparent 1.1px),
          linear-gradient(180deg, #F8FAFD 0%, #E8ECF4 55%, #E2E7F1 100%);
        background-repeat:no-repeat,no-repeat,no-repeat,no-repeat,repeat,no-repeat;
        background-size:100% 680px,100% 680px,100% 900px,100% 900px,22px 22px,100% 100%;
        background-position:top left,top right,center,bottom center,0 0,top;
        background-attachment:scroll,scroll,scroll,scroll,fixed,fixed;
      }
      .bc{font-family:'Barlow Condensed',sans-serif}
      .mo{font-family:'JetBrains Mono',monospace}
      button,input,select,textarea{font-family:'Barlow',sans-serif;font-size:14px}
      input:focus,select:focus,textarea:focus{outline:none;border-color:${T.flame}!important;box-shadow:0 0 0 3px ${T.flame}20}
      ::-webkit-scrollbar{width:6px;height:6px}
      ::-webkit-scrollbar-track{background:transparent}
      ::-webkit-scrollbar-thumb{background:${T.border2};border-radius:3px}
      ::-webkit-scrollbar-thumb:hover{background:${T.muted}}
      .rh{transition:background .12s ease}
      .rh:hover{background:${T.s3}!important}
      .ch{transition:transform .16s ease, box-shadow .16s ease, border-color .16s ease}
      .ch:hover{border-color:${T.flame}!important;transform:translateY(-3px);box-shadow:${T.shadowMd}}
      .card-shadow{box-shadow:${T.shadowSm}, 0 0 0 1px rgba(16,24,40,.02)}
      .panel{transition:box-shadow .18s ease, transform .18s ease}
      .lift-hover{transition:transform .16s ease, box-shadow .16s ease}
      .lift-hover:hover{transform:translateY(-3px);box-shadow:${T.shadowMd}}
      .btn-lift{transition:transform .12s ease, box-shadow .12s ease, filter .12s ease, opacity .12s ease}
      .btn-lift:hover{transform:translateY(-1px);filter:brightness(1.04)}
      .btn-lift:active{transform:translateY(0);filter:brightness(.97)}
      .tab-btn{transition:background .14s ease, color .14s ease, box-shadow .14s ease}
      .tab-btn:hover{background:${T.s3}}
      .pin-key{transition:transform .1s ease, box-shadow .1s ease, border-color .1s ease}
      .pin-key:hover{transform:translateY(-1px);box-shadow:${T.shadowMd};border-color:${T.border2}}
      .pin-key:active{transform:translateY(0)}
      @keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
      @keyframes spin{to{transform:rotate(360deg)}}
      @keyframes pulse{0%,100%{opacity:1}50%{opacity:.45}}
      @keyframes glow{0%,100%{opacity:.55}50%{opacity:1}}
      .fi{animation:fadeUp .2s ease both}
      .spin{animation:spin 1s linear infinite}
      .pulse{animation:pulse 1.5s ease infinite}
      select{appearance:none}
      .grid-2b{display:grid;grid-template-columns:1.4fr 1fr;gap:16px}
      .grid-3b{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}
      .grid-4b{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
      @media(max-width:860px){.grid-2b,.grid-3b{grid-template-columns:1fr}}
      @media(max-width:640px){.hide-sm{display:none!important}.grid-4b{grid-template-columns:repeat(2,1fr)}}
      @media(max-width:480px){.hide-xs{display:none!important}}
      @media(max-width:420px){.grid-4b{grid-template-columns:1fr}}
      .menu-item:hover{background:${T.s3}!important}
    `}</style>
  );
}

function BackgroundLayer() {
  return <div style={{
    position: "fixed", inset: 0, pointerEvents: "none", zIndex: 0,
    background: "radial-gradient(140% 100% at 50% 0%, transparent 55%, rgba(13,21,38,.05) 100%)",
  }} />;
}

/* ═══════════════════════════════════════════════════
   UI PRIMITIVES
═══════════════════════════════════════════════════ */
const iSt = {
  width: "100%", padding: "9.5px 12px", borderRadius: 9,
  border: `1px solid ${T.border2}`, background: T.s3,
  color: T.text, fontSize: 13, transition: "border-color .12s ease, box-shadow .12s ease",
};

function F({ label, children, col }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, gridColumn: col }}>
      <label style={{
        fontSize: 10, fontWeight: 700, letterSpacing: ".08em",
        textTransform: "uppercase", color: T.muted2,
      }}>{label}</label>
      {children}
    </div>
  );
}

function Sel({ value, onChange, options, style: s }) {
  const opts = options.map((o) => (o && typeof o === "object" ? o : { value: o, label: o }));
  return (
    <div style={{ position: "relative" }}>
      <select value={value} onChange={onChange} style={{ ...iSt, paddingRight: 30, ...s }}>
        {opts.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
      <ChevronDown size={13} style={{
        position: "absolute", right: 10, top: "50%",
        transform: "translateY(-50%)", color: T.muted, pointerEvents: "none",
      }} />
    </div>
  );
}

function CurrencyToggle({ value, onChange }) {
  return (
    <div style={{ display: "flex", borderRadius: 8, overflow: "hidden", border: `1px solid ${T.border2}` }}>
      {["SUM", "USD"].map((c) => (
        <button key={c} type="button" onClick={() => onChange(c)} style={{
          flex: 1, padding: "9px", border: "none", cursor: "pointer",
          fontSize: 12, fontWeight: 600,
          background: value === c ? T.flame : "transparent",
          color: value === c ? "#fff" : T.muted,
        }}>{c === "SUM" ? "SO'M" : "USD"}</button>
      ))}
    </div>
  );
}

function Modal({ title, onClose, children, wide, xwide }) {
  return (
    <div onClick={onClose} style={{
      position: "fixed", inset: 0, zIndex: 200,
      background: "rgba(10,16,28,.55)", backdropFilter: "blur(3px)", WebkitBackdropFilter: "blur(3px)",
      display: "flex", alignItems: "flex-start", justifyContent: "center",
      padding: "20px 14px", overflowY: "auto",
    }}>
      <div onClick={(e) => e.stopPropagation()} className="fi" style={{
        width: "100%", maxWidth: xwide ? 880 : wide ? 660 : 480,
        background: T.s1, border: `1px solid ${T.border}`,
        borderRadius: 16, overflow: "hidden", marginBottom: 20,
        boxShadow: T.shadowLg,
      }}>
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "16px 22px", borderBottom: `1px solid ${T.border}`,
          position: "sticky", top: 0, background: `linear-gradient(180deg,${T.s1},${T.s2})`, zIndex: 10,
        }}>
          <span className="bc" style={{ fontSize: 16.5, fontWeight: 700, letterSpacing: ".01em" }}>{title}</span>
          <button onClick={onClose} style={{
            background: T.s3, border: `1px solid ${T.border2}`, borderRadius: 8, cursor: "pointer",
            color: T.muted2, padding: 5, display: "flex",
          }}><X size={16} /></button>
        </div>
        <div style={{ padding: "20px 22px" }}>{children}</div>
      </div>
    </div>
  );
}

function SaveBtn({ onClick, disabled, children = "Saqlash", color }) {
  return (
    <button onClick={onClick} disabled={disabled} className={disabled ? "" : "btn-lift"} style={{
      width: "100%", marginTop: 18, padding: "12.5px", borderRadius: 10, border: "none",
      background: disabled ? T.s3 : `linear-gradient(135deg,${color || T.flame},${color ? color + "99" : "#D84315"})`,
      color: disabled ? T.muted : "#fff", fontWeight: 700, fontSize: 13.5,
      cursor: disabled ? "not-allowed" : "pointer",
      display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
      boxShadow: disabled ? "none" : `0 8px 20px -6px ${color || T.flame}88`,
    }}>{children}</button>
  );
}

/* ── Global tasdiqlash tizimi — window.confirm sandboxda ishlamaydi ── */
let _confirmSetter = null;
function askConfirm(message) {
  return new Promise((resolve) => {
    if (_confirmSetter) _confirmSetter({ message, resolve });
  });
}
function ConfirmHost() {
  const [state, setState] = useState(null);
  useEffect(() => { _confirmSetter = setState; return () => { _confirmSetter = null; }; }, []);
  if (!state) return null;
  return (
    <Modal title="Tasdiqlash" onClose={() => { state.resolve(false); setState(null); }}>
      <p style={{ fontSize: 13.5, color: T.text, lineHeight: 1.6, whiteSpace: "pre-line" }}>{state.message}</p>
      <div style={{ display: "flex", gap: 10, marginTop: 18 }}>
        <Btn variant="ghost" style={{ flex: 1, justifyContent: "center" }}
          onClick={() => { state.resolve(false); setState(null); }}>Bekor qilish</Btn>
        <Btn variant="red" style={{ flex: 1, justifyContent: "center" }}
          onClick={() => { state.resolve(true); setState(null); }}>Tasdiqlash</Btn>
      </div>
    </Modal>
  );
}

function Btn({ onClick, children, variant = "primary", size = "md", style: s }) {
  const sizes = { sm: "6px 11px", md: "9px 16px", lg: "12px 22px" };
  const variants = {
    primary: { background: `linear-gradient(135deg,${T.flame},#D84315)`, color: "#fff", border: "none", boxShadow: `0 6px 16px -6px ${T.flame}90` },
    ghost: { background: T.s1, color: T.muted2, border: `1px solid ${T.border2}`, boxShadow: T.shadowSm },
    teal: { background: T.tealD, color: T.teal, border: `1px solid ${T.teal}35` },
    red: { background: T.redD, color: T.red, border: `1px solid ${T.red}35` },
    gold: { background: T.goldD, color: T.gold, border: `1px solid ${T.gold}35` },
  };
  return (
    <button onClick={onClick} className="btn-lift" style={{
      display: "inline-flex", alignItems: "center", gap: 6,
      padding: sizes[size], borderRadius: 9, cursor: "pointer",
      fontSize: size === "sm" ? 11.5 : 13, fontWeight: 600,
      ...variants[variant], ...s,
    }}>{children}</button>
  );
}

function Badge({ children, color }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 4,
      padding: "3px 10px", borderRadius: 20, fontSize: 11, fontWeight: 700,
      background: color + "16", color, letterSpacing: ".02em",
      border: `1px solid ${color}30`,
    }}>{children}</span>
  );
}

/* ── HEADER DROPDOWN MENU — zamonaviy, tor ekranda ham qulay ── */
function JSONExportModal({ data, onClose }) {
  const [copied, setCopied] = useState(false);
  const jsonStr = JSON.stringify(data, null, 2);
  const filename = `avtogaz-backup-${todayISO()}.json`;

  function download() {
    const ok = tryDownloadFile(filename, jsonStr, "application/json");
    if (!ok) copyToClipboard();
  }

  function copyToClipboard() {
    const ta = document.createElement("textarea");
    ta.value = jsonStr;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); setCopied(true); setTimeout(() => setCopied(false), 2500); }
    catch (e) { /* jim o'tamiz, matn baribir ko'rinadi textarea'da */ }
    document.body.removeChild(ta);
  }

  return (
    <Modal title="JSON eksport — zaxira nusxa" onClose={onClose} wide>
      <p style={{ fontSize: 12.5, color: T.muted, marginBottom: 14 }}>
        Fayl avtomatik yuklab olinadi. Agar yuklab bo'lmasa, "Nusxalash" tugmasini bosib matnni
        oling va <b>{filename}</b> nomi bilan .json fayl sifatida saqlang.
      </p>

      <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
        <Btn onClick={download}><Download size={14} /> Faylni yuklash</Btn>
        <Btn variant={copied ? "teal" : "ghost"} onClick={copyToClipboard}>
          {copied ? <Check size={14} /> : <Upload size={14} />} {copied ? "Nusxalandi!" : "Nusxalash"}
        </Btn>
      </div>

      <textarea
        readOnly value={jsonStr}
        onClick={(e) => e.target.select()}
        style={{
          width: "100%", height: 260, padding: 12, borderRadius: 8,
          border: `1px solid ${T.border2}`, background: T.s3, color: T.text,
          fontFamily: "monospace", fontSize: 11, resize: "vertical",
        }}
      />
    </Modal>
  );
}

function JSONImportModal({ onClose, onImport }) {
  const [text, setText] = useState("");
  const [error, setError] = useState("");
  const [pendingData, setPendingData] = useState(null); // parse qilingan, tasdiq kutayotgan
  const fileRef = useRef(null);

  function handleFile(e) {
    const file = e.target.files?.[0]; if (!file) return;
    const r = new FileReader();
    r.onload = () => { setText(r.result); setPendingData(null); };
    r.onerror = () => setError("Faylni o'qib bo'lmadi.");
    r.readAsText(file);
  }

  function checkAndPreview() {
    setError("");
    try {
      const p = JSON.parse(text);
      if (!isValidData(p)) { setError("❌ Fayl noto'g'ri formatda — kerakli maydonlar topilmadi."); return; }
      setPendingData(p);
    } catch (e) {
      setError("❌ Matn to'g'ri JSON formatida emas.");
    }
  }

  function confirmImport() {
    onImport(pendingData);
    onClose();
  }

  if (pendingData) {
    const itemCount = (pendingData.serviceCards?.length || 0) + (pendingData.products?.length || 0);
    return (
      <Modal title="Importni tasdiqlash" onClose={onClose}>
        <div style={{ padding: "14px 16px", background: T.goldD, border: `1px solid ${T.gold}30`, borderRadius: 9, marginBottom: 16 }}>
          <p style={{ fontSize: 13.5, color: T.text, lineHeight: 1.6 }}>
            Joriy barcha ma'lumotlar shu fayldagi <b>{itemCount} ta yozuv</b> bilan almashtiriladi.
            <br /><br />
            <b style={{ color: T.red }}>Bu amalni ortga qaytarib bo'lmaydi.</b> Davom etasizmi?
          </p>
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <Btn variant="ghost" style={{ flex: 1, justifyContent: "center" }} onClick={() => setPendingData(null)}>Orqaga</Btn>
          <Btn variant="red" style={{ flex: 1, justifyContent: "center" }} onClick={confirmImport}>
            <Upload size={14} /> Ha, import qilish
          </Btn>
        </div>
      </Modal>
    );
  }

  return (
    <Modal title="JSON import — zaxiradan tiklash" onClose={onClose} wide>
      <p style={{ fontSize: 12.5, color: T.muted, marginBottom: 14 }}>
        Faylni tanlang yoki JSON matnini pastdagi maydonga to'g'ridan-to'g'ri joylashtiring (Ctrl+V).
      </p>

      <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
        <Btn variant="ghost" onClick={() => fileRef.current?.click()}><Upload size={14} /> Fayl tanlash</Btn>
        <input ref={fileRef} type="file" accept=".json" onChange={handleFile} style={{ display: "none" }} />
      </div>

      <textarea
        value={text} onChange={(e) => { setText(e.target.value); setError(""); }}
        placeholder="JSON matnini shu yerga joylang yoki faylni yuqoridan tanlang..."
        style={{
          width: "100%", height: 240, padding: 12, borderRadius: 8,
          border: `1px solid ${error ? T.red : T.border2}`, background: T.s3, color: T.text,
          fontFamily: "monospace", fontSize: 11, resize: "vertical",
        }}
      />
      {error && <p style={{ color: T.red, fontSize: 12, marginTop: 8 }}>{error}</p>}

      <SaveBtn disabled={!text.trim()} onClick={checkAndPreview} color={T.gold}>
        <Upload size={15} /> Import qilish
      </SaveBtn>
    </Modal>
  );
}

function RateChangeModal({ rate, onClose, onSave }) {
  const [val, setVal] = useState(String(rate));
  const num_ = Number(val);
  return (
    <Modal title="USD kursini o'zgartirish" onClose={onClose}>
      <F label="1 USD = ? so'm">
        <input
          type="number" style={iSt} value={val}
          onChange={(e) => setVal(e.target.value)}
          autoFocus onFocus={(e) => e.target.select()}
        />
      </F>
      <p style={{ fontSize: 11.5, color: T.muted, marginTop: 8 }}>
        Joriy kurs: <b className="mo">{fmtSum(rate)}</b>
      </p>
      <SaveBtn disabled={!num_ || num_ <= 0} onClick={() => { onSave(num_); onClose(); }}>
        Saqlash
      </SaveBtn>
    </Modal>
  );
}

function ResetAllConfirmModal({ data, onClose, onConfirm }) {
  const [step, setStep] = useState(1); // 1: ogohlantirish+eksport, 2: yozib tasdiqlash
  const [typed, setTyped] = useState("");
  const CONFIRM_WORD = "OCHIRISH";

  const counts = {
    cards: data.serviceCards?.length || 0,
    products: data.products?.length || 0,
    cashflow: data.cashflow?.length || 0,
    partners: data.partners?.length || 0,
  };

  function exportBeforeReset() {
    tryDownloadFile(`avtogaz-oxirgi-zaxira-${todayISO()}.json`, JSON.stringify(data, null, 2), "application/json");
  }

  if (step === 1) {
    return (
      <Modal title="⚠️ Barcha ma'lumotni tozalash" onClose={onClose} wide>
        <div style={{ padding: "14px 16px", background: T.redD, border: `1px solid ${T.red}40`, borderRadius: 9, marginBottom: 16 }}>
          <p style={{ fontSize: 13.5, color: T.text, lineHeight: 1.7 }}>
            Bu amal <b style={{ color: T.red }}>BARCHA</b> ma'lumotni butunlay o'chiradi:
          </p>
          <ul style={{ margin: "10px 0 0 18px", fontSize: 12.5, color: T.muted2, lineHeight: 1.9 }}>
            <li>{counts.cards} ta xizmat kartasi</li>
            <li>{counts.products} ta sklad mahsuloti</li>
            <li>{counts.cashflow} ta kassa yozuvi</li>
            <li>{counts.partners} ta hamkor</li>
            <li>Barcha qarzlar, xodimlar, kafolatlar va boshqa yozuvlar</li>
          </ul>
        </div>
        <p style={{ fontSize: 12.5, color: T.gold, marginBottom: 14, display: "flex", alignItems: "center", gap: 8 }}>
          <AlertTriangle size={14} /> Tavsiya: davom etishdan oldin zaxira nusxa oling.
        </p>
        <Btn variant="ghost" style={{ width: "100%", justifyContent: "center", marginBottom: 10 }} onClick={exportBeforeReset}>
          <Download size={14} /> Avval JSON zaxira yuklab olish
        </Btn>
        <SaveBtn color={T.red} onClick={() => setStep(2)}>
          Tushundim, davom etaman
        </SaveBtn>
      </Modal>
    );
  }

  return (
    <Modal title="Yakuniy tasdiqlash" onClose={onClose}>
      <p style={{ fontSize: 13, color: T.text, marginBottom: 14, lineHeight: 1.6 }}>
        Davom etish uchun pastdagi maydonga aynan <b className="mo" style={{ color: T.red }}>{CONFIRM_WORD}</b> so'zini yozing:
      </p>
      <input
        style={{ ...iSt, textAlign: "center", fontSize: 16, fontWeight: 700, letterSpacing: ".05em" }}
        value={typed} onChange={(e) => setTyped(e.target.value)}
        placeholder={CONFIRM_WORD} autoFocus
      />
      <SaveBtn color={T.red} disabled={typed !== CONFIRM_WORD} onClick={onConfirm}>
        <Trash2 size={15} /> Ha, barcha ma'lumotni butunlay o'chirish
      </SaveBtn>
    </Modal>
  );
}

function HeaderMenu({ role, rate, patch, data, onImport, onResetAll }) {
  const [open, setOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [importOpen, setImportOpen] = useState(false);
  const [rateOpen, setRateOpen] = useState(false);
  const [excelError, setExcelError] = useState(false);
  const [resetPinOpen, setResetPinOpen] = useState(false);
  const [resetConfirmOpen, setResetConfirmOpen] = useState(false);
  const menuRef = useRef(null);

  useEffect(() => {
    if (!open) return;
    function onDocClick(e) { if (menuRef.current && !menuRef.current.contains(e.target)) setOpen(false); }
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, [open]);

  return (
    <div ref={menuRef} style={{ position: "relative" }}>
      <button onClick={() => setOpen((s) => !s)} style={{
        display: "flex", alignItems: "center", gap: 6,
        background: T.s2, border: `1px solid ${T.border2}`,
        borderRadius: 8, padding: "7px 11px", cursor: "pointer",
      }}>
        <Settings2 size={13} color={T.flame} />
        <span className="mo hide-sm" style={{ fontSize: 11, fontWeight: 600, color: T.text }}>{fmtSum(rate)}</span>
        <ChevronDown size={11} color={T.muted} style={{ transform: open ? "rotate(180deg)" : "none", transition: "transform .15s" }} />
      </button>

      {open && (
        <div style={{
          position: "absolute", right: 0, top: "calc(100% + 8px)",
          width: 220, background: T.s1, border: `1px solid ${T.border2}`,
          borderRadius: 11, boxShadow: "0 12px 32px rgba(20,30,45,.18)",
          overflow: "hidden", zIndex: 300,
        }} className="fi">
          {role === "admin" && (
            <button onClick={() => { setRateOpen(true); setOpen(false); }}
              className="menu-item" style={menuItemSt}>
              <Settings2 size={14} color={T.flame} /> Kursni o'zgartirish
              <span className="mo" style={{ marginLeft: "auto", fontSize: 11, color: T.muted }}>{fmtSum(rate)}</span>
            </button>
          )}
          {(role === "admin" || role === "kassir") && (
            <>
              <button onClick={() => { setExportOpen(true); setOpen(false); }} className="menu-item" style={menuItemSt}>
                <Download size={14} color={T.blue} /> JSON eksport
              </button>
              <button onClick={() => { const ok = doExportExcel(data, rate); if (!ok) setExcelError(true); setOpen(false); }} className="menu-item" style={menuItemSt}>
                <Download size={14} color={T.teal} /> Excel eksport
              </button>
            </>
          )}
          {role === "admin" && (
            <button onClick={() => { setImportOpen(true); setOpen(false); }} className="menu-item" style={menuItemSt}>
              <Upload size={14} color={T.gold} /> JSON import
            </button>
          )}
          {role === "admin" && (
            <button onClick={() => { setResetPinOpen(true); setOpen(false); }}
              className="menu-item" style={{ ...menuItemSt, borderTop: `1px solid ${T.border}`, color: T.red }}>
              <Trash2 size={14} color={T.red} /> Barcha ma'lumotni tozalash
            </button>
          )}
        </div>
      )}

      {exportOpen && <JSONExportModal data={data} onClose={() => setExportOpen(false)} />}
      {importOpen && <JSONImportModal onClose={() => setImportOpen(false)} onImport={onImport} />}
      {rateOpen && (
        <RateChangeModal rate={rate} onClose={() => setRateOpen(false)}
          onSave={(v) => patch((d) => { d.settings.usdRate = v; return d; })} />
      )}
      {excelError && (
        <Modal title="Excel eksport muvaffaqiyatsiz" onClose={() => setExcelError(false)}>
          <p style={{ fontSize: 13.5, color: T.text, lineHeight: 1.6, marginBottom: 16 }}>
            Excel faylni yuklab bo'lmadi (brauzer cheklovi bo'lishi mumkin).
            Buning o'rniga <b>JSON eksport</b>dan foydalaning — u orqali ham barcha ma'lumot to'liq saqlanadi.
          </p>
          <Btn onClick={() => { setExcelError(false); setExportOpen(true); }}>
            <Download size={14} /> JSON eksportga o'tish
          </Btn>
        </Modal>
      )}
      {resetPinOpen && (
        <SimplePinModal onClose={() => setResetPinOpen(false)}
          onSuccess={() => { setResetPinOpen(false); setResetConfirmOpen(true); }} />
      )}
      {resetConfirmOpen && (
        <ResetAllConfirmModal data={data} onClose={() => setResetConfirmOpen(false)}
          onConfirm={() => { onResetAll(); setResetConfirmOpen(false); }} />
      )}
    </div>
  );
}
const menuItemSt = {
  display: "flex", alignItems: "center", gap: 10, width: "100%",
  padding: "11px 14px", background: "none", border: "none", cursor: "pointer",
  fontSize: 12.5, color: "inherit", textAlign: "left", fontFamily: "inherit",
};

function Empty({ Icon, text, sub }) {
  return (
    <div style={{ padding: "42px 20px", textAlign: "center" }}>
      <div style={{
        width: 46, height: 46, borderRadius: 13, margin: "0 auto 12px",
        background: T.s3, border: `1px solid ${T.border}`,
        display: "flex", alignItems: "center", justifyContent: "center",
      }}>
        <Icon size={20} color={T.muted} />
      </div>
      <p style={{ color: T.muted2, fontSize: 13, fontWeight: 600 }}>{text}</p>
      {sub && <p style={{ color: T.muted, fontSize: 11.5, marginTop: 4 }}>{sub}</p>}
    </div>
  );
}

function Tbl({ cols, rows, empty }) {
  if (!rows.length)
    return <Empty Icon={Package} text={empty || "Ma'lumot yo'q"} />;
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
        <thead>
          <tr style={{ borderBottom: `1px solid ${T.border}`, background: T.s2 }}>
            {cols.map((c) => (
              <th key={c.k} style={{
                padding: "10px 10px", textAlign: "left", color: T.muted2,
                fontWeight: 700, fontSize: 10, textTransform: "uppercase",
                letterSpacing: ".06em", whiteSpace: "nowrap",
              }}>{c.h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} className="rh" style={{
              borderBottom: `1px solid ${T.border}25`,
              background: i % 2 ? "transparent" : `${T.s2}90`,
            }}>
              {cols.map((c) => (
                <td key={c.k} style={{ padding: "10px 10px", color: T.text, whiteSpace: "nowrap" }}>
                  {c.r ? c.r(row) : row[c.k]}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Stat({ label, value, sub, color, Icon }) {
  return (
    <div className="card-shadow lift-hover" style={{
      position: "relative", background: `linear-gradient(165deg,${T.s1},${T.s2})`, border: `1px solid ${T.border}`, borderRadius: 14,
      padding: "16px 17px", overflow: "hidden",
    }}>
      <div style={{
        position: "absolute", top: 0, left: 0, right: 0, height: 3,
        background: `linear-gradient(90deg,${color},${color}55)`,
      }} />
      <div style={{
        position: "absolute", top: -30, right: -30, width: 90, height: 90, borderRadius: "50%",
        background: `radial-gradient(circle,${color}14,transparent 70%)`,
      }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, position: "relative" }}>
        <span style={{
          fontSize: 10, fontWeight: 700, letterSpacing: ".08em",
          textTransform: "uppercase", color: T.muted,
        }}>{label}</span>
        {Icon && <div style={{
          background: `linear-gradient(135deg,${color}22,${color}0d)`, border: `1px solid ${color}30`,
          borderRadius: 8, padding: 6, display: "flex",
        }}>
          <Icon size={13} color={color} />
        </div>}
      </div>
      <div className="mo bc" style={{ fontSize: 20, fontWeight: 700, color, lineHeight: 1.15, position: "relative" }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: T.muted, marginTop: 5, position: "relative" }}>{sub}</div>}
    </div>
  );
}

function Card({ title, children, action, pad = true, Icon, color = T.flame }) {
  return (
    <div className="card-shadow" style={{
      background: `linear-gradient(180deg,${T.s1},${T.s2})`, border: `1px solid ${T.border}`,
      borderRadius: 14, overflow: "hidden",
    }}>
      {title && (
        <div style={{
          padding: "13px 18px", borderBottom: `1px solid ${T.border}`,
          display: "flex", justifyContent: "space-between", alignItems: "center",
          background: `linear-gradient(180deg,${T.s2}CC,${T.s1}00)`,
        }}>
          <span style={{ display: "flex", alignItems: "center", gap: 9 }}>
            {Icon && (
              <span style={{
                display: "flex", alignItems: "center", justifyContent: "center",
                width: 24, height: 24, borderRadius: 7,
                background: `linear-gradient(135deg,${color}22,${color}0d)`, border: `1px solid ${color}30`,
              }}><Icon size={12.5} color={color} /></span>
            )}
            <span className="bc" style={{ fontSize: 14.5, fontWeight: 700, letterSpacing: ".01em" }}>{title}</span>
          </span>
          {action}
        </div>
      )}
      <div style={{ padding: pad ? "16px 18px" : 0 }}>{children}</div>
    </div>
  );
}

function PageHeader({ Icon, title, sub, color = T.flame, action }) {
  return (
    <div style={{
      display: "flex", justifyContent: "space-between", alignItems: "flex-start",
      marginBottom: 20, flexWrap: "wrap", gap: 12,
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 13 }}>
        {Icon && (
          <div style={{
            width: 42, height: 42, borderRadius: 12, flexShrink: 0,
            background: `linear-gradient(135deg,${color}20,${color}08)`,
            border: `1px solid ${color}30`, boxShadow: `0 6px 16px -8px ${color}70`,
            display: "flex", alignItems: "center", justifyContent: "center",
          }}><Icon size={19} color={color} /></div>
        )}
        <div>
          <h2 className="bc" style={{ fontSize: 22, fontWeight: 800, letterSpacing: ".01em" }}>{title}</h2>
          {sub && <p style={{ color: T.muted, fontSize: 12, marginTop: 3 }}>{sub}</p>}
        </div>
      </div>
      {action}
    </div>
  );
}

/* ═══════════════════════════════════════════════════
   LOGIN SCREEN
═══════════════════════════════════════════════════ */
function LoginScreen({ pins, onSuccess }) {
  const [digits, setDigits] = useState("");
  const [error, setError] = useState(false);

  function checkPin(v) {
    const match = Object.entries(pins).find(([, p]) => p === v);
    if (match) { setDigits(""); setError(false); onSuccess(match[0]); }
    else { setError(true); setTimeout(() => { setDigits(""); setError(false); }, 450); }
  }
  function press(d) {
    if (digits.length >= 4 || error) return;
    const next = digits + d;
    setDigits(next);
    if (next.length === 4) checkPin(next);
  }

  return (
    <div style={{
      minHeight: "100vh", display: "flex",
      alignItems: "center", justifyContent: "center", padding: 20,
      position: "relative", overflow: "hidden",
    }}>
      <GlobalStyles />
      <BackgroundLayer />
      <div style={{
        position: "absolute", inset: 0, pointerEvents: "none",
        background: `radial-gradient(700px 420px at 50% -10%, ${T.flame}1c, transparent 60%)`,
      }} />
      <div className="fi card-shadow" style={{
        position: "relative", width: "100%", maxWidth: 320, textAlign: "center",
        background: `linear-gradient(180deg,${T.s1},${T.s2})`,
        border: `1px solid ${T.border}`, borderRadius: 22,
        padding: "36px 28px 30px", boxShadow: T.shadowLg,
      }}>
        <div style={{
          width: 60, height: 60, borderRadius: 17, margin: "0 auto 18px",
          background: `linear-gradient(135deg,${T.flame},#BF360C)`,
          display: "flex", alignItems: "center", justifyContent: "center",
          boxShadow: `0 10px 24px -8px ${T.flame}90`,
        }}>
          <Zap size={28} color="#fff" />
        </div>
        <h1 className="bc" style={{ fontSize: 21, fontWeight: 800, letterSpacing: ".04em", marginBottom: 5 }}>
          AVTOGAZ SERVICE
        </h1>
        <p style={{
          fontSize: 12, color: T.muted, marginBottom: 26,
          display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
        }}>
          <Lock size={12} /> PIN-kodni kiriting
        </p>

        <div className={error ? "pulse" : ""} style={{ display: "flex", justifyContent: "center", gap: 12, marginBottom: 26 }}>
          {[0, 1, 2, 3].map((i) => (
            <span key={i} style={{
              width: 14, height: 14, borderRadius: "50%",
              border: `2px solid ${error ? T.red : digits.length > i ? T.flame : T.border2}`,
              background: error ? T.red : digits.length > i ? T.flame : "transparent",
              transition: "background .12s ease, border-color .12s ease",
            }} />
          ))}
        </div>
        {error && <p style={{ fontSize: 12, color: T.red, marginBottom: 14, fontWeight: 600 }}>PIN noto'g'ri</p>}

        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 11 }}>
          {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => (
            <button key={n} className="pin-key" onClick={() => press(String(n))} style={{
              background: T.s1, border: `1px solid ${T.border2}`, borderRadius: 14,
              color: T.text, fontSize: 19, fontWeight: 600, padding: "15px 0", cursor: "pointer",
              boxShadow: T.shadowSm,
            }}>{n}</button>
          ))}
          <div />
          <button className="pin-key" onClick={() => press("0")} style={{
            background: T.s1, border: `1px solid ${T.border2}`, borderRadius: 14,
            color: T.text, fontSize: 19, fontWeight: 600, padding: "15px 0", cursor: "pointer",
            boxShadow: T.shadowSm,
          }}>0</button>
          <button className="pin-key" onClick={() => !error && setDigits((s) => s.slice(0, -1))} style={{
            background: T.s1, border: `1px solid ${T.border2}`, borderRadius: 14,
            color: T.muted, padding: "15px 0", cursor: "pointer",
            boxShadow: T.shadowSm,
            display: "flex", alignItems: "center", justifyContent: "center",
          }}><Delete size={18} /></button>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════
   ROOT APP
═══════════════════════════════════════════════════ */
const LOCAL_BACKUP_KEY = "avtogaz-v2-local-backup";

function saveLocalBackup(data) {
  try { localStorage.setItem(LOCAL_BACKUP_KEY, JSON.stringify({ data, savedAt: Date.now() })); }
  catch (e) { /* localStorage to'liq bo'lishi mumkin, jim o'tamiz */ }
}

function loadLocalBackup() {
  try {
    const raw = localStorage.getItem(LOCAL_BACKUP_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed || null; // {data, savedAt} ni to'liq qaytaramiz
  } catch (e) { return null; }
}

function isValidData(p) {
  return p && typeof p === "object" && Array.isArray(p.products) && Array.isArray(p.serviceCards);
}

export default function App() {
  const [data, setData] = useState(emptyData());
  const [loaded, setLoaded] = useState(false);
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved | error
  const [tab, setTab] = useState("dashboard");
  const [role, setRole] = useState(null);
  const [dataSource, setDataSource] = useState(null); // "server" | "local" | "empty"
  const saveTimer = useRef(null);
  const retryCount = useRef(0);

  // ── YUKLASH: server va lokal zaxirani solishtirib, eng so'nggisini olamiz ──
  useEffect(() => {
    (async () => {
      let serverPayload = null; // {data, savedAt}
      try {
        const res = await storage.get(STORAGE_KEY, true);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          // Eski format (to'g'ridan-to'g'ri data) yoki yangi format ({data, savedAt}) bo'lishi mumkin
          if (parsed && parsed.data && parsed.savedAt) {
            if (isValidData(parsed.data)) serverPayload = parsed;
          } else if (isValidData(parsed)) {
            serverPayload = { data: parsed, savedAt: 0 }; // eski format — vaqt tamg'asi yo'q, eng past ustuvorlik
          }
        }
      } catch (e) {
        console.error("Server yuklash xatosi:", e);
      }

      const localPayload = loadLocalBackup(); // {data, savedAt} yoki null

      let chosen = null, source = "empty";
      if (serverPayload && localPayload) {
        // Ikkalasi ham bor — vaqti kattarog'ini tanlaymiz
        if (num(serverPayload.savedAt) >= num(localPayload.savedAt)) {
          chosen = serverPayload.data; source = "server";
        } else {
          chosen = localPayload.data; source = "local";
        }
      } else if (serverPayload) {
        chosen = serverPayload.data; source = "server";
      } else if (localPayload) {
        chosen = localPayload.data; source = "local";
      }

      if (chosen) {
        setData({
          ...emptyData(), ...chosen,
          settings: {
            ...emptyData().settings, ...(chosen.settings || {}),
            pins: { ...emptyData().settings.pins, ...((chosen.settings || {}).pins || {}) },
          },
        });
        // Har ikkala joyga ham eng yangi holatni yozib qo'yamiz — sinxronlashtirish
        saveLocalBackup(chosen);
      }
      setDataSource(source);
      setLoaded(true);
    })();
  }, []);

  // ── SAQLASH: har o'zgarishda darhol lokal, keyin serverga (retry bilan) ──
  useEffect(() => {
    if (!loaded) return;
    saveLocalBackup(data); // HAR DOIM darhol lokalga — bu hech qachon yo'qolmaydi

    setSaveState("saving");
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => saveToServer(data), 500);
    return () => clearTimeout(saveTimer.current);
  }, [data, loaded]);

  async function saveToServer(payload, attempt = 0) {
    setSaveState("saving");
    try {
      // Server tomonda ham {data, savedAt} formatida saqlaymiz — vaqt solishtirish uchun
      await storage.set(STORAGE_KEY, JSON.stringify({ data: payload, savedAt: Date.now() }), true);
      setSaveState("saved");
      retryCount.current = 0;
    } catch (e) {
      console.error("Saqlash xatosi:", e);
      if (attempt < 2) {
        setTimeout(() => saveToServer(payload, attempt + 1), 1200 * (attempt + 1));
      } else {
        setSaveState("error");
        retryCount.current += 1;
      }
    }
  }

  const patch = useCallback((fn) => setData((d) => fn(JSON.parse(JSON.stringify(d)))), []);
  const rate = data.settings.usdRate || 12650;

  const allTabs = [
    { id: "dashboard", label: "Bosh sahifa",  Icon: Zap,        roles: ["admin", "kassir"] },
    { id: "callcenter",label: "Qo'ng'iroqlar", Icon: PhoneCall,  roles: ["admin", "kassir"] },
    { id: "services",  label: "Kartalar",      Icon: Car,        roles: ["admin", "kassir"] },
    { id: "warehouse", label: "Sklad",         Icon: Package,    roles: ["admin", "kassir"] },
    { id: "cashier",   label: "Kassa",         Icon: Wallet,     roles: ["admin", "kassir"] },
    { id: "ustalar",   label: "Usta hisobi",   Icon: Wrench,     roles: ["admin", "kassir", "usta"] },
    { id: "warranty",  label: "Kafolat",       Icon: ShieldCheck,roles: ["admin", "kassir"] },
    { id: "partners",  label: "Hamkorlar",     Icon: Handshake,  roles: ["admin", "kassir"] },
    { id: "employees", label: "Xodimlar",      Icon: Users,      roles: ["admin"] },
    { id: "analytics", label: "Analitika",     Icon: BarChart3,  roles: ["admin"] },
  ];
  const tabs = allTabs.filter((t) => role && t.roles.includes(role));

  useEffect(() => {
    if (role && tabs.length && !tabs.some((t) => t.id === tab)) setTab(tabs[0].id);
  }, [role]);

  if (!loaded)
    return (
      <div style={{
        minHeight: "100vh", display: "flex",
        alignItems: "center", justifyContent: "center",
      }}>
        <GlobalStyles />
      <BackgroundLayer />
        <Loader2 size={30} color={T.flame} className="spin" />
      </div>
    );

  if (!role) return <LoginScreen pins={data.settings.pins} onSuccess={setRole} />;

  return (
    <div style={{ minHeight: "100vh", color: T.text }}>
      <GlobalStyles />
      <BackgroundLayer />

      {/* HEADER */}
      <header style={{
        background: `linear-gradient(180deg,${T.s1},${T.s1}F2)`,
        backdropFilter: "blur(10px)", WebkitBackdropFilter: "blur(10px)",
        borderBottom: `1px solid ${T.border}`,
        padding: "0 20px", height: 56, display: "flex",
        alignItems: "center", justifyContent: "space-between",
        position: "sticky", top: 0, zIndex: 100,
        boxShadow: T.shadowSm,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
          <div style={{
            width: 34, height: 34, borderRadius: 10,
            background: `linear-gradient(135deg,${T.flame},#BF360C)`,
            display: "flex", alignItems: "center", justifyContent: "center",
            boxShadow: `0 6px 14px -5px ${T.flame}90`,
          }}><Zap size={17} color="#fff" /></div>
          <div>
            <div className="bc" style={{ fontSize: 15.5, fontWeight: 800, letterSpacing: ".05em", lineHeight: 1, whiteSpace: "nowrap" }}>
              AVTOGAZ SERVICE
            </div>
            <div style={{ fontSize: 9.5, color: T.muted, letterSpacing: ".1em", fontWeight: 600, marginTop: 3, display: "flex", alignItems: "center", gap: 5 }}>
              <span style={{
                display: "inline-flex", alignItems: "center", gap: 4,
                color: T.flame, background: T.flameD, padding: "1.5px 7px", borderRadius: 20, fontWeight: 700,
              }}>{ROLE_LABELS[role].toUpperCase()}</span>
              <span className="hide-sm">{fmtDate(todayISO())}</span>
            </div>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {dataSource === "local" && (
            <span title="Serverdan yuklab bo'lmadi — mahalliy zaxiradan tiklandi" className="hide-xs" style={{
              fontSize: 10, color: T.gold, display: "flex", alignItems: "center", gap: 4,
              background: T.goldD, padding: "4px 9px", borderRadius: 20, fontWeight: 600,
            }}>
              <AlertTriangle size={10} /> <span className="hide-sm">Lokal zaxira</span>
            </span>
          )}

          <span className="hide-xs" style={{
            fontSize: 10.5, fontWeight: 600, color: T.muted,
            display: "flex", alignItems: "center", gap: 4,
          }}>
            {saveState === "saving" && <><Loader2 size={11} className="spin" color={T.gold} /> <span className="hide-sm" style={{ color: T.gold }}>Saqlanmoqda</span></>}
            {(saveState === "saved" || saveState === "idle") && <><Check size={11} color={T.teal} /> <span className="hide-sm" style={{ color: T.teal }}>Saqlangan</span></>}
            {saveState === "error" && <><AlertTriangle size={11} color={T.red} /> <span className="hide-sm" style={{ color: T.red }}>Saqlanmadi</span></>}
          </span>

          <button
            onClick={() => saveToServer(data)}
            className="btn-lift"
            style={{
              display: "flex", alignItems: "center", gap: 6,
              padding: "7px 13px", borderRadius: 9, border: "none", cursor: "pointer",
              background: saveState === "error"
                ? `linear-gradient(135deg,${T.red},#B71C1C)`
                : `linear-gradient(135deg,${T.teal},#00897B)`,
              color: "#fff", fontSize: 12, fontWeight: 700,
              boxShadow: `0 6px 14px -6px ${saveState === "error" ? T.red : T.teal}90`,
            }}
          >
            {saveState === "saving"
              ? <Loader2 size={13} className="spin" />
              : <Save size={13} />}
            <span className="hide-sm">{saveState === "error" ? "Qayta saqlash" : "Saqlash"}</span>
          </button>

          <div style={{ width: 1, height: 22, background: T.border }} className="hide-sm" />

          <HeaderMenu
            role={role} rate={rate} patch={patch} data={data}
            onImport={(p) => {
              const next = { ...emptyData(), ...p, settings: { ...emptyData().settings, ...(p.settings || {}) } };
              setData(next);
              saveLocalBackup(next);
            }}
            onResetAll={() => {
              const fresh = emptyData();
              setData(fresh);
              saveLocalBackup(fresh);
              saveToServer(fresh);
            }}
          />

          <Btn variant="ghost" size="sm" onClick={() => setRole(null)} style={{ color: T.red }}>
            <LogOut size={13} />
          </Btn>
        </div>
      </header>

      {/* TABS */}
      <div style={{
        background: T.s1, borderBottom: `1px solid ${T.border}`,
        padding: "10px 20px", position: "sticky", top: 56, zIndex: 90,
        boxShadow: "0 1px 0 rgba(16,24,40,.03)",
      }}>
        <div style={{
          display: "flex", overflowX: "auto", gap: 3,
          background: T.s3, border: `1px solid ${T.border}`,
          borderRadius: 11, padding: 4, width: "fit-content", maxWidth: "100%",
        }}>
          {tabs.map(({ id, label, Icon }) => {
            const a = tab === id;
            return (
              <button key={id} onClick={() => setTab(id)} className={a ? "" : "tab-btn"} style={{
                display: "flex", alignItems: "center", gap: 6,
                padding: "8px 14px", border: "none", borderRadius: 8,
                cursor: "pointer", fontSize: 12.5, fontWeight: a ? 700 : 500,
                color: a ? T.flame : T.muted2,
                background: a ? T.s1 : "transparent",
                boxShadow: a ? T.shadowSm : "none",
                whiteSpace: "nowrap", transition: "background .12s, color .12s, box-shadow .12s",
              }}>
                <Icon size={14} /> {label}
              </button>
            );
          })}
        </div>
      </div>

      {/* CONTENT */}
      <div style={{ padding: "24px 20px 40px", maxWidth: 1400, margin: "0 auto" }} className="fi">
        {tab === "dashboard"  && <DashboardTab  data={data} patch={patch} rate={rate} setTab={setTab} />}
        {tab === "callcenter" && <CallCenterTab data={data} patch={patch} />}
        {tab === "services"   && <ServicesTab   data={data} patch={patch} rate={rate} />}
        {tab === "warehouse"  && <WarehouseTab  data={data} patch={patch} rate={rate} />}
        {tab === "cashier"    && <CashierTab    data={data} patch={patch} rate={rate} />}
        {tab === "ustalar"    && <UstaTab       data={data} patch={patch} rate={rate} canManage={role !== "usta"} />}
        {tab === "warranty"   && <WarrantyTab   data={data} patch={patch} />}
        {tab === "partners"   && <PartnersTab   data={data} patch={patch} rate={rate} />}
        {tab === "employees"  && <EmployeesTab  data={data} patch={patch} rate={rate} />}
        {tab === "analytics"  && <AnalyticsTab  data={data} patch={patch} rate={rate} />}
      </div>
      <ConfirmHost />
    </div>
  );
}

function Placeholder({ label }) {
  return (
    <div style={{ padding: "70px 20px", textAlign: "center", color: T.muted }}>
      <Wrench size={38} color={T.border2} style={{ marginBottom: 14 }} />
      <div className="bc" style={{ fontSize: 19, fontWeight: 700, color: T.muted2, marginBottom: 6 }}>
        {label}
      </div>
      <p style={{ fontSize: 13 }}>Bu bo'lim 2-qismda qo'shiladi</p>
    </div>
  );
}

/* ═══════════════════════════════════════════════════
   DASHBOARD — 1 EKRAN
═══════════════════════════════════════════════════ */
function DashboardTab({ data, patch, rate, setTab }) {
  const [quickOpen, setQuickOpen] = useState(false);
  const [callOpen, setCallOpen] = useState(false);
  const [closeDayOpen, setCloseDayOpen] = useState(false);

  const todayCards = data.serviceCards.filter((c) => c.date === todayISO());
  const openCards = data.serviceCards.filter((c) => cardStatus(c) === "ochiq");
  const todayCF = data.cashflow.filter((c) => c.date === todayISO());
  const todayIncome = todayCF.filter((c) => c.type === "kirim").reduce((s, c) => s + num(c.amountSum), 0);
  const todayExpense = todayCF.filter((c) => c.type === "chiqim").reduce((s, c) => s + num(c.amountSum), 0);

  const allIncome = data.cashflow.filter((c) => c.type === "kirim").reduce((s, c) => s + num(c.amountSum), 0);
  const allExpense = data.cashflow.filter((c) => c.type === "chiqim").reduce((s, c) => s + num(c.amountSum), 0);

  const ustaPending = ustaPendingByName(data);
  const ustaPendingTotal = ustaPending.reduce((s, u) => s + u.amountSum, 0);

  const waitingLeads = (data.leads || []).filter((l) => l.stage === "yangi" || l.stage === "tasdiqlangan");
  const supDebt = supplierDebts(data).reduce((s, d) => s + Math.max(0, d.debtSum), 0);

  function closeDay() {
    patch((d) => {
      const groups = ustaPendingByName(d);
      d.ustaLedger.forEach((e) => { if (!e.paid) e.paid = true; });
      groups.forEach((g) => {
        d.cashflow.unshift({
          id: uid(), date: todayISO(), type: "chiqim", category: "Usta xizmat haqi",
          currency: "SUM", amount: g.amountSum, amountSum: g.amountSum, amountUsd: g.amountSum / rate,
          note: `${g.usta} — kun yakunida to'landi (${g.count} ta xizmat)`,
        });
      });
      return d;
    });
    setCloseDayOpen(false);
  }

  return (
    <div>
      {/* QUICK ACTIONS */}
      <div style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
        <Btn size="lg" onClick={() => setQuickOpen(true)}>
          <Plus size={17} /> Yangi karta ochish
        </Btn>
        <Btn size="lg" variant="ghost" onClick={() => setCallOpen(true)}>
          <PhoneCall size={16} /> Qo'ng'iroq qabul
        </Btn>
        {ustaPendingTotal > 0 && (
          <Btn size="lg" variant="gold" onClick={() => setCloseDayOpen(true)}>
            <Check size={16} /> Kun yakunlash ({fmtSum(ustaPendingTotal)})
          </Btn>
        )}
      </div>

      {/* MAIN STATS */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 13, marginBottom: 20 }}>
        <Stat label="Bugungi kartalar" value={todayCards.length + " ta"}
          sub={`${openCards.length} ta ochiq`} color={T.flame} Icon={Car} />
        <Stat label="Bugungi kirim" value={fmtSum(todayIncome)}
          sub={`Chiqim: ${fmtSum(todayExpense)}`} color={T.teal} Icon={TrendingUp} />
        <Stat label="Kassa balansi" value={fmtSum(allIncome - allExpense)}
          sub={fmtUsd((allIncome - allExpense) / rate)} color={T.gold} Icon={Wallet} />
        <Stat label="Usta haqi (to'lanmagan)" value={fmtSum(ustaPendingTotal)}
          sub={`${ustaPending.length} ta usta`} color={T.red} Icon={Wrench} />
      </div>

      <div className="grid-2b" style={{ marginBottom: 16 }}>
        {/* OCHIQ KARTALAR */}
        <Card title={`Ochiq kartalar (${openCards.length})`} Icon={Car} color={T.flame} pad={false}
          action={<Btn size="sm" variant="ghost" onClick={() => setTab("services")}>Barchasi <ArrowRight size={11} /></Btn>}>
          {openCards.length === 0 ? (
            <div style={{ padding: "30px 20px", textAlign: "center", color: T.muted, fontSize: 13 }}>
              Ochiq karta yo'q
            </div>
          ) : (
            <div style={{ padding: "12px 14px", display: "grid", gap: 8 }}>
              {openCards.slice(0, 5).map((c) => (
                <div key={c.id} style={{
                  background: T.s3, borderRadius: 9, padding: "11px 14px",
                  display: "flex", justifyContent: "space-between", alignItems: "center",
                  borderLeft: `3px solid ${SERVICE_COLORS[c.serviceType] || T.flame}`,
                }}>
                  <div>
                    <div className="mo" style={{ fontSize: 14, fontWeight: 700, color: T.flame }}>
                      {c.plate}
                    </div>
                    <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                      {c.carModel} · {c.serviceType} {c.usta && `· ${c.usta}`}
                    </div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div className="mo" style={{ fontSize: 12, color: T.muted2 }}>
                      {fmtSum(cardPartsCost(c) + cardUstaFeeSum(c))}
                    </div>
                    <div style={{ fontSize: 10, color: T.muted }}>hozircha</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>

        {/* KUTILAYOTGAN QO'NG'IROQLAR */}
        <Card title={`Kutilmoqda (${waitingLeads.length})`} Icon={PhoneCall} color={T.gold} pad={false}
          action={<Btn size="sm" variant="ghost" onClick={() => setTab("callcenter")}>Barchasi <ArrowRight size={11} /></Btn>}>
          {waitingLeads.length === 0 ? (
            <div style={{ padding: "30px 20px", textAlign: "center", color: T.muted, fontSize: 13 }}>
              Kutilayotgan mijoz yo'q
            </div>
          ) : (
            <div style={{ padding: "12px 14px", display: "grid", gap: 8 }}>
              {waitingLeads.slice(0, 5).map((l) => {
                const stage = LEAD_STAGES.find((s) => s.id === l.stage);
                return (
                  <div key={l.id} style={{
                    background: T.s3, borderRadius: 9, padding: "10px 13px",
                    borderLeft: `3px solid ${stage?.color}`,
                  }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: 13, fontWeight: 600 }}>{l.name || "Nomsiz"}</span>
                      <Badge color={stage?.color}>{stage?.label}</Badge>
                    </div>
                    <div className="mo" style={{ fontSize: 11, color: T.muted, marginTop: 3 }}>
                      {l.phone} {l.carModel && `· ${l.carModel}`}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </Card>
      </div>

      {/* BOTTOM ROW */}
      <div className="grid-3b">
        <Card title="Usta haqi — bugun" Icon={Wrench} color={T.red}>
          {ustaPending.length === 0 ? (
            <div style={{ color: T.teal, fontSize: 13, display: "flex", alignItems: "center", gap: 6 }}>
              <Check size={14} /> Hammasi to'langan
            </div>
          ) : ustaPending.map((u) => (
            <div key={u.usta} style={{
              display: "flex", justifyContent: "space-between", padding: "7px 0",
              borderBottom: `1px solid ${T.border}25`,
            }}>
              <span style={{ fontSize: 12.5, color: T.muted2 }}>{u.usta} ({u.count})</span>
              <span className="mo" style={{ fontSize: 12.5, fontWeight: 700, color: T.red }}>
                {fmtSum(u.amountSum)}
              </span>
            </div>
          ))}
        </Card>

        <Card title="Bugungi xizmatlar" Icon={Zap} color={T.teal}>
          {SERVICE_TYPES.map((t) => {
            const n = todayCards.filter((c) => c.serviceType === t).length;
            return (
              <div key={t} style={{
                display: "flex", justifyContent: "space-between", alignItems: "center",
                padding: "7px 0", borderBottom: `1px solid ${T.border}25`,
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ width: 6, height: 6, borderRadius: "50%", background: SERVICE_COLORS[t] }} />
                  <span style={{ fontSize: 12.5, color: T.muted2 }}>{t}</span>
                </div>
                <span className="mo" style={{ fontSize: 13, fontWeight: 700, color: n ? SERVICE_COLORS[t] : T.muted }}>
                  {n} ta
                </span>
              </div>
            );
          })}
        </Card>

        <Card title="Qarzlar" Icon={AlertTriangle} color={T.red}>
          {[
            ["Ta'minotchiga", supDebt, T.red],
            ["Hamkorlardan", partnerBalances(data).reduce((s, p) => s + Math.max(0, p.debtSum), 0), T.teal],
          ].map(([l, v, c]) => (
            <div key={l} style={{
              display: "flex", justifyContent: "space-between", padding: "9px 0",
              borderBottom: `1px solid ${T.border}25`,
            }}>
              <span style={{ fontSize: 12.5, color: T.muted2 }}>{l}</span>
              <span className="mo" style={{ fontSize: 13, fontWeight: 700, color: c }}>{fmtSum(v)}</span>
            </div>
          ))}
        </Card>
      </div>

      {quickOpen && <NewCardModal data={data} onClose={() => setQuickOpen(false)}
        onSave={(base) => { patch((d) => { d.serviceCards.unshift({ id: uid(), status: "ochiq", parts: [], ustaFeeEntries: [], ...base }); return d; }); setQuickOpen(false); }} />}
      {callOpen && <NewLeadModal onClose={() => setCallOpen(false)}
        onSave={(lead) => { patch((d) => { d.leads.unshift({ id: uid(), ...lead }); return d; }); setCallOpen(false); }} />}
      {closeDayOpen && (
        <Modal title="Kun yakunlash — usta haqi to'lash" onClose={() => setCloseDayOpen(false)}>
          <p style={{ color: T.muted, fontSize: 13, marginBottom: 14 }}>
            Quyidagi summalar to'landi deb belgilanadi va kassaga chiqim yoziladi:
          </p>
          <div style={{ display: "grid", gap: 8 }}>
            {ustaPending.map((u) => (
              <div key={u.usta} style={{
                display: "flex", justifyContent: "space-between",
                padding: "11px 14px", background: T.s3, borderRadius: 8,
              }}>
                <span style={{ fontWeight: 600 }}>{u.usta} <span style={{ color: T.muted, fontWeight: 400 }}>({u.count} ta)</span></span>
                <span className="mo" style={{ color: T.flame, fontWeight: 700 }}>{fmtSum(u.amountSum)}</span>
              </div>
            ))}
          </div>
          <div style={{
            marginTop: 12, padding: "12px 14px", background: T.redD,
            borderRadius: 8, display: "flex", justifyContent: "space-between",
          }}>
            <span style={{ fontWeight: 700 }}>JAMI</span>
            <span className="mo" style={{ color: T.red, fontWeight: 800, fontSize: 16 }}>{fmtSum(ustaPendingTotal)}</span>
          </div>
          <SaveBtn onClick={closeDay} color={T.teal}>
            <Check size={16} /> To'landi — kunni yopish
          </SaveBtn>
        </Modal>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════════════
   CALL CENTER — CRM VORONKA
═══════════════════════════════════════════════════ */
function CallCenterTab({ data, patch }) {
  const [addOpen, setAddOpen] = useState(false);
  const [search, setSearch] = useState("");
  const [stageFilter, setStageFilter] = useState("hammasi");

  const leads = (data.leads || [])
    .filter((l) => stageFilter === "hammasi" || l.stage === stageFilter)
    .filter((l) => (l.name + l.phone + l.carModel + l.plate).toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => (b.date + b.time).localeCompare(a.date + a.time));

  function moveStage(id, stage) {
    patch((d) => {
      const l = d.leads.find((x) => x.id === id);
      if (l) { l.stage = stage; l.updatedAt = todayISO(); }
      return d;
    });
  }
  async function deleteLead(id) {
    if (!(await askConfirm("Bu qo'ng'iroq yozuvi o'chirilsinmi?"))) return;
    patch((d) => { d.leads = d.leads.filter((l) => l.id !== id); return d; });
  }

  const counts = LEAD_STAGES.map((s) => ({
    ...s, n: (data.leads || []).filter((l) => l.stage === s.id).length,
  }));

  return (
    <div>
      <PageHeader Icon={PhoneCall} color={T.blue} title="Qo'ng'iroqlar markazi"
        sub={`Jami ${(data.leads || []).length} ta qo'ng'iroq`}
        action={<Btn onClick={() => setAddOpen(true)}><Plus size={15} /> Qo'ng'iroq qabul qilish</Btn>} />

      {/* VORONKA */}
      <div className="grid-4b" style={{ marginBottom: 18 }}>
        {counts.map((s) => (
          <div key={s.id} onClick={() => setStageFilter(stageFilter === s.id ? "hammasi" : s.id)} style={{
            background: stageFilter === s.id ? s.color + "18" : T.s1,
            border: `1px solid ${stageFilter === s.id ? s.color : T.border}`,
            borderRadius: 11, padding: "14px 16px", cursor: "pointer",
            borderTop: `3px solid ${s.color}`,
          }}>
            <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase", color: T.muted, marginBottom: 6 }}>
              {s.label}
            </div>
            <div className="mo bc" style={{ fontSize: 24, fontWeight: 800, color: s.color }}>{s.n}</div>
          </div>
        ))}
      </div>

      {/* SEARCH */}
      <div style={{ position: "relative", marginBottom: 14, maxWidth: 340 }}>
        <Search size={13} style={{ position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)", color: T.muted }} />
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Telefon, ism, mashina..."
          style={{ ...iSt, paddingLeft: 32 }} />
      </div>

      {/* LEAD CARDS */}
      {leads.length === 0 ? (
        <div style={{ padding: "50px 20px", textAlign: "center", color: T.muted, background: T.s1, borderRadius: 12, border: `1px solid ${T.border}` }}>
          <Phone size={32} color={T.border2} style={{ marginBottom: 12 }} />
          <p style={{ fontSize: 13 }}>Qo'ng'iroq yo'q</p>
        </div>
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(310px,1fr))", gap: 12 }}>
          {leads.map((l) => {
            const stage = LEAD_STAGES.find((s) => s.id === l.stage) || LEAD_STAGES[0];
            return (
              <div key={l.id} className="ch" style={{
                background: T.s1, border: `1px solid ${T.border}`,
                borderRadius: 12, overflow: "hidden",
                borderLeft: `3px solid ${stage.color}`,
                transition: "all .15s",
              }}>
                <div style={{ padding: "13px 15px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                    <div>
                      <div style={{ fontSize: 15, fontWeight: 700 }}>{l.name || "Nomsiz"}</div>
                      <a href={`tel:${l.phone}`} className="mo" style={{
                        fontSize: 13, color: T.flame, textDecoration: "none",
                        display: "flex", alignItems: "center", gap: 5, marginTop: 3,
                      }}>
                        <Phone size={11} /> {l.phone}
                      </a>
                    </div>
                    <Badge color={stage.color}>{stage.label}</Badge>
                  </div>

                  {(l.carModel || l.plate) && (
                    <div style={{ fontSize: 12, color: T.muted2, marginBottom: 6 }}>
                      🚗 {l.carModel} {l.plate && <span className="mo">· {l.plate}</span>}
                    </div>
                  )}
                  {l.issue && (
                    <div style={{
                      fontSize: 12, color: T.muted2, background: T.s3,
                      borderRadius: 7, padding: "8px 11px", marginBottom: 8,
                    }}>{l.issue}</div>
                  )}
                  <div className="mo" style={{ fontSize: 10.5, color: T.muted, display: "flex", alignItems: "center", gap: 5 }}>
                    <Clock size={10} /> {fmtDate(l.date)} {l.time}
                    {l.serviceType && <span style={{ color: SERVICE_COLORS[l.serviceType] }}>· {l.serviceType}</span>}
                  </div>
                </div>

                <div style={{
                  display: "flex", gap: 6, padding: "10px 14px",
                  borderTop: `1px solid ${T.border}`, background: T.s2, flexWrap: "wrap",
                }}>
                  {l.stage === "yangi" && (
                    <Btn size="sm" variant="gold" onClick={() => moveStage(l.id, "tasdiqlangan")}>
                      Tasdiqlash
                    </Btn>
                  )}
                  {(l.stage === "yangi" || l.stage === "tasdiqlangan") && (
                    <>
                      <Btn size="sm" variant="teal" onClick={() => moveStage(l.id, "keldi")}>
                        <Check size={11} /> Keldi
                      </Btn>
                      <Btn size="sm" variant="red" onClick={() => moveStage(l.id, "bekor")}>
                        Bekor
                      </Btn>
                    </>
                  )}
                  {(l.stage === "keldi" || l.stage === "bekor") && (
                    <Btn size="sm" variant="ghost" onClick={() => moveStage(l.id, "yangi")}>
                      <RefreshCw size={11} /> Qayta ochish
                    </Btn>
                  )}
                  <button onClick={() => deleteLead(l.id)} style={{
                    marginLeft: "auto", background: "none", border: "none",
                    cursor: "pointer", color: T.muted, padding: 4,
                  }}><Trash2 size={13} /></button>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {addOpen && <NewLeadModal onClose={() => setAddOpen(false)}
        onSave={(lead) => { patch((d) => { d.leads.unshift({ id: uid(), ...lead }); return d; }); setAddOpen(false); }} />}
    </div>
  );
}

function NewLeadModal({ onClose, onSave }) {
  const [f, setF] = useState({
    date: todayISO(), time: nowTime(), name: "", phone: "",
    carModel: "", plate: "", serviceType: "", issue: "", stage: "yangi",
  });
  const set = (k) => (e) => setF((s) => ({ ...s, [k]: e.target?.value ?? e }));

  return (
    <Modal title="Qo'ng'iroq qabul qilish" onClose={onClose} wide>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <F label="Mijoz ismi"><input style={iSt} value={f.name} onChange={set("name")} placeholder="Ism" autoFocus /></F>
        <F label="Telefon raqami *"><input style={iSt} value={f.phone} onChange={set("phone")} placeholder="+998 90 000 00 00" /></F>
        <F label="Avtomobil modeli"><input style={iSt} value={f.carModel} onChange={set("carModel")} placeholder="Nexia, Cobalt..." /></F>
        <F label="Davlat raqami"><input style={iSt} value={f.plate} onChange={set("plate")} placeholder="01 A 123 BC" /></F>
        <F label="Xizmat turi">
          <Sel value={f.serviceType} onChange={set("serviceType")}
            options={[{ value: "", label: "Aniqlanmagan" }, ...SERVICE_TYPES]} />
        </F>
        <F label="Sana / vaqt">
          <div style={{ display: "flex", gap: 8 }}>
            <input type="date" style={iSt} value={f.date} onChange={set("date")} />
            <input type="time" style={{ ...iSt, width: 100 }} value={f.time} onChange={set("time")} />
          </div>
        </F>
        <F label="Muammo / izoh" col="1/-1">
          <textarea rows={2} style={iSt} value={f.issue} onChange={set("issue")}
            placeholder="Mijoz nima dedi..." />
        </F>
      </div>
      <SaveBtn disabled={!f.phone.trim()} onClick={() => onSave(f)}>
        <PhoneCall size={15} /> Saqlash
      </SaveBtn>
    </Modal>
  );
}

/* ═══════════════════════════════════════════════════
   SERVICES TAB
═══════════════════════════════════════════════════ */
function ServicesTab({ data, patch, rate }) {
  const [newOpen, setNewOpen] = useState(false);
  const [workCard, setWorkCard] = useState(null);
  const [search, setSearch] = useState("");
  const [editPinOpen, setEditPinOpen] = useState(null); // card pending PIN check
  const [editCard, setEditCard] = useState(null); // card being edited after PIN ok

  const cards = data.serviceCards;
  const openCards = cards.filter((c) => cardStatus(c) === "ochiq");
  const closedCards = cards
    .filter((c) => cardStatus(c) !== "ochiq")
    .filter((c) => (c.plate + c.carModel + c.phone).toLowerCase().includes(search.toLowerCase()));

  const totalRevenue = closedCards.reduce((s, c) => s + num(c.finalTotal), 0);
  const totalProfit = closedCards.reduce((s, c) => s + num(c.profitSum), 0);

  function createCard(base) {
    patch((d) => {
      d.serviceCards.unshift({
        id: uid(), status: "ochiq", parts: [], ustaFeeEntries: [],
        partsCost: 0, ustaFee: 0, finalTotal: 0, profitSum: 0, ...base,
      });
      return d;
    });
    setNewOpen(false);
  }

  function addPart(cardId, part) {
    patch((d) => {
      const card = d.serviceCards.find((c) => c.id === cardId);
      const product = d.products.find((p) => p.id === part.productId);
      if (product) product.qty = Math.max(0, num(product.qty) - part.qty);
      if (card) { card.parts = card.parts || []; card.parts.push(part); }
      return d;
    });
  }
  function removePart(cardId, idx) {
    patch((d) => {
      const card = d.serviceCards.find((c) => c.id === cardId);
      if (!card) return d;
      const part = card.parts[idx];
      if (part) {
        const product = d.products.find((p) => p.id === part.productId);
        if (product) product.qty = num(product.qty) + part.qty;
      }
      card.parts = card.parts.filter((_, i) => i !== idx);
      return d;
    });
  }
  function addFee(cardId, amount, note) {
    patch((d) => {
      const card = d.serviceCards.find((c) => c.id === cardId);
      if (card) {
        card.ustaFeeEntries = card.ustaFeeEntries || [];
        card.ustaFeeEntries.push({ id: uid(), amount, note, date: todayISO() });
      }
      return d;
    });
  }
  function removeFee(cardId, feeId) {
    patch((d) => {
      const card = d.serviceCards.find((c) => c.id === cardId);
      if (card) card.ustaFeeEntries = (card.ustaFeeEntries || []).filter((e) => e.id !== feeId);
      return d;
    });
  }
  function closeCard(cardId, fin) {
    patch((d) => {
      const card = d.serviceCards.find((c) => c.id === cardId);
      if (!card) return d;
      Object.assign(card, fin, { status: "yakunlangan" });

      const isNasiya = card.paymentType === "Nasiya (qarzga)";
      const isContracted = (d.contractedMasters || []).some((m) => m.name === card.usta);
      card.ustaIsContracted = isContracted;

      if (fin.finalTotal > 0 && !isNasiya) {
        d.cashflow.unshift({
          id: uid(), date: todayISO(), type: "kirim", category: "Xizmat to'lovi",
          currency: "SUM", amount: fin.finalTotal, amountSum: fin.finalTotal,
          amountUsd: fin.finalTotal / rate, paymentType: card.paymentType,
          note: `${card.serviceType} — ${card.carModel || ""} (${card.plate})`,
        });
      }
      if (fin.finalTotal > 0 && isNasiya) {
        d.nasiyaDebts = d.nasiyaDebts || [];
        d.nasiyaDebts.push({
          id: uid(), date: todayISO(), cardId: card.id,
          plate: card.plate, phone: card.phone, carModel: card.carModel,
          serviceType: card.serviceType, amountSum: fin.finalTotal, paid: false,
        });
      }
      if (fin.docFee > 0) {
        d.cashflow.unshift({
          id: uid(), date: todayISO(), type: "chiqim", category: "Hujjat xarajati",
          currency: "SUM", amount: fin.docFee, amountSum: fin.docFee, amountUsd: fin.docFee / rate,
          note: `${card.plate} — hujjat xarajati`,
        });
      }
      if (fin.ustaFee > 0 && !isContracted) {
        // Oddiy usta — kunlik/xizmat bo'yicha haq, kassadan to'lanadi
        d.ustaLedger.unshift({
          id: uid(), date: todayISO(), usta: card.usta || "Noma'lum",
          cardId: card.id, amountSum: fin.ustaFee, paid: false,
        });
      }
      // Kelishilgan (oylik) usta bo'lsa — fin.ustaFee kassaga chiqim yozilmaydi,
      // profitSum hisobida allaqachon ayirilgan, lekin bu holda uni QAYTA qo'shib,
      // servis foydasiga aylantiramiz (chunki ustaga pul chiqmaydi).
      if (fin.ustaFee > 0 && isContracted) {
        card.profitSum = num(card.profitSum) + num(fin.ustaFee);
        card.contractedUstaBonus = num(fin.ustaFee); // hisobot uchun — qancha "qaytdi"
      }
      // Mijozdan B/U tovar qabul qilingan bo'lsa (chegirma evaziga) — bu tovar
      // hech qayerda ko'rinmay qolmasligi uchun skladga "B/U tovarlar" sifatida kiritamiz.
      (fin.buItems || []).forEach((b) => {
        const price = num(b.price);
        d.products.push({
          id: uid(), name: b.name, unit: "dona", category: "B/U tovarlar",
          costSum: price, priceSum: price, priceUsd: 0, qty: 1,
        });
      });
      return d;
    });
    setWorkCard(null);
  }
  async function deleteCard(id) {
    if (!(await askConfirm("Karta o'chirilsinmi?\nIchidagi mahsulotlar skladga qaytariladi.\nBu amalni ortga qaytarib bo'lmaydi."))) return;
    patch((d) => {
      const card = d.serviceCards.find((c) => c.id === id);
      if (card) {
        // Karta ochiq bo'lsimi, yakunlangan bo'lsimi — mahsulot skladdan
        // chiqarilgan bo'ladi, shuning uchun har doim qaytarib beramiz.
        (card.parts || []).forEach((p) => {
          const prod = d.products.find((x) => x.id === p.productId);
          if (prod) prod.qty = num(prod.qty) + p.qty;
        });
      }
      d.serviceCards = d.serviceCards.filter((x) => x.id !== id);
      d.ustaLedger = d.ustaLedger.filter((x) => !(x.cardId === id && !x.paid));
      return d;
    });
  }

  function saveEditedCard(cardId, updated) {
    patch((d) => {
      const card = d.serviceCards.find((c) => c.id === cardId);
      if (!card) return d;

      // Eski yakunlash yozuvlarini (kassa, usta hisobi, nasiya) bekor qilamiz —
      // faqat hali to'lanmagan (paid=false) izlarni, chunki to'langanini o'zgartirish
      // moliyaviy tarixni buzadi.
      d.ustaLedger = d.ustaLedger.filter((x) => !(x.cardId === cardId && !x.paid));
      d.nasiyaDebts = (d.nasiyaDebts || []).filter((n) => !(n.cardId === cardId && !n.paid));
      // Eski kassa yozuvini ham (agar hali "tuzatilmagan" bo'lsa) belgilab qo'yamiz —
      // xavfsizlik uchun o'chirmaymiz, faqat izoh bilan bekor qilingan deb belgilaymiz.
      d.cashflow.forEach((c) => {
        if (c.editedCardId === cardId) c.category = c.category + " (bekor qilingan)";
      });

      Object.assign(card, updated);

      const isNasiya = card.paymentType === "Nasiya (qarzga)";
      const isContracted = (d.contractedMasters || []).some((m) => m.name === card.usta);
      card.ustaIsContracted = isContracted;

      if (num(updated.finalTotal) > 0 && !isNasiya) {
        d.cashflow.unshift({
          id: uid(), date: card.date || todayISO(), type: "kirim", category: "Xizmat to'lovi (tahrirlangan)",
          currency: "SUM", amount: updated.finalTotal, amountSum: updated.finalTotal,
          amountUsd: updated.finalTotal / rate, paymentType: card.paymentType,
          editedCardId: cardId,
          note: `${card.serviceType} — ${card.carModel || ""} (${card.plate}) — TAHRIRLANGAN`,
        });
      }
      if (num(updated.finalTotal) > 0 && isNasiya) {
        d.nasiyaDebts = d.nasiyaDebts || [];
        d.nasiyaDebts.push({
          id: uid(), date: card.date || todayISO(), cardId: card.id,
          plate: card.plate, phone: card.phone, carModel: card.carModel,
          serviceType: card.serviceType, amountSum: updated.finalTotal, paidAmount: 0, paid: false,
        });
      }
      if (num(updated.ustaFee) > 0 && !isContracted) {
        d.ustaLedger.unshift({
          id: uid(), date: card.date || todayISO(), usta: card.usta || "Noma'lum",
          cardId: card.id, amountSum: updated.ustaFee, paid: false,
        });
      }
      if (num(updated.ustaFee) > 0 && isContracted) {
        card.profitSum = num(card.profitSum) + num(updated.ustaFee);
        card.contractedUstaBonus = num(updated.ustaFee);
      }
      return d;
    });
    setEditCard(null);
  }

  return (
    <div>
      <PageHeader Icon={Car} color={T.flame} title="Xizmat kartalari"
        sub={`${openCards.length} ta ochiq · ${closedCards.length} ta yakunlangan`}
        action={<Btn onClick={() => setNewOpen(true)}><Plus size={15} /> Yangi karta</Btn>} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 13, marginBottom: 20 }}>
        <Stat label="Ochiq kartalar" value={openCards.length + " ta"} color={T.gold} Icon={Clock} />
        <Stat label="Yakunlangan" value={closedCards.length + " ta"} color={T.blue} Icon={Car} />
        <Stat label="Jami tushum" value={fmtSum(totalRevenue)} sub={fmtUsd(totalRevenue / rate)} color={T.teal} Icon={Wallet} />
        <Stat label="Jami foyda" value={fmtSum(totalProfit)} color={T.flame} Icon={TrendingUp} />
      </div>

      {/* OCHIQ KARTALAR */}
      {openCards.length > 0 && (
        <div style={{ marginBottom: 20 }}>
          <div className="bc" style={{ fontSize: 15, fontWeight: 700, color: T.gold, marginBottom: 10, display: "flex", alignItems: "center", gap: 7 }}>
            <Clock size={15} /> Ochiq kartalar
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(290px,1fr))", gap: 12 }}>
            {openCards.map((c) => (
              <div key={c.id} className="ch" onClick={() => setWorkCard(c)} style={{
                background: T.s1, border: `1px solid ${T.border}`, borderRadius: 12,
                cursor: "pointer", overflow: "hidden", transition: "all .15s",
                borderLeft: `3px solid ${SERVICE_COLORS[c.serviceType] || T.flame}`,
              }}>
                <div style={{ padding: "13px 15px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                    <div>
                      <div className="mo" style={{ fontSize: 15, fontWeight: 700, color: T.flame }}>{c.plate}</div>
                      <div style={{ fontSize: 11.5, color: T.muted, marginTop: 2 }}>{c.carModel || "—"}</div>
                    </div>
                    <Badge color={SERVICE_COLORS[c.serviceType] || T.flame}>{c.serviceType}</Badge>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, color: T.muted, marginBottom: 8 }}>
                    <span>👤 {c.usta || "Usta yo'q"}</span>
                    <span className="mo">{fmtDate(c.date)}</span>
                  </div>
                  <div style={{ display: "flex", gap: 12, paddingTop: 9, borderTop: `1px solid ${T.border}` }}>
                    <div>
                      <div style={{ fontSize: 9.5, color: T.muted }}>MAHSULOT</div>
                      <div className="mo" style={{ fontSize: 12, fontWeight: 600 }}>{fmtSum(cardPartsCost(c))}</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 9.5, color: T.muted }}>USTA HAQI</div>
                      <div className="mo" style={{ fontSize: 12, fontWeight: 600, color: T.gold }}>{fmtSum(cardUstaFeeSum(c))}</div>
                    </div>
                  </div>
                </div>
                <div style={{
                  padding: "9px 15px", background: T.s2,
                  borderTop: `1px solid ${T.border}`, display: "flex",
                  alignItems: "center", justifyContent: "space-between",
                }}>
                  <span style={{ fontSize: 11.5, color: T.flame, fontWeight: 600, display: "flex", alignItems: "center", gap: 5 }}>
                    <PlayCircle size={12} /> Davom ettirish
                  </span>
                  <button onClick={(e) => { e.stopPropagation(); deleteCard(c.id); }} style={{
                    background: "none", border: "none", cursor: "pointer", color: T.muted, padding: 3,
                  }}><Trash2 size={12} /></button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* SEARCH + CLOSED */}
      <div style={{ position: "relative", marginBottom: 12, maxWidth: 320 }}>
        <Search size={13} style={{ position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)", color: T.muted }} />
        <input value={search} onChange={(e) => setSearch(e.target.value)}
          placeholder="Raqam, mashina, telefon..." style={{ ...iSt, paddingLeft: 32 }} />
      </div>

      <Card title={`Yakunlangan kartalar (${closedCards.length})`} Icon={Check} color={T.teal} pad={false}>
        <Tbl
          empty="Yakunlangan karta yo'q"
          cols={[
            { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
            { k: "plate", h: "Raqam", r: (r) => <span className="mo" style={{ fontWeight: 700, color: T.flame }}>{r.plate}</span> },
            { k: "phone", h: "Telefon", r: (r) => <span className="mo" style={{ fontSize: 12 }}>{r.phone || "—"}</span> },
            { k: "carModel", h: "Mashina" },
            { k: "serviceType", h: "Xizmat", r: (r) => <Badge color={SERVICE_COLORS[r.serviceType] || T.blue}>{r.serviceType}</Badge> },
            { k: "usta", h: "Usta", r: (r) => r.usta || "—" },
            { k: "partsCost", h: "Tan narx", r: (r) => <span style={{ color: T.muted2 }}>{fmtSum(cardPartsRealCost(r))}</span> },
            { k: "ustaFee", h: "Usta haqi", r: (r) => <span style={{ color: T.gold }}>{fmtSum(cardUstaFeeSum(r))}</span> },
            { k: "finalTotal", h: "Yakuniy", r: (r) => <span className="mo" style={{ fontWeight: 700 }}>{fmtSum(r.finalTotal)}</span> },
            { k: "profitSum", h: "Foyda", r: (r) => <span className="mo" style={{ fontWeight: 700, color: num(r.profitSum) >= 0 ? T.teal : T.red }}>{fmtSum(r.profitSum)}</span> },
            { k: "edit", h: "", r: (r) => <button onClick={() => setEditPinOpen(r)} style={{ background: "none", border: "none", cursor: "pointer", color: T.muted }}><Pencil size={13} /></button> },
            { k: "del", h: "", r: (r) => <button onClick={() => deleteCard(r.id)} style={{ background: "none", border: "none", cursor: "pointer", color: T.muted }}><Trash2 size={13} /></button> },
          ]}
          rows={closedCards}
        />
      </Card>

      {newOpen && <NewCardModal data={data} onClose={() => setNewOpen(false)} onSave={createCard} />}
      {workCard && (
        <CardWorkspace
          card={data.serviceCards.find((c) => c.id === workCard.id) || workCard}
          products={data.products}
          onClose={() => setWorkCard(null)}
          onAddPart={(p) => addPart(workCard.id, p)}
          onRemovePart={(i) => removePart(workCard.id, i)}
          onAddFee={(a, n) => addFee(workCard.id, a, n)}
          onRemoveFee={(id) => removeFee(workCard.id, id)}
          onFinalize={(fin) => closeCard(workCard.id, fin)}
        />
      )}
      {editPinOpen && (
        <SimplePinModal onClose={() => setEditPinOpen(null)}
          onSuccess={() => { setEditCard(editPinOpen); setEditPinOpen(null); }} />
      )}
      {editCard && (
        <EditFinishedCardModal
          card={editCard} products={data.products}
          onClose={() => setEditCard(null)}
          onSave={(updated) => saveEditedCard(editCard.id, updated)}
        />
      )}
    </div>
  );
}

function NewCardModal({ data, onClose, onSave }) {
  const ustaNames = ustaNameOptions(data);
  const [f, setF] = useState({
    date: todayISO(), plate: "", phone: "", carModel: "", usta: "",
    serviceType: "Servis", paymentType: PAYMENT_TYPES[0],
    hasWarranty: false, warrantyMonths: 6, oneTime: false,
  });
  const set = (k) => (e) => setF((s) => ({ ...s, [k]: e.target?.value ?? e }));
  const canSave = f.oneTime || f.plate.trim();

  return (
    <Modal title="Yangi xizmat kartasi" onClose={onClose} wide>
      <label style={{
        display: "flex", alignItems: "center", gap: 9, marginBottom: 14,
        padding: "11px 14px", background: T.s3, borderRadius: 9, cursor: "pointer", fontSize: 13,
      }}>
        <input type="checkbox" checked={f.oneTime}
          onChange={(e) => setF((s) => ({ ...s, oneTime: e.target.checked }))}
          style={{ accentColor: T.flame, width: 15, height: 15 }} />
        Bir martalik mijoz (davlat raqami shart emas)
      </label>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <F label="Sana"><input type="date" style={iSt} value={f.date} onChange={set("date")} /></F>
        <F label={f.oneTime ? "Davlat raqami" : "Davlat raqami *"}>
          <input style={iSt} value={f.plate} onChange={set("plate")} placeholder="01 A 123 BC" />
        </F>
        <F label="Telefon"><input style={iSt} value={f.phone} onChange={set("phone")} placeholder="+998 90 000 00 00" /></F>
        <F label="Avtomobil modeli"><input style={iSt} value={f.carModel} onChange={set("carModel")} /></F>
        <F label="Xizmat turi"><Sel value={f.serviceType} onChange={set("serviceType")} options={SERVICE_TYPES} /></F>
        <F label="To'lov turi"><Sel value={f.paymentType} onChange={set("paymentType")} options={PAYMENT_TYPES} /></F>
        <F label="Usta ismi" col="1/-1">
          <input style={iSt} value={f.usta} onChange={set("usta")} list="usta-list" placeholder="Usta ismi" />
          <datalist id="usta-list">{ustaNames.map((n) => <option key={n} value={n} />)}</datalist>
        </F>
      </div>

      <label style={{
        display: "flex", alignItems: "center", gap: 9, marginTop: 14,
        paddingTop: 14, borderTop: `1px solid ${T.border}`, cursor: "pointer", fontSize: 13,
      }}>
        <input type="checkbox" checked={f.hasWarranty}
          onChange={(e) => setF((s) => ({ ...s, hasWarranty: e.target.checked }))}
          style={{ accentColor: T.teal, width: 15, height: 15 }} />
        <ShieldCheck size={14} color={T.teal} /> Kafolat berilsin
        {f.hasWarranty && (
          <input type="number" style={{ ...iSt, width: 80, marginLeft: 8 }}
            value={f.warrantyMonths} onChange={set("warrantyMonths")} placeholder="oy" />
        )}
      </label>

      <p style={{ fontSize: 11.5, color: T.muted, marginTop: 12 }}>
        Karta "ochiq" holatda ochiladi. Ish davomida mahsulot va usta haqini qo'shib borasiz.
      </p>

      <SaveBtn disabled={!canSave} onClick={() => onSave({ ...f, usta: f.usta.trim(), plate: f.plate.trim() || "Bir martalik" })}>
        <Plus size={15} /> Kartani ochish
      </SaveBtn>
    </Modal>
  );
}

function EditFinishedCardModal({ card, products, onClose, onSave }) {
  const [f, setF] = useState({
    plate: card.plate || "", phone: card.phone || "", carModel: card.carModel || "",
    usta: card.usta || "", serviceType: card.serviceType, paymentType: card.paymentType,
    ustaFee: String(Math.round(num(card.ustaFee))),
    finalTotal: String(Math.round(num(card.finalTotal))),
    docFee: String(Math.round(num(card.docFee || 0))),
    hasWarranty: !!card.hasWarranty, warrantyMonths: card.warrantyMonths || 6,
  });
  const set = (k) => (e) => setF((s) => ({ ...s, [k]: e.target?.value ?? e }));

  const partsCost = cardPartsRealCost(card);
  const profitSum = num(f.finalTotal) - partsCost - num(f.ustaFee) - num(f.docFee);

  return (
    <Modal title={`Kartani tahrirlash — ${card.plate}`} onClose={onClose} wide>
      <div style={{ padding: "10px 14px", background: T.goldD, border: `1px solid ${T.gold}30`, borderRadius: 8, marginBottom: 16, fontSize: 12, color: T.gold, display: "flex", alignItems: "center", gap: 8 }}>
        <Lock size={13} /> Bu karta allaqachon yakunlangan. O'zgarishlar moliyaviy yozuvlarni yangilaydi.
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <F label="Davlat raqami"><input style={iSt} value={f.plate} onChange={set("plate")} /></F>
        <F label="Telefon"><input style={iSt} value={f.phone} onChange={set("phone")} /></F>
        <F label="Avtomobil modeli"><input style={iSt} value={f.carModel} onChange={set("carModel")} /></F>
        <F label="Usta ismi"><input style={iSt} value={f.usta} onChange={set("usta")} /></F>
        <F label="Xizmat turi"><Sel value={f.serviceType} onChange={set("serviceType")} options={SERVICE_TYPES} /></F>
        <F label="To'lov turi"><Sel value={f.paymentType} onChange={set("paymentType")} options={PAYMENT_TYPES} /></F>
      </div>

      <div style={{ marginTop: 14, paddingTop: 14, borderTop: `1px solid ${T.border}` }}>
        <div style={{ fontSize: 11, fontWeight: 700, color: T.muted2, textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 10 }}>
          Mahsulotlar (tan narx: {fmtSum(partsCost)}) — bu qismni o'zgartirish uchun avval kartani ochib qayta ishlang
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
          <F label="Usta xizmat haqi"><input type="number" style={iSt} value={f.ustaFee} onChange={set("ustaFee")} /></F>
          <F label="Hujjat xarajati"><input type="number" style={iSt} value={f.docFee} onChange={set("docFee")} /></F>
          <F label="Yakuniy summa"><input type="number" style={iSt} value={f.finalTotal} onChange={set("finalTotal")} /></F>
        </div>
      </div>

      <label style={{
        display: "flex", alignItems: "center", gap: 9, marginTop: 14,
        paddingTop: 14, borderTop: `1px solid ${T.border}`, cursor: "pointer", fontSize: 13,
      }}>
        <input type="checkbox" checked={f.hasWarranty}
          onChange={(e) => setF((s) => ({ ...s, hasWarranty: e.target.checked }))}
          style={{ accentColor: T.teal, width: 15, height: 15 }} />
        <ShieldCheck size={14} color={T.teal} /> Kafolat berilsin
        {f.hasWarranty && (
          <input type="number" style={{ ...iSt, width: 80, marginLeft: 8 }}
            value={f.warrantyMonths} onChange={set("warrantyMonths")} placeholder="oy" />
        )}
      </label>

      <div style={{ marginTop: 14, padding: "12px 14px", background: T.s3, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>
        <span style={{ fontSize: 12.5, color: T.muted }}>Yangi hisoblangan foyda</span>
        <span className="mo" style={{ fontWeight: 700, color: profitSum >= 0 ? T.teal : T.red }}>{fmtSum(profitSum)}</span>
      </div>

      <SaveBtn color={T.gold} onClick={() => onSave({
        plate: f.plate.trim(), phone: f.phone.trim(), carModel: f.carModel.trim(),
        usta: f.usta.trim(), serviceType: f.serviceType, paymentType: f.paymentType,
        ustaFee: num(f.ustaFee), docFee: num(f.docFee), finalTotal: num(f.finalTotal),
        profitSum, hasWarranty: f.hasWarranty, warrantyMonths: num(f.warrantyMonths),
      })}>
        <Save size={15} /> O'zgarishlarni saqlash
      </SaveBtn>
    </Modal>
  );
}

function CardWorkspace({ card, products, onClose, onAddPart, onRemovePart, onAddFee, onRemoveFee, onFinalize }) {
  const [productId, setProductId] = useState(products[0]?.id || "");
  const [qty, setQty] = useState(1);
  const [feeAmount, setFeeAmount] = useState("");
  const [feeNote, setFeeNote] = useState("");
  const [finalizing, setFinalizing] = useState(false);

  const parts = card.parts || [];
  const fees = card.ustaFeeEntries || [];
  const partsCost = cardPartsCost(card);
  const realPartsCost = cardPartsRealCost(card);
  const feeSum = cardUstaFeeSum(card);

  // Servis / Moy → sotish narxi; Ustanovka → tan narx
  const useSalePrice = card.serviceType === "Servis" || card.serviceType === "Moy bo'limi";

  function add() {
    const p = products.find((x) => x.id === productId);
    if (!p) return;
    const q = num(qty); if (q <= 0 || q > num(p.qty)) return;
    const unit = useSalePrice ? num(p.priceSum) : num(p.costSum);
    onAddPart({
      productId: p.id, name: p.name, qty: q,
      unitCost: unit, lineTotal: q * unit,
      costUnit: num(p.costSum), saleUnit: num(p.priceSum),
    });
    setQty(1);
  }

  function addFeeFn() {
    const a = num(feeAmount); if (a <= 0) return;
    onAddFee(a, feeNote.trim());
    setFeeAmount(""); setFeeNote("");
  }

  if (finalizing)
    return <FinalizeModal card={card} partsCost={partsCost} realPartsCost={realPartsCost} feeSum={feeSum}
      onBack={() => setFinalizing(false)} onClose={onClose} onSave={onFinalize} />;

  return (
    <Modal title={`${card.plate} — ${card.carModel || ""}`} onClose={onClose} wide>
      <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 16, flexWrap: "wrap" }}>
        <Badge color={T.gold}>OCHIQ</Badge>
        <Badge color={SERVICE_COLORS[card.serviceType]}>{card.serviceType}</Badge>
        <span style={{ fontSize: 12, color: T.muted }}>
          {card.usta ? `👤 ${card.usta}` : "Usta belgilanmagan"} · {fmtDate(card.date)}
        </span>
      </div>

      {/* MAHSULOT */}
      <div style={{ background: T.s2, border: `1px solid ${T.border}`, borderRadius: 11, padding: 15, marginBottom: 14 }}>
        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: ".08em", textTransform: "uppercase", color: T.muted2, marginBottom: 10 }}>
          Mahsulot qo'shish — {useSalePrice ? "sotish narxida" : "tan narxida"}
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: parts.length ? 12 : 0, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 170 }}>
            <Sel value={productId} onChange={(e) => setProductId(e.target.value)}
              options={products.length
                ? products.map((p) => ({ value: p.id, label: `${p.name} (${p.qty} ${p.unit})` }))
                : [{ value: "", label: "Sklad bo'sh" }]} />
          </div>
          <input type="number" style={{ ...iSt, width: 72 }} value={qty} onChange={(e) => setQty(e.target.value)} />
          <Btn onClick={add} size="md"><Plus size={14} /></Btn>
        </div>
        {parts.map((p, i) => (
          <div key={i} style={{
            display: "flex", justifyContent: "space-between", alignItems: "center",
            background: T.s3, borderRadius: 8, padding: "9px 13px", marginBottom: 6,
          }}>
            <span style={{ fontSize: 13 }}>{p.name} × {p.qty}</span>
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <span className="mo" style={{ fontSize: 12.5, color: T.muted2 }}>{fmtSum(p.lineTotal)}</span>
              <button onClick={() => onRemovePart(i)} style={{ background: "none", border: "none", cursor: "pointer", color: T.red }}>
                <Trash2 size={13} />
              </button>
            </div>
          </div>
        ))}
        <div style={{ display: "flex", justifyContent: "space-between", paddingTop: 11, marginTop: 6, borderTop: `1px solid ${T.border}` }}>
          <span style={{ fontSize: 12.5, color: T.muted }}>Jami mahsulot</span>
          <span className="mo" style={{ fontSize: 13, fontWeight: 700, color: T.flame }}>{fmtSum(partsCost)}</span>
        </div>
      </div>

      {/* USTA HAQI */}
      <div style={{ background: T.s2, border: `1px solid ${T.border}`, borderRadius: 11, padding: 15, marginBottom: 16 }}>
        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: ".08em", textTransform: "uppercase", color: T.muted2, marginBottom: 10 }}>
          Usta xizmat haqi
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: fees.length ? 12 : 0, flexWrap: "wrap" }}>
          <input type="number" style={{ ...iSt, flex: 1, minWidth: 110 }} placeholder="Summa"
            value={feeAmount} onChange={(e) => setFeeAmount(e.target.value)} />
          <input style={{ ...iSt, flex: 1, minWidth: 130 }} placeholder="Izoh (ixtiyoriy)"
            value={feeNote} onChange={(e) => setFeeNote(e.target.value)} />
          <Btn onClick={addFeeFn} size="md"><Plus size={14} /></Btn>
        </div>
        {fees.map((e) => (
          <div key={e.id} style={{
            display: "flex", justifyContent: "space-between", alignItems: "center",
            background: T.s3, borderRadius: 8, padding: "9px 13px", marginBottom: 6,
          }}>
            <span style={{ fontSize: 13 }}>{e.note || "Usta xizmat haqi"}</span>
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <span className="mo" style={{ fontSize: 12.5, color: T.gold }}>{fmtSum(e.amount)}</span>
              <button onClick={() => onRemoveFee(e.id)} style={{ background: "none", border: "none", cursor: "pointer", color: T.red }}>
                <Trash2 size={13} />
              </button>
            </div>
          </div>
        ))}
        <div style={{ display: "flex", justifyContent: "space-between", paddingTop: 11, marginTop: 6, borderTop: `1px solid ${T.border}` }}>
          <span style={{ fontSize: 12.5, color: T.muted }}>Jami usta haqi</span>
          <span className="mo" style={{ fontSize: 13, fontWeight: 700, color: T.gold }}>{fmtSum(feeSum)}</span>
        </div>
      </div>

      <SaveBtn onClick={() => setFinalizing(true)} color={T.teal}>
        <Check size={16} /> Kartani yakunlash
      </SaveBtn>
      <p style={{ fontSize: 11, color: T.muted, textAlign: "center", marginTop: 9 }}>
        Yopsangiz ham ma'lumotlar saqlanadi — keyinroq davom ettirasiz
      </p>
    </Modal>
  );
}

function FinalizeModal({ card, partsCost, realPartsCost, feeSum, onBack, onClose, onSave }) {
  const [agreedSum, setAgreedSum] = useState("");
  const [percent, setPercent] = useState(20);
  const [discount, setDiscount] = useState(0);
  const [docFee, setDocFee] = useState("");
  const [ustaFee, setUstaFee] = useState(String(Math.round(feeSum)));
  const [materialCost, setMaterialCost] = useState("");
  const [buItems, setBuItems] = useState([]);
  const [paymentType, setPaymentType] = useState(card.paymentType || PAYMENT_TYPES[0]);

  const uf = num(ustaFee);
  const disc = num(discount);
  const doc = num(docFee);
  const buTotal = buItems.reduce((s, b) => s + num(b.price), 0);

  let finalTotal = 0, effectiveParts = partsCost;

  if (card.serviceType === "Ustanovka") {
    finalTotal = num(agreedSum);
  } else if (card.serviceType === "Servis") {
    // ustama faqat usta haqiga
    finalTotal = Math.max(0, partsCost + uf * (1 + num(percent) / 100) - disc);
  } else if (card.serviceType === "Detailing") {
    effectiveParts = num(materialCost);
    finalTotal = Math.max(0, effectiveParts + uf * (1 + num(percent) / 100) - disc);
  } else {
    // Moy bo'limi — faqat mahsulot sotish narxi
    finalTotal = Math.max(0, partsCost - disc);
  }

  finalTotal = Math.max(0, finalTotal - buTotal);
  // Detailing uchun material xarajati qo'lda kiritiladi — bu allaqachon tan narx.
  // Boshqa turlarda (ayniqsa Servis/Moy bo'limi) mijozga sotish narxida yozilgan
  // bo'lishi mumkin, shuning uchun profit har doim haqiqiy tan narxdan hisoblanadi.
  const costForProfit = card.serviceType === "Detailing" ? effectiveParts : realPartsCost;
  const profitSum = finalTotal - costForProfit - uf - doc;

  return (
    <Modal title={`Yakunlash — ${card.plate}`} onClose={onClose} wide>
      <button onClick={onBack} style={{
        background: "none", border: "none", cursor: "pointer",
        color: T.muted, fontSize: 12, marginBottom: 14, padding: 0,
      }}>← Orqaga (davom ettirish)</button>

      <div style={{
        display: "flex", justifyContent: "space-between",
        padding: "11px 14px", background: T.s3, borderRadius: 8, marginBottom: 14,
      }}>
        <span style={{ fontSize: 12.5, color: T.muted }}>Jami mahsulot</span>
        <span className="mo" style={{ fontWeight: 700, color: T.flame }}>{fmtSum(partsCost)}</span>
      </div>

      <div style={{ marginBottom: 14 }}>
        <F label="To'lov turi — mijoz qanday to'laydi">
          <Sel value={paymentType} onChange={(e) => setPaymentType(e.target.value)} options={PAYMENT_TYPES} />
        </F>
        {paymentType === "Nasiya (qarzga)" && (
          <p style={{ fontSize: 11.5, color: T.gold, marginTop: 6 }}>
            ⚠ Bu summa kassaga yozilmaydi — Kassa → Nasiya qarzdorlar ro'yxatiga tushadi.
          </p>
        )}
        {paymentType === "Karta (Click/Payme)" && (
          <p style={{ fontSize: 11.5, color: T.purple, marginTop: 6 }}>
            ℹ Kassa balansiga kirmaydi — alohida Click/Payme statistikasida ko'rinadi.
          </p>
        )}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        {card.serviceType === "Ustanovka" && (
          <F label="Kelishilgan summa *">
            <input type="number" style={iSt} value={agreedSum} onChange={(e) => setAgreedSum(e.target.value)} autoFocus />
          </F>
        )}
        {card.serviceType === "Detailing" && (
          <F label="Material xarajati (qo'lda)">
            <input type="number" style={iSt} value={materialCost} onChange={(e) => setMaterialCost(e.target.value)} />
          </F>
        )}
        {card.serviceType !== "Moy bo'limi" && (
          <F label="Usta xizmat haqi">
            <input type="number" style={iSt} value={ustaFee} onChange={(e) => setUstaFee(e.target.value)} />
          </F>
        )}
        {(card.serviceType === "Servis" || card.serviceType === "Detailing") && (
          <F label="Ustama % (usta haqiga)">
            <input type="number" style={iSt} value={percent} onChange={(e) => setPercent(e.target.value)} />
          </F>
        )}
        <F label="Hujjat xarajati">
          <input type="number" style={iSt} value={docFee} onChange={(e) => setDocFee(e.target.value)} placeholder="0" />
        </F>
        <F label="Skidka (so'm)">
          <input type="number" style={iSt} value={discount} onChange={(e) => setDiscount(e.target.value)} />
        </F>
      </div>

      {/* B/U TOVAR */}
      {(card.serviceType === "Servis" || card.serviceType === "Ustanovka") && (
        <div style={{ marginTop: 16, paddingTop: 14, borderTop: `1px solid ${T.border}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <span style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: ".08em", textTransform: "uppercase", color: T.gold }}>
              B/U tovar qabul qilish
            </span>
            <Btn size="sm" variant="ghost" onClick={() => setBuItems((s) => [...s, { name: "", price: "" }])}>
              <Plus size={11} /> Qo'shish
            </Btn>
          </div>
          {buItems.map((b, i) => (
            <div key={i} style={{ display: "flex", gap: 8, marginBottom: 7 }}>
              <input style={{ ...iSt, flex: 2 }} placeholder="Mahsulot nomi" value={b.name}
                onChange={(e) => setBuItems((s) => s.map((x, j) => j === i ? { ...x, name: e.target.value } : x))} />
              <input type="number" style={{ ...iSt, flex: 1 }} placeholder="Qabul narxi" value={b.price}
                onChange={(e) => setBuItems((s) => s.map((x, j) => j === i ? { ...x, price: e.target.value } : x))} />
              <button onClick={() => setBuItems((s) => s.filter((_, j) => j !== i))} style={{
                background: "none", border: "none", cursor: "pointer", color: T.red, padding: "0 6px",
              }}><X size={15} /></button>
            </div>
          ))}
        </div>
      )}

      {/* SUMMARY */}
      <div style={{ background: T.s2, borderRadius: 10, padding: "14px 16px", marginTop: 16 }}>
        {[
          ["Mahsulot (mijozga)", effectiveParts, T.muted2],
          costForProfit !== effectiveParts && ["Mahsulot (tan narx)", costForProfit, T.muted],
          ["Usta haqi", uf, T.gold],
          ["Hujjat xarajati", doc, T.red],
          buTotal > 0 && ["B/U tovar chegirma", -buTotal, T.gold],
          ["━ YAKUNIY SUMMA", finalTotal, T.flame],
          ["Servis foydasi", profitSum, profitSum >= 0 ? T.teal : T.red],
        ].filter(Boolean).map(([l, v, c]) => (
          <div key={l} style={{
            display: "flex", justifyContent: "space-between", padding: "7px 0",
            borderBottom: `1px solid ${T.border}25`,
          }}>
            <span style={{
              fontSize: 12.5, color: String(l).includes("━") ? T.text : T.muted,
              fontWeight: String(l).includes("━") ? 700 : 400,
            }}>{String(l).replace("━ ", "")}</span>
            <span className="mo" style={{
              fontSize: String(l).includes("━") ? 15 : 12.5,
              fontWeight: 700, color: c,
            }}>{fmtSum(v)}</span>
          </div>
        ))}
      </div>

      <SaveBtn color={T.teal} onClick={() => onSave({
        partsCost: effectiveParts, ustaFee: uf, docFee: doc,
        finalTotal, profitSum, discount: disc, paymentType,
        buItems: buItems.filter((b) => b.name.trim()),
      })}>
        <Check size={16} /> Yakunlash va yopish
      </SaveBtn>
    </Modal>
  );
}

/* ═══════════════════════════════════════════════════
   AVTOGAZ v2 — QISM 2
   Part 1 fayliga shu komponentlarni qo'shing (App() dagi
   Placeholder qatorlarini shular bilan almashtiring):
   tab === "warehouse" -> <WarehouseTab .../>
   tab === "cashier"   -> <CashierTab .../>
   tab === "ustalar"   -> <UstaTab .../>
   tab === "warranty"  -> <WarrantyTab .../>
   tab === "partners"  -> <PartnersTab .../>
   tab === "analytics" -> <AnalyticsTab .../>
═══════════════════════════════════════════════════ */

/* ─── WAREHOUSE TAB — manba turi bilan ─── */
function WarehouseTab({ data, patch, rate }) {
  const [stockOpen, setStockOpen] = useState(false);
  const [saleOpen, setSaleOpen] = useState(false);
  const [catFilter, setCatFilter] = useState("barchasi");
  const [search, setSearch] = useState("");

  const categories = ["barchasi", ...(data.settings.categories || CATEGORIES_DEFAULT)];
  const products = data.products
    .filter((p) => catFilter === "barchasi" || p.category === catFilter)
    .filter((p) => p.name.toLowerCase().includes(search.toLowerCase()));

  const totalValueSum = data.products.reduce((s, p) => s + num(p.qty) * num(p.costSum), 0);

  function addStock(entry) {
    patch((d) => {
      let product = d.products.find((p) => p.id === entry.productId);
      const prevCost = product ? num(product.costSum) : 0;
      if (!product) {
        product = { id: uid(), name: entry.productName, unit: entry.unit, category: entry.category, costSum: 0, priceSum: entry.priceSum || 0, priceUsd: entry.priceUsd || 0, qty: 0 };
        d.products.push(product);
      }
      const prevQty = num(product.qty);
      const newQty = prevQty + entry.qty;
      product.costSum = newQty > 0 ? Math.round((prevQty * prevCost + entry.qty * entry.unitCostSum) / newQty) : entry.unitCostSum;
      product.qty = newQty;
      if (entry.updatePrice) { product.priceSum = entry.priceSum; product.priceUsd = entry.priceUsd; }

      d.stockIns.unshift({ id: uid(), date: entry.date, productId: product.id, productName: product.name, qty: entry.qty, unit: product.unit, currency: entry.currency, unitCostSum: entry.unitCostSum, totalSum: entry.totalSum, supplier: entry.supplier, paidSum: entry.paidSum, sourceType: entry.sourceType });

      if (entry.sourceType === "Ta'minotchi" && entry.paidSum > 0) {
        d.cashflow.unshift({ id: uid(), date: entry.date, type: "chiqim", category: "Ta'minotchiga to'lov", currency: "SUM", amount: entry.paidSum, amountSum: entry.paidSum, amountUsd: entry.paidSum / rate, supplier: entry.supplier, note: `${product.name} x${entry.qty} — kirim to'lovi` });
      }
      return d;
    });
  }

  function saveEdit(prod, log) {
    patch((d) => {
      d.editLog = d.editLog || [];
      d.editLog.push({ id: uid(), date: todayISO(), ...log });
      const idx = d.products.findIndex((p) => p.id === prod.id);
      if (idx >= 0) d.products[idx] = prod;
      return d;
    });
  }

  function addFreeSale(sale) {
    patch((d) => {
      sale.items.forEach((it) => {
        const product = d.products.find((p) => p.id === it.productId);
        if (product) product.qty = Math.max(0, num(product.qty) - it.qty);
      });
      d.cashflow.unshift({ id: uid(), date: sale.date, type: "kirim", category: "Erkin savdo", currency: "SUM", amount: sale.totalSum, amountSum: sale.totalSum, amountUsd: sale.totalSum / rate, note: `${sale.items.map((i) => `${i.name} x${i.qty}`).join(", ")} — ${sale.customer}` });
      return d;
    });
  }

  return (
    <div>
      <PageHeader Icon={Package} color={T.gold} title="Sklad"
        sub={<>Qiymati: <span style={{ color: T.gold, fontWeight: 600 }}>{fmtSum(totalValueSum)}</span></>}
        action={<div style={{ display: "flex", gap: 8 }}>
          <Btn variant="ghost" onClick={() => setSaleOpen(true)}><ShoppingCart size={14} /> Erkin savdo</Btn>
          <Btn onClick={() => setStockOpen(true)}><Plus size={15} /> Kirim qilish</Btn>
        </div>} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 13, marginBottom: 18 }}>
        <Stat label="Mahsulot turlari" value={data.products.length + " ta"} color={T.blue} Icon={Package} />
        <Stat label="Sklad qiymati" value={fmtSum(totalValueSum)} sub={fmtUsd(totalValueSum / rate)} color={T.gold} Icon={Wallet} />
        <Stat label="Kam qolgan (5<)" value={data.products.filter((p) => num(p.qty) < 5).length + " ta"} color={T.red} Icon={AlertTriangle} />
      </div>

      <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: 1, minWidth: 180 }}>
          <Search size={13} style={{ position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)", color: T.muted }} />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Mahsulot qidirish..." style={{ ...iSt, paddingLeft: 32 }} />
        </div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {categories.map((c) => (
            <button key={c} className={catFilter === c ? "" : "tab-btn"} onClick={() => setCatFilter(c)} style={{
              padding: "7px 12px", borderRadius: 8, cursor: "pointer", fontSize: 11.5, fontWeight: 500,
              border: `1px solid ${catFilter === c ? T.flame : T.border2}`,
              background: catFilter === c ? T.flameD : T.s1,
              color: catFilter === c ? T.flame : T.muted,
              boxShadow: T.shadowSm,
            }}>{c === "barchasi" ? "Barchasi" : c}</button>
          ))}
        </div>
      </div>

      <Card title={`Mahsulotlar (${products.length})`} Icon={Package} color={T.flame} pad={false}>
        <Tbl
          empty="Mahsulot yo'q"
          cols={[
            { k: "name", h: "Nomi", r: (r) => <span style={{ fontWeight: 500 }}>{r.name}</span> },
            { k: "category", h: "Kategoriya", r: (r) => <Badge color={T.blue}>{r.category || "—"}</Badge> },
            { k: "costSum", h: "Kelish narxi", r: (r) => fmtSum(r.costSum) },
            { k: "price", h: "Sotish narxi", r: (r) => <span className="mo" style={{ color: T.teal, fontSize: 12 }}>{fmtSum(r.priceSum)}{r.priceUsd ? ` · ${fmtUsd(r.priceUsd)}` : ""}</span> },
            { k: "qty", h: "Qoldiq", r: (r) => <span className="mo" style={{ fontWeight: 700, color: num(r.qty) < 5 ? T.red : T.text }}>{r.qty} {r.unit}</span> },
            { k: "val", h: "Qiymati", r: (r) => <span style={{ color: T.gold }}>{fmtSum(num(r.qty) * num(r.costSum))}</span> },
            { k: "edit", h: "", r: (r) => <PinGuardEdit item={r} onSave={saveEdit} /> },
          ]}
          rows={products}
        />
      </Card>

      <div style={{ marginTop: 16 }}>
        <Card title={`Kirim tarixi (${data.stockIns.length})`} Icon={Download} color={T.teal} pad={false}>
          <Tbl
            empty="Kirim yo'q"
            cols={[
              { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
              { k: "productName", h: "Mahsulot" },
              { k: "qty", h: "Miqdor" },
              { k: "sourceType", h: "Manba", r: (r) => <Badge color={r.sourceType === "Ta'minotchi" ? T.red : r.sourceType === "Insider servis" ? T.purple : T.teal}>{r.sourceType || "Ta'minotchi"}</Badge> },
              { k: "totalSum", h: "Jami", r: (r) => <span style={{ color: T.flame, fontWeight: 600 }}>{fmtSum(r.totalSum)}</span> },
              { k: "supplier", h: "Ta'minotchi" },
              { k: "debt", h: "Qarz", r: (r) => {
                  if (r.sourceType === "O'z mahsuloti") return <span style={{ color: T.teal }}>Qarzsiz</span>;
                  const d = num(r.totalSum) - num(r.paidSum);
                  return <span style={{ color: d > 0 ? T.red : T.teal, fontWeight: 600 }}>{fmtSum(d)}</span>;
                } },
            ]}
            rows={[...data.stockIns].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 60)}
          />
        </Card>
      </div>

      {(data.stockOuts || []).length > 0 && (
        <div style={{ marginTop: 16 }}>
          <Card title={`Chiqim tarixi — Insider servis (${data.stockOuts.length})`} Icon={Upload} color={T.gold} pad={false}>
            <Tbl
              empty="Chiqim yo'q"
              cols={[
                { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
                { k: "productName", h: "Mahsulot" },
                { k: "qty", h: "Miqdor", r: (r) => <span className="mo">{r.qty}</span> },
                { k: "partnerName", h: "Insider servis", r: (r) => <span style={{ fontWeight: 600 }}>{r.partnerName}</span> },
                { k: "amountSum", h: "Qiymati", r: (r) => <span style={{ color: T.purple, fontWeight: 600 }}>{fmtSum(r.amountSum)}</span> },
                { k: "reason", h: "Sabab", r: (r) => <span style={{ color: T.muted, fontSize: 12 }}>{r.reason}</span> },
              ]}
              rows={[...data.stockOuts].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 60)}
            />
          </Card>
        </div>
      )}

      {stockOpen && <StockInModal data={data} rate={rate} onClose={() => setStockOpen(false)} onSave={(e) => { addStock(e); setStockOpen(false); }} />}
      {saleOpen && <FreeSaleModal products={data.products} onClose={() => setSaleOpen(false)} onSave={(s) => { addFreeSale(s); setSaleOpen(false); }} />}
    </div>
  );
}

function PinGuardEdit({ item, onSave }) {
  const [pinOpen, setPinOpen] = useState(false);
  const [editOpen, setEditOpen] = useState(false);
  return (
    <>
      <button onClick={() => setPinOpen(true)} style={{ background: "none", border: "none", cursor: "pointer", color: T.muted, padding: 4 }}>
        <Pencil size={13} />
      </button>
      {pinOpen && <SimplePinModal onClose={() => setPinOpen(false)} onSuccess={() => { setPinOpen(false); setEditOpen(true); }} />}
      {editOpen && <EditProductModal item={item} onClose={() => setEditOpen(false)} onSave={(p, log) => { onSave(p, log); setEditOpen(false); }} />}
    </>
  );
}

function SimplePinModal({ onClose, onSuccess }) {
  const [pin, setPin] = useState("");
  const [err, setErr] = useState(false);
  function check() { if (pin === "9999") onSuccess(); else { setErr(true); setPin(""); } }
  return (
    <Modal title="🔐 Tasdiqlash kodi" onClose={onClose}>
      <p style={{ color: T.muted, fontSize: 13, marginBottom: 14 }}>Tahrirlash uchun kodni kiriting:</p>
      <F label="Kod">
        <input type="password" style={iSt} value={pin}
          onChange={(e) => { setPin(e.target.value); setErr(false); }}
          onKeyDown={(e) => e.key === "Enter" && check()} autoFocus />
      </F>
      {err && <p style={{ color: T.red, fontSize: 12, marginTop: 8 }}>Noto'g'ri kod</p>}
      <SaveBtn onClick={check} disabled={!pin}>Tasdiqlash</SaveBtn>
    </Modal>
  );
}

function EditProductModal({ item, onClose, onSave }) {
  const [f, setF] = useState({ ...item });
  const [reason, setReason] = useState("");
  const set = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));
  return (
    <Modal title={`Tahrirlash — ${item.name}`} onClose={onClose}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <F label="Nomi" col="1/-1"><input style={iSt} value={f.name} onChange={set("name")} /></F>
        <F label="Kelish narxi"><input type="number" style={iSt} value={f.costSum} onChange={set("costSum")} /></F>
        <F label="Sotish (so'm)"><input type="number" style={iSt} value={f.priceSum} onChange={set("priceSum")} /></F>
        <F label="Sotish (USD)"><input type="number" style={iSt} value={f.priceUsd || 0} onChange={set("priceUsd")} /></F>
        <F label="Qoldiq"><input type="number" style={iSt} value={f.qty} onChange={set("qty")} /></F>
        <F label="Sabab *" col="1/-1"><input style={iSt} value={reason} onChange={(e) => setReason(e.target.value)} placeholder="Narx oshdi, xato tuzatildi..." /></F>
      </div>
      <SaveBtn disabled={!reason.trim()} onClick={() => onSave(f, { before: item, after: f, reason, user: "Kassir" })}>Saqlash</SaveBtn>
    </Modal>
  );
}

function StockInModal({ data, rate, onClose, onSave }) {
  const [mode, setMode] = useState("existing");
  const [sourceType, setSourceType] = useState("Ta'minotchi");
  const [productId, setProductId] = useState(data.products[0]?.id || "");
  const [newName, setNewName] = useState("");
  const [unit, setUnit] = useState(UNITS[0]);
  const [category, setCategory] = useState((data.settings.categories || CATEGORIES_DEFAULT)[0]);
  const [qty, setQty] = useState(1);
  const [unitCost, setUnitCost] = useState("");
  const [priceSum, setPriceSum] = useState("");
  const [priceUsd, setPriceUsd] = useState("");
  const [updatePrice, setUpdatePrice] = useState(false);
  const [supplier, setSupplier] = useState("");
  const [paid, setPaid] = useState(0);
  const [currency, setCurrency] = useState("SUM");

  const unitCostSum = toSum(unitCost, currency, rate);
  const totalSum = unitCostSum * num(qty);
  const paidSum = toSum(paid, currency, rate);
  const debt = totalSum - paidSum;

  const existingProd = data.products.find((p) => p.id === productId);
  const prevCost = existingProd ? num(existingProd.costSum) : 0;
  const priceRose = prevCost > 0 && unitCostSum > prevCost;

  const canSave = (mode === "existing" ? !!productId : !!newName.trim()) && (sourceType === "O'z mahsuloti" || sourceType === "Insider servis" || supplier.trim());

  return (
    <Modal title="Skladga kirim" onClose={onClose} wide>
      <F label="Mahsulot manbai" col="1/-1">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
          {SOURCE_TYPES.map((s) => (
            <button key={s} onClick={() => setSourceType(s)} style={{
              padding: "10px 8px", borderRadius: 8, cursor: "pointer", fontSize: 12, fontWeight: 600,
              border: `1.5px solid ${sourceType === s ? T.flame : T.border2}`,
              background: sourceType === s ? T.flameD : T.s3,
              color: sourceType === s ? T.flame : T.muted2,
            }}>{s}</button>
          ))}
        </div>
        {sourceType === "O'z mahsuloti" && <p style={{ fontSize: 11, color: T.teal, marginTop: 6 }}>✓ Qarz hosil bo'lmaydi</p>}
        {sourceType === "Insider servis" && <p style={{ fontSize: 11, color: T.purple, marginTop: 6 }}>↔ Tovar bilan hisob-kitob</p>}
      </F>

      <div style={{ display: "flex", borderRadius: 8, overflow: "hidden", border: `1px solid ${T.border2}`, margin: "14px 0" }}>
        {[["existing", "Mavjud mahsulot"], ["new", "Yangi mahsulot"]].map(([id, l]) => (
          <button key={id} onClick={() => setMode(id)} style={{ flex: 1, padding: "9px", border: "none", cursor: "pointer", fontSize: 12.5, fontWeight: 500, background: mode === id ? T.flame : "transparent", color: mode === id ? "#fff" : T.muted }}>{l}</button>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        {mode === "existing" ? (
          <div style={{ gridColumn: "1/-1" }}>
            <F label="Mahsulot">
              <Sel value={productId} onChange={(e) => setProductId(e.target.value)} options={data.products.length ? data.products.map((p) => ({ value: p.id, label: `${p.name} (${p.qty} ${p.unit})` })) : [{ value: "", label: "Mahsulot yo'q" }]} />
            </F>
          </div>
        ) : (
          <>
            <F label="Nomi" col="1/-1"><input style={iSt} value={newName} onChange={(e) => setNewName(e.target.value)} /></F>
            <F label="Birlik"><Sel value={unit} onChange={(e) => setUnit(e.target.value)} options={UNITS} /></F>
            <F label="Kategoriya"><Sel value={category} onChange={(e) => setCategory(e.target.value)} options={data.settings.categories || CATEGORIES_DEFAULT} /></F>
          </>
        )}
        <F label="Miqdor"><input type="number" style={iSt} value={qty} onChange={(e) => setQty(e.target.value)} /></F>
        <F label="Valyuta"><CurrencyToggle value={currency} onChange={setCurrency} /></F>
        <F label={`Kelish narxi (${currency})`}><input type="number" style={iSt} value={unitCost} onChange={(e) => setUnitCost(e.target.value)} /></F>
        <F label="Sotish narxi (SO'M)"><input type="number" style={iSt} value={priceSum} onChange={(e) => setPriceSum(e.target.value)} /></F>
        <F label="Sotish narxi (USD)"><input type="number" style={iSt} value={priceUsd} onChange={(e) => setPriceUsd(e.target.value)} /></F>
        {sourceType !== "O'z mahsuloti" && (
          <F label={sourceType === "Insider servis" ? "Insider servis nomi" : "Ta'minotchi"}>
            <input style={iSt} value={supplier} onChange={(e) => setSupplier(e.target.value)} />
          </F>
        )}
        {sourceType === "Ta'minotchi" && (
          <F label="Hozir to'landi"><input type="number" style={iSt} value={paid} onChange={(e) => setPaid(e.target.value)} /></F>
        )}
      </div>

      {priceRose && (
        <div style={{ marginTop: 12, padding: "10px 14px", background: T.goldD, border: `1px solid ${T.gold}30`, borderRadius: 8 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: T.gold, display: "flex", alignItems: "center", gap: 6 }}>
            <AlertTriangle size={13} /> Narx oshgan! {fmtSum(prevCost)} → {fmtSum(unitCostSum)}
          </div>
          <label style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 8, cursor: "pointer", fontSize: 12 }}>
            <input type="checkbox" checked={updatePrice} onChange={(e) => setUpdatePrice(e.target.checked)} style={{ accentColor: T.gold }} />
            Sotish narxini ham yangilash
          </label>
        </div>
      )}

      {sourceType === "Ta'minotchi" && (
        <div style={{ marginTop: 14, padding: "12px 14px", background: T.s3, borderRadius: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0" }}>
            <span style={{ fontSize: 12, color: T.muted }}>Jami</span>
            <span className="mo" style={{ fontSize: 12, fontWeight: 600 }}>{fmtSum(totalSum)}</span>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0" }}>
            <span style={{ fontSize: 12, color: T.muted }}>Qarz</span>
            <span className="mo" style={{ fontSize: 12, fontWeight: 600, color: debt > 0 ? T.red : T.teal }}>{fmtSum(debt)}</span>
          </div>
        </div>
      )}

      <SaveBtn disabled={!canSave} onClick={() => onSave({
        productId: mode === "existing" ? productId : null,
        productName: mode === "existing" ? existingProd?.name : newName.trim(),
        unit: mode === "existing" ? existingProd?.unit : unit,
        category, qty: num(qty), currency, unitCostSum, totalSum,
        priceSum: num(priceSum), priceUsd: num(priceUsd), updatePrice,
        supplier: supplier.trim() || (sourceType === "O'z mahsuloti" ? "—" : "Noma'lum"),
        paidSum: sourceType === "Ta'minotchi" ? paidSum : totalSum,
        sourceType, date: todayISO(),
      })}>Kirim qilish</SaveBtn>
    </Modal>
  );
}

function FreeSaleModal({ products, onClose, onSave }) {
  const [customer, setCustomer] = useState("");
  const [cart, setCart] = useState([]);
  const [productId, setProductId] = useState(products[0]?.id || "");
  const [qty, setQty] = useState(1);
  const product = products.find((p) => p.id === productId);

  function add() {
    if (!product) return;
    const q = num(qty); if (q <= 0) return;
    setCart((s) => [...s, { productId: product.id, name: product.name, qty: q, lineTotalSum: q * num(product.priceSum) }]);
    setQty(1);
  }
  const total = cart.reduce((s, i) => s + i.lineTotalSum, 0);

  return (
    <Modal title="Erkin savdo" onClose={onClose} wide>
      <F label="Mijoz"><input style={iSt} value={customer} onChange={(e) => setCustomer(e.target.value)} placeholder="Mijoz ismi" /></F>
      <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
        <div style={{ flex: 2 }}>
          <Sel value={productId} onChange={(e) => setProductId(e.target.value)} options={products.length ? products.map((p) => ({ value: p.id, label: `${p.name} (${p.qty})` })) : [{ value: "", label: "Sklad bo'sh" }]} />
        </div>
        <input type="number" style={{ ...iSt, width: 70 }} value={qty} onChange={(e) => setQty(e.target.value)} />
        <Btn onClick={add}><Plus size={14} /></Btn>
      </div>
      {cart.map((i, idx) => (
        <div key={idx} style={{ display: "flex", justifyContent: "space-between", background: T.s3, borderRadius: 7, padding: "8px 12px", marginTop: 8 }}>
          <span style={{ fontSize: 13 }}>{i.name} x{i.qty}</span>
          <span className="mo">{fmtSum(i.lineTotalSum)}</span>
        </div>
      ))}
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 14, paddingTop: 14, borderTop: `1px solid ${T.border}` }}>
        <span style={{ fontWeight: 700 }}>Jami</span>
        <span className="mo" style={{ color: T.teal, fontWeight: 700 }}>{fmtSum(total)}</span>
      </div>
      <SaveBtn disabled={!cart.length} onClick={() => onSave({ date: todayISO(), customer: customer.trim() || "Mijoz", items: cart, totalSum: total })}>Saqlash</SaveBtn>
    </Modal>
  );
}

/* ─── CASHIER TAB — qarz turi + muddat eslatma + Click alohida ─── */
const DEBT_TX_TYPES = ["Ta'minotchi to'lovi", "Mahsulot uchun", "Xizmat uchun", "Avans", "Boshqa"];

function CashierTab({ data, patch, rate }) {
  const [open, setOpen] = useState(false);
  const [payStockIn, setPayStockIn] = useState(null);
  const [reportOpen, setReportOpen] = useState(false);
  const [payNasiya, setPayNasiya] = useState(null);
  const [givePersonalOpen, setGivePersonalOpen] = useState(false);
  const [payPersonal, setPayPersonal] = useState(null);
  const [editEntry, setEditEntry] = useState(null);

  const cf = data.cashflow;
  const clickEntries = cf.filter((c) => c.paymentType === "Karta (Click/Payme)");
  const clickTotal = clickEntries.reduce((s, c) => s + num(c.amountSum), 0);

  // Nasiya (kassadan tashqari) — kassa yozuvlaridan + xizmat kartalaridan
  const nasiyaCashEntries = cf.filter((c) => c.paymentType === "Nasiya (qarzga)");
  const nasiyaDebts = data.nasiyaDebts || [];
  const unpaidNasiya = nasiyaDebts.filter((n) => !n.paid);
  const nasiyaRemainingTotal = unpaidNasiya.reduce((s, n) => s + (num(n.amountSum) - num(n.paidAmount || 0)), 0);
  const nasiyaTotal = nasiyaCashEntries.reduce((s, c) => s + num(c.amountSum), 0) + nasiyaRemainingTotal;

  const cashFlow = cf.filter((c) => c.paymentType !== "Karta (Click/Payme)" && c.paymentType !== "Nasiya (qarzga)");
  const incomeSUM = cashFlow.filter((c) => c.type === "kirim" && c.currency !== "USD").reduce((s, c) => s + num(c.amountSum), 0);
  const expenseSUM = cashFlow.filter((c) => c.type === "chiqim" && c.currency !== "USD").reduce((s, c) => s + num(c.amountSum), 0);
  const incomeUSD = cashFlow.filter((c) => c.type === "kirim" && c.currency === "USD").reduce((s, c) => s + num(c.amount), 0);
  const expenseUSD = cashFlow.filter((c) => c.type === "chiqim" && c.currency === "USD").reduce((s, c) => s + num(c.amount), 0);

  const debts = supplierDebts(data);
  const totalDebt = debts.reduce((s, d) => s + Math.max(0, d.debtSum), 0);

  const dueSoon = (data.debtReminders || []).filter((r) => {
    const days = Math.ceil((new Date(r.dueDate) - new Date()) / 86400000);
    return days <= 2 && !r.resolved;
  });

  function addEntry(entry) {
    patch((d) => {
      d.cashflow.unshift({ id: uid(), ...entry });
      if (entry.reminderDate) {
        d.debtReminders = d.debtReminders || [];
        d.debtReminders.push({ id: uid(), dueDate: entry.reminderDate, note: entry.note, amountSum: entry.amountSum, txType: entry.debtTxType, resolved: false });
      }
      return d;
    });
  }

  function payDebt(supplierName, amountSum, meta = {}) {
    patch((d) => {
      let remaining = amountSum;
      const open = d.stockIns.filter((s) => s.supplier === supplierName && num(s.totalSum) - num(s.paidSum) > 0.5).sort((a, b) => (a.date || "").localeCompare(b.date || ""));
      for (const s of open) {
        if (remaining <= 0) break;
        const debt = num(s.totalSum) - num(s.paidSum);
        const pay = Math.min(debt, remaining);
        s.paidSum = num(s.paidSum) + pay;
        remaining -= pay;
      }
      d.cashflow.unshift({
        id: uid(), date: todayISO(), type: "chiqim", category: "Ta'minotchiga to'lov",
        currency: meta.currency || "SUM", amount: meta.amountOriginal ?? amountSum, amountSum,
        amountUsd: (meta.currency === "USD" ? meta.amountOriginal : amountSum / rate) || amountSum / rate,
        paymentType: meta.paymentType, supplier: supplierName,
        note: `Qarz to'lovi — ${supplierName}${meta.paymentType ? ` (${meta.paymentType})` : ""}`,
      });
      return d;
    });
  }

  function receiveNasiyaPayment(nasiyaId, amountSum) {
    patch((d) => {
      const n = (d.nasiyaDebts || []).find((x) => x.id === nasiyaId);
      if (n) {
        n.paidAmount = num(n.paidAmount) + amountSum;
        if (n.paidAmount >= num(n.amountSum) - 0.5) n.paid = true;
      }
      d.cashflow.unshift({
        id: uid(), date: todayISO(), type: "kirim", category: "Xizmat to'lovi",
        currency: "SUM", amount: amountSum, amountSum, amountUsd: amountSum / rate,
        paymentType: "Naqd pul",
        note: `Nasiya to'lovi (${n?.paid ? "to'liq" : "qisman"}) — ${n?.plate || ""} (${n?.serviceType || ""})`,
      });
      return d;
    });
  }

  function givePersonalDebt(item) {
    patch((d) => {
      d.personalDebts = d.personalDebts || [];
      d.personalDebts.push({
        id: uid(), date: todayISO(), name: item.name, note: item.note,
        amountSum: item.amountSum, paidAmount: 0, paid: false,
      });
      d.cashflow.unshift({
        id: uid(), date: todayISO(), type: "chiqim", category: "Shaxsiy qarz berildi",
        currency: "SUM", amount: item.amountSum, amountSum: item.amountSum, amountUsd: item.amountSum / rate,
        note: `${item.name} — naqd qarz${item.note ? ` (${item.note})` : ""}`,
      });
      return d;
    });
  }

  function receivePersonalPayment(debtId, amountSum) {
    patch((d) => {
      const p = (d.personalDebts || []).find((x) => x.id === debtId);
      if (p) {
        p.paidAmount = num(p.paidAmount) + amountSum;
        if (p.paidAmount >= num(p.amountSum) - 0.5) p.paid = true;
      }
      d.cashflow.unshift({
        id: uid(), date: todayISO(), type: "kirim", category: "Shaxsiy qarz qaytdi",
        currency: "SUM", amount: amountSum, amountSum, amountUsd: amountSum / rate,
        note: `${p?.name || ""} — qarz qaytdi (${p?.paid ? "to'liq" : "qisman"})`,
      });
      return d;
    });
  }

  return (
    <div>
      <PageHeader Icon={Wallet} color={T.teal} title="Kassa" sub="Naqd pul harakati va qarzlar"
        action={<div style={{ display: "flex", gap: 8 }}>
          <Btn variant="ghost" onClick={() => setReportOpen(true)}><Calendar size={14} /> Kun hisoboti</Btn>
          <Btn variant="gold" onClick={() => setGivePersonalOpen(true)}><Users size={14} /> Shaxsiy qarz berish</Btn>
          <Btn onClick={() => setOpen(true)}><Plus size={15} /> Yozuv qo'shish</Btn>
        </div>} />

      {dueSoon.length > 0 && (
        <div style={{ background: T.redD, border: `1px solid ${T.red}40`, borderRadius: 10, padding: "12px 16px", marginBottom: 16, display: "flex", alignItems: "center", gap: 10 }}>
          <AlertTriangle size={16} color={T.red} />
          <span style={{ fontSize: 13, color: T.red, fontWeight: 600 }}>{dueSoon.length} ta qarz muddati yaqinlashmoqda yoki o'tgan!</span>
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 13, marginBottom: 20 }}>
        <Stat label="Kassadagi SO'M" value={fmtSum(incomeSUM - expenseSUM)} color={T.flame} Icon={Wallet} />
        <Stat label="Kassadagi USD" value={fmtUsd(incomeUSD - expenseUSD)} color={T.gold} Icon={Wallet} />
        <Stat label="Click/Payme savdosi" value={fmtSum(clickTotal)} sub="balansdan tashqari" color={T.purple} Icon={TrendingUp} />
        <Stat label="Nasiya qarz (qolgan)" value={fmtSum(nasiyaTotal)} sub={`${unpaidNasiya.length} ta to'lanmagan · balansdan tashqari`} color={T.gold} Icon={Clock} />
        <Stat label="Ta'minotchi qarzi" value={fmtSum(totalDebt)} color={T.red} Icon={AlertTriangle} />
      </div>

      <Card title={`Kassa harakati (${cashFlow.length})`} Icon={Wallet} color={T.gold} pad={false}>
        <Tbl
          empty="Yozuv yo'q"
          cols={[
            { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
            { k: "type", h: "Turi", r: (r) => <Badge color={r.type === "kirim" ? T.teal : T.red}>{r.type === "kirim" ? "Kirim" : "Chiqim"}</Badge> },
            { k: "category", h: "Turkum" },
            { k: "currency", h: "Val", r: (r) => <span style={{ fontSize: 11, color: T.muted }}>{r.currency === "USD" ? "USD" : "SUM"}</span> },
            { k: "amount", h: "Summa", r: (r) => <span className="mo" style={{ fontWeight: 700, color: r.type === "kirim" ? T.teal : T.red }}>{r.type === "kirim" ? "+" : "-"}{r.currency === "USD" ? fmtUsd(r.amount) : fmtSum(r.amountSum)}</span> },
            { k: "note", h: "Izoh", r: (r) => <span style={{ color: T.muted, fontSize: 12 }}>{r.note || "—"}</span> },
            { k: "edit", h: "", r: (r) => <button onClick={() => setEditEntry(r)} style={{ background: "none", border: "none", cursor: "pointer", color: T.muted }}><Pencil size={13} /></button> },
            { k: "del", h: "", r: (r) => <button onClick={() => patch((d) => { d.cashflow = d.cashflow.filter((x) => x.id !== r.id); return d; })} style={{ background: "none", border: "none", cursor: "pointer", color: T.muted }}><Trash2 size={13} /></button> },
          ]}
          rows={[...cashFlow].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 80)}
        />
      </Card>

      <div style={{ marginTop: 16 }}>
        <Card title="Ta'minotchi qarzlari" Icon={Handshake} color={T.red} pad={false}>
          <Tbl
            empty="Qarz yo'q"
            cols={[
              { k: "name", h: "Ta'minotchi" },
              { k: "totalSum", h: "Jami", r: (r) => fmtSum(r.totalSum) },
              { k: "paidSum", h: "To'landi", r: (r) => <span style={{ color: T.teal }}>{fmtSum(r.paidSum)}</span> },
              { k: "debtSum", h: "Qarz", r: (r) => <span style={{ color: T.red, fontWeight: 700 }}>{fmtSum(r.debtSum)}</span> },
              { k: "act", h: "", r: (r) => <Btn size="sm" variant="teal" onClick={() => setPayStockIn(r)}>To'lov</Btn> },
            ]}
            rows={debts}
          />
        </Card>
      </div>

      <div style={{ marginTop: 16 }}>
        <Card title={`Nasiya qarzdorlar (${unpaidNasiya.length})`} Icon={AlertTriangle} color={T.red} pad={false}>
          <Tbl
            empty="Nasiya qarzdorlik yo'q"
            cols={[
              { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
              { k: "plate", h: "Raqam", r: (r) => <span className="mo" style={{ fontWeight: 700, color: T.flame }}>{r.plate}</span> },
              { k: "phone", h: "Telefon", r: (r) => <span className="mo" style={{ fontSize: 12 }}>{r.phone || "—"}</span> },
              { k: "carModel", h: "Mashina" },
              { k: "serviceType", h: "Xizmat", r: (r) => <Badge color={SERVICE_COLORS[r.serviceType] || T.blue}>{r.serviceType}</Badge> },
              { k: "amountSum", h: "Jami summa", r: (r) => <span className="mo" style={{ color: T.muted2 }}>{fmtSum(r.amountSum)}</span> },
              { k: "paidAmount", h: "To'langan", r: (r) => <span style={{ color: T.teal }}>{fmtSum(r.paidAmount || 0)}</span> },
              { k: "remaining", h: "Qolgan qarz", r: (r) => <span style={{ color: T.gold, fontWeight: 700 }}>{fmtSum(num(r.amountSum) - num(r.paidAmount || 0))}</span> },
              { k: "act", h: "", r: (r) => <Btn size="sm" variant="teal" onClick={() => setPayNasiya(r)}>To'lov qabul</Btn> },
            ]}
            rows={unpaidNasiya}
          />
        </Card>
      </div>

      <div style={{ marginTop: 16 }}>
        <Card title={`Shaxsiy qarzdorlar (${(data.personalDebts || []).filter((p) => !p.paid).length})`} Icon={Users} color={T.red} pad={false}>
          <Tbl
            empty="Shaxsiy qarzdorlik yo'q"
            cols={[
              { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
              { k: "name", h: "Ismi", r: (r) => <span style={{ fontWeight: 700 }}>{r.name}</span> },
              { k: "note", h: "Izoh", r: (r) => <span style={{ color: T.muted, fontSize: 12 }}>{r.note || "—"}</span> },
              { k: "amountSum", h: "Jami summa", r: (r) => <span className="mo" style={{ color: T.muted2 }}>{fmtSum(r.amountSum)}</span> },
              { k: "paidAmount", h: "To'langan", r: (r) => <span style={{ color: T.teal }}>{fmtSum(r.paidAmount || 0)}</span> },
              { k: "remaining", h: "Qolgan qarz", r: (r) => <span style={{ color: T.red, fontWeight: 700 }}>{fmtSum(num(r.amountSum) - num(r.paidAmount || 0))}</span> },
              { k: "act", h: "", r: (r) => <Btn size="sm" variant="teal" onClick={() => setPayPersonal(r)}>To'lov qabul</Btn> },
            ]}
            rows={(data.personalDebts || []).filter((p) => !p.paid).sort((a, b) => b.date.localeCompare(a.date))}
          />
        </Card>
      </div>

      <div style={{ marginTop: 16 }}>
        <DocFeesSection data={data} />
      </div>

      {open && <NewCashflowModal data={data} debts={debts} rate={rate} onClose={() => setOpen(false)} onSave={(e) => { addEntry(e); setOpen(false); }} />}
      {payStockIn && (
        <Modal title={`To'lov — ${payStockIn.name}`} onClose={() => setPayStockIn(null)}>
          <PaySupplierForm supplier={payStockIn} rate={rate}
            onSave={(amountSum, meta) => { payDebt(payStockIn.name, amountSum, meta); setPayStockIn(null); }} />
        </Modal>
      )}
      {payNasiya && (
        <NasiyaPayModal nasiya={payNasiya} onClose={() => setPayNasiya(null)}
          onSave={(amountSum) => { receiveNasiyaPayment(payNasiya.id, amountSum); setPayNasiya(null); }} />
      )}
      {givePersonalOpen && (
        <GivePersonalDebtModal onClose={() => setGivePersonalOpen(false)}
          onSave={(item) => { givePersonalDebt(item); setGivePersonalOpen(false); }} />
      )}
      {payPersonal && (
        <PayPersonalDebtModal debt={payPersonal} onClose={() => setPayPersonal(null)}
          onSave={(amountSum) => { receivePersonalPayment(payPersonal.id, amountSum); setPayPersonal(null); }} />
      )}
      {editEntry && (
        <EditCashEntryModal entry={editEntry} rate={rate} onClose={() => setEditEntry(null)}
          onSave={(updated) => {
            patch((d) => {
              const idx = d.cashflow.findIndex((c) => c.id === editEntry.id);
              if (idx >= 0) d.cashflow[idx] = { ...d.cashflow[idx], ...updated };
              return d;
            });
            setEditEntry(null);
          }} />
      )}
      {reportOpen && <DailyReport data={data} onClose={() => setReportOpen(false)} />}
    </div>
  );
}

function EditCashEntryModal({ entry, rate, onClose, onSave }) {
  const [type, setType] = useState(entry.type);
  const cats = type === "kirim" ? INCOME_CATEGORIES : EXPENSE_CATEGORIES;
  const [category, setCategory] = useState(cats.includes(entry.category) ? entry.category : cats[0]);
  const [currency, setCurrency] = useState(entry.currency || "SUM");
  const [amount, setAmount] = useState(String(entry.currency === "USD" ? entry.amount : Math.round(entry.amountSum)));
  const [paymentType, setPaymentType] = useState(entry.paymentType || "Naqd pul");
  const [note, setNote] = useState(entry.note || "");

  useEffect(() => {
    if (!cats.includes(category)) setCategory(cats[0]);
  }, [type]);

  const amountSum = toSum(amount, currency, rate);

  return (
    <Modal title="Kassa yozuvini tahrirlash" onClose={onClose} wide>
      <div style={{ padding: "10px 14px", background: T.goldD, border: `1px solid ${T.gold}30`, borderRadius: 8, marginBottom: 16, fontSize: 12, color: T.gold }}>
        Sana: {fmtDate(entry.date)} — asl yozuv o'zgartiriladi
      </div>

      <div style={{ display: "flex", borderRadius: 8, overflow: "hidden", border: `1px solid ${T.border2}`, marginBottom: 14 }}>
        <button onClick={() => setType("kirim")} style={{ flex: 1, padding: 9, border: "none", cursor: "pointer", fontWeight: 600, fontSize: 13, background: type === "kirim" ? T.teal : "transparent", color: type === "kirim" ? "#fff" : T.muted }}>Kirim</button>
        <button onClick={() => setType("chiqim")} style={{ flex: 1, padding: 9, border: "none", cursor: "pointer", fontWeight: 600, fontSize: 13, background: type === "chiqim" ? T.red : "transparent", color: type === "chiqim" ? "#fff" : T.muted }}>Chiqim</button>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <F label="Turkum"><Sel value={category} onChange={(e) => setCategory(e.target.value)} options={cats} /></F>
        <F label="To'lov usuli"><Sel value={paymentType} onChange={(e) => setPaymentType(e.target.value)} options={PAYMENT_TYPES} /></F>
        <F label="Valyuta"><CurrencyToggle value={currency} onChange={setCurrency} /></F>
        <F label={`Summa (${currency})`}><input type="number" style={iSt} value={amount} onChange={(e) => setAmount(e.target.value)} /></F>
        <F label="Izoh" col="1/-1"><input style={iSt} value={note} onChange={(e) => setNote(e.target.value)} /></F>
      </div>

      <div style={{ marginTop: 14, padding: "10px 14px", background: T.s3, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>
        <span style={{ fontSize: 12, color: T.muted }}>Yangi summa (so'm)</span>
        <span className="mo" style={{ fontWeight: 700 }}>{fmtSum(amountSum)}</span>
      </div>

      <SaveBtn color={T.gold} onClick={() => onSave({
        type, category, currency, amount: num(amount), amountSum, amountUsd: amountSum / rate,
        paymentType, note: note.trim(),
      })}>
        <Save size={15} /> O'zgarishlarni saqlash
      </SaveBtn>
    </Modal>
  );
}

function GivePersonalDebtModal({ onClose, onSave }) {
  const [name, setName] = useState("");
  const [note, setNote] = useState("");
  const [amount, setAmount] = useState("");
  const amtNum = num(amount);

  return (
    <Modal title="Shaxsiy qarz berish" onClose={onClose}>
      <p style={{ fontSize: 12.5, color: T.muted, marginBottom: 14 }}>
        Istalgan odamga (usta, ishchi, tanish) naqd pul berilganda shu yerdan kiriting.
        Summa darhol kassadan ayiriladi va qarzdorlar ro'yxatiga tushadi.
      </p>
      <div style={{ display: "grid", gap: 12 }}>
        <F label="Ismi *"><input style={iSt} value={name} onChange={(e) => setName(e.target.value)} autoFocus /></F>
        <F label="Izoh (ixtiyoriy)">
          <input style={iSt} value={note} onChange={(e) => setNote(e.target.value)} placeholder="Masalan: avans, shaxsiy ehtiyoj uchun..." />
        </F>
        <F label="Summa (so'm) *">
          <input type="number" style={iSt} value={amount} onChange={(e) => setAmount(e.target.value)} />
        </F>
      </div>
      {amtNum > 0 && (
        <div style={{ marginTop: 12, padding: "10px 14px", background: T.redD, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>
          <span style={{ fontSize: 12, color: T.red }}>Kassadan ayiriladi</span>
          <span className="mo" style={{ fontWeight: 700, color: T.red }}>{fmtSum(amtNum)}</span>
        </div>
      )}
      <SaveBtn disabled={!name.trim() || !amtNum} color={T.gold}
        onClick={() => onSave({ name: name.trim(), note: note.trim(), amountSum: amtNum })}>
        <Users size={15} /> Qarz berish
      </SaveBtn>
    </Modal>
  );
}

function PayPersonalDebtModal({ debt, onClose, onSave }) {
  const remaining = num(debt.amountSum) - num(debt.paidAmount || 0);
  const [amt, setAmt] = useState(String(Math.round(remaining)));
  const amtNum = num(amt);
  const after = Math.max(0, remaining - amtNum);

  return (
    <Modal title={`Qarz qaytarish — ${debt.name}`} onClose={onClose}>
      <div style={{ marginBottom: 14, padding: "10px 14px", background: T.redD, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>
        <span style={{ fontSize: 12, color: T.red }}>Qolgan qarz</span>
        <span className="mo" style={{ fontWeight: 700, color: T.red }}>{fmtSum(remaining)}</span>
      </div>
      {debt.note && <p style={{ fontSize: 12, color: T.muted, marginBottom: 14 }}>{debt.note}</p>}

      <F label="To'lov summasi (so'm)">
        <input type="number" style={iSt} value={amt} onChange={(e) => setAmt(e.target.value)} autoFocus onFocus={(e) => e.target.select()} />
      </F>

      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <Btn variant="ghost" size="sm" onClick={() => setAmt(String(Math.round(remaining)))}>To'liq summa</Btn>
        <Btn variant="ghost" size="sm" onClick={() => setAmt(String(Math.round(remaining / 2)))}>Yarmi</Btn>
      </div>

      <div style={{ marginTop: 14, padding: "10px 14px", background: T.s3, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>
        <span style={{ fontSize: 12, color: T.muted }}>To'lovdan keyin qoladigan qarz</span>
        <span className="mo" style={{ fontWeight: 700, color: after > 0 ? T.red : T.teal }}>{fmtSum(after)}</span>
      </div>

      <SaveBtn disabled={!amtNum || amtNum <= 0} color={T.teal}
        onClick={() => onSave(Math.min(amtNum, remaining))}>
        <Check size={15} /> To'lovni qabul qilish
      </SaveBtn>
    </Modal>
  );
}

function NasiyaPayModal({ nasiya, onClose, onSave }) {
  const remaining = num(nasiya.amountSum) - num(nasiya.paidAmount || 0);
  const [amt, setAmt] = useState(String(Math.round(remaining)));
  const amtNum = num(amt);
  const isFull = amtNum >= remaining - 0.5;
  const after = Math.max(0, remaining - amtNum);

  return (
    <Modal title={`Nasiya to'lovi — ${nasiya.plate}`} onClose={onClose}>
      <div style={{ marginBottom: 14, padding: "10px 14px", background: T.goldD, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>
        <span style={{ fontSize: 12, color: T.gold }}>Qolgan qarz</span>
        <span className="mo" style={{ fontWeight: 700, color: T.gold }}>{fmtSum(remaining)}</span>
      </div>
      <p style={{ fontSize: 12.5, color: T.muted, marginBottom: 16 }}>
        {nasiya.carModel} · {nasiya.serviceType} · {fmtDate(nasiya.date)}
        {num(nasiya.paidAmount) > 0 && <> · avval to'langan: <b>{fmtSum(nasiya.paidAmount)}</b></>}
      </p>

      <F label="To'lov summasi (so'm)">
        <input type="number" style={iSt} value={amt} onChange={(e) => setAmt(e.target.value)} autoFocus onFocus={(e) => e.target.select()} />
      </F>

      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <Btn variant="ghost" size="sm" onClick={() => setAmt(String(Math.round(remaining)))}>To'liq summa</Btn>
        <Btn variant="ghost" size="sm" onClick={() => setAmt(String(Math.round(remaining / 2)))}>Yarmi</Btn>
      </div>

      <div style={{ marginTop: 14, padding: "10px 14px", background: T.s3, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>
        <span style={{ fontSize: 12, color: T.muted }}>To'lovdan keyin qoladigan qarz</span>
        <span className="mo" style={{ fontWeight: 700, color: after > 0 ? T.red : T.teal }}>{fmtSum(after)}</span>
      </div>

      <SaveBtn disabled={!amtNum || amtNum <= 0} color={T.teal}
        onClick={() => onSave(Math.min(amtNum, remaining))}>
        <Check size={15} /> {isFull ? "To'liq to'lov qabul qilindi" : "Qisman to'lov qabul qilindi"}
      </SaveBtn>
    </Modal>
  );
}

/* ─── HUJJAT XARAJATLARI — alohida jadval ─── */
function DocFeesSection({ data }) {
  const [monthFilter, setMonthFilter] = useState("hammasi");

  const docRows = data.serviceCards
    .filter((c) => num(c.docFee) > 0)
    .map((c) => ({
      id: c.id, date: c.date, plate: c.plate, carModel: c.carModel,
      serviceType: c.serviceType, docFee: num(c.docFee),
    }))
    .sort((a, b) => b.date.localeCompare(a.date));

  const months = ["hammasi", ...new Set(docRows.map((r) => r.date.slice(0, 7)))].sort().reverse();
  const filtered = monthFilter === "hammasi" ? docRows : docRows.filter((r) => r.date.startsWith(monthFilter));
  const total = filtered.reduce((s, r) => s + r.docFee, 0);

  return (
    <Card title={`Hujjat xarajatlari (${filtered.length})`} Icon={Pencil} color={T.blue} pad={false}
      action={
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {months.length > 1 && (
            <Sel value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}
              style={{ padding: "5px 8px", fontSize: 11.5, width: 130 }}
              options={months.map((m) => ({ value: m, label: m === "hammasi" ? "Barcha oylar" : m }))} />
          )}
          <span className="mo" style={{ fontSize: 13, fontWeight: 700, color: T.flame }}>{fmtSum(total)}</span>
        </div>
      }>
      <Tbl
        empty="Hujjat xarajati kiritilmagan"
        cols={[
          { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
          { k: "plate", h: "Raqam", r: (r) => <span className="mo" style={{ fontWeight: 700, color: T.flame }}>{r.plate}</span> },
          { k: "carModel", h: "Mashina" },
          { k: "serviceType", h: "Xizmat", r: (r) => <Badge color={SERVICE_COLORS[r.serviceType] || T.blue}>{r.serviceType}</Badge> },
          { k: "docFee", h: "Summa", r: (r) => <span className="mo" style={{ fontWeight: 700, color: T.red }}>{fmtSum(r.docFee)}</span> },
        ]}
        rows={filtered}
      />
    </Card>
  );
}

function PaySupplierForm({ supplier, rate, onSave }) {
  const [currency, setCurrency] = useState("SUM");
  const [paymentType, setPaymentType] = useState("Naqd pul");
  const [amt, setAmt] = useState(String(Math.round(supplier.debtSum)));
  const amountSum = toSum(amt, currency, rate);

  return (
    <div>
      <div style={{ marginBottom: 14, padding: "10px 14px", background: T.redD, borderRadius: 8 }}>
        <span style={{ fontSize: 12, color: T.red }}>Joriy qarz: <b>{fmtSum(supplier.debtSum)}</b></span>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <F label="Valyuta"><CurrencyToggle value={currency} onChange={setCurrency} /></F>
        <F label="To'lov turi"><Sel value={paymentType} onChange={(e) => setPaymentType(e.target.value)} options={PAYMENT_TYPES} /></F>
      </div>

      <div style={{ marginTop: 12 }}>
        <F label={`Summa (${currency})`}>
          <input type="number" style={iSt} value={amt} onChange={(e) => setAmt(e.target.value)} autoFocus />
        </F>
      </div>

      {currency === "USD" && (
        <div style={{ marginTop: 10, padding: "9px 13px", background: T.s3, borderRadius: 7, display: "flex", justifyContent: "space-between" }}>
          <span style={{ fontSize: 11.5, color: T.muted }}>Ekvivalenti (so'm)</span>
          <span className="mo" style={{ fontSize: 12, fontWeight: 700 }}>{fmtSum(amountSum)}</span>
        </div>
      )}

      <SaveBtn disabled={!amt} onClick={() => onSave(amountSum, { currency, paymentType, amountOriginal: num(amt) })}>
        To'lovni saqlash
      </SaveBtn>
    </div>
  );
}

function NewCashflowModal({ data, debts, rate, onClose, onSave }) {
  const [type, setType] = useState("kirim");
  const cats = type === "kirim" ? INCOME_CATEGORIES : EXPENSE_CATEGORIES;
  const [category, setCategory] = useState(cats[0]);
  const [currency, setCurrency] = useState("SUM");
  const [amount, setAmount] = useState("");
  const [paymentType, setPaymentType] = useState("Naqd pul");
  const [supplier, setSupplier] = useState(debts[0]?.name || "");
  const [debtTxType, setDebtTxType] = useState(DEBT_TX_TYPES[0]);
  const [reminderOn, setReminderOn] = useState(false);
  const [reminderDate, setReminderDate] = useState("");
  const [note, setNote] = useState("");

  useEffect(() => setCategory((type === "kirim" ? INCOME_CATEGORIES : EXPENSE_CATEGORIES)[0]), [type]);

  const amountSum = toSum(amount, currency, rate);
  const isSupplierPay = category === "Ta'minotchiga to'lov";
  const isDebtRelated = isSupplierPay || category === "Rahbarga chiqim" || category === "Rahbardan kirim";

  return (
    <Modal title="Kassa yozuvi" onClose={onClose} wide>
      <div style={{ display: "flex", borderRadius: 8, overflow: "hidden", border: `1px solid ${T.border2}`, marginBottom: 14 }}>
        <button onClick={() => setType("kirim")} style={{ flex: 1, padding: 9, border: "none", cursor: "pointer", fontWeight: 600, fontSize: 13, background: type === "kirim" ? T.teal : "transparent", color: type === "kirim" ? "#fff" : T.muted }}>Kirim</button>
        <button onClick={() => setType("chiqim")} style={{ flex: 1, padding: 9, border: "none", cursor: "pointer", fontWeight: 600, fontSize: 13, background: type === "chiqim" ? T.red : "transparent", color: type === "chiqim" ? "#fff" : T.muted }}>Chiqim</button>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <F label="Turkum"><Sel value={category} onChange={(e) => setCategory(e.target.value)} options={cats} /></F>
        <F label="To'lov usuli"><Sel value={paymentType} onChange={(e) => setPaymentType(e.target.value)} options={PAYMENT_TYPES} /></F>
        <F label="Valyuta"><CurrencyToggle value={currency} onChange={setCurrency} /></F>
        <F label={`Summa (${currency})`}><input type="number" style={iSt} value={amount} onChange={(e) => setAmount(e.target.value)} /></F>

        {isSupplierPay && (
          <F label="Ta'minotchi" col="1/-1">
            {debts.length ? <Sel value={supplier} onChange={(e) => setSupplier(e.target.value)} options={debts.map((d) => d.name)} /> : <input style={iSt} value={supplier} onChange={(e) => setSupplier(e.target.value)} placeholder="Ta'minotchi nomi" />}
          </F>
        )}
        {isDebtRelated && (
          <F label="Qarz muomila turi" col="1/-1">
            <Sel value={debtTxType} onChange={(e) => setDebtTxType(e.target.value)} options={DEBT_TX_TYPES} />
          </F>
        )}
        <F label="Izoh" col="1/-1"><input style={iSt} value={note} onChange={(e) => setNote(e.target.value)} /></F>
      </div>

      {isDebtRelated && (
        <div style={{ marginTop: 14, paddingTop: 14, borderTop: `1px solid ${T.border}` }}>
          <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", fontSize: 13, marginBottom: reminderOn ? 10 : 0 }}>
            <input type="checkbox" checked={reminderOn} onChange={(e) => setReminderOn(e.target.checked)} style={{ accentColor: T.gold }} />
            <Calendar size={14} color={T.gold} /> Muddat eslatmasi qo'shish
          </label>
          {reminderOn && (
            <F label="Qachongacha to'lash kerak"><input type="date" style={iSt} value={reminderDate} onChange={(e) => setReminderDate(e.target.value)} /></F>
          )}
        </div>
      )}

      {paymentType === "Karta (Click/Payme)" && (
        <p style={{ fontSize: 11.5, color: T.purple, marginTop: 10 }}>
          ℹ Bu yozuv kassa balansiga kirmaydi — faqat Click/Payme savdosi sifatida hisoblanadi.
        </p>
      )}

      <SaveBtn onClick={() => onSave({
        date: todayISO(), type, category, currency, amount: num(amount), amountSum, amountUsd: amountSum / rate,
        paymentType, supplier: isSupplierPay ? supplier : undefined,
        debtTxType: isDebtRelated ? debtTxType : undefined,
        reminderDate: reminderOn ? reminderDate : undefined, note: note.trim(),
      })}>Saqlash</SaveBtn>
    </Modal>
  );
}

function DailyReport({ data, onClose }) {
  const todayCF = data.cashflow.filter((c) => c.date === todayISO());
  const todaySC = data.serviceCards.filter((c) => c.date === todayISO());
  const income = todayCF.filter((c) => c.type === "kirim" && c.paymentType !== "Karta (Click/Payme)").reduce((s, c) => s + num(c.amountSum), 0);
  const expense = todayCF.filter((c) => c.type === "chiqim" && c.paymentType !== "Karta (Click/Payme)").reduce((s, c) => s + num(c.amountSum), 0);
  const click = todayCF.filter((c) => c.paymentType === "Karta (Click/Payme)").reduce((s, c) => s + num(c.amountSum), 0);

  return (
    <Modal title={`Kun hisoboti — ${fmtDate(todayISO())}`} onClose={onClose} wide>
      <div style={{ display: "grid", gap: 7 }}>
        {[
          ["Jami kartalar", todaySC.length + " ta", T.text],
          ["Servis", todaySC.filter((c) => c.serviceType === "Servis").length + " ta", T.teal],
          ["Ustanovka", todaySC.filter((c) => c.serviceType === "Ustanovka").length + " ta", T.flame],
          ["Detailing", todaySC.filter((c) => c.serviceType === "Detailing").length + " ta", T.purple],
          ["Moy bo'limi", todaySC.filter((c) => c.serviceType === "Moy bo'limi").length + " ta", T.gold],
          ["━ Bugungi naqd kirim", fmtSum(income), T.teal],
          ["━ Bugungi naqd chiqim", fmtSum(expense), T.red],
          ["━ Click/Payme savdosi", fmtSum(click), T.purple],
          ["━ Kun balansi", fmtSum(income - expense), T.flame],
        ].map(([l, v, c]) => (
          <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "9px 13px", borderRadius: 7, background: String(l).includes("━") ? T.flameD : T.s3 }}>
            <span style={{ fontSize: 12.5, fontWeight: String(l).includes("━") ? 700 : 400, color: String(l).includes("━") ? T.flame : T.muted }}>{String(l).replace("━ ", "")}</span>
            <span className="mo" style={{ fontSize: 13, fontWeight: 700, color: c }}>{v}</span>
          </div>
        ))}
      </div>
      <p style={{ fontSize: 11, color: T.muted, textAlign: "center", marginTop: 14 }}>Skrinshot qilib rahbarga yuboring</p>
    </Modal>
  );
}

/* ─── USTA HISOB-KITOBI TAB ─── */
function UstaTab({ data, patch, rate, canManage = true }) {
  const [addContractedOpen, setAddContractedOpen] = useState(false);
  const pendingByName = ustaPendingByName(data);
  const pendingGrouped = ustaPendingGrouped(data);
  const totalPending = pendingByName.reduce((s, u) => s + u.amountSum, 0);
  const paidHistory = data.ustaLedger.filter((e) => e.paid).sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const contractedMasters = data.contractedMasters || [];

  function closeGroup(g) {
    patch((d) => {
      d.ustaLedger.forEach((e) => { if (g.ids.includes(e.id)) e.paid = true; });
      d.cashflow.unshift({ id: uid(), date: todayISO(), type: "chiqim", category: "Usta xizmat haqi", currency: "SUM", amount: g.amountSum, amountSum: g.amountSum, amountUsd: g.amountSum / rate, note: g.date ? `${g.usta} — ${fmtDate(g.date)} (${g.count} ta)` : `${g.usta} — (${g.count} ta)` });
      return d;
    });
  }

  function addContractedMaster(name) {
    patch((d) => {
      d.contractedMasters = d.contractedMasters || [];
      if (!d.contractedMasters.some((m) => m.name === name)) {
        d.contractedMasters.push({ id: uid(), name });
      }
      return d;
    });
  }

  function removeContractedMaster(id) {
    patch((d) => { d.contractedMasters = (d.contractedMasters || []).filter((m) => m.id !== id); return d; });
  }

  // shu ustalarning bu oyda servisga qo'shgan foydasi (bonus)
  const contractedBonusThisMonth = data.serviceCards
    .filter((c) => c.ustaIsContracted && (c.date || "").startsWith(currentMonthKey()))
    .reduce((s, c) => s + num(c.contractedUstaBonus), 0);

  return (
    <div>
      <PageHeader Icon={Wrench} color={T.gold} title="Usta hisobi" sub="Ustalarning xizmat haqi va to'lovlar" />
      {!canManage && (
        <div style={{ background: T.s2, border: `1px solid ${T.border2}`, borderRadius: 10, padding: "11px 15px", marginBottom: 16, fontSize: 12, color: T.muted, display: "flex", alignItems: "center", gap: 8 }}>
          <Lock size={13} color={T.flame} /> Faqat ko'rish — to'lovlarni yopish uchun Kassirga murojaat qiling
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 13, marginBottom: 20 }}>
        <Stat label="Kutilayotgan jami" value={fmtSum(totalPending)} sub={fmtUsd(totalPending / rate)} color={T.gold} Icon={Wrench} />
        <Stat label="Qarzdor ustalar" value={pendingByName.length + " ta"} color={T.blue} Icon={Users} />
        <Stat label="To'langan yozuvlar" value={paidHistory.length + " ta"} color={T.teal} Icon={Check} />
        <Stat label="Kelishilgan bonus (bu oy)" value={fmtSum(contractedBonusThisMonth)} sub="servis foydasiga qo'shildi" color={T.purple} Icon={Star} />
      </div>

      {canManage && (
        <div style={{ marginBottom: 16 }}>
          <Card title="Kelishilgan (oylik) ustalar" Icon={Wrench} color={T.purple}
            action={<Btn size="sm" onClick={() => setAddContractedOpen(true)}><Plus size={12} /> Qo'shish</Btn>}>
            <p style={{ fontSize: 12, color: T.muted, marginBottom: contractedMasters.length ? 12 : 0 }}>
              Bu ro'yxatdagi ustalarning xizmat haqi kassadan to'lanmaydi — <b>servis foydasi</b> sifatida hisoblanadi
              (chunki ular oylik maosh oladi, kunlik emas).
            </p>
            {contractedMasters.length > 0 && (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {contractedMasters.map((m) => (
                  <div key={m.id} style={{
                    display: "flex", alignItems: "center", gap: 8,
                    background: T.purpleD, border: `1px solid ${T.purple}30`,
                    borderRadius: 20, padding: "6px 8px 6px 14px", fontSize: 12.5, fontWeight: 600,
                  }}>
                    {m.name}
                    <button onClick={() => removeContractedMaster(m.id)} style={{
                      background: "none", border: "none", cursor: "pointer", color: T.muted, padding: 2, display: "flex",
                    }}><X size={13} /></button>
                  </div>
                ))}
              </div>
            )}
          </Card>
        </div>
      )}

      <Card title="Usta bo'yicha holat" Icon={BarChart3} color={T.teal} pad={false}>
        <Tbl
          empty="Kutilayotgan to'lov yo'q"
          cols={[
            { k: "usta", h: "Usta", r: (r) => <span style={{ fontWeight: 600 }}>{r.usta}</span> },
            { k: "count", h: "Xizmatlar" },
            { k: "amountSum", h: "Jami", r: (r) => <span style={{ color: T.flame, fontWeight: 700 }}>{fmtSum(r.amountSum)}</span> },
            ...(canManage ? [{ k: "act", h: "", r: (r) => <Btn size="sm" variant="teal" onClick={() => closeGroup(r)}>Yopish</Btn> }] : []),
          ]}
          rows={pendingByName}
        />
      </Card>
      <div style={{ marginTop: 16 }}>
        <Card title="Kun bo'yicha yig'ilgan" Icon={Calendar} color={T.gold} pad={false}>
          <Tbl
            empty="Yozuv yo'q"
            cols={[
              { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
              { k: "usta", h: "Usta" },
              { k: "count", h: "Soni" },
              { k: "amountSum", h: "Summa", r: (r) => <span style={{ fontWeight: 700 }}>{fmtSum(r.amountSum)}</span> },
              ...(canManage ? [{ k: "act", h: "", r: (r) => <Btn size="sm" variant="gold" onClick={() => closeGroup(r)}>Kunni yopish</Btn> }] : []),
            ]}
            rows={pendingGrouped}
          />
        </Card>
      </div>
      <div style={{ marginTop: 16 }}>
        <Card title="To'langan tarix" Icon={Check} color={T.teal} pad={false}>
          <Tbl
            empty="Ma'lumot yo'q"
            cols={[
              { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
              { k: "usta", h: "Usta" },
              { k: "amountSum", h: "Summa", r: (r) => <span style={{ color: T.teal, fontWeight: 700 }}>{fmtSum(r.amountSum)}</span> },
            ]}
            rows={paidHistory.slice(0, 50)}
          />
        </Card>
      </div>

      {addContractedOpen && (
        <AddContractedMasterModal ustaNames={ustaNameOptions(data)}
          onClose={() => setAddContractedOpen(false)}
          onSave={(name) => { addContractedMaster(name); setAddContractedOpen(false); }} />
      )}
    </div>
  );
}

function AddContractedMasterModal({ ustaNames, onClose, onSave }) {
  const [name, setName] = useState("");
  return (
    <Modal title="Kelishilgan usta qo'shish" onClose={onClose}>
      <p style={{ fontSize: 12.5, color: T.muted, marginBottom: 14 }}>
        Usta ismini kiriting yoki mavjud ustalar ro'yxatidan tanlang. Bu usta uchun xizmat haqi
        endi kassadan to'lanmaydi — servis foydasiga qo'shiladi.
      </p>
      <F label="Usta ismi">
        <input style={iSt} value={name} onChange={(e) => setName(e.target.value)}
          list="contracted-usta-list" placeholder="Masalan: Ravshan aka" autoFocus />
        <datalist id="contracted-usta-list">{ustaNames.map((n) => <option key={n} value={n} />)}</datalist>
      </F>
      <SaveBtn disabled={!name.trim()} color={T.purple} onClick={() => onSave(name.trim())}>
        <Plus size={15} /> Qo'shish
      </SaveBtn>
    </Modal>
  );
}

/* ─── WARRANTY TAB ─── */
function warrantyStatus(card) {
  if (!card.hasWarranty) return null;
  const expiry = new Date(card.date);
  expiry.setMonth(expiry.getMonth() + num(card.warrantyMonths));
  return { expiryISO: expiry.toISOString().slice(0, 10), active: expiry >= new Date(todayISO()) };
}

function WarrantyTab({ data, patch }) {
  const [claimOpen, setClaimOpen] = useState(null);
  const [manualOpen, setManualOpen] = useState(false);

  const cards = data.serviceCards.filter((c) => c.hasWarranty).map((c) => ({ ...c, w: warrantyStatus(c) }));
  const manual = (data.warranty || []).map((w) => ({ ...w, w: warrantyStatus(w) }));
  const all = [...cards, ...manual];
  const active = all.filter((c) => c.w?.active).length;

  function addClaim(claim) {
    patch((d) => {
      const product = d.products.find((p) => p.id === claim.replacementProductId);
      if (product) product.qty = Math.max(0, num(product.qty) - claim.qty);
      d.brokenItems = d.brokenItems || [];
      d.brokenItems.push({ id: uid(), date: todayISO(), name: claim.brokenProduct, qty: claim.qty, fromPlate: claim.plate, status: "Tekshirilmoqda" });
      d.warrantyClaims.unshift({ id: uid(), ...claim });
      if (claim.ustaFeeCharged > 0) {
        d.cashflow.unshift({ id: uid(), date: todayISO(), type: "kirim", category: "Xizmat to'lovi", currency: "SUM", amount: claim.ustaFeeCharged, amountSum: claim.ustaFeeCharged, amountUsd: 0, note: `Kafolat — usta haqi — ${claim.plate}` });
      }
      return d;
    });
  }

  function saveManual(w) {
    patch((d) => { d.warranty = d.warranty || []; d.warranty.push({ id: uid(), isManual: true, ...w }); return d; });
  }

  const broken = data.brokenItems || [];

  return (
    <div>
      <PageHeader Icon={ShieldCheck} color={T.teal} title="Kafolat" sub={`${active} ta faol`}
        action={<Btn variant="ghost" onClick={() => setManualOpen(true)}><Plus size={14} /> Eski mijoz kafolati</Btn>} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 13, marginBottom: 20 }}>
        <Stat label="Faol kafolatlar" value={active + " ta"} color={T.teal} Icon={ShieldCheck} />
        <Stat label="Jami kafolatlangan" value={all.length + " ta"} color={T.blue} Icon={ShieldCheck} />
        <Stat label="Almashtirishlar" value={data.warrantyClaims.length + " ta"} color={T.red} Icon={AlertTriangle} />
      </div>

      <Card title="Kafolatlangan kartalar" Icon={ShieldCheck} color={T.teal} pad={false}>
        <Tbl
          empty="Karta yo'q"
          cols={[
            { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
            { k: "plate", h: "Raqam", r: (r) => <span className="mo" style={{ fontWeight: 700, color: T.flame }}>{r.plate}</span> },
            { k: "carModel", h: "Mashina" },
            { k: "warrantyMonths", h: "Muddat", r: (r) => `${r.warrantyMonths} oy` },
            { k: "expiry", h: "Tugaydi", r: (r) => fmtDate(r.w?.expiryISO) },
            { k: "active", h: "Holat", r: (r) => <Badge color={r.w?.active ? T.teal : T.muted}>{r.w?.active ? "Faol" : "Tugagan"}</Badge> },
            { k: "manual", h: "", r: (r) => r.isManual ? <Badge color={T.gold}>Eski</Badge> : null },
            { k: "act", h: "", r: (r) => <Btn size="sm" variant="ghost" onClick={() => setClaimOpen(r)}>Almashtirish</Btn> },
          ]}
          rows={all}
        />
      </Card>

      {broken.length > 0 && (
        <div style={{ marginTop: 16 }}>
          <Card title={`Yaroqsiz tovarlar (${broken.length})`} Icon={AlertTriangle} color={T.red} pad={false}>
            <Tbl
              empty=""
              cols={[
                { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
                { k: "name", h: "Mahsulot" },
                { k: "qty", h: "Miqdor" },
                { k: "fromPlate", h: "Manba" },
                { k: "status", h: "Holat", r: (r) => <Badge color={T.gold}>{r.status}</Badge> },
              ]}
              rows={broken}
            />
          </Card>
        </div>
      )}

      {claimOpen && <WarrantyClaimModal card={claimOpen} products={data.products} onClose={() => setClaimOpen(null)} onSave={(c) => { addClaim(c); setClaimOpen(null); }} />}
      {manualOpen && <ManualWarrantyModal onClose={() => setManualOpen(false)} onSave={(w) => { saveManual(w); setManualOpen(false); }} />}
    </div>
  );
}

function WarrantyClaimModal({ card, products, onClose, onSave }) {
  const [brokenProduct, setBrokenProduct] = useState("");
  const [replId, setReplId] = useState(products[0]?.id || "");
  const [qty, setQty] = useState(1);
  const [chargeUsta, setChargeUsta] = useState(false);
  const [ustaFee, setUstaFee] = useState(0);
  const [docConfirmed, setDocConfirmed] = useState(false);
  const repl = products.find((p) => p.id === replId);

  return (
    <Modal title={`Kafolat almashtirish — ${card.plate}`} onClose={onClose} wide>
      <F label="Brak mahsulot"><input style={iSt} value={brokenProduct} onChange={(e) => setBrokenProduct(e.target.value)} /></F>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginTop: 12 }}>
        <F label="O'rniga beriladigan"><Sel value={replId} onChange={(e) => setReplId(e.target.value)} options={products.map((p) => ({ value: p.id, label: `${p.name} (${p.qty})` }))} /></F>
        <F label="Miqdor"><input type="number" style={iSt} value={qty} onChange={(e) => setQty(e.target.value)} /></F>
      </div>
      <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 14, cursor: "pointer", fontSize: 13 }}>
        <input type="checkbox" checked={docConfirmed} onChange={(e) => setDocConfirmed(e.target.checked)} style={{ accentColor: T.teal }} />
        Kafolat hujjati ko'rsatildi
      </label>
      <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 10, cursor: "pointer", fontSize: 13 }}>
        <input type="checkbox" checked={chargeUsta} onChange={(e) => setChargeUsta(e.target.checked)} style={{ accentColor: T.flame }} />
        Usta xizmat haqi olinsin
      </label>
      {chargeUsta && <F label="Summa"><input type="number" style={iSt} value={ustaFee} onChange={(e) => setUstaFee(e.target.value)} /></F>}
      <SaveBtn disabled={!brokenProduct.trim()} onClick={() => onSave({ date: todayISO(), plate: card.plate, brokenProduct: brokenProduct.trim(), replacementProductId: replId, replacementName: repl?.name, qty: num(qty), ustaFeeCharged: chargeUsta ? num(ustaFee) : 0, docConfirmed })}>Saqlash</SaveBtn>
    </Modal>
  );
}

function ManualWarrantyModal({ onClose, onSave }) {
  const [f, setF] = useState({ plate: "", phone: "", carModel: "", date: "", warrantyMonths: 12, product: "" });
  const set = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));
  return (
    <Modal title="Eski mijoz kafolati" onClose={onClose} wide>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <F label="Davlat raqami"><input style={iSt} value={f.plate} onChange={set("plate")} /></F>
        <F label="Telefon"><input style={iSt} value={f.phone} onChange={set("phone")} /></F>
        <F label="Mashina"><input style={iSt} value={f.carModel} onChange={set("carModel")} /></F>
        <F label="O'rnatilgan sana"><input type="date" style={iSt} value={f.date} onChange={set("date")} /></F>
        <F label="Muddat (oy)"><input type="number" style={iSt} value={f.warrantyMonths} onChange={set("warrantyMonths")} /></F>
        <F label="Mahsulot" col="1/-1"><input style={iSt} value={f.product} onChange={set("product")} /></F>
      </div>
      <SaveBtn disabled={!f.plate.trim() || !f.date} onClick={() => onSave({ ...f, hasWarranty: true })}>Saqlash</SaveBtn>
    </Modal>
  );
}

/* ─── PARTNERS TAB ─── */
function PartnersTab({ data, patch, rate }) {
  const [addOpen, setAddOpen] = useState(false);
  const [giveOpen, setGiveOpen] = useState(null);
  const [payOpen, setPayOpen] = useState(null);
  const [bonusRulesOpen, setBonusRulesOpen] = useState(false);

  const balances = partnerBalances(data);
  const totalDebt = balances.reduce((s, p) => s + Math.max(0, p.debtSum), 0);
  const bonusRules = data.bonusRules || [];

  function addPartner(p) { patch((d) => { d.partners.unshift({ id: uid(), ...p }); return d; }); }

  function giveProduct(partnerId, item) {
    patch((d) => {
      const product = d.products.find((p) => p.id === item.productId);
      if (product) product.qty = Math.max(0, num(product.qty) - item.qty);
      d.partnerTx.unshift({ id: uid(), partnerId, date: todayISO(), type: "mahsulot", ...item });

      // Sklad harakati — insider servisga qarzga berilgan mahsulot (chiqim)
      const partnerObj = d.partners.find((p) => p.id === partnerId);
      d.stockOuts = d.stockOuts || [];
      d.stockOuts.unshift({
        id: uid(), date: todayISO(), productId: item.productId, productName: item.name,
        qty: item.qty, amountSum: item.amountSum, reason: "Insider servisga qarzga berildi",
        partnerId, partnerName: partnerObj?.name || "Noma'lum",
      });

      // Bonus limit tekshiruvi — shu mahsulot bo'yicha jami olingan miqdor
      const rule = (d.bonusRules || []).find((r) => r.productId === item.productId);
      if (rule) {
        const totalTaken = d.partnerTx
          .filter((t) => t.partnerId === partnerId && t.type === "mahsulot" && t.productId === item.productId)
          .reduce((s, t) => s + num(t.qty), 0);
        const alreadyEarned = (d.bonusAwards || []).filter(
          (a) => a.partnerId === partnerId && a.ruleId === rule.id
        ).length;
        const earnedCount = Math.floor(totalTaken / rule.limitQty);
        if (earnedCount > alreadyEarned) {
          d.bonusAwards = d.bonusAwards || [];
          for (let i = alreadyEarned; i < earnedCount; i++) {
            d.bonusAwards.push({
              id: uid(), partnerId, ruleId: rule.id, date: todayISO(),
              productName: rule.productName, bonusText: rule.bonusText,
              claimed: false,
            });
          }
        }
      }
      return d;
    });
  }

  function receivePay(partnerId, payment) {
    patch((d) => {
      d.partnerTx.unshift({ id: uid(), partnerId, date: todayISO(), type: "tolov", amountSum: payment.amountSum, currency: payment.currency, paymentType: payment.paymentType });
      d.cashflow.unshift({
        id: uid(), date: todayISO(), type: "kirim", category: "Hamkordan to'lov",
        currency: payment.currency, amount: payment.amount, amountSum: payment.amountSum,
        amountUsd: payment.currency === "USD" ? payment.amount : payment.amountSum / rate,
        paymentType: payment.paymentType,
        note: `Insider servis to'lovi — ${payment.paymentType}`,
      });
      return d;
    });
  }

  function claimBonus(awardId) {
    patch((d) => {
      const a = (d.bonusAwards || []).find((x) => x.id === awardId);
      if (a) a.claimed = true;
      return d;
    });
  }

  const unclaimedBonuses = (data.bonusAwards || []).filter((a) => !a.claimed);

  return (
    <div>
      <PageHeader Icon={Handshake} color={T.purple} title="Hamkorlar"
        sub={<>Jami qarz: <span style={{ color: T.red, fontWeight: 600 }}>{fmtSum(totalDebt)}</span></>}
        action={<div style={{ display: "flex", gap: 8 }}>
          <Btn variant="gold" onClick={() => setBonusRulesOpen(true)}>
            <Star size={14} /> Bonus qoidalari {bonusRules.length > 0 && `(${bonusRules.length})`}
          </Btn>
          <Btn onClick={() => setAddOpen(true)}><Plus size={15} /> Hamkor qo'shish</Btn>
        </div>} />

      {unclaimedBonuses.length > 0 && (
        <div style={{
          background: T.goldD, border: `1px solid ${T.gold}40`, borderRadius: 10,
          padding: "13px 16px", marginBottom: 16,
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
            <Star size={15} color={T.gold} />
            <span style={{ fontSize: 13, fontWeight: 700, color: T.gold }}>
              {unclaimedBonuses.length} ta hamkor bonus olishga haqli!
            </span>
          </div>
          <div style={{ display: "grid", gap: 6 }}>
            {unclaimedBonuses.map((a) => {
              const partner = data.partners.find((p) => p.id === a.partnerId);
              return (
                <div key={a.id} style={{
                  display: "flex", justifyContent: "space-between", alignItems: "center",
                  background: T.s1, borderRadius: 7, padding: "9px 13px",
                }}>
                  <span style={{ fontSize: 12.5 }}>
                    <b>{partner?.name || "Noma'lum"}</b> — {a.productName} limitiga yetdi:
                    <span style={{ color: T.gold, fontWeight: 600 }}> {a.bonusText}</span>
                  </span>
                  <Btn size="sm" variant="gold" onClick={() => claimBonus(a.id)}>
                    <Check size={11} /> Berildi
                  </Btn>
                </div>
              );
            })}
          </div>
        </div>
      )}

      <Card title={`Hamkorlar (${balances.length})`} Icon={Handshake} color={T.purple} pad={false}>
        <Tbl
          empty="Hamkor yo'q"
          cols={[
            { k: "name", h: "Nomi", r: (r) => <span style={{ fontWeight: 600 }}>{r.name}</span> },
            { k: "phone", h: "Telefon" },
            { k: "given", h: "Olingan", r: (r) => fmtSum(r.given) },
            { k: "paid", h: "To'langan", r: (r) => <span style={{ color: T.teal }}>{fmtSum(r.paid)}</span> },
            { k: "debtSum", h: "Qarz", r: (r) => <span style={{ fontWeight: 700, color: r.debtSum > 0 ? T.red : T.teal }}>{fmtSum(r.debtSum)}</span> },
            { k: "act", h: "", r: (r) => (
                <div style={{ display: "flex", gap: 6 }}>
                  <Btn size="sm" variant="ghost" onClick={() => setGiveOpen(r)}>Berish</Btn>
                  <Btn size="sm" variant="teal" onClick={() => setPayOpen(r)}>To'lov</Btn>
                </div>
              ) },
          ]}
          rows={balances}
        />
      </Card>

      {bonusRules.length > 0 && (
        <div style={{ marginTop: 16 }}>
          <Card title="Bonus limit qoidalari" Icon={Star} color={T.gold} pad={false}>
            <Tbl
              empty=""
              cols={[
                { k: "productName", h: "Mahsulot" },
                { k: "limitQty", h: "Limit", r: (r) => <span className="mo">{r.limitQty} dona</span> },
                { k: "bonusText", h: "Bonus", r: (r) => <span style={{ color: T.gold, fontWeight: 600 }}>{r.bonusText}</span> },
                { k: "del", h: "", r: (r) => (
                    <button onClick={() => patch((d) => { d.bonusRules = d.bonusRules.filter((x) => x.id !== r.id); return d; })}
                      style={{ background: "none", border: "none", cursor: "pointer", color: T.muted }}>
                      <Trash2 size={13} />
                    </button>
                  ) },
              ]}
              rows={bonusRules}
            />
          </Card>
        </div>
      )}

      {addOpen && <NewPartnerModal onClose={() => setAddOpen(false)} onSave={(p) => { addPartner(p); setAddOpen(false); }} />}
      {giveOpen && (
        <GiveProductModal
          partner={giveOpen} products={data.products} data={data}
          onClose={() => setGiveOpen(null)}
          onSave={(item) => { giveProduct(giveOpen.id, item); setGiveOpen(null); }}
        />
      )}
      {payOpen && (
        <Modal title={`${payOpen.name} — to'lov qabul qilish`} onClose={() => setPayOpen(null)}>
          <PaySupplierForm supplier={{ name: payOpen.name, debtSum: payOpen.debtSum }} rate={rate}
            onSave={(amountSum, meta) => { receivePay(payOpen.id, { ...meta, amountSum, amount: meta.amountOriginal }); setPayOpen(null); }} />
        </Modal>
      )}
      {bonusRulesOpen && (
        <BonusRulesModal data={data} onClose={() => setBonusRulesOpen(false)}
          onAdd={(rule) => patch((d) => { d.bonusRules = d.bonusRules || []; d.bonusRules.push({ id: uid(), ...rule }); return d; })}
        />
      )}
    </div>
  );
}

function BonusRulesModal({ data, onClose, onAdd }) {
  const [productId, setProductId] = useState(data.products[0]?.id || "");
  const [limitQty, setLimitQty] = useState(60);
  const [bonusText, setBonusText] = useState("");
  const product = data.products.find((p) => p.id === productId);

  return (
    <Modal title="Bonus limit qoidasi qo'shish" onClose={onClose} wide>
      <p style={{ fontSize: 12.5, color: T.muted, marginBottom: 16 }}>
        Insider servis (hamkor) belgilangan mahsulotdan jami necha dona olganida bonus/sovg'a berilishini belgilang.
        Masalan: <b>60 dona moy</b> olganda → <b>"Kalit to'plami"</b> bonus.
      </p>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <F label="Mahsulot" col="1/-1">
          <Sel value={productId} onChange={(e) => setProductId(e.target.value)}
            options={data.products.length ? data.products.map((p) => ({ value: p.id, label: p.name })) : [{ value: "", label: "Mahsulot yo'q" }]} />
        </F>
        <F label="Limit (necha dona)">
          <input type="number" style={iSt} value={limitQty} onChange={(e) => setLimitQty(e.target.value)} />
        </F>
        <F label="Bonus / sovg'a matni">
          <input style={iSt} value={bonusText} onChange={(e) => setBonusText(e.target.value)} placeholder="Masalan: Kalit to'plami" />
        </F>
      </div>
      <p style={{ fontSize: 11.5, color: T.muted, marginTop: 12 }}>
        Har {limitQty || "?"} donadan keyin (ya'ni {limitQty}, {num(limitQty) * 2}, {num(limitQty) * 3}...) tizim avtomatik bonusni ko'rsatadi.
      </p>
      <SaveBtn disabled={!product || !bonusText.trim() || !num(limitQty)} color={T.gold}
        onClick={() => { onAdd({ productId, productName: product.name, limitQty: num(limitQty), bonusText: bonusText.trim() }); onClose(); }}>
        <Star size={15} /> Qoidani saqlash
      </SaveBtn>
    </Modal>
  );
}

function NewPartnerModal({ onClose, onSave }) {
  const [name, setName] = useState("");
  const [phone, setPhone] = useState("");
  return (
    <Modal title="Yangi hamkor" onClose={onClose}>
      <F label="Nomi"><input style={iSt} value={name} onChange={(e) => setName(e.target.value)} autoFocus /></F>
      <div style={{ marginTop: 12 }}>
        <F label="Telefon"><input style={iSt} value={phone} onChange={(e) => setPhone(e.target.value)} /></F>
      </div>
      <SaveBtn disabled={!name.trim()} onClick={() => onSave({ name: name.trim(), phone: phone.trim() })}>Saqlash</SaveBtn>
    </Modal>
  );
}

function GiveProductModal({ partner, products, data, onClose, onSave }) {
  const [productId, setProductId] = useState(products[0]?.id || "");
  const [qty, setQty] = useState(1);
  const product = products.find((p) => p.id === productId);
  const total = num(product?.priceSum) * num(qty);

  // Shu mahsulot uchun bonus qoidasi bormi va progress qancha
  const rule = (data.bonusRules || []).find((r) => r.productId === productId);
  const alreadyTaken = rule
    ? data.partnerTx.filter((t) => t.partnerId === partner.id && t.type === "mahsulot" && t.productId === productId).reduce((s, t) => s + num(t.qty), 0)
    : 0;
  const afterTaken = alreadyTaken + num(qty);
  const progress = rule ? Math.min(100, ((afterTaken % rule.limitQty) / rule.limitQty) * 100) : 0;
  const willEarnBonus = rule && Math.floor(afterTaken / rule.limitQty) > Math.floor(alreadyTaken / rule.limitQty);

  return (
    <Modal title={`${partner.name} — mahsulot berish`} onClose={onClose}>
      <F label="Mahsulot"><Sel value={productId} onChange={(e) => setProductId(e.target.value)} options={products.map((p) => ({ value: p.id, label: `${p.name} (${p.qty})` }))} /></F>
      <div style={{ marginTop: 12 }}>
        <F label="Miqdor"><input type="number" style={iSt} value={qty} onChange={(e) => setQty(e.target.value)} /></F>
      </div>

      {rule && (
        <div style={{ marginTop: 14, padding: "12px 14px", background: T.goldD, border: `1px solid ${T.gold}30`, borderRadius: 9 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, marginBottom: 7 }}>
            <span style={{ color: T.muted2 }}>Bonus: <b style={{ color: T.gold }}>{rule.bonusText}</b> — har {rule.limitQty} donada</span>
            <span className="mo" style={{ color: T.gold, fontWeight: 700 }}>{afterTaken} / {Math.ceil(afterTaken / rule.limitQty) * rule.limitQty || rule.limitQty}</span>
          </div>
          <div style={{ height: 6, background: T.s3, borderRadius: 3, overflow: "hidden" }}>
            <div style={{ width: `${progress}%`, height: "100%", background: T.gold, borderRadius: 3, transition: "width .3s" }} />
          </div>
          {willEarnBonus && (
            <div style={{ marginTop: 8, fontSize: 12, fontWeight: 700, color: T.gold, display: "flex", alignItems: "center", gap: 6 }}>
              <Star size={13} /> Bu yuk bilan limitga yetadi — bonus avtomatik qo'shiladi!
            </div>
          )}
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 14, padding: "10px 14px", background: T.s3, borderRadius: 8 }}>
        <span style={{ fontSize: 12, color: T.muted }}>Qarzga qo'shiladi</span>
        <span className="mo" style={{ color: T.red, fontWeight: 700 }}>{fmtSum(total)}</span>
      </div>
      <SaveBtn disabled={!product} onClick={() => onSave({ productId, name: product.name, qty: num(qty), amountSum: total })}>Berish</SaveBtn>
    </Modal>
  );
}

/* ═══════════════════════════════════════════════════
   XODIMLAR TAB — erkin lavozim, moslashuvchan oylik
═══════════════════════════════════════════════════ */
function currentMonthKey() { return todayISO().slice(0, 7); }

function EmployeesTab({ data, patch, rate }) {
  const [addOpen, setAddOpen] = useState(false);
  const [payOpen, setPayOpen] = useState(null);
  const [monthFilter, setMonthFilter] = useState(currentMonthKey());

  const employees = data.employees || [];
  const payments = data.employeePayments || [];

  const monthPayments = payments.filter((p) => p.month === monthFilter);
  const totalPaidThisMonth = monthPayments.reduce((s, p) => s + num(p.amountSum), 0);

  const months = [...new Set([currentMonthKey(), ...payments.map((p) => p.month)])].sort().reverse();

  function addEmployee(emp) {
    patch((d) => { d.employees.push({ id: uid(), ...emp }); return d; });
  }

  function payEmployee(employeeId, amountSum, month, note) {
    patch((d) => {
      d.employeePayments = d.employeePayments || [];
      d.employeePayments.push({ id: uid(), employeeId, month, date: todayISO(), amountSum, note });
      d.cashflow.unshift({
        id: uid(), date: todayISO(), type: "chiqim", category: "Ish haqi",
        currency: "SUM", amount: amountSum, amountSum, amountUsd: amountSum / rate,
        note: `${d.employees.find((e) => e.id === employeeId)?.name || ""} — ${d.employees.find((e) => e.id === employeeId)?.position || ""} oyligi (${month})`,
      });
      return d;
    });
  }

  function deleteEmployee(id) {
    patch((d) => { d.employees = d.employees.filter((e) => e.id !== id); return d; });
  }

  function isPaidThisMonth(employeeId) {
    return monthPayments.some((p) => p.employeeId === employeeId);
  }

  return (
    <div>
      <PageHeader Icon={Users} color={T.blue} title="Xodimlar" sub="Erkin lavozim va moslashuvchan oylik"
        action={<Btn onClick={() => setAddOpen(true)}><Plus size={15} /> Xodim qo'shish</Btn>} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 13, marginBottom: 20 }}>
        <Stat label="Jami xodimlar" value={employees.length + " ta"} color={T.blue} Icon={Users} />
        <Stat label={`To'langan (${monthFilter})`} value={fmtSum(totalPaidThisMonth)} color={T.teal} Icon={Check} />
        <Stat label="Bu oy to'lanmagan" value={employees.filter((e) => !isPaidThisMonth(e.id)).length + " ta"} color={T.gold} Icon={Clock} />
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
        <span style={{ fontSize: 12, color: T.muted, fontWeight: 600 }}>Oy:</span>
        <Sel value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}
          style={{ width: 150 }}
          options={months.map((m) => ({ value: m, label: m }))} />
      </div>

      <Card title={`Xodimlar ro'yxati (${employees.length})`} Icon={Users} color={T.blue} pad={false}>
        <Tbl
          empty="Hali xodim qo'shilmagan"
          cols={[
            { k: "name", h: "Ismi", r: (r) => <span style={{ fontWeight: 600 }}>{r.name}</span> },
            { k: "position", h: "Lavozim", r: (r) => <Badge color={T.purple}>{r.position}</Badge> },
            { k: "phone", h: "Telefon", r: (r) => <span className="mo" style={{ fontSize: 12 }}>{r.phone || "—"}</span> },
            { k: "standardSalary", h: "Standart oylik", r: (r) => <span className="mo" style={{ color: T.muted2 }}>{fmtSum(r.standardSalary)}</span> },
            {
              k: "status", h: `Holat (${monthFilter})`,
              r: (r) => isPaidThisMonth(r.id)
                ? <Badge color={T.teal}>✓ To'langan</Badge>
                : <Badge color={T.gold}>Kutilmoqda</Badge>,
            },
            {
              k: "act", h: "",
              r: (r) => (
                <div style={{ display: "flex", gap: 6 }}>
                  <Btn size="sm" variant="teal" onClick={() => setPayOpen(r)}>Oylik to'lash</Btn>
                  <button onClick={async () => { if (await askConfirm(`${r.name} o'chirilsinmi?`)) deleteEmployee(r.id); }}
                    style={{ background: "none", border: "none", cursor: "pointer", color: T.muted }}>
                    <Trash2 size={13} />
                  </button>
                </div>
              ),
            },
          ]}
          rows={employees}
        />
      </Card>

      <div style={{ marginTop: 16 }}>
        <Card title={`To'lovlar tarixi — ${monthFilter}`} Icon={Wallet} color={T.teal} pad={false}>
          <Tbl
            empty="Bu oyda to'lov yo'q"
            cols={[
              { k: "date", h: "Sana", r: (r) => fmtDate(r.date) },
              { k: "employeeId", h: "Xodim", r: (r) => {
                  const e = employees.find((x) => x.id === r.employeeId);
                  return e ? `${e.name} (${e.position})` : "Noma'lum";
                } },
              { k: "amountSum", h: "Summa", r: (r) => <span style={{ color: T.teal, fontWeight: 700 }}>{fmtSum(r.amountSum)}</span> },
              { k: "note", h: "Izoh", r: (r) => <span style={{ color: T.muted, fontSize: 12 }}>{r.note || "—"}</span> },
            ]}
            rows={monthPayments}
          />
        </Card>
      </div>

      {addOpen && <AddEmployeeModal onClose={() => setAddOpen(false)} onSave={(emp) => { addEmployee(emp); setAddOpen(false); }} />}
      {payOpen && (
        <PayEmployeeModal
          employee={payOpen} month={monthFilter}
          onClose={() => setPayOpen(null)}
          onSave={(amountSum, note) => { payEmployee(payOpen.id, amountSum, monthFilter, note); setPayOpen(null); }}
        />
      )}
    </div>
  );
}

function AddEmployeeModal({ onClose, onSave }) {
  const [name, setName] = useState("");
  const [position, setPosition] = useState("");
  const [phone, setPhone] = useState("");
  const [standardSalary, setStandardSalary] = useState("");

  return (
    <Modal title="Yangi xodim qo'shish" onClose={onClose}>
      <div style={{ display: "grid", gap: 12 }}>
        <F label="Ismi *"><input style={iSt} value={name} onChange={(e) => setName(e.target.value)} autoFocus /></F>
        <F label="Lavozim *">
          <input style={iSt} value={position} onChange={(e) => setPosition(e.target.value)} placeholder="Masalan: Oshpaz, SMM, Tozalash xodimi..." />
        </F>
        <F label="Telefon"><input style={iSt} value={phone} onChange={(e) => setPhone(e.target.value)} /></F>
        <F label="Standart oylik (so'm)">
          <input type="number" style={iSt} value={standardSalary} onChange={(e) => setStandardSalary(e.target.value)} placeholder="Kelishilgan summa" />
        </F>
      </div>
      <p style={{ fontSize: 11.5, color: T.muted, marginTop: 10 }}>
        Bu — standart oylik. Har oy to'lov paytida bu summani xohlagancha o'zgartirish mumkin.
      </p>
      <SaveBtn disabled={!name.trim() || !position.trim()} onClick={() => onSave({ name: name.trim(), position: position.trim(), phone: phone.trim(), standardSalary: num(standardSalary) })}>
        Saqlash
      </SaveBtn>
    </Modal>
  );
}

function PayEmployeeModal({ employee, month, onClose, onSave }) {
  const [amount, setAmount] = useState(String(Math.round(employee.standardSalary || 0)));
  const [note, setNote] = useState("");
  const amtNum = num(amount);
  const diff = amtNum - num(employee.standardSalary || 0);

  return (
    <Modal title={`Oylik to'lash — ${employee.name}`} onClose={onClose}>
      <div style={{ marginBottom: 14, padding: "10px 14px", background: T.s3, borderRadius: 8 }}>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12 }}>
          <span style={{ color: T.muted }}>Lavozim</span>
          <Badge color={T.purple}>{employee.position}</Badge>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginTop: 6 }}>
          <span style={{ color: T.muted }}>Standart oylik</span>
          <span className="mo">{fmtSum(employee.standardSalary)}</span>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginTop: 6 }}>
          <span style={{ color: T.muted }}>Oy</span>
          <span className="mo">{month}</span>
        </div>
      </div>

      <F label="To'lov summasi (so'm) — o'zgartirish mumkin">
        <input type="number" style={iSt} value={amount} onChange={(e) => setAmount(e.target.value)} autoFocus onFocus={(e) => e.target.select()} />
      </F>

      {diff !== 0 && (
        <p style={{ fontSize: 11.5, color: diff > 0 ? T.teal : T.gold, marginTop: 8 }}>
          {diff > 0 ? `+${fmtSum(diff)} standartdan ko'p (bonus)` : `${fmtSum(diff)} standartdan kam (chegirma)`}
        </p>
      )}

      <div style={{ marginTop: 12 }}>
        <F label="Izoh (ixtiyoriy)">
          <input style={iSt} value={note} onChange={(e) => setNote(e.target.value)} placeholder="Masalan: bonus qo'shildi, kechikish uchun kamaytirildi..." />
        </F>
      </div>

      <SaveBtn disabled={!amtNum || amtNum <= 0} color={T.teal} onClick={() => onSave(amtNum, note.trim())}>
        <Check size={15} /> Oylikni to'lash
      </SaveBtn>
    </Modal>
  );
}

/* ─── ANALYTICS TAB ─── */
function AnalyticsTab({ data, patch, rate }) {
  const cards = data.serviceCards.filter((c) => cardStatus(c) !== "ochiq");
  const byType = (t) => cards.filter((c) => c.serviceType === t);
  const servis = byType("Servis"), ustanovka = byType("Ustanovka"), detailing = byType("Detailing"), moy = byType("Moy bo'limi");
  const sumF = (arr, f) => arr.reduce((s, c) => s + num(c[f]), 0);

  const rows = [
    { label: "Servis", count: servis.length, revenue: sumF(servis, "finalTotal"), color: T.teal },
    { label: "Ustanovka", count: ustanovka.length, revenue: sumF(ustanovka, "finalTotal"), color: T.flame },
    { label: "Detailing", count: detailing.length, revenue: sumF(detailing, "finalTotal"), color: T.purple },
    { label: "Moy bo'limi", count: moy.length, revenue: sumF(moy, "finalTotal"), color: T.gold },
  ];
  const maxRev = Math.max(1, ...rows.map((r) => r.revenue));
  const azimKpi = data.settings.azimKpi || 0;

  return (
    <div>
      <PageHeader Icon={BarChart3} color={T.purple} title="Analitika" sub="Umumiy moliyaviy ko'rinish" />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 13, marginBottom: 20 }}>
        <Stat label="Ustanovka soni" value={ustanovka.length + " ta"} sub={fmtSum(sumF(ustanovka, "finalTotal"))} color={T.flame} Icon={Car} />
        <Stat label="Servis soni" value={servis.length + " ta"} sub={fmtSum(sumF(servis, "finalTotal"))} color={T.teal} Icon={Car} />
        <Stat label="Detailing soni" value={detailing.length + " ta"} sub={fmtSum(sumF(detailing, "profitSum"))} color={T.purple} Icon={Car} />
        <Stat label="Moy bo'limi" value={moy.length + " ta"} sub={fmtSum(sumF(moy, "profitSum"))} color={T.gold} Icon={Droplets} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 13, marginBottom: 20 }}>
        <Stat label="Jami aylanma" value={fmtSum(sumF(cards, "finalTotal"))} sub={fmtUsd(sumF(cards, "finalTotal") / rate)} color={T.blue} Icon={TrendingUp} />
        <Stat label="Jami foyda" value={fmtSum(sumF(cards, "profitSum"))} color={T.teal} Icon={TrendingUp} />
        <KpiEditCard value={azimKpi} onSave={(v) => patch((d) => { d.settings.azimKpi = v; return d; })} />
      </div>

      <Card title="Xizmat turlari bo'yicha tushum" Icon={BarChart3} color={T.purple}>
        <div style={{ display: "grid", gap: 14 }}>
          {rows.map((r) => (
            <div key={r.label}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                <span style={{ fontSize: 13, fontWeight: 600 }}>{r.label} <span style={{ color: T.muted, fontWeight: 400 }}>({r.count})</span></span>
                <span className="mo" style={{ fontSize: 12.5, fontWeight: 600 }}>{fmtSum(r.revenue)}</span>
              </div>
              <div style={{ height: 6, background: T.s3, borderRadius: 3, overflow: "hidden" }}>
                <div style={{ width: `${(r.revenue / maxRev * 100).toFixed(0)}%`, height: "100%", background: r.color, borderRadius: 3 }} />
              </div>
            </div>
          ))}
        </div>
      </Card>

      <div style={{ marginTop: 20 }}>
        <OwnerMonthlyReport data={data} rate={rate} />
      </div>
    </div>
  );
}

/* ── RAHBAR OYLIK HISOBOTI — jami kirim va barcha chiqimlar ── */
function OwnerMonthlyReport({ data, rate }) {
  const [monthFilter, setMonthFilter] = useState(currentMonthKey());

  const cf = data.cashflow || [];
  const monthCF = cf.filter((c) => (c.date || "").startsWith(monthFilter));

  const monthIncome = monthCF
    .filter((c) => c.type === "kirim" && c.paymentType !== "Karta (Click/Payme)" && c.paymentType !== "Nasiya (qarzga)")
    .reduce((s, c) => s + num(c.amountSum), 0);

  const supplierPay = monthCF.filter((c) => c.category === "Ta'minotchiga to'lov").reduce((s, c) => s + num(c.amountSum), 0);
  const ustaPay = monthCF.filter((c) => c.category === "Usta xizmat haqi").reduce((s, c) => s + num(c.amountSum), 0);
  const docFeePay = monthCF.filter((c) => c.category === "Hujjat xarajati").reduce((s, c) => s + num(c.amountSum), 0);
  const employeePay = monthCF.filter((c) => c.category === "Ish haqi").reduce((s, c) => s + num(c.amountSum), 0);
  const otherExpense = monthCF
    .filter((c) => c.type === "chiqim" && !["Ta'minotchiga to'lov", "Usta xizmat haqi", "Hujjat xarajati", "Ish haqi"].includes(c.category))
    .reduce((s, c) => s + num(c.amountSum), 0);

  const totalExpense = supplierPay + ustaPay + docFeePay + employeePay + otherExpense;
  const netProfit = monthIncome - totalExpense;

  const months = [...new Set([currentMonthKey(), ...cf.map((c) => (c.date || "").slice(0, 7)).filter(Boolean)])].sort().reverse();

  const rows = [
    ["Jami kirim (naqd)", monthIncome, T.teal, false],
    ["Ta'minotchi to'lovlari", -supplierPay, T.red, false],
    ["Usta xizmat haqi", -ustaPay, T.red, false],
    ["Xodimlar oyligi", -employeePay, T.red, false],
    ["Hujjat xarajatlari", -docFeePay, T.red, false],
    ["Boshqa xarajatlar", -otherExpense, T.red, false],
    ["SOF FOYDA", netProfit, netProfit >= 0 ? T.teal : T.red, true],
  ];

  return (
    <Card
      title="Rahbar oylik hisoboti"
      Icon={Wallet} color={T.flame}
      action={
        <Sel value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}
          style={{ width: 150 }}
          options={months.map((m) => ({ value: m, label: m }))} />
      }
    >
      <div style={{ display: "grid", gap: 2 }}>
        {rows.map(([label, val, color, isBold]) => (
          <div key={label} style={{
            display: "flex", justifyContent: "space-between", alignItems: "center",
            padding: isBold ? "14px 14px" : "9px 14px",
            marginTop: isBold ? 8 : 0,
            borderRadius: 8,
            background: isBold ? T.flameD : "transparent",
            borderTop: isBold ? `2px solid ${T.flame}` : "none",
          }}>
            <span style={{ fontSize: isBold ? 14 : 12.5, fontWeight: isBold ? 800 : 400, color: isBold ? T.text : T.muted }}>
              {label}
            </span>
            <span className="mo" style={{ fontSize: isBold ? 18 : 13, fontWeight: isBold ? 800 : 600, color }}>
              {val < 0 ? "−" : ""}{fmtSum(Math.abs(val))}
            </span>
          </div>
        ))}
      </div>
      <p style={{ fontSize: 11, color: T.muted, marginTop: 14, textAlign: "center" }}>
        USD ekvivalenti: {fmtUsd(netProfit / rate)} · Click/Payme va Nasiya bu hisobga kirmaydi (alohida kuzatiladi)
      </p>
    </Card>
  );
}

function KpiEditCard({ value, onSave }) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(value);
  if (editing) return (
    <div style={{ background: T.s1, border: `1px solid ${T.flame}60`, borderRadius: 12, padding: "15px 17px" }}>
      <div style={{ fontSize: 10, fontWeight: 700, color: T.muted, marginBottom: 8, textTransform: "uppercase" }}>Azim Avazovich KPI</div>
      <div style={{ display: "flex", gap: 8 }}>
        <input autoFocus type="number" style={iSt} value={val} onChange={(e) => setVal(e.target.value)} />
        <Btn size="sm" onClick={() => { onSave(num(val)); setEditing(false); }}>✓</Btn>
      </div>
    </div>
  );
  return (
    <div onClick={() => setEditing(true)} style={{ background: T.s1, border: `1px solid ${T.border}`, borderRadius: 12, padding: "15px 17px", cursor: "pointer", borderTop: `2px solid ${T.gold}` }}>
      <div style={{ fontSize: 10, fontWeight: 700, color: T.muted, marginBottom: 8, textTransform: "uppercase" }}>Azim Avazovich KPI</div>
      <div className="mo bc" style={{ fontSize: 19, fontWeight: 700, color: T.gold }}>{fmtSum(value)}</div>
      <div style={{ fontSize: 10.5, color: T.muted, marginTop: 4 }}>Bosing va o'zgartiring</div>
    </div>
  );
}
