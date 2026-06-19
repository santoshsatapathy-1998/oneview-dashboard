import React, { useState, useEffect, useMemo, useCallback, useContext, createContext, useRef } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, ResponsiveContainer, XAxis, YAxis, Tooltip, ReferenceLine, CartesianGrid,
  BarChart, Bar, Cell, PieChart, Pie, Legend, AreaChart, Area,
} from "recharts";
import {
  LayoutDashboard, Factory, Flame, Truck, FlaskConical, Layers, Wrench, Hammer,
  AlertTriangle, ShieldCheck, TrendingUp, TrendingDown, Minus, Search, Plus, Save,
  ChevronRight, ChevronDown, ChevronUp, X, Users, ClipboardList, ArrowUpRight, ArrowDownRight,
  CheckCircle2, Download, RefreshCw, Menu, Loader2, CalendarPlus, Trash2, Info,
  Sun, Moon, Printer, FileText, Bell, Star, Target, Activity, BarChart2,
  Zap, Award, Eye, Filter, SortAsc, SortDesc, Check, AlertCircle,
} from "lucide-react";

/* ============================================================================
   THEME CONTEXT
   ============================================================================ */
const ThemeContext = createContext(false);
function useDark() { return useContext(ThemeContext); }
function chartTheme(dark) {
  return {
    grid: dark ? "#1e293b" : "#f1f5f9",
    axis: dark ? "#475569" : "#94a3b8",
    tooltip: { fontSize: 12, borderRadius: 8, backgroundColor: dark ? "#1e293b" : "#fff", border: dark ? "1px solid #334155" : "1px solid #e2e8f0", color: dark ? "#e2e8f0" : "#0f172a" },
  };
}

/* ============================================================================
   TOAST CONTEXT
   ============================================================================ */
const ToastContext = createContext(null);
function useToast() { return useContext(ToastContext); }

function ToastProvider({ children }) {
  const [toasts, setToasts] = useState([]);
  const push = useCallback((msg, type = "success") => {
    const id = Date.now();
    setToasts((t) => [...t, { id, msg, type }]);
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 3500);
  }, []);
  const iconMap = { success: <Check className="h-4 w-4" />, error: <AlertCircle className="h-4 w-4" />, info: <Info className="h-4 w-4" /> };
  const colorMap = { success: "bg-emerald-600 text-white", error: "bg-red-600 text-white", info: "bg-slate-800 text-white" };
  return (
    <ToastContext.Provider value={push}>
      {children}
      <div className="fixed bottom-6 right-6 z-[100] flex flex-col gap-2 no-print pointer-events-none">
        {toasts.map((t) => (
          <div key={t.id} className={`flex items-center gap-2 px-4 py-2.5 rounded-xl shadow-xl text-sm font-medium ${colorMap[t.type]}`}
            style={{ animation: "slideInRight 0.25s ease-out" }}>
            {iconMap[t.type]}{t.msg}
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
}

/* ============================================================================
   GLOBAL STYLES
   ============================================================================ */
const GLOBAL_STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap');

@keyframes slideInRight {
  from { opacity: 0; transform: translateX(20px); }
  to   { opacity: 1; transform: translateX(0); }
}
@keyframes fadeIn {
  from { opacity: 0; transform: translateY(6px); }
  to   { opacity: 1; transform: translateY(0); }
}
.animate-fadeIn { animation: fadeIn 0.2s ease-out; }

/* ---- Dark mode scoped overrides ---- */
.dark-mode .bg-white        { background-color: #111827 !important; }
.dark-mode .bg-slate-50     { background-color: #0f172a !important; }
.dark-mode .bg-slate-100    { background-color: #1e293b !important; }
.dark-mode .bg-slate-100\\/70 { background-color: rgba(30,41,59,0.7) !important; }
.dark-mode .bg-slate-50\\/50  { background-color: rgba(15,23,42,0.5)  !important; }
.dark-mode .bg-slate-50\\/60  { background-color: rgba(15,23,42,0.6)  !important; }
.dark-mode .bg-orange-50\\/40 { background-color: rgba(234,88,12,0.12) !important; }
.dark-mode .border-slate-50,
.dark-mode .border-slate-100,
.dark-mode .border-slate-200 { border-color: #1e293b !important; }
.dark-mode .border-slate-300  { border-color: #334155 !important; }
.dark-mode .divide-slate-100 > * + * { border-color: #1e293b !important; }
.dark-mode .text-slate-400  { color: #64748b  !important; }
.dark-mode .text-slate-500  { color: #94a3b8  !important; }
.dark-mode .text-slate-600  { color: #cbd5e1  !important; }
.dark-mode .text-slate-700  { color: #e2e8f0  !important; }
.dark-mode .text-slate-800  { color: #f1f5f9  !important; }
.dark-mode .text-slate-900  { color: #f8fafc  !important; }
.dark-mode .hover\\:bg-slate-50:hover     { background-color: #1e293b !important; }
.dark-mode .hover\\:bg-slate-50\\/60:hover { background-color: rgba(30,41,59,0.6) !important; }
.dark-mode .hover\\:border-slate-300:hover { border-color: #475569 !important; }
.dark-mode .hover\\:text-slate-700:hover  { color: #f1f5f9 !important; }
.dark-mode .bg-orange-50,.dark-mode .bg-sky-50,.dark-mode .bg-indigo-50,.dark-mode .bg-amber-50,
.dark-mode .bg-red-50,.dark-mode .bg-rose-50,.dark-mode .bg-teal-50,.dark-mode .bg-emerald-50,
.dark-mode .bg-violet-50,.dark-mode .bg-fuchsia-50 { background-color: rgba(255,255,255,0.06) !important; }
.dark-mode .border-orange-200,.dark-mode .border-sky-200,.dark-mode .border-indigo-200,
.dark-mode .border-amber-200,.dark-mode .border-red-200,.dark-mode .border-rose-200,
.dark-mode .border-teal-200,.dark-mode .border-emerald-200,.dark-mode .border-violet-200,
.dark-mode .border-fuchsia-200 { border-color: rgba(255,255,255,0.14) !important; }
.dark-mode .text-amber-600  { color: #fbbf24 !important; }
.dark-mode .text-amber-700, .dark-mode .text-amber-800 { color: #fde68a !important; }
.dark-mode .text-red-600    { color: #f87171 !important; }
.dark-mode .text-red-700, .dark-mode .text-red-800 { color: #fca5a5 !important; }
.dark-mode .text-emerald-600 { color: #34d399 !important; }
.dark-mode .text-emerald-700,.dark-mode .text-emerald-800 { color: #6ee7b7 !important; }
.dark-mode .text-sky-600    { color: #38bdf8 !important; }
.dark-mode .text-indigo-600 { color: #818cf8 !important; }
.dark-mode .text-rose-600   { color: #fb7185 !important; }
.dark-mode .text-teal-600   { color: #2dd4bf !important; }
.dark-mode .text-violet-600 { color: #a78bfa !important; }
.dark-mode .text-fuchsia-600{ color: #e879f9 !important; }
.dark-mode .text-orange-600 { color: #fb923c !important; }
.dark-mode input, .dark-mode select, .dark-mode textarea {
  background-color: #0f172a !important; border-color: #334155 !important; color: #e2e8f0 !important;
}
.dark-mode input::placeholder { color: #64748b !important; }

/* ---- Print ---- */
@media print {
  .no-print  { display: none !important; }
  .print-only{ display: block !important; }
  .print-area{ padding: 0 !important; overflow: visible !important; }
  .scroll-area{ max-height: none !important; overflow: visible !important; }
}
@media screen { .print-only { display: none; } }

/* ---- Scrollbar ---- */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
.dark-mode ::-webkit-scrollbar-thumb { background: #334155; }
`;

/* ============================================================================
   BRANDING
   ============================================================================ */
const APP_NAME    = "Jindal OneView";
const APP_TAGLINE = "Unified KPI Intelligence Platform";
const LOGO_DATA_URI = "data:image/png;base64,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UASj73R54TBRCff7GFtClXYxDuNlqC8D0INtd0C6q+XfZZrHL7T3wf41gJ91eXzeJyZmwES7od48g2zt8lQEr+JRZGYxgBcB+M+WC+ni6z+PsV2COs38gB3fAenFGNvEvvi7b6O7VpQ6yfFPbnlNxQ5IP0D3XbHaTbYD0lZ019GLiBkw0W6kCHp/AOBWAHcA+Gz+5+PjZI4LAFQAfCAPeL7lOE8C+ATGSqmdaM/8HsX0dkAqSrhpy4XC/cgmhHW6i5Agm+Htxskq2ztg/RzZto1ldD5Tu1if3P53z8dYf+jW7L3URfZejHnCjzQxABPtPhRjGyqcnj8251nYA3lgLSPr1HQSgINagmaRcQqAS/Kg1EnbRNeWAbdmfg1033rx1HEC2393GHiL97AVWeOL37YEQoesd/PL2s5bRyK7x9xJabhYtvRVABe3XBC0ts+cj23XK68AsLzD46fIbgtcmVcg2LaSiGiWK07uByLr95wCGMFYl6eJHknLo/i71S2ZcKfPuz+yRh4+D7oOQH+XF+jFsb6TH6fYNnE6GzHc33IxUWTXQHafusiMHaa/ycOH2t5b8ee78uM2d+C1K8ZmgXPZEnVUOiKiXRuAJQ+ED+cn7nKeTRWZVTMPjo08MBdZWvG4L8+YB7rIvIrf/ePyzE8wViruZt1u6yb2x2HsnrVgejOTv5sfz7a9zpPz4xatOKc76/lb42TeAPDq/LjRDrz234Bds6gLLEET7VrFifqHyEqspwF4HbL7wUflwXG839OnkM0A/hyAL+SBeTo7IB0N4AmMddaKMTYBq5sdkI7MM+mnpnlhn2Jsp6LxHJW/znQa562ivJ3kFyvFuLcuOdo/P76fxusvjvE/yHZy6vR+N/Hqm4hmye9ie8BbiKyl5EJkE68EWZn6MWTdn37b8rXTvedYxtiEqWI50cg0jhO2ZO07cl4ZmSB49WDHK3YTvTdBNgN9uoqxi8EJWEREu20Qtl1keCb/el5IEzEDJqIZ/v00bb+r2pLN6U46D+gMHWe62eTOPFftquMTERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERHRLPP/A+sO2WJ7hWLfAAAAAElFTkSuQmCC";

/* ============================================================================
   SEED DATA
   ============================================================================ */
const SEED_DATA = [{"id":"kpi-1","department":"PlateMill","workstation":"CTL","leader":"D Kamlambe","inceptionDate":"1.10.2025","kpiName":"CTL 1 Machine Availability >90%","uom":"%","goodnessIndicator":80.0,"baseline":75.0,"commitment":80.0,"monthly":{"2025-04":65.0,"2025-05":65.0,"2025-06":59.0,"2025-07":58.0,"2025-08":75.0,"2025-09":51.0,"2025-10":73.0,"2025-11":75.0,"2025-12":74.0,"2026-01":74.0,"2026-02":75.0,"2026-03":75.0,"2026-04":60.0,"2026-05":60.3}},{"id":"kpi-2","department":"PlateMill","workstation":"CTL","leader":"D Kamlambe","inceptionDate":"1.10.2025","kpiName":"CTL-1 , 10 Coils/Shift cutting","uom":"Nos","goodnessIndicator":10.0,"baseline":5.0,"commitment":10.0,"monthly":{"2025-04":5.0,"2025-05":5.0,"2025-06":5.0,"2025-07":4.0,"2025-08":6.0,"2025-09":4.0,"2025-10":6.0,"2025-11":6.0,"2025-12":6.0,"2026-01":6.0,"2026-02":6.0,"2026-03":6.0,"2026-04":4.0,"2026-05":4.0}},{"id":"kpi-3","department":"PlateMill","workstation":"CTL","leader":"D Kamlambe","inceptionDate":"1.10.2025","kpiName":"CTL 3 Machine Availability >90%","uom":"%","goodnessIndicator":0.85,"baseline":0.77,"commitment":0.85,"monthly":{"2025-04":74.0,"2025-05":78.0,"2025-06":74.0,"2025-07":77.0,"2025-08":78.0,"2025-09":54.0,"2025-10":71.0,"2025-11":79.0,"2025-12":73.0,"2026-01":74.0,"2026-02":78.0,"2026-03":78.0,"2026-04":65.0,"2026-05":75.4}},{"id":"kpi-4","department":"PlateMill","workstation":"CTL","leader":"D Kamlambe","inceptionDate":"1.10.2025","kpiName":"CTL-3 , 18 Coils/Shift cutting","uom":"Nos","goodnessIndicator":18.0,"baseline":15.0,"commitment":"16-17","monthly":{"2025-04":15.0,"2025-05":15.0,"2025-06":15.0,"2025-07":15.0,"2025-08":16.0,"2025-09":13.0,"2025-10":14.0,"2025-11":16.0,"2025-12":15.0,"2026-01":15.0,"2026-02":16.0,"2026-03":16.0,"2026-04":12.0,"2026-05":10.0}},{"id":"kpi-5","department":"PlateMill","workstation":"CTL","leader":"D Kamlambe","inceptionDate":"1.10.2025","kpiName":"Absenteeism 0%","uom":"%","goodnessIndicator":0.0,"baseline":10.0,"commitment":0.0,"monthly":{"2025-04":12.0,"2025-05":15.0,"2025-06":14.0,"2025-07":14.0,"2025-08":12.0,"2025-09":12.0,"2025-10":14.0,"2025-11":14.0,"2025-12":12.0,"2026-01":10.0,"2026-02":10.0,"2026-03":8.0,"2026-04":10.0,"2026-05":8.0}},{"id":"kpi-6","department":"PlateMill","workstation":"Finishing Mill","leader":"Vipul Pandey","inceptionDate":null,"kpiName":"Production 3300 Mt/Day","uom":"Mt/Day","goodnessIndicator":3175.0,"baseline":3084.0,"commitment":3300.0,"monthly":{"2025-11":3085.0,"2025-12":2682.0,"2026-01":3084.0,"2026-02":3175.0,"2026-03":2945.6,"2026-04":1585.0}},{"id":"kpi-7","department":"PlateMill","workstation":"Finishing Mill","leader":"Vipul Pandey","inceptionDate":null,"kpiName":"Equipment Availability 90%","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-8","department":"PlateMill","workstation":"Finishing Mill","leader":"Vipul Pandey","inceptionDate":null,"kpiName":"Rejection 3% max","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-9","department":"PlateMill","workstation":"Finishing Mill","leader":"Vipul Pandey","inceptionDate":null,"kpiName":"Zero Harm","uom":"Nos","goodnessIndicator":0.0,"baseline":0.0,"commitment":0.0,"monthly":{"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-10","department":"PlateMill","workstation":"Finishing Area","leader":"Rupak Ranjan","inceptionDate":null,"kpiName":"Discrete plates per day","uom":"Nos","goodnessIndicator":null,"baseline":null,"commitment":400.0,"monthly":{"2025-04":206.0,"2025-05":234.0,"2025-06":182.0,"2025-07":250.0,"2025-08":269.0,"2025-09":240.0,"2025-10":284.0,"2026-01":291.0,"2026-02":298.0,"2026-03":238.0,"2026-04":176.0,"2026-05":261.0}},{"id":"kpi-11","department":"PlateMill","workstation":"Finishing Area","leader":"Rupak Ranjan","inceptionDate":null,"kpiName":"Machine utilization","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":0.9,"monthly":{}},{"id":"kpi-12","department":"PlateMill","workstation":"Finishing Area","leader":"Rupak Ranjan","inceptionDate":null,"kpiName":"Rejection","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":0.0,"monthly":{}},{"id":"kpi-13","department":"PlateMill","workstation":"Finishing Area","leader":"Rupak Ranjan","inceptionDate":null,"kpiName":"Rework","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":0.0,"monthly":{}},{"id":"kpi-14","department":"PlateMill","workstation":"Finishing Area","leader":"Rupak Ranjan","inceptionDate":null,"kpiName":"Plates/shift at AC Bay","uom":"Nos","goodnessIndicator":null,"baseline":null,"commitment":55.0,"monthly":{}},{"id":"kpi-15","department":"PlateMill","workstation":"Finishing Area","leader":"Rupak Ranjan","inceptionDate":null,"kpiName":"Plates/shift at EF Bay","uom":"Nos","goodnessIndicator":null,"baseline":null,"commitment":78.0,"monthly":{}},{"id":"kpi-16","department":"PlateMill","workstation":"Finishing Area","leader":"Rupak Ranjan","inceptionDate":null,"kpiName":"Machine availability","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":95.0,"monthly":{}},{"id":"kpi-17","department":"PlateMill","workstation":"Finishing Area","leader":"Rupak Ranjan","inceptionDate":null,"kpiName":"Crane availability","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":95.0,"monthly":{}},{"id":"kpi-18","department":"PlateMill","workstation":"Finishing Area","leader":"Rupak Ranjan","inceptionDate":null,"kpiName":"Absenteeism","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":0.0,"monthly":{}},{"id":"kpi-19","department":"Raipur","workstation":"CNC Horizontal Broring Machine -11 PAMA -01","leader":"Manoj Kumar","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-20","department":"Raipur","workstation":"CNC Horizontal Broring Machine -14 Juaristi -01","leader":"Manoj Kumar","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-21","department":"Raipur","workstation":"CNC Plano Milling Machine -07 SNK","leader":"Manoj Kumar","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-22","department":"Raipur","workstation":"CNC Plano Milling Machine -06 SANCO","leader":"Raja Bagh","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-23","department":"Raipur","workstation":"Horizonatl Boring Machine -09 Wotan","leader":"Raja Bagh","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-24","department":"Raipur","workstation":"Horizonatl Boring Machine -23 HMT","leader":"Raja Bagh","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-25","department":"Raipur","workstation":"Induction Furnace 01 (4T)","leader":"Amit Srivas","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-26","department":"Raipur","workstation":"Centrifugal Casting Machine CCMCO","leader":"Amit Srivas","inceptionDate":"01.04.2025","kpiName":"Quality","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{"2025-04":100.0,"2025-05":100.0}},{"id":"kpi-27","department":"Raipur","workstation":"Centrifugal Casting Machine CCMCO","leader":"Amit Srivas","inceptionDate":"01.04.2025","kpiName":"In process breakdown","uom":"Zero","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{"2025-04":"Zero","2025-05":1.0}},{"id":"kpi-28","department":"Raipur","workstation":"Core Shooter Machine -01 Galaxy","leader":"Sandeep Kumar","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-29","department":"Raipur","workstation":"Hydraulic Press -03 (1500T)","leader":"Sanket Shimpi","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-30","department":"Raipur","workstation":"CNC Oxyfuel & Plasma Cutting Machine -MGM","leader":"Manjeet Singh","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-31","department":"Raipur","workstation":"Plate Bending Machine -03 Wuxi","leader":"Anuj Yadav","inceptionDate":"01.04.2025","kpiName":"OEE","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-32","department":"SPM","workstation":"Furnace Area","leader":"Devendra Sao","inceptionDate":"24.04.2026","kpiName":"Break Down","uom":"Min/Month","goodnessIndicator":null,"baseline":177.0,"commitment":120.0,"monthly":{"2025-04":126.0,"2025-05":210.0,"2025-06":126.0,"2025-07":67.8,"2025-08":127.2,"2025-09":61.2,"2025-10":114.0,"2025-11":34.2,"2025-12":184.8,"2026-01":52.8,"2026-02":397.2,"2026-03":105,"2026-04":151.8,"2026-05":123.0}},{"id":"kpi-33","department":"SPM","workstation":"Furnace Area","leader":"Devendra Sao","inceptionDate":"24.04.2026","kpiName":"Fuel","uom":"MCAL/MT","goodnessIndicator":null,"baseline":421.0,"commitment":410.0,"monthly":{"2025-04":383.0,"2025-05":405.0,"2025-06":410.0,"2025-07":404.0,"2025-08":407.0,"2025-09":422.0,"2025-10":466.0,"2025-11":428.0,"2025-12":420.0,"2026-01":437.0,"2026-02":434.0,"2026-03":433.0,"2026-04":407.0,"2026-05":439.0}},{"id":"kpi-34","department":"SPM","workstation":"Furnace Area","leader":"Devendra Sao","inceptionDate":"24.04.2026","kpiName":"Zero Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-35","department":"SPM","workstation":"Furnace Area","leader":"Devendra Sao","inceptionDate":"24.04.2026","kpiName":"Scale Cleaning","uom":"%","goodnessIndicator":null,"baseline":50.0,"commitment":100.0,"monthly":{"2025-04":50.0,"2025-05":50.0,"2025-06":50.0,"2025-07":50.0,"2025-08":50.0,"2025-09":50.0,"2025-10":50.0,"2025-11":50.0,"2025-12":50.0,"2026-01":50.0,"2026-02":50.0,"2026-03":50.0,"2026-04":50.0,"2026-05":80.0}},{"id":"kpi-36","department":"SPM","workstation":"Mill Area","leader":"Ashish Mendhe","inceptionDate":"24.04.2026","kpiName":"Break Down","uom":"Min/Month","goodnessIndicator":null,"baseline":2163.0,"commitment":1800.0,"monthly":{"2025-04":1920.0,"2025-05":1704.0,"2025-06":1704.0,"2025-07":1800.0,"2025-08":2484.0,"2025-09":1686.0,"2025-10":2136.0,"2025-11":2076.0,"2025-12":1896.0,"2026-01":2154.0,"2026-02":1854.0,"2026-03":1650,"2026-04":1954.8,"2026-05":2137}},{"id":"kpi-37","department":"SPM","workstation":"Mill Area","leader":"Ashish Mendhe","inceptionDate":"24.04.2026","kpiName":"Production","uom":"MT/Day","goodnessIndicator":null,"baseline":1661.0,"commitment":1915.0,"monthly":{"2025-04":1745.4615,"2025-05":1615.871,"2025-06":1693.5,"2025-07":1662.3226,"2025-08":1638.4839,"2025-09":1699.0667,"2025-10":1459.9355,"2025-11":1807.6333,"2025-12":1639.5484,"2026-01":1800.2258,"2026-02":1702.4286,"2026-03":1727.0645,"2026-04":1333.8261,"2026-05":1741.75}},{"id":"kpi-38","department":"SPM","workstation":"Mill Area","leader":"Ashish Mendhe","inceptionDate":"24.04.2026","kpiName":"Section change time","uom":"Min","goodnessIndicator":null,"baseline":55.0,"commitment":35.0,"monthly":{"2025-04":54.0,"2025-05":57.0,"2025-06":43.0,"2025-07":38.0,"2025-08":40.0,"2025-09":38.0,"2025-10":40.0,"2025-11":46.0,"2025-12":45.0,"2026-01":51.0,"2026-02":46.0,"2026-03":64.0,"2026-04":62.0,"2026-05":54.0}},{"id":"kpi-39","department":"SPM","workstation":"Mill Area","leader":"Ashish Mendhe","inceptionDate":"24.04.2026","kpiName":"Zero Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-40","department":"SPM","workstation":"Mill Area","leader":"Ashish Mendhe","inceptionDate":"24.04.2026","kpiName":"Zero Rejection & Rework","uom":"%","goodnessIndicator":null,"baseline":0.3,"commitment":0.02,"monthly":{"2025-04":0.478,"2025-05":0.29,"2025-06":0.934,"2025-07":0.475,"2025-08":0.399,"2025-09":0.238,"2025-10":0.627,"2025-11":0.428,"2025-12":0.193,"2026-01":0.126,"2026-02":0.163,"2026-03":0.27,"2026-04":0.038,"2026-05":0.83}},{"id":"kpi-41","department":"SPM","workstation":"Cold Saw Bay","leader":"Satya Patel","inceptionDate":"24.04.2026","kpiName":"Break Down","uom":"Min/Month","goodnessIndicator":null,"baseline":687.0,"commitment":480.0,"monthly":{"2025-04":344.0,"2025-05":587.0,"2025-06":224.0,"2025-07":444.0,"2025-08":486.6,"2025-09":587.0,"2025-10":365.7816,"2025-11":372.972,"2025-12":380.1623,"2026-01":387.3527,"2026-02":394.543,"2026-03":706.8,"2026-04":1047,"2026-05":425.0}},{"id":"kpi-42","department":"SPM","workstation":"Cold Saw Bay","leader":"Satya Patel","inceptionDate":"24.04.2026","kpiName":"Production","uom":"MT/Day","goodnessIndicator":null,"baseline":1661.0,"commitment":1915.0,"monthly":{"2025-04":1745.4615,"2025-05":1615.871,"2025-06":1693.5,"2025-07":1662.3226,"2025-08":1638.4839,"2025-09":1699.0667,"2025-10":1459.9355,"2025-11":1807.6333,"2025-12":1639.5484,"2026-01":1800.2258,"2026-02":1702.4286,"2026-03":1727.0645,"2026-04":1333.8261,"2026-05":2137}},{"id":"kpi-43","department":"SPM","workstation":"Cold Saw Bay","leader":"Satya Patel","inceptionDate":"24.04.2026","kpiName":"Scale Cleaning","uom":"%","goodnessIndicator":null,"baseline":50.0,"commitment":70.0,"monthly":{"2025-04":60.0,"2025-05":60.0,"2025-06":60.0,"2025-07":60.0,"2025-08":60.0,"2025-09":60.0,"2025-10":60.0,"2025-11":60.0,"2025-12":60.0,"2026-01":60.0,"2026-02":60.0,"2026-03":60.0,"2026-04":60.0,"2026-05":60.0}},{"id":"kpi-44","department":"SPM","workstation":"Cold Saw Bay","leader":"Satya Patel","inceptionDate":"24.04.2026","kpiName":"Zero Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-45","department":"SPM","workstation":"Cold Saw Bay","leader":"Satya Patel","inceptionDate":"24.04.2026","kpiName":"Zero Rejection & Rework","uom":"%","goodnessIndicator":null,"baseline":0.3,"commitment":0.0,"monthly":{"2025-04":0.478,"2025-05":0.29,"2025-06":0.934,"2025-07":0.475,"2025-08":0.399,"2025-09":0.238,"2025-10":0.627,"2025-11":0.428,"2025-12":0.193,"2026-01":0.126,"2026-02":0.163,"2026-03":0.27,"2026-04":0.038,"2026-05":0.83}},{"id":"kpi-46","department":"SPM","workstation":"Logistic Area","leader":"Arjun Patel","inceptionDate":"24.04.2026","kpiName":"Break Down","uom":"Min/Month","goodnessIndicator":null,"baseline":35.0,"commitment":0.0,"monthly":{"2025-04":100.2,"2025-05":49.8,"2025-06":139.2,"2025-07":388.2,"2025-08":136.8,"2025-09":223.2,"2025-10":298.2,"2025-11":72.0,"2025-12":64.2,"2026-01":28.2,"2026-02":40.2,"2026-03":55.8,"2026-04":15.0,"2026-05":72.0}},{"id":"kpi-47","department":"SPM","workstation":"Logistic Area","leader":"Arjun Patel","inceptionDate":"24.04.2026","kpiName":"Dispatch","uom":"MT/Day","goodnessIndicator":null,"baseline":1607.0,"commitment":1915.0,"monthly":{"2025-04":1460.9667,"2025-05":1520.4516,"2025-06":1514.6333,"2025-07":1768,"2025-08":1591.7419,"2025-09":1473.4,"2025-10":1556.2581,"2025-11":1757.6667,"2025-12":1876.5161,"2026-01":1786.4516,"2026-02":1841.8214,"2026-03":1760.3226,"2026-04":1040.2667,"2026-05":1746.1071}},{"id":"kpi-48","department":"SPM","workstation":"Logistic Area","leader":"Arjun Patel","inceptionDate":"24.04.2026","kpiName":"TAT","uom":"Hrs","goodnessIndicator":null,"baseline":14.0,"commitment":10.0,"monthly":{"2025-04":10.0,"2025-05":9.0,"2025-06":12.0,"2025-07":16.0,"2025-08":12.0,"2025-09":13.5,"2025-10":12.0,"2025-11":14.0,"2025-12":13.0,"2026-01":11.0,"2026-02":13.0,"2026-03":14.0,"2026-04":19.0,"2026-05":17.0}},{"id":"kpi-49","department":"SPM","workstation":"Logistic Area","leader":"Arjun Patel","inceptionDate":"24.04.2026","kpiName":"Zero Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-50","department":"SPM","workstation":"Logistic Area","leader":"Arjun Patel","inceptionDate":"24.04.2026","kpiName":"OTIF","uom":"%","goodnessIndicator":null,"baseline":97.0,"commitment":100.0,"monthly":{"2025-04":97.9,"2025-05":98.7,"2025-06":98.7,"2025-07":99.2,"2025-08":99.2,"2025-09":99.8,"2025-10":99.3,"2025-11":99.5,"2025-12":99.6,"2026-01":99.48,"2026-02":94.19,"2026-03":97.65,"2026-04":97.23,"2026-05":90.9}},{"id":"kpi-51","department":"SPM","workstation":"Rollshop Area","leader":"Yogesh Nishad","inceptionDate":"24.04.2026","kpiName":"Break Down","uom":"Min/Month","goodnessIndicator":null,"baseline":512.0,"commitment":450.0,"monthly":{}},{"id":"kpi-52","department":"SPM","workstation":"Rollshop Area","leader":"Yogesh Nishad","inceptionDate":"24.04.2026","kpiName":"Production","uom":"MT/Day","goodnessIndicator":null,"baseline":1661.0,"commitment":1915.0,"monthly":{"2025-04":1745.4615,"2025-05":1615.871,"2025-06":1693.5,"2025-07":1662.3226,"2025-08":1638.4839,"2025-09":1699.0667,"2025-10":1459.9355,"2025-11":1807.6333,"2025-12":1639.5484,"2026-01":1800.2258,"2026-02":1702.4286,"2026-03":1727.0645,"2026-04":1333.8261,"2026-05":2137}},{"id":"kpi-53","department":"SPM","workstation":"Rollshop Area","leader":"Yogesh Nishad","inceptionDate":"24.04.2026","kpiName":"Zero Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-54","department":"SPM","workstation":"Utility Area","leader":"Surendra Choudhary","inceptionDate":"24.04.2026","kpiName":"Break Down","uom":"Min/Month","goodnessIndicator":null,"baseline":229.0,"commitment":150.0,"monthly":{"2025-04":27.0,"2025-05":23.0,"2025-06":0.0,"2025-07":154.0,"2025-08":34.0,"2025-09":107.0,"2025-10":117.0,"2025-11":415.0,"2025-12":139.0,"2026-01":229.0,"2026-02":556.0,"2026-03":0.0,"2026-04":408,"2026-05":107.0}},{"id":"kpi-55","department":"SPM","workstation":"Utility Area","leader":"Surendra Choudhary","inceptionDate":"24.04.2026","kpiName":"Water Consumption","uom":"M3/MT","goodnessIndicator":null,"baseline":0.3,"commitment":0.25,"monthly":{"2025-04":0.352,"2025-05":0.356,"2025-06":0.32,"2025-07":0.231,"2025-08":0.273,"2025-09":0.339,"2025-10":0.323,"2025-11":0.277,"2025-12":0.273,"2026-01":0.31,"2026-02":0.271,"2026-03":0.306,"2026-04":0.388,"2026-05":0.3215}},{"id":"kpi-56","department":"SPM","workstation":"Utility Area","leader":"Surendra Choudhary","inceptionDate":"24.04.2026","kpiName":"Zero Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-57","department":"SPM","workstation":"Utility Area","leader":"Surendra Choudhary","inceptionDate":"24.04.2026","kpiName":"TDS","uom":"PPM","goodnessIndicator":null,"baseline":350.0,"commitment":300.0,"monthly":{"2026-04":350.0,"2026-05":345.0}},{"id":"kpi-58","department":"SPM","workstation":"Hydraulic Area","leader":"Yashwant Rathore","inceptionDate":"24.04.2026","kpiName":"Break Down Min/Month","uom":"Min/Month","goodnessIndicator":null,"baseline":111.0,"commitment":80.0,"monthly":{"2025-04":0.0,"2025-05":27.0,"2025-06":35.0,"2025-07":76.0,"2025-08":24.0,"2025-09":35.0,"2025-10":59.0,"2025-11":116.0,"2025-12":108.0,"2026-01":109.0,"2026-02":64.0,"2026-03":52.8,"2026-04":217.2,"2026-05":188.0}},{"id":"kpi-59","department":"SPM","workstation":"Hydraulic Area","leader":"Yashwant Rathore","inceptionDate":"24.04.2026","kpiName":"Oil Consumption","uom":"ml/MT","goodnessIndicator":null,"baseline":128.0,"commitment":100.0,"monthly":{"2025-04":106.4299,"2025-05":65.5793,"2025-06":88.8692,"2025-07":85.5779,"2025-08":99.2263,"2025-09":102.9977,"2025-10":162.4022,"2025-11":100.6841,"2025-12":90.8984,"2026-01":112.8891,"2026-02":105.7313,"2026-03":141.2055,"2026-04":150.5965,"2026-05":116.2624}},{"id":"kpi-60","department":"SPM","workstation":"Hydraulic Area","leader":"Yashwant Rathore","inceptionDate":"24.04.2026","kpiName":"Zero Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-61","department":"SPM","workstation":"SPM Service Center","leader":"Gourav kumar","inceptionDate":"24.04.2026","kpiName":"Production","uom":"MT/Day","goodnessIndicator":null,"baseline":5751.0,"commitment":6200.0,"monthly":{"2025-04":3099.064,"2025-05":3664.181,"2025-06":3465.824,"2025-07":3287.904,"2025-08":2940.67,"2025-09":3804.535,"2025-10":2580.113,"2025-11":3285.0,"2025-12":3129.0,"2026-01":4254.0,"2026-02":4201.0,"2026-03":5633.0,"2026-04":5521.0,"2026-05":6098.0}},{"id":"kpi-62","department":"SPM","workstation":"SPM Service Center","leader":"Gourav kumar","inceptionDate":"24.04.2026","kpiName":"Zero Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-63","department":"Rail Mill","workstation":"Tandem Mill","leader":"Ranjeet Biswal","inceptionDate":null,"kpiName":"Mill Availability","uom":"%","goodnessIndicator":null,"baseline":98.0,"commitment":98.0,"monthly":{"2025-04":0.95}},{"id":"kpi-64","department":"Rail Mill","workstation":"Tandem Mill","leader":"Ranjeet Biswal","inceptionDate":null,"kpiName":"Mechanical Delay","uom":"%","goodnessIndicator":null,"baseline":1.3,"commitment":"<1","monthly":{"2026-02":0.009,"2026-04":0.0095}},{"id":"kpi-65","department":"Rail Mill","workstation":"Tandem Mill","leader":"Ranjeet Biswal","inceptionDate":null,"kpiName":"Electrical Delay","uom":"%","goodnessIndicator":null,"baseline":0.8,"commitment":"<1","monthly":{"2026-02":0.007,"2026-04":0.0075}},{"id":"kpi-66","department":"Rail Mill","workstation":"Tandem Mill","leader":"Ranjeet Biswal","inceptionDate":null,"kpiName":"Zero Process Delay","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{"2025-04":0.95}},{"id":"kpi-67","department":"Rail Mill","workstation":"Tandem Mill","leader":"Ranjeet Biswal","inceptionDate":null,"kpiName":"Zero Incident","uom":"nos","goodnessIndicator":null,"baseline":2.0,"commitment":0.0,"monthly":{"2025-04":1.0,"2025-05":2.0,"2025-06":1.0,"2025-07":1.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":1.0,"2025-12":1.0,"2026-01":2.0,"2026-02":1.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-68","department":"Rail Mill","workstation":"Hot Saw","leader":"Dayal Singh","inceptionDate":null,"kpiName":"Vibration Control <=5mm","uom":"mm","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-69","department":"Rail Mill","workstation":"Hot Saw","leader":"Dayal Singh","inceptionDate":null,"kpiName":"Blade life to be maintain for Normal rail","uom":"no of cut","goodnessIndicator":null,"baseline":">300","commitment":null,"monthly":{}},{"id":"kpi-70","department":"Rail Mill","workstation":"Hot Saw","leader":"Dayal Singh","inceptionDate":null,"kpiName":"Blade life to be maintain for Beams section <500 mm","uom":"no of cut","goodnessIndicator":null,"baseline":">500","commitment":null,"monthly":{}},{"id":"kpi-71","department":"Rail Mill","workstation":"Hot Saw","leader":"Dayal Singh","inceptionDate":null,"kpiName":"Blade life to be maintain for Beams section >500 mm","uom":"no of cut","goodnessIndicator":null,"baseline":">400","commitment":null,"monthly":{}},{"id":"kpi-72","department":"Rail Mill","workstation":"Hot Saw","leader":"Dayal Singh","inceptionDate":null,"kpiName":"Hotsaw availability >=99%","uom":"%","goodnessIndicator":null,"baseline":96.7,"commitment":99.0,"monthly":{"2025-10":96.6,"2025-11":98.6,"2025-12":97.8,"2026-01":98.7,"2026-02":98.7,"2026-03":98.6,"2026-04":98.1}},{"id":"kpi-73","department":"Rail Mill","workstation":"Hot Saw","leader":"Dayal Singh","inceptionDate":null,"kpiName":"Zero Harm","uom":"Nos","goodnessIndicator":null,"baseline":2.0,"commitment":0.0,"monthly":{"2025-04":1.0,"2025-05":1.0,"2025-06":2.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":1.0,"2025-11":3.0,"2025-12":1.0,"2026-01":0.0,"2026-02":2.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-74","department":"Rail Mill","workstation":"Rail Finishing","leader":"Akhilesh Shrivastav","inceptionDate":null,"kpiName":"1200 Mt of Rail FG Materil","uom":"MT","goodnessIndicator":null,"baseline":900.0,"commitment":1200.0,"monthly":{"2026-01":912.0,"2026-02":1375.0,"2026-03":1175.0,"2026-04":1430.8,"2026-05":1092.0}},{"id":"kpi-75","department":"Rail Mill","workstation":"Rail Finishing","leader":"Akhilesh Shrivastav","inceptionDate":null,"kpiName":"Equipment and Bed Utilization","uom":"%","goodnessIndicator":null,"baseline":80.0,"commitment":98.0,"monthly":{"2026-01":60.8,"2026-02":91.6,"2026-03":78.3,"2026-04":95.3,"2026-05":91.0}},{"id":"kpi-76","department":"Rail Mill","workstation":"Rail Finishing","leader":"Akhilesh Shrivastav","inceptionDate":null,"kpiName":"Rejection (Over length, Short length and Out of Square)","uom":"%","goodnessIndicator":null,"baseline":"Zero","commitment":"Zero","monthly":{"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-77","department":"Rail Mill","workstation":"Rail Finishing","leader":"Akhilesh Shrivastav","inceptionDate":null,"kpiName":"Rework (MDF)","uom":"%","goodnessIndicator":null,"baseline":"Zero","commitment":"Zero","monthly":{"2026-01":0.6,"2026-02":0.8,"2026-03":0.8,"2026-04":0.28,"2026-05":0.14}},{"id":"kpi-78","department":"Rail Mill","workstation":"Rail Finishing","leader":"Akhilesh Shrivastav","inceptionDate":null,"kpiName":"Zero Breakdown in CS1 TO CS5","uom":"%","goodnessIndicator":null,"baseline":"Zero","commitment":"Zero","monthly":{"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-79","department":"Rail Mill","workstation":"Rail Finishing","leader":"Akhilesh Shrivastav","inceptionDate":null,"kpiName":"Gag Press Availability","uom":"%","goodnessIndicator":null,"baseline":50.0,"commitment":98.0,"monthly":{"2026-01":74.0,"2026-02":92.0,"2026-03":93.0,"2026-04":95.0,"2026-05":98.0}},{"id":"kpi-80","department":"BF-2","workstation":"Cast House","leader":"N G Pattnaik","inceptionDate":"24.11.25","kpiName":"Runner preparation time","uom":"Min","goodnessIndicator":20.0,"baseline":"30 Min","commitment":"20 Min","monthly":{"2026-01":30.0,"2026-02":28.0,"2026-03":28.0,"2026-04":25.0,"2026-05":25.0}},{"id":"kpi-81","department":"BF-2","workstation":"Cast House","leader":"N G Pattnaik","inceptionDate":"24.11.25","kpiName":"Equipment availability","uom":"Percentage","goodnessIndicator":1.0,"baseline":0.9,"commitment":1.0,"monthly":{"2026-01":0.9,"2026-02":0.93,"2026-03":0.95,"2026-04":0.95,"2026-05":0.95}},{"id":"kpi-82","department":"BF-2","workstation":"Cast House","leader":"N G Pattnaik","inceptionDate":"24.11.25","kpiName":"Mudgun mass consumption","uom":"Rs / Thm","goodnessIndicator":135.0,"baseline":"140 Rs/THM","commitment":"135 Rs/THM","monthly":{"2026-01":140.0,"2026-02":140.0,"2026-03":138.0,"2026-04":137.0,"2026-05":136.0}},{"id":"kpi-83","department":"BF-2","workstation":"Cast House","leader":"N G Pattnaik","inceptionDate":"24.11.25","kpiName":"Zero LTI zero FAC","uom":"Nos.","goodnessIndicator":0.0,"baseline":"0 Nos","commitment":"0 Nos","monthly":{"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-84","department":"BF-2","workstation":"Mudgun Stove & Hydraulic","leader":"Utkal Kesari","inceptionDate":"24.11.25","kpiName":"Zero leakage","uom":"Percentage","goodnessIndicator":0.0,"baseline":0.1,"commitment":0.0,"monthly":{"2026-01":5.0,"2026-02":5.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-85","department":"BF-2","workstation":"Mudgun Stove & Hydraulic","leader":"Utkal Kesari","inceptionDate":"24.11.25","kpiName":"Maintaining NAS value","uom":"Value","goodnessIndicator":5.0,"baseline":"6 Level","commitment":"5 Level","monthly":{"2026-01":6.0,"2026-02":6.0,"2026-03":6.0,"2026-04":6.0,"2026-05":5.0}},{"id":"kpi-86","department":"BF-2","workstation":"Mudgun Stove & Hydraulic","leader":"Utkal Kesari","inceptionDate":"24.11.25","kpiName":"Equipment availability","uom":"Percentage","goodnessIndicator":1.0,"baseline":0.9,"commitment":1.0,"monthly":{"2026-01":0.95,"2026-02":0.95,"2026-03":0.95,"2026-04":1.0,"2026-05":1.0}},{"id":"kpi-87","department":"BF-2","workstation":"Mudgun Stove & Hydraulic","leader":"Utkal Kesari","inceptionDate":"24.11.25","kpiName":"Zero LTI zero FAC","uom":"Nos.","goodnessIndicator":0.0,"baseline":"0 Nos","commitment":"0 Nos","monthly":{"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-88","department":"BF-2","workstation":"PCI Mill 3","leader":"Manish Deshmukh","inceptionDate":"24.11.25","kpiName":"Zero Rejection","uom":"%","goodnessIndicator":0.005,"baseline":0.02,"commitment":0.005,"monthly":{"2026-01":0.02,"2026-02":0.01,"2026-03":0.005,"2026-04":0.005,"2026-05":0.005}},{"id":"kpi-89","department":"BF-2","workstation":"PCI Mill 3","leader":"Manish Deshmukh","inceptionDate":"24.11.25","kpiName":"Grinding Rate","uom":"T/Hr","goodnessIndicator":23.0,"baseline":"22.5 T/Hr","commitment":"23 T/Hr","monthly":{"2026-01":22.5,"2026-02":23.0,"2026-03":23.0,"2026-04":23.0,"2026-05":23.0}},{"id":"kpi-90","department":"BF-2","workstation":"PCI Mill 3","leader":"Manish Deshmukh","inceptionDate":"24.11.25","kpiName":"Equipment availability","uom":"Percentage","goodnessIndicator":0.95,"baseline":0.92,"commitment":0.95,"monthly":{"2026-01":0.92,"2026-02":0.95,"2026-03":0.95,"2026-04":0.95,"2026-05":0.95}},{"id":"kpi-91","department":"BF-2","workstation":"PCI Mill 3","leader":"Manish Deshmukh","inceptionDate":"24.11.25","kpiName":"Zero LTI zero FAC","uom":"Nos.","goodnessIndicator":0.0,"baseline":"0 Nos","commitment":"0 Nos","monthly":{"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-92","department":"BF-1","workstation":"Cast House","leader":"Munnu Singh","inceptionDate":"24.11.2025","kpiName":"Runner preparation time","uom":"Mins","goodnessIndicator":20.0,"baseline":"30 Min","commitment":"20 Min","monthly":{"2026-01":30.0,"2026-02":25.0,"2026-03":25.0,"2026-04":25.0,"2026-05":25.0}},{"id":"kpi-93","department":"BF-1","workstation":"Cast House","leader":"Munnu Singh","inceptionDate":"24.11.2025","kpiName":"Taphole Equipment Availability","uom":"%","goodnessIndicator":1.0,"baseline":0.95,"commitment":1.0,"monthly":{"2026-01":0.95,"2026-02":0.95,"2026-03":0.95,"2026-04":0.95,"2026-05":0.95}},{"id":"kpi-94","department":"BF-1","workstation":"Cast House","leader":"Munnu Singh","inceptionDate":"24.11.2025","kpiName":"Mudgun Consumption","uom":"Rs/Thm","goodnessIndicator":160.0,"baseline":"165 Rs/Thm","commitment":"160Rs/Thm","monthly":{"2026-01":165.0,"2026-02":165.0,"2026-03":165.0,"2026-04":165.0,"2026-05":165.0}},{"id":"kpi-95","department":"BF-1","workstation":"Cast House","leader":"Munnu Singh","inceptionDate":"24.11.2025","kpiName":"Zero LTI / Zero First-aid cases","uom":"Nos","goodnessIndicator":0.0,"baseline":"0 No","commitment":"0 No","monthly":{"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-96","department":"BF-1","workstation":"Integrated Pump House","leader":"Ajit Das","inceptionDate":"24.11.2025","kpiName":"Equipment Availability","uom":"%","goodnessIndicator":1.0,"baseline":0.95,"commitment":1.0,"monthly":{"2026-01":1.0,"2026-02":1.0,"2026-03":1.0,"2026-04":1.0,"2026-05":1.0}},{"id":"kpi-97","department":"BF-1","workstation":"Integrated Pump House","leader":"Ajit Das","inceptionDate":"24.11.2025","kpiName":"Inlet/outlet temperature difference (ΔT)","uom":"Deg C","goodnessIndicator":5.0,"baseline":"4 Deg C","commitment":"5 Deg C","monthly":{"2026-01":7.0,"2026-02":7.0,"2026-03":7.0,"2026-04":7.0,"2026-05":7.0}},{"id":"kpi-98","department":"BF-1","workstation":"Integrated Pump House","leader":"Ajit Das","inceptionDate":"24.11.2025","kpiName":"Cooling circuit leak-free","uom":"Nos","goodnessIndicator":0.0,"baseline":15.0,"commitment":0.0,"monthly":{"2026-01":10.0,"2026-02":5.0,"2026-03":6.0,"2026-04":4.0,"2026-05":4.0}},{"id":"kpi-99","department":"BF-1","workstation":"Integrated Pump House","leader":"Ajit Das","inceptionDate":"24.11.2025","kpiName":"Zero LTI / Zero First-aid cases","uom":"Nos","goodnessIndicator":0.0,"baseline":"0 Nos","commitment":"0 Nos","monthly":{"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-100","department":"BF-1","workstation":"BPRT","leader":"Amit Arya","inceptionDate":"24.11.2025","kpiName":"Equipment Availability","uom":"%","goodnessIndicator":1.0,"baseline":100.0,"commitment":100.0,"monthly":{"2026-01":95.0,"2026-02":97.0,"2026-03":97.0,"2026-04":100.0,"2026-05":100.0}},{"id":"kpi-101","department":"BF-1","workstation":"BPRT","leader":"Amit Arya","inceptionDate":"24.11.2025","kpiName":"Oil Leakage","uom":"Nos","goodnessIndicator":0.0,"baseline":10.0,"commitment":0.0,"monthly":{"2026-01":3.0,"2026-02":2.0,"2026-03":3.0,"2026-04":4.0,"2026-05":4.0}},{"id":"kpi-102","department":"BF-1","workstation":"BPRT","leader":"Amit Arya","inceptionDate":"24.11.2025","kpiName":"Air leakage in blower house","uom":"%","goodnessIndicator":0.0,"baseline":0.02,"commitment":0.0,"monthly":{"2026-01":2.0,"2026-02":2.0,"2026-03":2.0,"2026-04":2.0,"2026-05":2.0}},{"id":"kpi-103","department":"BF-1","workstation":"BPRT","leader":"Amit Arya","inceptionDate":"24.11.2025","kpiName":"Zero LTI / Zero First-aid cases","uom":"Nos","goodnessIndicator":0.0,"baseline":"0 Nos","commitment":"0 Nos","monthly":{"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-104","department":"SMS-2","workstation":"PQ Bay","leader":"Vinay Patnakar","inceptionDate":"01.12.2025","kpiName":"Ensuring Structural integrity 100%","uom":"No of Failure","goodnessIndicator":null,"baseline":1.0,"commitment":0.0,"monthly":{"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-105","department":"SMS-2","workstation":"PQ Bay","leader":"Vinay Patnakar","inceptionDate":"01.12.2025","kpiName":"Turn around time of torpedo <15 min","uom":"Min","goodnessIndicator":null,"baseline":"20 min","commitment":"15 min","monthly":{"2025-12":20.0,"2026-01":19.0,"2026-02":19.0,"2026-03":18.0,"2026-04":17.0}},{"id":"kpi-106","department":"SMS-2","workstation":"PQ Bay","leader":"Vinay Patnakar","inceptionDate":"01.12.2025","kpiName":"Crane Availability>95%","uom":"%","goodnessIndicator":null,"baseline":0.85,"commitment":0.95,"monthly":{"2025-12":85.0,"2026-01":87.0,"2026-02":87.0,"2026-03":90.0,"2026-04":92.0}},{"id":"kpi-107","department":"SMS-2","workstation":"RH Deggaser","leader":"Deepak Gupta","inceptionDate":"01.12.2025","kpiName":"Vessel changeover time<480 min","uom":"Min","goodnessIndicator":null,"baseline":"600 min","commitment":"480 min","monthly":{"2025-12":600.0,"2026-01":560.0,"2026-02":540.0,"2026-03":480.0,"2026-04":470.0}},{"id":"kpi-108","department":"SMS-2","workstation":"RH Deggaser","leader":"Deepak Gupta","inceptionDate":"01.12.2025","kpiName":"H2 PPM<1.5 ppm (Rail heats)","uom":"%","goodnessIndicator":null,"baseline":0.98,"commitment":1.0,"monthly":{"2025-12":98.0,"2026-01":98.7,"2026-02":98.9,"2026-03":99.2,"2026-04":99.4}},{"id":"kpi-109","department":"SMS-2","workstation":"RH Deggaser","leader":"Deepak Gupta","inceptionDate":"01.12.2025","kpiName":"MFL Functionality=100%","uom":"%","goodnessIndicator":null,"baseline":0.8,"commitment":1.0,"monthly":{"2025-12":0.8,"2026-01":82.0,"2026-02":0.85,"2026-03":0.9,"2026-04":0.95}},{"id":"kpi-110","department":"SMS-2","workstation":"RH Deggaser","leader":"Deepak Gupta","inceptionDate":"01.12.2025","kpiName":"RH vessel life>65 heats","uom":"Nos of Heat","goodnessIndicator":null,"baseline":"50 heats","commitment":"65 heats","monthly":{"2025-12":50.0,"2026-01":55.0,"2026-02":65.0,"2026-03":72.0,"2026-04":75.0}},{"id":"kpi-111","department":"SMS-2","workstation":"Slab Caster","leader":"V. Thodge","inceptionDate":"01.12.2025","kpiName":"Hot Charging>90%","uom":"%","goodnessIndicator":null,"baseline":0.765,"commitment":0.9,"monthly":{"2025-12":0.76,"2026-01":0.75,"2026-02":0.8,"2026-03":0.81,"2026-04":0.81}},{"id":"kpi-112","department":"SMS-2","workstation":"Slab Caster","leader":"V. Thodge","inceptionDate":"01.12.2025","kpiName":"Diversion<0.5%(UT Fail, crack)","uom":"%","goodnessIndicator":null,"baseline":0.01,"commitment":0.05,"monthly":{"2025-12":0.07,"2026-01":0.07,"2026-02":0.06,"2026-03":0.054,"2026-04":0.0535}},{"id":"kpi-113","department":"SMS-2","workstation":"Slab Caster","leader":"V. Thodge","inceptionDate":"01.12.2025","kpiName":"Machine Availability>98%","uom":"%","goodnessIndicator":null,"baseline":0.95,"commitment":0.98,"monthly":{"2025-12":95.0,"2026-01":85.0,"2026-02":96.0,"2026-03":96.5,"2026-04":97.0}},{"id":"kpi-114","department":"SMS-2","workstation":"Segment Shop","leader":"Nitigya Ranjan","inceptionDate":"01.04.2026","kpiName":"Availability (Slab Caster)>100%","uom":"%","goodnessIndicator":null,"baseline":0.97,"commitment":1.0,"monthly":{"2026-04":97.0}},{"id":"kpi-115","department":"SMS-3","workstation":"EAF","leader":"Anupam Thakre","inceptionDate":"01.12.2025","kpiName":"Zero Harm","uom":"Nos of incidents","goodnessIndicator":null,"baseline":"Zero","commitment":"Zero","monthly":{"2025-12":"Zero","2026-01":"Zero","2026-02":"Zero","2026-03":"Zero","2026-04":"Zero","2026-05":"Zero"}},{"id":"kpi-116","department":"SMS-3","workstation":"EAF","leader":"Anupam Thakre","inceptionDate":"01.12.2025","kpiName":"Electrode Consumption","uom":"Kg per Ton","goodnessIndicator":null,"baseline":1.0,"commitment":0.98,"monthly":{"2025-12":1.03,"2026-01":0.94,"2026-02":1.07,"2026-03":1.06,"2026-04":1.11,"2026-05":1.09}},{"id":"kpi-117","department":"SMS-3","workstation":"EAF","leader":"Anupam Thakre","inceptionDate":"01.12.2025","kpiName":"Power Consumption","uom":"KWh per Ton","goodnessIndicator":null,"baseline":332.0,"commitment":330.0,"monthly":{"2025-12":340.0,"2026-01":358.0,"2026-02":338.0,"2026-03":339.0,"2026-04":343.0,"2026-05":347.0}},{"id":"kpi-118","department":"SMS-3","workstation":"EAF","leader":"Anupam Thakre","inceptionDate":"01.12.2025","kpiName":"Flux Consumption","uom":"Kg per Ton","goodnessIndicator":null,"baseline":112.0,"commitment":105.0,"monthly":{"2025-12":115.0,"2026-01":127.0,"2026-02":120.0,"2026-03":122.0,"2026-04":128.0,"2026-05":119.0}},{"id":"kpi-119","department":"SMS-3","workstation":"EAF","leader":"Anupam Thakre","inceptionDate":"01.12.2025","kpiName":"Equipment Availability","uom":"%","goodnessIndicator":null,"baseline":0.96,"commitment":0.96,"monthly":{"2025-12":94.9,"2026-01":95.8,"2026-02":96.0,"2026-03":96.7,"2026-04":96.8,"2026-05":95.2}},{"id":"kpi-120","department":"SMS-3","workstation":"Billet Caster","leader":"Suresh Dhiman","inceptionDate":"01.12.2025","kpiName":"Zero Harm","uom":"Nos of incidents","goodnessIndicator":null,"baseline":0.0,"commitment":"Zero","monthly":{"2025-12":"Zero","2026-01":"Zero","2026-02":"Zero","2026-03":"Zero","2026-04":"Zero","2026-05":"Zero"}},{"id":"kpi-121","department":"SMS-3","workstation":"Billet Caster","leader":"Suresh Dhiman","inceptionDate":"01.12.2025","kpiName":"Strand Availability","uom":"%","goodnessIndicator":null,"baseline":98.0,"commitment":98.5,"monthly":{"2025-12":99.4,"2026-01":98.25,"2026-02":98.2,"2026-03":98.25,"2026-04":98.7,"2026-05":98.6}},{"id":"kpi-122","department":"SMS-3","workstation":"Billet Caster","leader":"Suresh Dhiman","inceptionDate":"01.12.2025","kpiName":"FTR","uom":"%","goodnessIndicator":null,"baseline":94.0,"commitment":94.5,"monthly":{"2025-12":96.5,"2026-01":97.0,"2026-02":97.3,"2026-03":97.43,"2026-04":97.81,"2026-05":97.3}},{"id":"kpi-123","department":"SMS-3","workstation":"Billet Caster","leader":"Suresh Dhiman","inceptionDate":"01.12.2025","kpiName":"Qualilty of Product","uom":"%","goodnessIndicator":null,"baseline":99.6,"commitment":99.7,"monthly":{"2025-12":99.98,"2026-01":99.96,"2026-02":99.97,"2026-03":99.97,"2026-04":99.8,"2026-05":99.9}},{"id":"kpi-124","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"Zero Harm","uom":"Nos of incidents","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-12":"Zero","2026-01":"Zero","2026-02":"Zero","2026-03":"Zero","2026-04":"Zero","2026-05":"Zero"}},{"id":"kpi-125","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"Make up Water Consumption","uom":"m3/hr/day","goodnessIndicator":null,"baseline":145.0,"commitment":140.0,"monthly":{"2025-12":129.39,"2026-01":137.71,"2026-02":148.11,"2026-03":149.81,"2026-04":162.0,"2026-05":133.534}},{"id":"kpi-126","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"Equipment Availabitity","uom":"%","goodnessIndicator":null,"baseline":100.0,"commitment":100.0,"monthly":{"2025-12":100.0,"2026-01":100.0,"2026-02":100.0,"2026-03":100.0,"2026-04":100.0,"2026-05":100.0}},{"id":"kpi-127","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"PH","uom":"-","goodnessIndicator":null,"baseline":8.1,"commitment":"7.2-8.3","monthly":{"2025-12":8.6,"2026-01":8.5,"2026-02":8.5,"2026-03":8.6,"2026-04":8.5,"2026-05":8.7}},{"id":"kpi-128","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"Conductivity","uom":"µS/Cm","goodnessIndicator":null,"baseline":654.0,"commitment":"<1500","monthly":{"2025-12":851.0,"2026-01":1020.0,"2026-02":865.0,"2026-03":851.0,"2026-04":965.0,"2026-05":891.0}},{"id":"kpi-129","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"Turbidity","uom":"NTU","goodnessIndicator":null,"baseline":5.0,"commitment":"<10","monthly":{"2025-12":6.0,"2026-01":9.0,"2026-02":5.0,"2026-03":4.0,"2026-04":6.0,"2026-05":4.0}},{"id":"kpi-130","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"Total Hardness","uom":"PPM","goodnessIndicator":null,"baseline":8.0,"commitment":"<10","monthly":{"2025-12":10.0,"2026-01":12.0,"2026-02":10.0,"2026-03":10.0,"2026-04":8.0,"2026-05":8.0}},{"id":"kpi-131","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"TDS","uom":"PPM","goodnessIndicator":null,"baseline":634.0,"commitment":"<1000","monthly":{"2025-12":582.0,"2026-01":661.0,"2026-02":620.0,"2026-03":564.0,"2026-04":657.0,"2026-05":647.0}},{"id":"kpi-132","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"TSS","uom":"PPM","goodnessIndicator":null,"baseline":6.0,"commitment":"<10","monthly":{"2025-12":6.0,"2026-01":8.0,"2026-02":6.0,"2026-03":4.0,"2026-04":5.0,"2026-05":5.0}},{"id":"kpi-133","department":"SMS-3","workstation":"Water Complex","leader":"Sanjiv Pandey","inceptionDate":"01.12.2025","kpiName":"NITRITE","uom":"PPM","goodnessIndicator":null,"baseline":311.0,"commitment":"250-350","monthly":{"2025-12":301.0,"2026-01":306.0,"2026-02":291.0,"2026-03":262.0,"2026-04":307.0,"2026-05":297.0}},{"id":"kpi-134","department":"SMS-3","workstation":"LRF","leader":"Onkarnath","inceptionDate":"11.05.2026","kpiName":"Zero Harm","uom":"Nos of incidents","goodnessIndicator":null,"baseline":"Zero","commitment":"Zero","monthly":{}},{"id":"kpi-135","department":"SMS-3","workstation":"LRF","leader":"Onkarnath","inceptionDate":"11.05.2026","kpiName":"Electrode Consumption","uom":"Kg per Ton","goodnessIndicator":null,"baseline":0.65,"commitment":0.55,"monthly":{}},{"id":"kpi-136","department":"SMS-3","workstation":"LRF","leader":"Onkarnath","inceptionDate":"11.05.2026","kpiName":"Power Consumption","uom":"KWh per Ton","goodnessIndicator":null,"baseline":65.0,"commitment":65.0,"monthly":{}},{"id":"kpi-137","department":"SMS-3","workstation":"LRF","leader":"Onkarnath","inceptionDate":"11.05.2026","kpiName":"Flux Consumption","uom":"Kg per Ton","goodnessIndicator":null,"baseline":null,"commitment":null,"monthly":{}},{"id":"kpi-138","department":"SMS-3","workstation":"LRF","leader":"Onkarnath","inceptionDate":"11.05.2026","kpiName":"Equipment Availability","uom":"%","goodnessIndicator":null,"baseline":1.0,"commitment":1.0,"monthly":{}},{"id":"kpi-139","department":"SMS-3","workstation":"Combi Caster","leader":"Suresh Dhiman","inceptionDate":"11.05.2026","kpiName":"Zero Harm","uom":"Nos of incidents","goodnessIndicator":null,"baseline":null,"commitment":"Zero","monthly":{}},{"id":"kpi-140","department":"SMS-3","workstation":"Combi Caster","leader":"Suresh Dhiman","inceptionDate":"11.05.2026","kpiName":"Strand Availability","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":99.0,"monthly":{}},{"id":"kpi-141","department":"SMS-3","workstation":"Combi Caster","leader":"Suresh Dhiman","inceptionDate":"11.05.2026","kpiName":"FTR","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":95.0,"monthly":{}},{"id":"kpi-142","department":"SMS-3","workstation":"Combi Caster","leader":"Suresh Dhiman","inceptionDate":"11.05.2026","kpiName":"Qualilty of Product","uom":"%","goodnessIndicator":null,"baseline":null,"commitment":99.75,"monthly":{"2026-05":99.74}},{"id":"kpi-143","department":"RMH-1","workstation":"WT-1","leader":"Anand Dubey","inceptionDate":null,"kpiName":"Zero Harm","uom":"Nos","goodnessIndicator":0.0,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-144","department":"RMH-1","workstation":"WT-1","leader":"Anand Dubey","inceptionDate":null,"kpiName":"Wagon tippler availability >90%","uom":"%","goodnessIndicator":">90","baseline":"<80","commitment":">90","monthly":{"2025-04":">90","2025-12":80.0,"2026-01":80.0,"2026-02":80.0,"2026-03":80.0,"2026-04":80.0,"2026-05":80.0}},{"id":"kpi-145","department":"RMH-1","workstation":"WT-1","leader":"Anand Dubey","inceptionDate":null,"kpiName":"Unloading Time of 3.4 Mintus/Wagon","uom":"Minutes","goodnessIndicator":3.4,"baseline":3.4,"commitment":3.4,"monthly":{"2025-04":3.4,"2025-12":3.4,"2026-01":3.4,"2026-02":3.4,"2026-03":3.4,"2026-04":3.4,"2026-05":3.4}},{"id":"kpi-146","department":"RMH-1","workstation":"WT-1","leader":"Anand Dubey","inceptionDate":null,"kpiName":"Unloading Wagon 16box/ hr","uom":"Nos","goodnessIndicator":15.0,"baseline":16.0,"commitment":15.0,"monthly":{"2025-04":15.0,"2025-12":16.0,"2026-01":16.0,"2026-02":16.0,"2026-03":16.0,"2026-04":16.0,"2026-05":16.0}},{"id":"kpi-147","department":"RMH-1","workstation":"WT-1","leader":"Anand Dubey","inceptionDate":null,"kpiName":"Zero Contamination (mixing)","uom":"Nos","goodnessIndicator":0.0,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-148","department":"RMH-1","workstation":"WT-2","leader":"Rishikesh Thawait","inceptionDate":null,"kpiName":"Zero Harm","uom":"Nos","goodnessIndicator":0.0,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-149","department":"RMH-1","workstation":"WT-2","leader":"Rishikesh Thawait","inceptionDate":null,"kpiName":"Wagon tippler availability >90%","uom":"%","goodnessIndicator":">90","baseline":"<85","commitment":">90","monthly":{"2025-04":">90","2025-12":85.0,"2026-01":85.0,"2026-02":85.0,"2026-03":87.0,"2026-04":87.0,"2026-05":88.0}},{"id":"kpi-150","department":"RMH-1","workstation":"WT-2","leader":"Rishikesh Thawait","inceptionDate":null,"kpiName":"Unloading Time of 3.4 Mintus/Wagon","uom":"Minutes","goodnessIndicator":3.4,"baseline":4.0,"commitment":3.4,"monthly":{"2025-04":4.0,"2025-12":4.0,"2026-01":4.0,"2026-02":4.0,"2026-03":4.0,"2026-04":4.0,"2026-05":4.0}},{"id":"kpi-151","department":"RMH-1","workstation":"WT-2","leader":"Rishikesh Thawait","inceptionDate":null,"kpiName":"Unloading Wagon 16box/ hr","uom":"Nos","goodnessIndicator":16.0,"baseline":14.0,"commitment":16.0,"monthly":{"2025-04":14.0,"2025-12":14.0,"2026-01":14.0,"2026-02":14.0,"2026-03":14.0,"2026-04":14.0,"2026-05":14.0}},{"id":"kpi-152","department":"RMH-1","workstation":"WT-2","leader":"Rishikesh Thawait","inceptionDate":null,"kpiName":"Zero Contamination (mixing)","uom":"Nos","goodnessIndicator":0.0,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-153","department":"RMH-1","workstation":"WT-3","leader":"Abhishek Bhol","inceptionDate":null,"kpiName":"Zero Harm","uom":"Nos","goodnessIndicator":0.0,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-154","department":"RMH-1","workstation":"WT-3","leader":"Abhishek Bhol","inceptionDate":null,"kpiName":"Wagon tippler availability >90%","uom":"%","goodnessIndicator":">90","baseline":"<82","commitment":">90","monthly":{"2025-04":85.0,"2025-12":82.0,"2026-01":82.0,"2026-02":82.0,"2026-03":84.0,"2026-04":84.0,"2026-05":84.0}},{"id":"kpi-155","department":"RMH-1","workstation":"WT-3","leader":"Abhishek Bhol","inceptionDate":null,"kpiName":"Unloading Time of 3.4 Mintus/Wagon","uom":"Minutes","goodnessIndicator":3.4,"baseline":4.0,"commitment":3.4,"monthly":{"2025-04":3.5,"2025-12":4.0,"2026-01":4.0,"2026-02":4.0,"2026-03":4.0,"2026-04":4.0,"2026-05":4.0}},{"id":"kpi-156","department":"RMH-1","workstation":"WT-3","leader":"Abhishek Bhol","inceptionDate":null,"kpiName":"Unloading Wagon 16box/ hr","uom":"Nos","goodnessIndicator":16.0,"baseline":15.0,"commitment":16.0,"monthly":{"2025-04":15.0,"2025-12":16.0,"2026-01":16.0,"2026-02":16.0,"2026-03":16.0,"2026-04":16.0,"2026-05":16.0}},{"id":"kpi-157","department":"RMH-1","workstation":"WT-3","leader":"Abhishek Bhol","inceptionDate":null,"kpiName":"Zero Contamination (mixing)","uom":"Nos","goodnessIndicator":0.0,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-158","department":"RMH-1","workstation":"BC-2 (WT-1)","leader":"Anand Dubey","inceptionDate":null,"kpiName":"Zero Harm","uom":"Nos","goodnessIndicator":0.0,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-159","department":"RMH-1","workstation":"BC-2 (WT-1)","leader":"Anand Dubey","inceptionDate":null,"kpiName":"BC-2 Conveyor Beltavailability >90%","uom":"%","goodnessIndicator":">90","baseline":"<85","commitment":">90","monthly":{"2025-04":">90","2025-12":85.0,"2026-01":85.0,"2026-02":85.0,"2026-03":85.0,"2026-04":85.0,"2026-05":85.0}},{"id":"kpi-160","department":"RMH-1","workstation":"BC-2 (WT-1)","leader":"Anand Dubey","inceptionDate":null,"kpiName":"Conveyor Belt Life>547","uom":"Days","goodnessIndicator":547.0,"baseline":365.0,"commitment":547.0,"monthly":{"2025-04":370.0,"2025-12":365.0,"2026-01":365.0,"2026-02":365.0,"2026-03":365.0,"2026-04":365.0,"2026-05":365.0}},{"id":"kpi-161","department":"RMH-1","workstation":"BC-2 (WT-1)","leader":"Anand Dubey","inceptionDate":null,"kpiName":"Zero Contamination (mixing)","uom":"Nos","goodnessIndicator":0.0,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0,"2026-05":0.0}},{"id":"kpi-162","department":"RMH-3","workstation":"Flux Route","leader":"Prabhat Jha","inceptionDate":"2026-12-25","kpiName":"ZERO Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-163","department":"RMH-3","workstation":"Flux Route","leader":"Prabhat Jha","inceptionDate":"2026-12-25","kpiName":"Belt life >3month","uom":"months","goodnessIndicator":null,"baseline":3.0,"commitment":3.5,"monthly":{"2025-04":3.0,"2025-05":3.0,"2025-06":3.0,"2025-07":3.0,"2025-08":3.0,"2025-09":3.0,"2025-10":3.0,"2025-11":3.0,"2025-12":3.0,"2026-01":3.0,"2026-02":3.0,"2026-03":3.0,"2026-04":3.0}},{"id":"kpi-164","department":"RMH-3","workstation":"Flux Route","leader":"Prabhat Jha","inceptionDate":"2026-12-25","kpiName":"Equipment availability >90%","uom":"%","goodnessIndicator":null,"baseline":82.0,"commitment":85.0,"monthly":{"2025-04":86.0,"2025-05":88.0,"2025-06":86.0,"2025-07":79.0,"2025-08":81.0,"2025-09":70.0,"2025-10":79.0,"2025-11":86.0,"2025-12":83.0,"2026-01":78.0,"2026-02":85.0,"2026-03":86.0,"2026-04":86.0}},{"id":"kpi-165","department":"RMH-3","workstation":"Flux Route","leader":"Prabhat Jha","inceptionDate":"2026-12-25","kpiName":"Zero Contamination","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-166","department":"RMH-3","workstation":"Flux Route","leader":"Prabhat Jha","inceptionDate":"2026-12-25","kpiName":"Crushing Index (-3mm) >=90 %","uom":"%","goodnessIndicator":null,"baseline":90.0,"commitment":"91 to 93","monthly":{"2025-04":90.95,"2025-05":90.82,"2025-06":91.245,"2025-07":91.09,"2025-08":90.86,"2025-09":90.98,"2025-10":90.555,"2025-11":91.465,"2025-12":90.94,"2026-01":91.66,"2026-02":90.99,"2026-03":91.125,"2026-04":91.545}},{"id":"kpi-167","department":"RMH-3","workstation":"Blast Furnace-2 Route","leader":"Gurdeep Singh","inceptionDate":"2026-12-25","kpiName":"ZERO Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-168","department":"RMH-3","workstation":"Blast Furnace-2 Route","leader":"Gurdeep Singh","inceptionDate":"2026-12-25","kpiName":"Zero Contamination","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-169","department":"RMH-3","workstation":"Blast Furnace-2 Route","leader":"Gurdeep Singh","inceptionDate":"2026-12-25","kpiName":"changeover time <35 minutes","uom":"minutes","goodnessIndicator":null,"baseline":35.0,"commitment":10.0,"monthly":{"2025-04":20.0,"2025-05":20.0,"2025-06":25.0,"2025-07":30.0,"2025-08":35.0,"2025-09":35.0,"2025-10":35.0,"2025-11":25.0,"2025-12":25.0,"2026-01":20.0,"2026-02":20.0,"2026-03":20.0,"2026-04":20.0}},{"id":"kpi-170","department":"RMH-3","workstation":"Blast Furnace-2 Route","leader":"Gurdeep Singh","inceptionDate":"2026-12-25","kpiName":"Equipment availability >90%","uom":"%","goodnessIndicator":null,"baseline":94.0,"commitment":94.0,"monthly":{"2025-04":94.2,"2025-05":94.2,"2025-06":94.2,"2025-07":94.7,"2025-08":94.8,"2025-09":94.2,"2025-10":94.4,"2025-11":94.2,"2025-12":94.7,"2026-01":94.2,"2026-02":94.0,"2026-03":94.7,"2026-04":94.6}},{"id":"kpi-171","department":"RMH-3","workstation":"Alternate Feeding Route","leader":"PS Ayach","inceptionDate":"2026-12-25","kpiName":"ZERO Harm","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-172","department":"RMH-3","workstation":"Alternate Feeding Route","leader":"PS Ayach","inceptionDate":"2026-12-25","kpiName":"Zero Contamination","uom":"Nos.","goodnessIndicator":null,"baseline":0.0,"commitment":0.0,"monthly":{"2025-04":0.0,"2025-05":0.0,"2025-06":0.0,"2025-07":0.0,"2025-08":0.0,"2025-09":0.0,"2025-10":0.0,"2025-11":0.0,"2025-12":0.0,"2026-01":0.0,"2026-02":0.0,"2026-03":0.0,"2026-04":0.0}},{"id":"kpi-173","department":"RMH-3","workstation":"Alternate Feeding Route","leader":"PS Ayach","inceptionDate":"2026-12-25","kpiName":"Dust emission <60mg/cu.mtr","uom":"mg/cu.mtr","goodnessIndicator":null,"baseline":">1000","commitment":"<60","monthly":{"2025-12":">1000","2026-01":">1000","2026-02":">1000","2026-03":">1000","2026-04":"<60"}},{"id":"kpi-174","department":"RMH-3","workstation":"Alternate Feeding Route","leader":"PS Ayach","inceptionDate":"2026-12-25","kpiName":"Equipment availability >90%","uom":"%","goodnessIndicator":null,"baseline":90.0,"commitment":95.0,"monthly":{"2025-12":100.0,"2026-01":100.0,"2026-02":100.0,"2026-03":100.0,"2026-04":100.0}}];

/* ============================================================================
   DEPARTMENT METADATA
   ============================================================================ */
const DEPARTMENT_IDS = ["PlateMill","Raipur","SPM","Rail Mill","BF-2","BF-1","SMS-2","SMS-3","RMH-1","RMH-3"];
const DEPT_META = {
  PlateMill:   { name: "Plate Mill",               short: "Plate Mill", icon: Layers,      accent: "orange"  },
  Raipur:      { name: "Raipur Workshop",           short: "Raipur",    icon: Wrench,      accent: "sky"     },
  SPM:         { name: "SPM (Special Profile Mill)", short: "SPM",       icon: Factory,     accent: "indigo"  },
  "Rail Mill": { name: "Rail Mill",                 short: "Rail Mill", icon: Hammer,      accent: "amber"   },
  "BF-2":      { name: "Blast Furnace – 2",         short: "BF-2",      icon: Flame,       accent: "red"     },
  "BF-1":      { name: "Blast Furnace – 1",         short: "BF-1",      icon: Flame,       accent: "rose"    },
  "SMS-2":     { name: "Steel Melting Shop – 2",    short: "SMS-2",     icon: FlaskConical,accent: "teal"    },
  "SMS-3":     { name: "Steel Melting Shop – 3",    short: "SMS-3",     icon: FlaskConical,accent: "emerald" },
  "RMH-1":     { name: "Raw Material Handling – 1", short: "RMH-1",     icon: Truck,       accent: "violet"  },
  "RMH-3":     { name: "Raw Material Handling – 3", short: "RMH-3",     icon: Truck,       accent: "fuchsia" },
};
const ACCENT_CLASSES = {
  orange:  { bg:"bg-orange-500",  text:"text-orange-600",  ring:"ring-orange-500",  soft:"bg-orange-50",  border:"border-orange-200",  dot:"bg-orange-500"  },
  sky:     { bg:"bg-sky-500",     text:"text-sky-600",     ring:"ring-sky-500",     soft:"bg-sky-50",     border:"border-sky-200",     dot:"bg-sky-500"     },
  indigo:  { bg:"bg-indigo-500",  text:"text-indigo-600",  ring:"ring-indigo-500",  soft:"bg-indigo-50",  border:"border-indigo-200",  dot:"bg-indigo-500"  },
  amber:   { bg:"bg-amber-500",   text:"text-amber-600",   ring:"ring-amber-500",   soft:"bg-amber-50",   border:"border-amber-200",   dot:"bg-amber-500"   },
  red:     { bg:"bg-red-500",     text:"text-red-600",     ring:"ring-red-500",     soft:"bg-red-50",     border:"border-red-200",     dot:"bg-red-500"     },
  rose:    { bg:"bg-rose-500",    text:"text-rose-600",    ring:"ring-rose-500",    soft:"bg-rose-50",    border:"border-rose-200",    dot:"bg-rose-500"    },
  teal:    { bg:"bg-teal-500",    text:"text-teal-600",    ring:"ring-teal-500",    soft:"bg-teal-50",    border:"border-teal-200",    dot:"bg-teal-500"    },
  emerald: { bg:"bg-emerald-500", text:"text-emerald-600", ring:"ring-emerald-500", soft:"bg-emerald-50", border:"border-emerald-200", dot:"bg-emerald-500" },
  violet:  { bg:"bg-violet-500",  text:"text-violet-600",  ring:"ring-violet-500",  soft:"bg-violet-50",  border:"border-violet-200",  dot:"bg-violet-500"  },
  fuchsia: { bg:"bg-fuchsia-500", text:"text-fuchsia-600", ring:"ring-fuchsia-500", soft:"bg-fuchsia-50", border:"border-fuchsia-200", dot:"bg-fuchsia-500" },
};
const ACCENT_HEX = {
  orange:"#f97316", sky:"#0ea5e9", indigo:"#6366f1", amber:"#f59e0b", red:"#ef4444",
  rose:"#f43f5e", teal:"#14b8a6", emerald:"#10b981", violet:"#8b5cf6", fuchsia:"#d946ef",
};

/* ============================================================================
   PARSING / STATUS ENGINE
   ============================================================================ */
function parseNumeric(raw) {
  if (raw === null || raw === undefined) return { num: null, qualifier: null };
  if (typeof raw === "number") return { num: raw, qualifier: null };
  let s = String(raw).trim();
  if (!s) return { num: null, qualifier: null };
  const lower = s.toLowerCase();
  if (lower === "zero") return { num: 0, qualifier: null };
  if (["nil","na","n/a","-"].includes(lower)) return { num: null, qualifier: null };
  let qualifier = null;
  if (s.startsWith(">") || s.startsWith("≥")) { qualifier = ">"; s = s.slice(1); }
  else if (s.startsWith("<") || s.startsWith("≤")) { qualifier = "<"; s = s.slice(1); }
  const rangeMatch = s.match(/(-?\d+\.?\d*)\s*(?:-|to)\s*(-?\d+\.?\d*)/i);
  if (rangeMatch) { const a = parseFloat(rangeMatch[1]), b = parseFloat(rangeMatch[2]); if (!isNaN(a)&&!isNaN(b)) return { num:(a+b)/2, qualifier }; }
  const numMatch = s.match(/-?\d+\.?\d*/);
  if (numMatch) { const v = parseFloat(numMatch[0]); if (!isNaN(v)) return { num:v, qualifier }; }
  return { num: null, qualifier };
}
function median(arr) {
  if (!arr.length) return null;
  const s = [...arr].sort((a,b)=>a-b); const mid = Math.floor(s.length/2);
  return s.length%2 ? s[mid] : (s[mid-1]+s[mid])/2;
}
const LOWER_BETTER = ["breakdown","rejection","rework","consumption","delay","leakage","contamination","tat","turnaround","changeover","downtime","absenteeism","dust emission","harm","incident","lti","first-aid","first aid","diversion","spillage","damage","complaint","defect","scrap","waste","idle","stoppage","vibration"];
const HIGHER_BETTER = ["availability","production","utilization","oee","dispatch","otif","quality","ftr","uptime","output","efficiency","life","grinding rate","hot charging","strand availability"];
function inferDirection(kpi) {
  const name = (kpi.kpiName||"").toLowerCase();
  if (/[<≤]\s*\d/.test(name)||/\b(max|maximum)\b/.test(name)) return "lower";
  if (/[>≥]\s*\d/.test(name)||/\b(min|minimum)\b/.test(name)) return "higher";
  if (LOWER_BETTER.some(k=>name.includes(k))) return "lower";
  if (HIGHER_BETTER.some(k=>name.includes(k))) return "higher";
  const b=parseNumeric(kpi.baseline).num, c=parseNumeric(kpi.commitment).num;
  if (b!==null&&c!==null&&b!==c) return c>b?"higher":"lower";
  return "unknown";
}
function monthKeysSorted(monthly) { return Object.keys(monthly||{}).sort(); }
function computeStatus(kpi, monthOverride) {
  const months = monthKeysSorted(kpi.monthly);
  if (!months.length) return { status:"nodata",latest:null,latestMonth:null,target:null,direction:null };
  const latestMonth = monthOverride||months[months.length-1];
  if (kpi.monthly[latestMonth]===undefined) return { status:"nodata",latest:null,latestMonth,target:null,direction:null };
  const latestParsed = parseNumeric(kpi.monthly[latestMonth]);
  if (latestParsed.num===null) return { status:"nodata",latest:null,latestMonth,target:null,direction:null };
  let targetParsed = parseNumeric(kpi.commitment);
  if (targetParsed.num===null) targetParsed = parseNumeric(kpi.baseline);
  const direction = inferDirection(kpi);
  if (targetParsed.num===null) return { status:"neutral",latest:latestParsed.num,latestMonth,target:null,direction };
  const allVals = months.map(m=>parseNumeric(kpi.monthly[m]).num).filter(v=>v!==null);
  const med = median(allVals); let targetNum = targetParsed.num;
  if (med!==null) {
    if (med>3&&targetNum<=1.5&&targetNum>=-1.5) targetNum*=100;
    else if (med<=1.5&&med>=-1.5&&targetNum>3) targetNum/=100;
  }
  if (direction==="unknown") return { status:"neutral",latest:latestParsed.num,latestMonth,target:targetNum,direction };
  const tolFloor = med!==null&&Math.abs(med)<=1.5?0.02:0.5;
  const tolerance = Math.max(Math.abs(targetNum)*0.08,tolFloor);
  let status;
  if (direction==="higher") status = latestParsed.num>=targetNum?"green":latestParsed.num>=targetNum-tolerance*2.5?"amber":"red";
  else if (direction==="lower") status = latestParsed.num<=targetNum?"green":latestParsed.num<=targetNum+tolerance*2.5?"amber":"red";
  else status = Math.abs(latestParsed.num-targetNum)<=tolerance?"green":"red";
  return { status,latest:latestParsed.num,latestMonth,target:targetNum,direction };
}
function computeTrend(kpi) {
  const months = monthKeysSorted(kpi.monthly);
  if (months.length<2) return "flat";
  const last=parseNumeric(kpi.monthly[months[months.length-1]]).num;
  const prev=parseNumeric(kpi.monthly[months[months.length-2]]).num;
  if (last===null||prev===null||last===prev) return "flat";
  return last>prev?"up":"down";
}
function momChange(kpi) {
  const months = monthKeysSorted(kpi.monthly);
  if (months.length<2) return null;
  const last=parseNumeric(kpi.monthly[months[months.length-1]]).num;
  const prev=parseNumeric(kpi.monthly[months[months.length-2]]).num;
  if (last===null||prev===null||prev===0) return null;
  return ((last-prev)/Math.abs(prev)*100);
}
function allMonthsAcross(allData) {
  const set=new Set();
  DEPARTMENT_IDS.forEach(d=>(allData[d]||[]).forEach(k=>Object.keys(k.monthly||{}).forEach(m=>set.add(m))));
  return Array.from(set).sort();
}
function plantHealthHistory(allData,n) {
  return allMonthsAcross(allData).slice(-n).map(m=>{
    let green=0,evaluated=0;
    DEPARTMENT_IDS.forEach(d=>(allData[d]||[]).forEach(k=>{
      const s=computeStatus(k,m).status;
      if(s==="green"||s==="amber"||s==="red"){evaluated++;if(s==="green")green++;}
    }));
    return { month:m,label:monthLabel(m),pct:evaluated?Math.round(green/evaluated*100):null,evaluated };
  });
}
function deptHealthHistory(kpis,n) {
  const set=new Set(); kpis.forEach(k=>Object.keys(k.monthly||{}).forEach(m=>set.add(m)));
  return Array.from(set).sort().slice(-n).map(m=>{
    let green=0,evaluated=0;
    kpis.forEach(k=>{const s=computeStatus(k,m).status;if(s==="green"||s==="amber"||s==="red"){evaluated++;if(s==="green")green++;}});
    return { month:m,label:monthLabel(m),pct:evaluated?Math.round(green/evaluated*100):null };
  });
}
function tierColor(pct) { return pct===null||pct===undefined?"#cbd5e1":pct>=80?"#10b981":pct>=50?"#f59e0b":"#ef4444"; }
function truncateLabel(s,n) { return s&&s.length>n?s.slice(0,n-1)+"…":s; }

/* ============================================================================
   FORMATTING HELPERS
   ============================================================================ */
const MONTH_NAMES=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
function monthLabel(m){if(!m)return"—";const[y,mo]=m.split("-");return`${MONTH_NAMES[parseInt(mo,10)-1]} '${y.slice(2)}`;}
function currentMonthStr(){const d=new Date();return`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;}
function addMonths(monthStr,n){const[y,m]=monthStr.split("-").map(Number);const d=new Date(y,m-1+n,1);return`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;}
function fmtVal(v){if(v===null||v===undefined||v==="")return"—";if(typeof v==="number")return Number.isInteger(v)?String(v):v.toFixed(2).replace(/\.00$/,"");return String(v);}
function fmtPct(v){if(v===null)return null;return(v>0?"+":"")+v.toFixed(1)+"%";}
function relTime(iso){if(!iso)return"";const ms=Date.now()-new Date(iso).getTime();const m=Math.floor(ms/60000);if(m<1)return"Just now";if(m<60)return`${m}m ago`;const h=Math.floor(m/60);if(h<24)return`${h}h ago`;return`${Math.floor(h/24)}d ago`;}

function downloadCSV(filename,headerRow,rows){
  const esc=c=>`"${String(c??"").replace(/"/g,'""')}"`;
  const csv=[headerRow,...rows].map(r=>r.map(esc).join(",")).join("\r\n");
  const blob=new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8;"});
  const url=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=url;a.download=filename;document.body.appendChild(a);a.click();
  document.body.removeChild(a);URL.revokeObjectURL(url);
}
function exportExcelReport(allData){
  const wb=XLSX.utils.book_new();
  const summaryRows=[["Department","Health %","On Track","Watch","Off Track","No Data","Leaders"]];
  DEPARTMENT_IDS.forEach(id=>{
    const kpis=allData[id]||[];
    const c={green:0,amber:0,red:0,neutral:0,nodata:0};
    kpis.forEach(k=>c[computeStatus(k).status]++);
    const tracked=c.green+c.amber+c.red,health=tracked?Math.round(c.green/tracked*100):"";
    const leaders=Array.from(new Set(kpis.map(k=>k.leader).filter(Boolean))).join(", ");
    summaryRows.push([DEPT_META[id].name,health,c.green,c.amber,c.red,c.nodata+c.neutral,leaders]);
  });
  const sw=XLSX.utils.aoa_to_sheet(summaryRows);
  sw["!cols"]=[{wch:26},{wch:10},{wch:10},{wch:8},{wch:10},{wch:9},{wch:40}];
  XLSX.utils.book_append_sheet(wb,sw,"Summary");
  DEPARTMENT_IDS.forEach(id=>{
    const kpis=allData[id]||[];
    const monthSet=new Set();kpis.forEach(k=>Object.keys(k.monthly||{}).forEach(m=>monthSet.add(m)));
    const months=Array.from(monthSet).sort();
    const header=["Workstation","Leader","KPI","UoM","Baseline","Commitment",...months.map(monthLabel),"Latest Status"];
    const rows=kpis.map(k=>{const s=computeStatus(k);return[k.workstation,k.leader,k.kpiName,k.uom,fmtVal(k.baseline),fmtVal(k.commitment),...months.map(m=>fmtVal(k.monthly[m])),STATUS_STYLES[s.status].label];});
    const ws=XLSX.utils.aoa_to_sheet([header,...rows]);
    ws["!cols"]=[{wch:18},{wch:16},{wch:34},{wch:8},{wch:10},{wch:10},...months.map(()=>({wch:9})),{wch:12}];
    XLSX.utils.book_append_sheet(wb,ws,DEPT_META[id].short.slice(0,31));
  });
  XLSX.writeFile(wb,`JindalOneView_Full_Report_${currentMonthStr()}.xlsx`);
}

/* ============================================================================
   STORAGE
   ============================================================================ */
async function loadDepartment(deptId){
  try{const res=await window.storage.get(`dept:${deptId}`,true);if(res&&res.value)return JSON.parse(res.value);}catch(e){}
  return SEED_DATA.filter(k=>k.department===deptId);
}
async function saveDepartment(deptId,data){await window.storage.set(`dept:${deptId}`,JSON.stringify(data),true);}

/* ============================================================================
   STATUS STYLES
   ============================================================================ */
const STATUS_STYLES={
  green:  {label:"On Track", dot:"bg-emerald-500",text:"text-emerald-700",bg:"bg-emerald-50",border:"border-emerald-200"},
  amber:  {label:"Watch",    dot:"bg-amber-500",  text:"text-amber-700",  bg:"bg-amber-50",  border:"border-amber-200"  },
  red:    {label:"Off Track",dot:"bg-red-500",    text:"text-red-700",    bg:"bg-red-50",    border:"border-red-200"    },
  neutral:{label:"Tracking", dot:"bg-slate-400",  text:"text-slate-600",  bg:"bg-slate-50",  border:"border-slate-200"  },
  nodata: {label:"No Data",  dot:"bg-slate-300",  text:"text-slate-400",  bg:"bg-slate-50",  border:"border-slate-200"  },
};

/* ============================================================================
   SMALL PRIMITIVES
   ============================================================================ */
function StatusBadge({status,compact}){
  const s=STATUS_STYLES[status]||STATUS_STYLES.nodata;
  return(
    <span className={`inline-flex items-center gap-1.5 rounded-full border ${s.border} ${s.bg} px-2.5 py-1 text-xs font-semibold ${s.text} whitespace-nowrap`}>
      <span className={`h-1.5 w-1.5 rounded-full ${s.dot}`}/>
      {!compact&&s.label}
    </span>
  );
}
function TrendIcon({trend,direction}){
  if(trend==="up") return<ArrowUpRight className={`h-3.5 w-3.5 ${direction==="higher"?"text-emerald-500":direction==="lower"?"text-red-500":"text-slate-400"}`}/>;
  if(trend==="down") return<ArrowDownRight className={`h-3.5 w-3.5 ${direction==="lower"?"text-emerald-500":direction==="higher"?"text-red-500":"text-slate-400"}`}/>;
  return<Minus className="h-3.5 w-3.5 text-slate-300"/>;
}
function Sparkline({kpi}){
  const dark=useDark();
  const months=monthKeysSorted(kpi.monthly).slice(-8);
  const data=months.map(m=>({m,v:parseNumeric(kpi.monthly[m]).num}));
  if(data.filter(d=>d.v!==null).length<2) return<div className="h-8 w-20 flex items-center text-[10px] text-slate-300">–</div>;
  const s=computeStatus(kpi);
  const lineColor=s.status==="green"?"#10b981":s.status==="amber"?"#f59e0b":s.status==="red"?"#ef4444":dark?"#64748b":"#94a3b8";
  return(
    <div className="h-8 w-20 shrink-0">
      <ResponsiveContainer width="100%" height="100%">
        <LineChart data={data}>
          <Line type="monotone" dataKey="v" stroke={lineColor} strokeWidth={1.75} dot={false} isAnimationActive={false}/>
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}

function HealthGauge({percent,size=116,label}){
  const dark=useDark();
  const p=percent===null?0:Math.max(0,Math.min(100,percent));
  const r=size/2-14,cx=size/2,cy=size/2+2;
  const sw=Math.max(6,Math.round(size*0.085));
  const fs=Math.max(13,Math.round(size*0.185));
  const toXY=(deg,radius)=>{const a=(deg*Math.PI)/180;return{x:cx+radius*Math.cos(a),y:cy-radius*Math.sin(a)};};
  const startPt=toXY(180,r),endPt=toXY(0,r),fillDeg=180-(p/100)*180,fillPt=toXY(fillDeg,r);
  const bgPath=`M ${startPt.x} ${startPt.y} A ${r} ${r} 0 0 1 ${endPt.x} ${endPt.y}`;
  const fillPath=`M ${startPt.x} ${startPt.y} A ${r} ${r} 0 0 1 ${fillPt.x} ${fillPt.y}`;
  const color=percent===null?(dark?"#475569":"#cbd5e1"):p>=80?"#10b981":p>=50?"#f59e0b":"#ef4444";
  const ticks=[0,25,50,75,100].map(t=>{const deg=180-(t/100)*180;return{inner:toXY(deg,r-6),outer:toXY(deg,r+2)};});
  return(
    <div className="flex flex-col items-center">
      <svg width={size} height={size/2+22} viewBox={`0 0 ${size} ${size/2+22}`}>
        <path d={bgPath} fill="none" stroke={dark?"#1e293b":"#e2e8f0"} strokeWidth={sw} strokeLinecap="round"/>
        {percent!==null&&<path d={fillPath} fill="none" stroke={color} strokeWidth={sw} strokeLinecap="round"/>}
        {ticks.map((tk,i)=><line key={i} x1={tk.inner.x} y1={tk.inner.y} x2={tk.outer.x} y2={tk.outer.y} stroke={dark?"#334155":"#94a3b8"} strokeWidth={1.5}/>)}
        <text x={cx} y={cy-6} textAnchor="middle" fontSize={fs} fontWeight="700" fill={dark?"#f1f5f9":"#1e293b"}>
          {percent===null?"–":`${Math.round(p)}%`}
        </text>
      </svg>
      {label&&<div className="text-[11px] font-medium text-slate-500 -mt-1">{label}</div>}
    </div>
  );
}

function StatCard({label,value,sub,color,icon:Icon}){
  const dark=useDark();
  const colorMap={emerald:"text-emerald-600",amber:"text-amber-600",red:"text-red-600",slate:"text-slate-500"};
  return(
    <div className={`rounded-xl border border-slate-200 bg-white p-5 shadow-sm flex flex-col justify-between gap-2`}>
      <div className="flex items-start justify-between">
        <div className={`h-9 w-9 rounded-lg ${color==="emerald"?"bg-emerald-50":color==="red"?"bg-red-50":color==="amber"?"bg-amber-50":"bg-slate-100"} flex items-center justify-center`}>
          {Icon&&<Icon className={`h-4.5 w-4.5 ${colorMap[color]}`}/>}
        </div>
      </div>
      <div>
        <div className={`text-3xl font-bold ${colorMap[color]}`}>{value}</div>
        <div className="text-sm font-medium text-slate-600 mt-0.5">{label}</div>
        {sub&&<div className="text-xs text-slate-400 mt-1">{sub}</div>}
      </div>
    </div>
  );
}

function MoMChip({kpi}){
  const change=momChange(kpi);
  if(change===null) return null;
  const dir=inferDirection(kpi);
  const isGood=(dir==="higher"&&change>0)||(dir==="lower"&&change<0);
  const isNeutral=dir==="unknown";
  const color=isNeutral?"text-slate-400":isGood?"text-emerald-600":"text-red-500";
  return(
    <span className={`text-[11px] font-semibold ${color} whitespace-nowrap`}>
      {change>0?"+":""}{change.toFixed(1)}% MoM
    </span>
  );
}

function ModalShell({onClose,children,maxW="max-w-lg"}){
  useEffect(()=>{
    const handler=(e)=>{if(e.key==="Escape")onClose();};
    window.addEventListener("keydown",handler);
    return()=>window.removeEventListener("keydown",handler);
  },[onClose]);
  return(
    <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4 z-50 animate-fadeIn" onClick={onClose}>
      <div className={`bg-white rounded-2xl shadow-2xl w-full ${maxW} max-h-[88vh] overflow-y-auto`} onClick={e=>e.stopPropagation()}>
        {children}
      </div>
    </div>
  );
}

function MethodologyModal({onClose}){
  return(
    <ModalShell onClose={onClose} maxW="max-w-md">
      <div className="p-6 space-y-4">
        <div className="flex items-center justify-between">
          <h3 className="text-lg font-semibold text-slate-800">How Status is Calculated</h3>
          <button onClick={onClose} className="p-1 rounded-lg hover:bg-slate-100"><X className="h-4 w-4 text-slate-400"/></button>
        </div>
        <div className="space-y-3">
          {[
            {color:"bg-emerald-500",label:"On Track",desc:"Latest value meets or exceeds the commitment target."},
            {color:"bg-amber-500",label:"Watch",desc:"Close to target — within ~8% margin. Needs monitoring."},
            {color:"bg-red-500",label:"Off Track",desc:"Significant gap from commitment. Immediate action needed."},
            {color:"bg-slate-300",label:"No Data",desc:"No monthly value logged yet, or target direction unclear."},
          ].map(({color,label,desc})=>(
            <div key={label} className="flex items-start gap-3 p-3 bg-slate-50 rounded-lg">
              <span className={`h-3 w-3 rounded-full ${color} mt-0.5 shrink-0`}/>
              <div><div className="text-sm font-semibold text-slate-700">{label}</div><div className="text-xs text-slate-500 mt-0.5">{desc}</div></div>
            </div>
          ))}
        </div>
        <p className="text-xs text-slate-400 border-t border-slate-100 pt-3">
          Direction (higher vs lower is better) is auto-detected from KPI name and baseline vs commitment comparison.
          For complex KPIs, verify the raw numbers — treat status as a quick first read.
        </p>
      </div>
    </ModalShell>
  );
}

/* ============================================================================
   COMMON DASHBOARD
   ============================================================================ */
function CommonDashboard({allData,onOpenDept}){
  const dark=useDark(); const ct=chartTheme(dark);
  const [search,setSearch]=useState("");
  const [statusFilter,setStatusFilter]=useState("all");
  const [deptFilter,setDeptFilter]=useState("all");
  const [showInfo,setShowInfo]=useState(false);
  const [sortCol,setSortCol]=useState("department");
  const [sortDir,setSortDir]=useState("asc");
  const searchRef=useRef();
  useEffect(()=>{
    const handler=e=>{if(e.key==="/"&&document.activeElement.tagName!=="INPUT"){e.preventDefault();searchRef.current?.focus();}};
    window.addEventListener("keydown",handler);return()=>window.removeEventListener("keydown",handler);
  },[]);
  const flat=useMemo(()=>{const out=[];for(const dept of DEPARTMENT_IDS)for(const kpi of allData[dept]||[])out.push({...kpi,_s:computeStatus(kpi)});return out;},[allData]);
  const counts=useMemo(()=>{const c={green:0,amber:0,red:0,neutral:0,nodata:0};flat.forEach(k=>c[k._s.status]++);return c;},[flat]);
  const trackedTotal=counts.green+counts.amber+counts.red;
  const overallHealth=trackedTotal?Math.round(counts.green/trackedTotal*100):null;
  const safetyKpis=useMemo(()=>flat.filter(k=>/harm|incident|lti|first.aid/i.test(k.kpiName)),[flat]);
  const safetyBreaches=safetyKpis.filter(k=>k._s.latest!==null&&k._s.latest>0);
  const deptStats=useMemo(()=>DEPARTMENT_IDS.map(id=>{
    const kpis=(allData[id]||[]).map(k=>({...k,_s:computeStatus(k)}));
    const c={green:0,amber:0,red:0,neutral:0,nodata:0};kpis.forEach(k=>c[k._s.status]++);
    const tracked=c.green+c.amber+c.red,health=tracked?Math.round(c.green/tracked*100):null;
    const leaders=Array.from(new Set(kpis.map(k=>k.leader).filter(Boolean)));
    return{id,kpis,counts:c,health,leaders};
  }),[allData]);
  const criticalAlerts=useMemo(()=>flat.filter(k=>k._s.status==="red").sort((a,b)=>a.department.localeCompare(b.department)),[flat]);
  const today=currentMonthStr();
  const staleDepts=useMemo(()=>DEPARTMENT_IDS.filter(id=>{const kpis=allData[id]||[];return kpis.length&&!kpis.some(k=>k.monthly&&k.monthly[today]!==undefined);}),[allData,today]);
  const trendHistory=useMemo(()=>plantHealthHistory(allData,12),[allData]);
  const deptBarData=useMemo(()=>deptStats.map(d=>({id:d.id,name:DEPT_META[d.id].short,health:d.health,fill:tierColor(d.health)})).sort((a,b)=>(a.health??-1)-(b.health??-1)),[deptStats]);
  const STATUS_NAME_TO_KEY={"On Track":"green","Watch":"amber","Off Track":"red","No Data":"nodata"};
  const statusPieData=useMemo(()=>[
    {name:"On Track",value:counts.green,color:"#10b981"},
    {name:"Watch",value:counts.amber,color:"#f59e0b"},
    {name:"Off Track",value:counts.red,color:"#ef4444"},
    {name:"No Data",value:counts.nodata+counts.neutral,color:"#cbd5e1"},
  ].filter(d=>d.value>0),[counts]);
  const handleSort=(col)=>{if(sortCol===col)setSortDir(d=>d==="asc"?"desc":"asc");else{setSortCol(col);setSortDir("asc");}};
  const SortBtn=({col})=>{const active=sortCol===col;return(
    <button onClick={()=>handleSort(col)} className="ml-1 inline-flex items-center">
      {active&&sortDir==="asc"?<SortAsc className="h-3 w-3 text-orange-500"/>:active&&sortDir==="desc"?<SortDesc className="h-3 w-3 text-orange-500"/>:<SortAsc className="h-3 w-3 text-slate-300"/>}
    </button>
  );};
  const filteredTable=useMemo(()=>{
    let rows=flat.filter(k=>{
      if(deptFilter!=="all"&&k.department!==deptFilter) return false;
      if(statusFilter!=="all"&&k._s.status!==statusFilter) return false;
      if(search){const hay=`${k.kpiName} ${k.leader} ${k.workstation} ${k.department}`.toLowerCase();if(!hay.includes(search.toLowerCase()))return false;}
      return true;
    });
    rows=[...rows].sort((a,b)=>{
      let av,bv;
      if(sortCol==="department"){av=a.department;bv=b.department;}
      else if(sortCol==="kpi"){av=a.kpiName;bv=b.kpiName;}
      else if(sortCol==="leader"){av=a.leader||"";bv=b.leader||"";}
      else if(sortCol==="latest"){av=a._s.latest??-Infinity;bv=b._s.latest??-Infinity;}
      else if(sortCol==="status"){const o={green:0,amber:1,red:2,neutral:3,nodata:4};av=o[a._s.status]??5;bv=o[b._s.status]??5;}
      else{av=a.department;bv=b.department;}
      if(typeof av==="string")return sortDir==="asc"?av.localeCompare(bv):bv.localeCompare(av);
      return sortDir==="asc"?av-bv:bv-av;
    });
    return rows;
  },[flat,deptFilter,statusFilter,search,sortCol,sortDir]);
  const exportAll=()=>{
    const rows=flat.map(k=>[DEPT_META[k.department]?.name||k.department,k.workstation,k.leader,k.kpiName,k.uom,fmtVal(k.baseline),fmtVal(k.commitment),monthLabel(k._s.latestMonth),fmtVal(k._s.latest),STATUS_STYLES[k._s.status].label]);
    downloadCSV("JindalOneView_Plantwide_Export.csv",["Department","Workstation","Leader","KPI","UoM","Baseline","Commitment","Latest Month","Latest Value","Status"],rows);
  };
  return(
    <div className="space-y-5 animate-fadeIn">
      {showInfo&&<MethodologyModal onClose={()=>setShowInfo(false)}/>}
      {/* Page header */}
      <div className="flex flex-col lg:flex-row lg:items-end lg:justify-between gap-4">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <div className="h-8 w-8 rounded-lg bg-orange-600 flex items-center justify-center"><Activity className="h-4 w-4 text-white"/></div>
            <h1 className="text-2xl lg:text-3xl font-bold text-slate-900" style={{fontFamily:"'Barlow Condensed',sans-serif"}}>Plant-Wide Command Center</h1>
            <button onClick={()=>setShowInfo(true)} className="text-slate-300 hover:text-slate-500 ml-1"><Info className="h-4 w-4"/></button>
          </div>
          <p className="text-sm text-slate-500">Live KPI snapshot across all 10 departments · {flat.length} KPIs · {new Date().toLocaleDateString("en-IN",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</p>
        </div>
        <div className="flex items-center gap-2 no-print flex-wrap">
          <button onClick={()=>window.print()} className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 shadow-sm"><Printer className="h-4 w-4"/>Print Report</button>
          <button onClick={()=>exportExcelReport(allData)} className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 shadow-sm"><FileText className="h-4 w-4"/>Excel Export</button>
          <button onClick={exportAll} className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 shadow-sm"><Download className="h-4 w-4"/>CSV</button>
        </div>
      </div>
      <div className="print-only space-y-1">
        <img src={LOGO_DATA_URI} alt="Jindal Steel" className="h-10"/>
        <div className="text-lg font-bold text-slate-800" style={{fontFamily:"'Barlow Condensed',sans-serif"}}>{APP_NAME} — KPI Report</div>
        <div className="text-xs text-slate-500">Generated {new Date().toLocaleString()} · JSPL Raigarh · TPM</div>
        <hr className="mt-2"/>
      </div>
      {/* Stat strip */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm flex items-center gap-4 lg:col-span-1">
          <HealthGauge percent={overallHealth} size={88}/>
          <div>
            <div className="text-sm font-semibold text-slate-700">Overall Health</div>
            <div className="text-xs text-slate-400 mt-0.5">{trackedTotal} KPIs evaluated</div>
            <div className="text-xs text-slate-400">{counts.nodata+counts.neutral} pending data</div>
          </div>
        </div>
        <StatCard label="On Track" value={counts.green} sub={`of ${trackedTotal} evaluated`} color="emerald" icon={CheckCircle2}/>
        <StatCard label="Watch" value={counts.amber} sub="within tolerance" color="amber" icon={Eye}/>
        <StatCard label="Off Track" value={counts.red} sub="needs attention" color="red" icon={AlertCircle}/>
      </div>
      {/* Safety + stale alerts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className={`rounded-xl border p-4 flex items-center gap-3 shadow-sm ${safetyBreaches.length?"border-red-200 bg-red-50":"border-emerald-200 bg-emerald-50"}`}>
          <div className={`h-10 w-10 rounded-xl flex items-center justify-center shrink-0 ${safetyBreaches.length?"bg-red-100":"bg-emerald-100"}`}>
            <ShieldCheck className={`h-5 w-5 ${safetyBreaches.length?"text-red-600":"text-emerald-600"}`}/>
          </div>
          <div className="flex-1 min-w-0">
            <div className={`text-sm font-semibold ${safetyBreaches.length?"text-red-800":"text-emerald-800"}`}>
              {safetyBreaches.length?`${safetyBreaches.length} safety KPI(s) reporting incidents`:"Zero Harm — no incidents this month"}
            </div>
            <div className="text-xs text-slate-500 mt-0.5">{safetyKpis.length} safety/Zero-Harm KPIs tracked plant-wide</div>
          </div>
        </div>
        {staleDepts.length>0?(
          <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 flex items-center gap-3 shadow-sm">
            <div className="h-10 w-10 rounded-xl bg-amber-100 flex items-center justify-center shrink-0"><Bell className="h-5 w-5 text-amber-600"/></div>
            <div className="flex-1 min-w-0">
              <div className="text-sm font-semibold text-amber-800">{staleDepts.length} dept(s) missing {monthLabel(today)} data</div>
              <div className="text-xs text-amber-700 mt-0.5 truncate">{staleDepts.map(id=>DEPT_META[id].short).join(", ")}</div>
            </div>
            <button onClick={()=>onOpenDept(staleDepts[0])} className="shrink-0 text-xs font-semibold text-amber-800 underline hover:text-amber-900 no-print">Log →</button>
          </div>
        ):(
          <div className="rounded-xl border border-slate-200 bg-white p-4 flex items-center gap-3 shadow-sm">
            <div className="h-10 w-10 rounded-xl bg-slate-100 flex items-center justify-center shrink-0"><CheckCircle2 className="h-5 w-5 text-slate-500"/></div>
            <div><div className="text-sm font-semibold text-slate-700">All departments up to date</div><div className="text-xs text-slate-400 mt-0.5">{monthLabel(today)} data logged across all departments</div></div>
          </div>
        )}
      </div>
      {/* Main grid: trend+bar | pie+right-rail */}
      <div className="grid grid-cols-1 xl:grid-cols-3 gap-4">
        <div className="xl:col-span-2 space-y-4">
          {/* Area chart for plant health trend */}
          {trendHistory.filter(t=>t.pct!==null).length>=2&&(
            <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="flex items-center justify-between mb-3">
                <div>
                  <h2 className="text-sm font-semibold text-slate-800 uppercase tracking-wide">Plant Health Trend</h2>
                  <p className="text-xs text-slate-400 mt-0.5">% KPIs on track, month-over-month</p>
                </div>
                {trendHistory.length>0&&trendHistory[trendHistory.length-1].pct!==null&&(
                  <div className="text-2xl font-bold text-orange-600">{trendHistory[trendHistory.length-1].pct}%</div>
                )}
              </div>
              <div className="h-52">
                <ResponsiveContainer width="100%" height="100%">
                  <AreaChart data={trendHistory} margin={{top:8,right:16,left:-10,bottom:8}}>
                    <defs>
                      <linearGradient id="healthGrad" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%"  stopColor="#ea580c" stopOpacity={0.2}/>
                        <stop offset="95%" stopColor="#ea580c" stopOpacity={0}/>
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={ct.grid}/>
                    <XAxis dataKey="label" tick={{fontSize:10,fill:ct.axis}} angle={-25} textAnchor="end" height={36} interval={trendHistory.length>10?1:0}/>
                    <YAxis tick={{fontSize:11,fill:ct.axis}} domain={[0,100]} width={34}/>
                    <Tooltip contentStyle={ct.tooltip} formatter={v=>[`${v}%`,"On Track"]}/>
                    <Area type="monotone" dataKey="pct" stroke="#ea580c" strokeWidth={2.25} fill="url(#healthGrad)" dot={{r:3,fill:"#ea580c"}} isAnimationActive={false} connectNulls/>
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}
          {/* Department comparison bar */}
          <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
            <div className="flex items-center justify-between mb-3">
              <div>
                <h2 className="text-sm font-semibold text-slate-800 uppercase tracking-wide">Department Comparison</h2>
                <p className="text-xs text-slate-400 mt-0.5">Click any bar to drill into that department</p>
              </div>
            </div>
            <div style={{height:Math.max(220,deptBarData.length*32)}}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={deptBarData} layout="vertical" margin={{top:4,right:28,left:4,bottom:4}}>
                  <CartesianGrid strokeDasharray="3 3" stroke={ct.grid} horizontal={false}/>
                  <XAxis type="number" domain={[0,100]} tick={{fontSize:10,fill:ct.axis}} tickFormatter={v=>`${v}%`}/>
                  <YAxis type="category" dataKey="name" tick={{fontSize:11,fill:dark?"#cbd5e1":"#475569"}} width={88}/>
                  <Tooltip contentStyle={ct.tooltip} formatter={v=>[v===null?"No data":`${v}%`,"On Track"]} cursor={{fill:dark?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.04)"}}/>
                  <Bar dataKey="health" radius={[0,6,6,0]} barSize={16} className="cursor-pointer"
                    onClick={(data)=>{if(data?.id) onOpenDept(data.id);}}>
                    {deptBarData.map((d,i)=><Cell key={i} fill={d.fill}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
        {/* Right rail */}
        <div className="space-y-4">
          {/* Status donut */}
          <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
            <div className="mb-2">
              <h2 className="text-sm font-semibold text-slate-800 uppercase tracking-wide">Status Distribution</h2>
              <p className="text-xs text-slate-400 mt-0.5 no-print">Click a slice to filter the table</p>
            </div>
            <div className="h-52">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={statusPieData} dataKey="value" nameKey="name" innerRadius={50} outerRadius={76} paddingAngle={3} className="cursor-pointer"
                    onClick={d=>{ if(d?.name) setStatusFilter(STATUS_NAME_TO_KEY[d.name]||"all"); }}>
                    {statusPieData.map((d,i)=><Cell key={i} fill={d.color}/>)}
                  </Pie>
                  <Tooltip contentStyle={ct.tooltip}/>
                  <Legend verticalAlign="bottom" height={36} wrapperStyle={{fontSize:11}}/>
                </PieChart>
              </ResponsiveContainer>
            </div>
            {statusFilter!=="all"&&(
              <button onClick={()=>setStatusFilter("all")} className="mt-1 text-xs text-orange-600 hover:text-orange-700 font-medium no-print">
                ✕ Clear filter
              </button>
            )}
          </div>
          {/* Critical alerts */}
          {criticalAlerts.length>0&&(
            <div className="rounded-xl border border-red-200 bg-white shadow-sm overflow-hidden">
              <div className="flex items-center gap-2 bg-red-50 px-4 py-2.5 border-b border-red-200">
                <AlertTriangle className="h-4 w-4 text-red-600 shrink-0"/>
                <span className="text-sm font-semibold text-red-800">Critical Alerts ({criticalAlerts.length})</span>
              </div>
              <div className="divide-y divide-slate-100 max-h-72 overflow-y-auto scroll-area">
                {criticalAlerts.map(k=>(
                  <button key={k.id} onClick={()=>onOpenDept(k.department)} className="w-full flex items-center justify-between gap-3 px-4 py-2.5 hover:bg-slate-50 text-left">
                    <div className="min-w-0">
                      <div className="text-sm font-medium text-slate-800 truncate">{k.kpiName}</div>
                      <div className="text-xs text-slate-400 truncate">{DEPT_META[k.department]?.short} · {k.workstation}</div>
                    </div>
                    <div className="text-right shrink-0">
                      <div className="text-sm font-semibold text-red-600">{fmtVal(k._s.latest)} {k.uom}</div>
                      <div className="text-[11px] text-slate-400">target {fmtVal(k._s.target)}</div>
                    </div>
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
      {/* Department grid */}
      <div>
        <h2 className="text-sm font-semibold text-slate-700 mb-3 uppercase tracking-wide">Departments at a Glance</h2>
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
          {deptStats.map(({id,counts:c,health,leaders})=>{
            const meta=DEPT_META[id],accent=ACCENT_CLASSES[meta.accent],Icon=meta.icon;
            return(
              <button key={id} onClick={()=>onOpenDept(id)}
                className="rounded-xl border border-slate-200 bg-white p-4 text-left hover:shadow-lg hover:border-slate-300 transition-all group shadow-sm">
                <div className="flex items-center justify-between mb-2">
                  <div className={`flex items-center gap-1.5 text-xs font-semibold ${accent.text}`}>
                    <Icon className="h-3.5 w-3.5"/>{meta.short}
                  </div>
                  <ChevronRight className="h-4 w-4 text-slate-200 group-hover:text-slate-400 transition-colors"/>
                </div>
                <div className="flex items-center justify-center -mx-2">
                  <HealthGauge percent={health} size={88}/>
                </div>
                <div className="flex items-center justify-center gap-2 text-[11px] mt-1">
                  <span className="text-emerald-600 font-bold">{c.green}G</span>
                  <span className="text-amber-600 font-bold">{c.amber}A</span>
                  <span className="text-red-600 font-bold">{c.red}R</span>
                </div>
                <div className="text-[11px] text-slate-400 mt-1.5 truncate">{leaders.slice(0,2).join(", ")}{leaders.length>2?` +${leaders.length-2}`:""}</div>
              </button>
            );
          })}
        </div>
      </div>
      {/* Full KPI table */}
      <div className="rounded-xl border border-slate-200 bg-white shadow-sm overflow-hidden">
        <div className="flex flex-col sm:flex-row gap-2 p-3 border-b border-slate-100 bg-slate-50">
          <div className="relative flex-1">
            <Search className="h-4 w-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2"/>
            <input ref={searchRef} value={search} onChange={e=>setSearch(e.target.value)} placeholder='Search KPI, leader, workstation… (press "/" to focus)'
              className="w-full pl-9 pr-3 py-2 text-sm rounded-lg border border-slate-200 focus:outline-none focus:ring-2 focus:ring-orange-400"/>
          </div>
          <select value={deptFilter} onChange={e=>setDeptFilter(e.target.value)} className="text-sm rounded-lg border border-slate-200 px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-400">
            <option value="all">All Departments</option>
            {DEPARTMENT_IDS.map(id=><option key={id} value={id}>{DEPT_META[id].short}</option>)}
          </select>
          <select value={statusFilter} onChange={e=>setStatusFilter(e.target.value)} className="text-sm rounded-lg border border-slate-200 px-3 py-2 focus:outline-none focus:ring-2 focus:ring-orange-400">
            <option value="all">All Status</option>
            <option value="green">On Track</option>
            <option value="amber">Watch</option>
            <option value="red">Off Track</option>
            <option value="nodata">No Data</option>
          </select>
        </div>
        <div className="overflow-x-auto max-h-[520px] scroll-area">
          <table className="w-full text-sm">
            <thead className="bg-white sticky top-0 border-b border-slate-200 z-10">
              <tr className="text-left text-xs text-slate-500 uppercase tracking-wide">
                <th className="px-4 py-2.5 font-medium cursor-pointer select-none" onClick={()=>handleSort("department")}>Dept<SortBtn col="department"/></th>
                <th className="px-4 py-2.5 font-medium">Workstation</th>
                <th className="px-4 py-2.5 font-medium cursor-pointer select-none" onClick={()=>handleSort("kpi")}>KPI<SortBtn col="kpi"/></th>
                <th className="px-4 py-2.5 font-medium cursor-pointer select-none" onClick={()=>handleSort("leader")}>Leader<SortBtn col="leader"/></th>
                <th className="px-4 py-2.5 font-medium cursor-pointer select-none" onClick={()=>handleSort("latest")}>Latest<SortBtn col="latest"/></th>
                <th className="px-4 py-2.5 font-medium">Target</th>
                <th className="px-4 py-2.5 font-medium">MoM</th>
                <th className="px-4 py-2.5 font-medium cursor-pointer select-none" onClick={()=>handleSort("status")}>Status<SortBtn col="status"/></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredTable.map(k=>(
                <tr key={k.id} className="hover:bg-slate-50 cursor-pointer transition-colors" onClick={()=>onOpenDept(k.department)}>
                  <td className="px-4 py-2.5 text-slate-600 whitespace-nowrap font-medium">{DEPT_META[k.department]?.short}</td>
                  <td className="px-4 py-2.5 text-slate-500 whitespace-nowrap text-xs">{k.workstation}</td>
                  <td className="px-4 py-2.5 text-slate-800 font-medium max-w-xs"><div className="truncate" title={k.kpiName}>{k.kpiName}</div></td>
                  <td className="px-4 py-2.5 text-slate-500 whitespace-nowrap text-xs">{k.leader}</td>
                  <td className="px-4 py-2.5 text-slate-700 whitespace-nowrap font-semibold">{fmtVal(k._s.latest)} {k._s.latest!==null?k.uom:""}</td>
                  <td className="px-4 py-2.5 text-slate-400 whitespace-nowrap text-xs">{fmtVal(k._s.target)}</td>
                  <td className="px-4 py-2.5 whitespace-nowrap"><MoMChip kpi={k}/></td>
                  <td className="px-4 py-2.5"><StatusBadge status={k._s.status}/></td>
                </tr>
              ))}
              {filteredTable.length===0&&(
                <tr><td colSpan={8} className="px-4 py-12 text-center text-slate-400">
                  <Search className="h-8 w-8 mx-auto mb-2 text-slate-200"/>No KPIs match this filter.
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
        <div className="px-4 py-2 border-t border-slate-100 bg-slate-50 text-xs text-slate-400">
          Showing {filteredTable.length} of {flat.length} KPIs · Click column header to sort · Press "/" to search
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
   DEPARTMENT DASHBOARD
   ============================================================================ */
function DepartmentDashboard({deptId,kpis,onLogData}){
  const dark=useDark(); const ct=chartTheme(dark);
  const meta=DEPT_META[deptId],accent=ACCENT_CLASSES[meta.accent],Icon=meta.icon;
  const [expanded,setExpanded]=useState(null);
  const enriched=useMemo(()=>kpis.map(k=>({...k,_s:computeStatus(k),_t:computeTrend(k)})),[kpis]);
  const counts=useMemo(()=>{const c={green:0,amber:0,red:0,neutral:0,nodata:0};enriched.forEach(k=>c[k._s.status]++);return c;},[enriched]);
  const tracked=counts.green+counts.amber+counts.red;
  const health=tracked?Math.round(counts.green/tracked*100):null;
  const latestDataMonth=useMemo(()=>{let max=null;enriched.forEach(k=>{const m=monthKeysSorted(k.monthly).pop();if(m&&(!max||m>max))max=m;});return max;},[enriched]);
  const workstations=useMemo(()=>{const order=[],map={};enriched.forEach(k=>{const ws=k.workstation||"General";if(!map[ws]){map[ws]=[];order.push(ws);}map[ws].push(k);});return order.map(ws=>({name:ws,leader:map[ws][0]?.leader,kpis:map[ws]}));},[enriched]);
  const deptTrend=useMemo(()=>deptHealthHistory(kpis,12),[kpis]);
  const wsBarData=useMemo(()=>workstations.map(ws=>{const c={green:0,amber:0,red:0};ws.kpis.forEach(k=>{if(c[k._s.status]!==undefined)c[k._s.status]++;});const t=c.green+c.amber+c.red,h=t?Math.round(c.green/t*100):null;return{name:truncateLabel(ws.name,16),health:h,fill:tierColor(h)};}).sort((a,b)=>(a.health??-1)-(b.health??-1)),[workstations]);
  const exportDept=()=>{
    const rows=enriched.map(k=>[k.workstation,k.leader,k.kpiName,k.uom,fmtVal(k.baseline),fmtVal(k.commitment),monthLabel(k._s.latestMonth),fmtVal(k._s.latest),STATUS_STYLES[k._s.status].label]);
    downloadCSV(`JindalOneView_${meta.short.replace(/\s+/g,"_")}_KPI.csv`,["Workstation","Leader","KPI","UoM","Baseline","Commitment","Latest Month","Latest Value","Status"],rows);
  };
  return(
    <div className="space-y-5 animate-fadeIn">
      {/* Header card */}
      <div className="rounded-xl border border-slate-200 bg-white shadow-sm p-5">
        <div className="flex flex-col lg:flex-row lg:items-center gap-4">
          <div className={`h-14 w-14 rounded-2xl ${accent.soft} flex items-center justify-center shrink-0`}>
            <Icon className={`h-7 w-7 ${accent.text}`}/>
          </div>
          <div className="flex-1">
            <h1 className="text-xl lg:text-2xl font-bold text-slate-900" style={{fontFamily:"'Barlow Condensed',sans-serif"}}>{meta.name}</h1>
            <p className="text-sm text-slate-500 mt-0.5">{workstations.length} workstations · {enriched.length} KPIs tracked · Latest data: {monthLabel(latestDataMonth)}</p>
            <div className="flex items-center gap-3 mt-2">
              <span className="inline-flex items-center gap-1 text-xs font-semibold text-emerald-600 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full">{counts.green} On Track</span>
              <span className="inline-flex items-center gap-1 text-xs font-semibold text-amber-600 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded-full">{counts.amber} Watch</span>
              <span className="inline-flex items-center gap-1 text-xs font-semibold text-red-600 bg-red-50 border border-red-200 px-2 py-0.5 rounded-full">{counts.red} Off Track</span>
            </div>
          </div>
          <HealthGauge percent={health} size={96}/>
          <div className="flex flex-col gap-2 no-print">
            <button onClick={()=>onLogData(deptId)} className={`inline-flex items-center gap-2 rounded-xl ${accent.bg} px-4 py-2.5 text-sm font-semibold text-white hover:opacity-90 shadow-sm`}>
              <ClipboardList className="h-4 w-4"/>Log Monthly Data
            </button>
            <button onClick={exportDept} className="inline-flex items-center justify-center gap-1.5 text-xs text-slate-500 hover:text-slate-700 px-2 py-1.5 rounded-lg hover:bg-slate-100">
              <Download className="h-3.5 w-3.5"/>Export CSV
            </button>
          </div>
        </div>
      </div>
      {/* Analytics: dept trend + workstation bar */}
      {(deptTrend.filter(t=>t.pct!==null).length>=2||wsBarData.length>1)&&(
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {deptTrend.filter(t=>t.pct!==null).length>=2&&(
            <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
              <h2 className="text-sm font-semibold text-slate-800 mb-1 uppercase tracking-wide">{meta.short} Health Trend</h2>
              <div className="h-52">
                <ResponsiveContainer width="100%" height="100%">
                  <AreaChart data={deptTrend} margin={{top:8,right:16,left:-10,bottom:8}}>
                    <defs>
                      <linearGradient id={`grad-${deptId}`} x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor={ACCENT_HEX[meta.accent]} stopOpacity={0.2}/>
                        <stop offset="95%" stopColor={ACCENT_HEX[meta.accent]} stopOpacity={0}/>
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={ct.grid}/>
                    <XAxis dataKey="label" tick={{fontSize:10,fill:ct.axis}} angle={-25} textAnchor="end" height={36} interval={deptTrend.length>10?1:0}/>
                    <YAxis tick={{fontSize:11,fill:ct.axis}} domain={[0,100]} width={34}/>
                    <Tooltip contentStyle={ct.tooltip} formatter={v=>[`${v}%`,"On Track"]}/>
                    <Area type="monotone" dataKey="pct" stroke={ACCENT_HEX[meta.accent]} strokeWidth={2.25} fill={`url(#grad-${deptId})`} dot={{r:3,fill:ACCENT_HEX[meta.accent]}} isAnimationActive={false} connectNulls/>
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}
          {wsBarData.length>1&&(
            <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
              <h2 className="text-sm font-semibold text-slate-800 mb-1 uppercase tracking-wide">Workstation Comparison</h2>
              <div style={{height:Math.max(180,wsBarData.length*36)}}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={wsBarData} layout="vertical" margin={{top:4,right:28,left:4,bottom:4}}>
                    <CartesianGrid strokeDasharray="3 3" stroke={ct.grid} horizontal={false}/>
                    <XAxis type="number" domain={[0,100]} tick={{fontSize:10,fill:ct.axis}} tickFormatter={v=>`${v}%`}/>
                    <YAxis type="category" dataKey="name" tick={{fontSize:11,fill:dark?"#cbd5e1":"#475569"}} width={114}/>
                    <Tooltip contentStyle={ct.tooltip} formatter={v=>[v===null?"No data":`${v}%`,"On Track"]}/>
                    <Bar dataKey="health" radius={[0,6,6,0]} barSize={16}>
                      {wsBarData.map((d,i)=><Cell key={i} fill={d.fill}/>)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}
        </div>
      )}
      {/* Workstation KPI lists */}
      {workstations.map(ws=>(
        <div key={ws.name} className="rounded-xl border border-slate-200 bg-white shadow-sm overflow-hidden">
          <div className="flex items-center justify-between px-4 py-3 bg-slate-50 border-b border-slate-200">
            <div className="flex items-center gap-2">
              <div className={`h-2 w-2 rounded-full ${accent.dot}`}/>
              <span className="text-sm font-semibold text-slate-800">{ws.name}</span>
              {ws.leader&&<span className="inline-flex items-center gap-1 text-xs text-slate-500 bg-white border border-slate-200 px-2 py-0.5 rounded-full"><Users className="h-3 w-3"/>{ws.leader}</span>}
            </div>
            <span className="text-xs font-medium text-slate-400 bg-slate-100 px-2 py-0.5 rounded-full">{ws.kpis.length} KPIs</span>
          </div>
          <div className="divide-y divide-slate-100">
            {ws.kpis.map(k=>(
              <div key={k.id}>
                <button onClick={()=>setExpanded(expanded===k.id?null:k.id)} className="w-full flex items-start sm:items-center gap-3 px-4 py-3 hover:bg-slate-50 text-left transition-colors">
                  <ChevronDown className={`h-4 w-4 text-slate-300 shrink-0 mt-0.5 sm:mt-0 transition-transform ${expanded===k.id?"rotate-0":"-rotate-90"}`}/>
                  <div className="flex-1 min-w-0 flex flex-col sm:flex-row sm:items-center gap-2 sm:gap-3">
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-medium text-slate-800 truncate" title={k.kpiName}>{k.kpiName}</div>
                      <div className="flex items-center gap-2 mt-0.5">
                        <span className="text-xs text-slate-400">Baseline {fmtVal(k.baseline)} · Target {fmtVal(k.commitment)} {k.uom}</span>
                        <MoMChip kpi={k}/>
                      </div>
                    </div>
                    <div className="hidden md:block shrink-0"><Sparkline kpi={k}/></div>
                    <div className="flex items-center justify-between sm:justify-end gap-3 shrink-0">
                      <div className="text-right sm:w-24">
                        <div className="text-sm font-bold text-slate-800">{fmtVal(k._s.latest)} {k._s.latest!==null?k.uom:""}</div>
                        <div className="flex items-center justify-end gap-0.5 text-[11px] text-slate-400">
                          <TrendIcon trend={k._t} direction={inferDirection(k)}/>{monthLabel(k._s.latestMonth)}
                        </div>
                      </div>
                      <div className="shrink-0"><StatusBadge status={k._s.status}/></div>
                    </div>
                  </div>
                </button>
                {expanded===k.id&&<KPIDetail kpi={k}/>}
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function KPIDetail({kpi}){
  const dark=useDark(); const ct=chartTheme(dark);
  const months=monthKeysSorted(kpi.monthly);
  const data=months.map(m=>({month:monthLabel(m),v:parseNumeric(kpi.monthly[m]).num}));
  const target=parseNumeric(kpi.commitment).num,baseline=parseNumeric(kpi.baseline).num;
  const labelsClose=target!==null&&baseline!==null&&Math.abs(target-baseline)<(Math.max(Math.abs(target),Math.abs(baseline),1)*0.12);
  const s=computeStatus(kpi);
  const lineColor=s.status==="green"?"#10b981":s.status==="amber"?"#f59e0b":s.status==="red"?"#ef4444":"#94a3b8";
  return(
    <div className="px-5 pb-5 pt-2 bg-slate-50/50">
      {data.length>=2?(
        <div className="h-64 bg-white rounded-xl border border-slate-100 shadow-sm p-3">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={data} margin={{top:10,right:20,left:-10,bottom:16}}>
              <CartesianGrid strokeDasharray="3 3" stroke={ct.grid}/>
              <XAxis dataKey="month" tick={{fontSize:10,fill:ct.axis}} angle={-30} textAnchor="end" height={34} interval={data.length>12?Math.ceil(data.length/12)-1:0}/>
              <YAxis tick={{fontSize:11,fill:ct.axis}} width={36}/>
              <Tooltip contentStyle={ct.tooltip}/>
              {target!==null&&<ReferenceLine y={target} stroke="#f59e0b" strokeDasharray="4 4" label={{value:"Target",fontSize:10,fill:"#f59e0b",position:"insideTopRight"}}/>}
              {baseline!==null&&<ReferenceLine y={baseline} stroke="#94a3b8" strokeDasharray="2 2" label={labelsClose?undefined:{value:"Baseline",fontSize:10,fill:"#94a3b8",position:"insideBottomLeft"}}/>}
              <Line type="monotone" dataKey="v" stroke={lineColor} strokeWidth={2.5} dot={{r:4,fill:lineColor,strokeWidth:0}} isAnimationActive={false}/>
            </LineChart>
          </ResponsiveContainer>
        </div>
      ):(
        <div className="h-24 flex items-center justify-center text-sm text-slate-400 bg-white rounded-xl border border-slate-100">
          <div className="text-center"><BarChart2 className="h-6 w-6 mx-auto mb-1 text-slate-200"/>Not enough history yet — log data to build the trend.</div>
        </div>
      )}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-3">
        {[["Inception",kpi.inceptionDate||"—"],["UoM",kpi.uom||"—"],["Baseline",fmtVal(kpi.baseline)],["Commitment",fmtVal(kpi.commitment)]].map(([label,value])=>(
          <div key={label} className="bg-white rounded-lg border border-slate-100 shadow-sm px-3 py-2">
            <div className="text-xs text-slate-400">{label}</div>
            <div className="text-sm font-semibold text-slate-700 mt-0.5">{value}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================================================================
   ADD KPI MODAL
   ============================================================================ */
function AddKpiModal({deptId,workstations,onAdd,onClose}){
  const [ws,setWs]=useState(workstations[0]||"");
  const [newWs,setNewWs]=useState("");
  const [leader,setLeader]=useState("");
  const [kpiName,setKpiName]=useState("");
  const [uom,setUom]=useState("");
  const [baseline,setBaseline]=useState("");
  const [commitment,setCommitment]=useState("");
  const finalWs=ws==="__new__"?newWs:ws;
  const submit=()=>{
    if(!finalWs.trim()||!kpiName.trim())return;
    onAdd({id:`kpi-new-${Date.now()}`,department:deptId,workstation:finalWs.trim(),leader:leader.trim(),inceptionDate:null,kpiName:kpiName.trim(),uom:uom.trim(),goodnessIndicator:null,baseline:baseline.trim()||null,commitment:commitment.trim()||null,monthly:{}});
  };
  return(
    <ModalShell onClose={onClose} maxW="max-w-md">
      <div className="p-6 space-y-4">
        <div className="flex items-center justify-between">
          <h3 className="text-lg font-semibold text-slate-800">Add New KPI</h3>
          <button onClick={onClose} className="p-1 rounded-lg hover:bg-slate-100"><X className="h-4 w-4 text-slate-400"/></button>
        </div>
        <div className="space-y-3">
          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1">Workstation</label>
            <select value={ws} onChange={e=>setWs(e.target.value)} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm">
              {workstations.map(w=><option key={w} value={w}>{w}</option>)}
              <option value="__new__">➕ New workstation…</option>
            </select>
            {ws==="__new__"&&<input value={newWs} onChange={e=>setNewWs(e.target.value)} placeholder="New workstation name" className="w-full mt-2 rounded-lg border border-slate-200 px-3 py-2 text-sm"/>}
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1">Leader / Owner</label>
            <input value={leader} onChange={e=>setLeader(e.target.value)} placeholder="e.g. Ramesh Kumar" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"/>
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1">KPI Name</label>
            <input value={kpiName} onChange={e=>setKpiName(e.target.value)} placeholder="e.g. Equipment Availability >90%" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"/>
          </div>
          <div className="grid grid-cols-3 gap-2">
            <div><label className="block text-xs font-medium text-slate-600 mb-1">UoM</label><input value={uom} onChange={e=>setUom(e.target.value)} placeholder="%, Nos" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"/></div>
            <div><label className="block text-xs font-medium text-slate-600 mb-1">Baseline</label><input value={baseline} onChange={e=>setBaseline(e.target.value)} placeholder="0" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"/></div>
            <div><label className="block text-xs font-medium text-slate-600 mb-1">Commitment</label><input value={commitment} onChange={e=>setCommitment(e.target.value)} placeholder="0" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"/></div>
          </div>
        </div>
        <div className="flex justify-end gap-2 pt-1">
          <button onClick={onClose} className="px-4 py-2 text-sm text-slate-500 hover:text-slate-700">Cancel</button>
          <button onClick={submit} className="px-5 py-2 text-sm font-semibold text-white bg-orange-600 rounded-xl hover:bg-orange-700 shadow-sm">Add KPI</button>
        </div>
      </div>
    </ModalShell>
  );
}

/* ============================================================================
   LOG DATA PAGE
   ============================================================================ */
function LogDataPage({deptId,kpis,onSave,onPickDept,allData}){
  const toast=useToast();
  const [draft,setDraft]=useState(()=>[...kpis]);
  const [extraMonths,setExtraMonths]=useState(0);
  const [showAddModal,setShowAddModal]=useState(false);
  const [saving,setSaving]=useState(false);
  const [savedFlash,setSavedFlash]=useState(false);
  const [dirty,setDirty]=useState(false);
  // kpis removed from deps — allData[null]||[] creates new [] every render → infinite loop
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(()=>{setDraft([...kpis]);setExtraMonths(0);setDirty(false);},[deptId]);
  if(!deptId){
    return(
      <div className="space-y-4 animate-fadeIn">
        <div>
          <h1 className="text-xl lg:text-2xl font-bold text-slate-900" style={{fontFamily:"'Barlow Condensed',sans-serif"}}>Log Monthly Data</h1>
          <p className="text-sm text-slate-500 mt-0.5">Select your department to enter this month's actuals.</p>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
          {DEPARTMENT_IDS.map(id=>{
            const meta=DEPT_META[id],accent=ACCENT_CLASSES[meta.accent],Icon=meta.icon;
            const kpis=allData[id]||[];
            const today=currentMonthStr();
            const hasData=kpis.some(k=>k.monthly&&k.monthly[today]!==undefined);
            return(
              <button key={id} onClick={()=>onPickDept(id)} className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm hover:shadow-md hover:border-slate-300 text-left transition-all group">
                <div className="flex items-start justify-between mb-3">
                  <div className={`h-10 w-10 rounded-xl ${accent.soft} flex items-center justify-center`}><Icon className={`h-5 w-5 ${accent.text}`}/></div>
                  {hasData?<span className="text-[10px] bg-emerald-100 text-emerald-600 font-semibold px-1.5 py-0.5 rounded-full">Logged ✓</span>
                           :<span className="text-[10px] bg-amber-100 text-amber-600 font-semibold px-1.5 py-0.5 rounded-full">Pending</span>}
                </div>
                <div className="text-sm font-semibold text-slate-800 group-hover:text-orange-600 transition-colors">{meta.short}</div>
                <div className="text-xs text-slate-400 mt-0.5">{kpis.length} KPIs</div>
              </button>
            );
          })}
        </div>
      </div>
    );
  }
  const meta=DEPT_META[deptId];
  const allMonthsSet=new Set();
  draft.forEach(k=>Object.keys(k.monthly||{}).forEach(m=>allMonthsSet.add(m)));
  let monthsArr=Array.from(allMonthsSet).sort();
  const today=currentMonthStr();
  if(!monthsArr.length)monthsArr=[today];
  let lastMonth=monthsArr[monthsArr.length-1];
  if(lastMonth<today){monthsArr.push(today);lastMonth=today;}
  for(let i=1;i<=extraMonths;i++)monthsArr.push(addMonths(lastMonth,i));
  const workstations=[],wsMap={};
  draft.forEach((k,idx)=>{const ws=k.workstation||"General";if(!wsMap[ws]){wsMap[ws]=[];workstations.push(ws);}wsMap[ws].push(idx);});
  const updateCell=(idx,month,value)=>{
    setDraft(prev=>{const next=[...prev];const rec={...next[idx],monthly:{...next[idx].monthly}};if(value==="")delete rec.monthly[month];else rec.monthly[month]=value;next[idx]=rec;return next;});
    setDirty(true);
  };
  const addKpi=rec=>{setDraft(prev=>[...prev,rec]);setDirty(true);setShowAddModal(false);};
  const removeKpi=idx=>{setDraft(prev=>prev.filter((_,i)=>i!==idx));setDirty(true);};
  const save=async()=>{
    setSaving(true);
    try{await onSave(deptId,draft);setSavedFlash(true);toast("Data saved successfully","success");setTimeout(()=>setSavedFlash(false),2500);}
    catch(e){toast("Save failed — please try again","error");}
    setSaving(false);setDirty(false);
  };
  const exportDept=()=>{
    const rows=draft.map(k=>[k.workstation,k.leader,k.kpiName,k.uom,fmtVal(k.baseline),fmtVal(k.commitment),...monthsArr.map(m=>fmtVal(k.monthly[m]))]);
    downloadCSV(`JindalOneView_${meta.short.replace(/\s+/g,"_")}_LogSheet.csv`,["Workstation","Leader","KPI","UoM","Baseline","Commitment",...monthsArr.map(monthLabel)],rows);
  };
  return(
    <div className="space-y-4 animate-fadeIn">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <button onClick={()=>onPickDept(null)} className="text-xs text-slate-400 hover:text-orange-600 transition-colors mb-1">← All Departments</button>
          <h1 className="text-xl lg:text-2xl font-bold text-slate-900" style={{fontFamily:"'Barlow Condensed',sans-serif"}}>Log Data — {meta.name}</h1>
        </div>
        <div className="flex items-center gap-2 flex-wrap no-print">
          <button onClick={()=>setShowAddModal(true)} className="inline-flex items-center gap-1.5 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 shadow-sm"><Plus className="h-4 w-4"/>Add KPI</button>
          <button onClick={()=>setExtraMonths(n=>n+1)} className="inline-flex items-center gap-1.5 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 shadow-sm"><CalendarPlus className="h-4 w-4"/>Add Month</button>
          <button onClick={exportDept} className="inline-flex items-center gap-1.5 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 shadow-sm"><Download className="h-4 w-4"/>CSV</button>
        </div>
      </div>
      <div className="text-xs text-slate-400 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2">
        💡 Type values directly — "Zero", "NA", "&gt;90" all work. Current month is highlighted in orange. Click <b>Save Changes</b> when done.
      </div>
      <div className="rounded-xl border border-slate-200 bg-white shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="text-sm border-collapse min-w-full">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200">
                <th className="sticky left-0 bg-slate-50 px-3 py-2.5 text-left font-semibold text-slate-600 text-xs uppercase z-10 min-w-[240px] shadow-[1px_0_0_0_#e2e8f0]">KPI</th>
                <th className="px-2 py-2.5 text-left font-medium text-slate-500 text-xs uppercase min-w-[60px]">UoM</th>
                <th className="px-2 py-2.5 text-left font-medium text-slate-500 text-xs uppercase min-w-[80px]">Target</th>
                {monthsArr.map(m=>(
                  <th key={m} className={`px-1 py-2.5 text-center font-medium text-xs uppercase min-w-[90px] ${m===today?"text-orange-600 bg-orange-50":""}`}>
                    {monthLabel(m)}{m===today&&<span className="ml-1 text-[9px] bg-orange-500 text-white px-1 py-0.5 rounded-full">NOW</span>}
                  </th>
                ))}
                <th className="px-2 py-2.5 min-w-[40px]"></th>
              </tr>
            </thead>
            <tbody>
              {workstations.map(ws=>(
                <React.Fragment key={ws}>
                  <tr className="bg-slate-100/60">
                    <td colSpan={monthsArr.length+4} className="px-3 py-1.5 text-xs font-bold text-slate-500 uppercase tracking-wider sticky left-0">
                      {ws} {wsMap[ws][0]!==undefined&&draft[wsMap[ws][0]].leader?`· ${draft[wsMap[ws][0]].leader}`:""}
                    </td>
                  </tr>
                  {wsMap[ws].map(idx=>{
                    const k=draft[idx];
                    return(
                      <tr key={k.id} className="border-b border-slate-50 hover:bg-orange-50/30 transition-colors">
                        <td className="sticky left-0 bg-white px-3 py-2 text-slate-700 text-sm shadow-[1px_0_0_0_#e2e8f0]" title={k.kpiName}><div className="truncate max-w-[220px]">{k.kpiName}</div></td>
                        <td className="px-2 py-2 text-slate-400 text-xs">{k.uom}</td>
                        <td className="px-2 py-2 text-slate-500 text-xs font-medium whitespace-nowrap">{fmtVal(k.commitment)}</td>
                        {monthsArr.map(m=>(
                          <td key={m} className={`px-1 py-1 ${m===today?"bg-orange-50/40":""}`}>
                            <input value={k.monthly[m]??""} onChange={e=>updateCell(idx,m,e.target.value)}
                              placeholder="—" className="w-20 rounded-lg border border-slate-200 px-2 py-1.5 text-xs text-center focus:outline-none focus:ring-2 focus:ring-orange-400 focus:border-orange-400 transition-all"/>
                          </td>
                        ))}
                        <td className="px-2 py-2 text-center">
                          <button onClick={()=>removeKpi(idx)} title="Remove KPI" className="p-1 rounded hover:bg-red-50"><Trash2 className="h-3.5 w-3.5 text-slate-300 hover:text-red-500"/></button>
                        </td>
                      </tr>
                    );
                  })}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      <div className="sticky bottom-4 flex justify-end no-print">
        <button onClick={save} disabled={!dirty||saving}
          className={`inline-flex items-center gap-2 rounded-xl px-6 py-3 text-sm font-bold shadow-xl transition-all ${dirty?"bg-orange-600 text-white hover:bg-orange-700 hover:shadow-2xl":"bg-slate-200 text-slate-400 cursor-not-allowed"}`}>
          {saving?<Loader2 className="h-4 w-4 animate-spin"/>:savedFlash?<Check className="h-4 w-4"/>:<Save className="h-4 w-4"/>}
          {savedFlash?"Saved!":saving?"Saving…":"Save Changes"}
        </button>
      </div>
      {showAddModal&&<AddKpiModal deptId={deptId} workstations={workstations} onAdd={addKpi} onClose={()=>setShowAddModal(false)}/>}
    </div>
  );
}

/* ============================================================================
   APP SHELL
   ============================================================================ */
export default function App(){
  const [allData,setAllData]=useState(null);
  const [view,setView]=useState({type:"common"});
  const [logDeptId,setLogDeptId]=useState(null);
  const [mobileNavOpen,setMobileNavOpen]=useState(false);
  const [roleMode,setRoleMode]=useState("management");
  const [leaderDept,setLeaderDept]=useState(null);
  const [dark,setDark]=useState(false);
  const [refreshing,setRefreshing]=useState(false);
  const [lastRefreshed,setLastRefreshed]=useState(null);

  const overallPlantHealth=useMemo(()=>{
    if(!allData)return null;
    let green=0,evaluated=0;
    DEPARTMENT_IDS.forEach(id=>(allData[id]||[]).forEach(k=>{const s=computeStatus(k).status;if(s==="green"||s==="amber"||s==="red"){evaluated++;if(s==="green")green++;}}));
    return evaluated?Math.round(green/evaluated*100):null;
  },[allData]);

  const fetchAll=useCallback(async()=>{
    const result={};
    for(const dept of DEPARTMENT_IDS)result[dept]=await loadDepartment(dept);
    setAllData(result);setLastRefreshed(new Date().toISOString());
  },[]);

  useEffect(()=>{fetchAll();},[fetchAll]);

  useEffect(()=>{
    (async()=>{
      try{const res=await window.storage.get("ui:darkMode",false);if(res&&res.value==="1")setDark(true);}catch(e){}
      try{const res=await window.storage.get("ui:role",false);if(res&&res.value){const p=JSON.parse(res.value);if(p.roleMode)setRoleMode(p.roleMode);if(p.leaderDept)setLeaderDept(p.leaderDept);}}catch(e){}
    })();
  },[]);

  const toggleDark=()=>setDark(d=>{const next=!d;window.storage.set("ui:darkMode",next?"1":"0",false).catch(()=>{});return next;});
  const persistRole=(rm,ld)=>window.storage.set("ui:role",JSON.stringify({roleMode:rm,leaderDept:ld}),false).catch(()=>{});
  const chooseRoleMode=m=>{setRoleMode(m);persistRole(m,leaderDept);};
  const chooseLeaderDept=id=>{setLeaderDept(id);persistRole(roleMode,id);if(id){setLogDeptId(id);setView({type:"log"});}};
  const openDept=id=>{setView({type:"department",id});setMobileNavOpen(false);};
  const openLog=id=>{setLogDeptId(id);setView({type:"log"});setMobileNavOpen(false);};
  const openLogDefault=()=>{if(roleMode==="leader"&&leaderDept){setLogDeptId(leaderDept);setView({type:"log"});}else setView({type:"log"});setMobileNavOpen(false);};
  const handleSaveDept=useCallback(async(deptId,data)=>{await saveDepartment(deptId,data);setAllData(prev=>({...prev,[deptId]:data}));},[]);
  const refreshAll=async()=>{setRefreshing(true);await fetchAll();setRefreshing(false);};

  const deptHealthDot=id=>{const kpis=allData?allData[id]||[]:[];let worst="green";for(const k of kpis){const s=computeStatus(k).status;if(s==="red"){worst="red";break;}if(s==="amber")worst="amber";}return worst;};
  const deptRedCount=id=>(allData?allData[id]||[]:[]).filter(k=>computeStatus(k).status==="red").length;

  const NavContent=(
    <>
      <div className="px-4 py-5 flex items-center gap-3 border-b border-slate-800">
        <div className="h-10 w-10 rounded-xl bg-white flex items-center justify-center shrink-0 p-1.5 shadow-sm">
          <img src={LOGO_DATA_URI} alt="Jindal Steel" className="max-h-full max-w-full object-contain"/>
        </div>
        <div className="min-w-0">
          <div className="text-white font-bold text-sm leading-tight truncate" style={{fontFamily:"'Barlow Condensed',sans-serif"}}>{APP_NAME}</div>
          <div className="text-slate-500 text-[10.5px] leading-tight truncate mt-0.5">{APP_TAGLINE}</div>
        </div>
      </div>
      <nav className="flex-1 overflow-y-auto py-3 px-2 space-y-0.5">
        <NavItem icon={LayoutDashboard} label="Command Center" active={view.type==="common"} onClick={()=>{setView({type:"common"});setMobileNavOpen(false);}}/>
        <div className="px-3 pt-4 pb-1.5 text-[10px] font-bold text-slate-600 uppercase tracking-widest">Departments</div>
        {DEPARTMENT_IDS.map(id=>{
          const meta=DEPT_META[id],Icon=meta.icon;
          const dot=deptHealthDot(id),dotClass=dot==="red"?"bg-red-500":dot==="amber"?"bg-amber-500":"bg-emerald-500";
          const isMine=roleMode==="leader"&&leaderDept===id;
          const redCount=deptRedCount(id);
          return(
            <button key={id} onClick={()=>openDept(id)}
              className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm transition-colors ${view.type==="department"&&view.id===id?"bg-slate-700 text-white":isMine?"bg-orange-600/15 text-orange-300 ring-1 ring-orange-600/40":"text-slate-400 hover:bg-slate-800 hover:text-slate-200"}`}>
              <Icon className="h-4 w-4 shrink-0"/>
              <span className="flex-1 text-left truncate text-xs">{meta.short}</span>
              {isMine&&<span className="text-[9px] uppercase tracking-wide text-orange-400 shrink-0 font-bold">You</span>}
              {redCount>0&&<span className="h-4 min-w-4 px-0.5 rounded-full bg-red-500 text-white text-[10px] font-bold flex items-center justify-center shrink-0">{redCount}</span>}
              <span className={`h-1.5 w-1.5 rounded-full ${dotClass} shrink-0`}/>
            </button>
          );
        })}
        <div className="px-3 pt-4 pb-1.5 text-[10px] font-bold text-slate-600 uppercase tracking-widest">Data Entry</div>
        <NavItem icon={ClipboardList} label="Log Monthly Data" active={view.type==="log"} onClick={openLogDefault}/>
      </nav>
      {lastRefreshed&&<div className="px-4 py-2 text-[10px] text-slate-600 border-t border-slate-800">Refreshed {relTime(lastRefreshed)}</div>}
      <div className="px-4 py-3 border-t border-slate-800 text-[10.5px] text-slate-600 flex items-center gap-1.5">
        <Zap className="h-3 w-3 text-orange-500 shrink-0"/> Browser trial mode — shared data link
      </div>
    </>
  );

  if(!allData){
    return(
      <ThemeContext.Provider value={dark}>
        <div className={`min-h-[640px] flex items-center justify-center ${dark?"bg-slate-950":"bg-slate-50"}`}>
          <style>{GLOBAL_STYLE}</style>
          <div className="flex flex-col items-center gap-4">
            <img src={LOGO_DATA_URI} alt="Jindal Steel" className="h-12 opacity-80"/>
            <div className="flex items-center gap-2 text-slate-400">
              <Loader2 className="h-5 w-5 animate-spin"/><span className="text-sm">Loading plant KPI data…</span>
            </div>
          </div>
        </div>
      </ThemeContext.Provider>
    );
  }

  let body;
  if(view.type==="common") body=<CommonDashboard allData={allData} onOpenDept={openDept}/>;
  else if(view.type==="department") body=<DepartmentDashboard deptId={view.id} kpis={allData[view.id]||[]} onLogData={openLog}/>;
  else body=<LogDataPage key={logDeptId??'__picker__'} deptId={logDeptId} kpis={logDeptId?allData[logDeptId]||[]:[]} onSave={handleSaveDept} onPickDept={id=>setLogDeptId(id)} allData={allData}/>;

  return(
    <ThemeContext.Provider value={dark}>
    <ToastProvider>
    <div className={`flex h-full min-h-[640px] ${dark?"dark-mode bg-slate-950":"bg-slate-100"}`} style={{fontFamily:"'Inter',sans-serif"}}>
      <style>{GLOBAL_STYLE}</style>
      <aside className="hidden md:flex md:w-64 lg:w-72 flex-col bg-slate-900 shrink-0 no-print shadow-2xl z-10">{NavContent}</aside>
      {mobileNavOpen&&(
        <div className="fixed inset-0 z-40 md:hidden no-print">
          <div className="absolute inset-0 bg-black/50 backdrop-blur-sm" onClick={()=>setMobileNavOpen(false)}/>
          <aside className="absolute left-0 top-0 h-full w-72 bg-slate-900 flex flex-col animate-fadeIn">{NavContent}</aside>
        </div>
      )}
      <div className="flex-1 flex flex-col min-w-0">
        {/* Mobile top bar */}
        <div className="md:hidden flex items-center gap-3 px-4 py-3 bg-slate-900 text-white no-print shrink-0">
          <button onClick={()=>setMobileNavOpen(true)}><Menu className="h-5 w-5"/></button>
          <div className="h-7 w-7 rounded-lg bg-white flex items-center justify-center shrink-0 p-1">
            <img src={LOGO_DATA_URI} alt="Jindal Steel" className="max-h-full max-w-full object-contain"/>
          </div>
          <span className="font-bold text-sm flex-1 truncate" style={{fontFamily:"'Barlow Condensed',sans-serif"}}>{APP_NAME}</span>
        </div>
        {/* Top bar */}
        <div className={`flex items-center justify-between gap-2 px-4 sm:px-6 py-2.5 border-b flex-wrap shrink-0 ${dark?"border-slate-800 bg-slate-900":"border-slate-200 bg-white"} no-print shadow-sm`}>
          <div className="flex items-center gap-2">
            <button onClick={refreshAll} title="Refresh data" className={`p-2 rounded-lg transition-colors ${dark?"text-slate-400 hover:bg-slate-800":"text-slate-400 hover:bg-slate-100"}`}>
              <RefreshCw className={`h-4 w-4 ${refreshing?"animate-spin":""}`}/>
            </button>
            <button onClick={toggleDark} title="Toggle dark mode" className={`p-2 rounded-lg transition-colors ${dark?"text-amber-300 hover:bg-slate-800":"text-slate-500 hover:bg-slate-100"}`}>
              {dark?<Sun className="h-4 w-4"/>:<Moon className="h-4 w-4"/>}
            </button>
            {overallPlantHealth!==null&&(
              <button onClick={()=>setView({type:"common"})} title="Go to Command Center"
                className={`hidden sm:inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold transition-colors ${overallPlantHealth>=80?dark?"bg-emerald-500/15 text-emerald-300":"bg-emerald-50 text-emerald-700 border border-emerald-200":overallPlantHealth>=50?dark?"bg-amber-500/15 text-amber-300":"bg-amber-50 text-amber-700 border border-amber-200":dark?"bg-red-500/15 text-red-300":"bg-red-50 text-red-700 border border-red-200"}`}>
                <span className={`h-1.5 w-1.5 rounded-full ${overallPlantHealth>=80?"bg-emerald-500":overallPlantHealth>=50?"bg-amber-500":"bg-red-500"}`}/>
                Plant Health {overallPlantHealth}%
              </button>
            )}
          </div>
          <div className="flex items-center gap-2">
            <span className="text-xs text-slate-400 hidden sm:inline">Viewing as</span>
            <div className={`inline-flex rounded-xl border overflow-hidden text-xs shadow-sm ${dark?"border-slate-700":"border-slate-200"}`}>
              <button onClick={()=>chooseRoleMode("management")} className={`px-3 py-1.5 font-semibold transition-colors ${roleMode==="management"?"bg-orange-600 text-white":dark?"bg-slate-900 text-slate-400 hover:bg-slate-800":"bg-white text-slate-600 hover:bg-slate-50"}`}>Top Management</button>
              <button onClick={()=>chooseRoleMode("leader")} className={`px-3 py-1.5 font-semibold transition-colors ${roleMode==="leader"?"bg-orange-600 text-white":dark?"bg-slate-900 text-slate-400 hover:bg-slate-800":"bg-white text-slate-600 hover:bg-slate-50"}`}>Dept Leader</button>
            </div>
            {roleMode==="leader"&&(
              <select value={leaderDept||""} onChange={e=>chooseLeaderDept(e.target.value||null)}
                className={`text-xs rounded-xl border px-2.5 py-1.5 font-medium shadow-sm ${dark?"bg-slate-900 border-slate-700 text-slate-200":"bg-white border-slate-200 text-slate-700"}`}>
                <option value="">Select department…</option>
                {DEPARTMENT_IDS.map(id=><option key={id} value={id}>{DEPT_META[id].short}</option>)}
              </select>
            )}
          </div>
        </div>
        {/* Main content */}
        <main className="flex-1 overflow-y-auto p-4 sm:p-6 lg:p-8 print-area">
          <div className="max-w-[1920px] mx-auto">{body}</div>
        </main>
      </div>
    </div>
    </ToastProvider>
    </ThemeContext.Provider>
  );
}

function NavItem({icon:Icon,label,active,onClick}){
  return(
    <button onClick={onClick} className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-xs font-medium transition-colors ${active?"bg-slate-700 text-white":"text-slate-400 hover:bg-slate-800 hover:text-slate-200"}`}>
      <Icon className="h-4 w-4 shrink-0"/><span className="flex-1 text-left">{label}</span>
    </button>
  );
}
