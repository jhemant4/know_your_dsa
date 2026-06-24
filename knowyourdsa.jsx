import { useState, useRef, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell
} from "recharts";
import {
  Upload, X, MapPin, Building2, Activity, TrendingUp, TrendingDown,
  Home, Plus, Minus, Filter, RefreshCw, ChevronDown, Layers
} from "lucide-react";

// ─── CITY COORDINATES ─────────────────────────────────────────────────────────
const CITY_COORDINATES = {
  "hyderabad": { lat: 17.3850, lng: 78.4867 },
  "chennai": { lat: 13.0827, lng: 80.2707 },
  "mumbai": { lat: 19.0760, lng: 72.8777 },
  "bengaluru": { lat: 12.9716, lng: 77.5946 },
  "bangalore": { lat: 12.9716, lng: 77.5946 },
  "delhi": { lat: 28.6139, lng: 77.2090 },
  "thane": { lat: 19.2183, lng: 72.9781 },
  "pune": { lat: 18.5204, lng: 73.8567 },
  "ahmedabad": { lat: 23.0225, lng: 72.5714 },
  "kolkata": { lat: 22.5726, lng: 88.3639 },
  "kolkatta": { lat: 22.5726, lng: 88.3639 },
  "navi mumbai": { lat: 19.0330, lng: 73.0297 },
  "faridabad": { lat: 28.4089, lng: 77.3178 },
  "ghaziabad": { lat: 28.6692, lng: 77.4538 },
  "gurugram": { lat: 28.4595, lng: 77.0266 },
  "gurgaon": { lat: 28.4595, lng: 77.0266 },
  "chandigarh": { lat: 30.7333, lng: 76.7794 },
  "jaipur": { lat: 26.9124, lng: 75.7873 },
  "goregaon": { lat: 19.1663, lng: 72.8526 },
  "indore": { lat: 22.7196, lng: 75.8577 },
  "nagpur": { lat: 21.1458, lng: 79.0882 },
  "bhopal": { lat: 23.2599, lng: 77.4126 },
  "surat": { lat: 21.1702, lng: 72.8311 },
  "vadodara": { lat: 22.3072, lng: 73.1812 },
  "rajkot": { lat: 22.3039, lng: 70.8022 },
  "lucknow": { lat: 26.8467, lng: 80.9462 },
  "kanpur": { lat: 26.4499, lng: 80.3319 },
  "agra": { lat: 27.1767, lng: 78.0081 },
  "nashik": { lat: 19.9975, lng: 73.7898 },
  "aurangabad": { lat: 19.8762, lng: 75.3433 },
  "kolhapur": { lat: 16.7050, lng: 74.2433 },
  "solapur": { lat: 17.6868, lng: 75.9064 },
  "coimbatore": { lat: 11.0168, lng: 76.9558 },
  "salem": { lat: 11.6643, lng: 78.1460 },
  "madurai": { lat: 9.9252, lng: 78.1198 },
  "tiruchirappalli": { lat: 10.7905, lng: 78.7047 },
  "vellore": { lat: 12.9165, lng: 79.1325 },
  "kanchipuram": { lat: 12.8185, lng: 79.6947 },
  "patna": { lat: 25.5941, lng: 85.1376 },
  "raipur": { lat: 21.2514, lng: 81.6296 },
  "bhubaneswar": { lat: 20.2961, lng: 85.8245 },
  "guwahati": { lat: 26.1445, lng: 91.7362 },
  "dehradun": { lat: 30.3165, lng: 78.0322 },
  "ludhiana": { lat: 30.9010, lng: 75.8573 },
  "amritsar": { lat: 31.6340, lng: 74.8723 },
  "jalandhar": { lat: 31.3260, lng: 75.5762 },
  "jodhpur": { lat: 26.2389, lng: 73.0243 },
  "udaipur": { lat: 24.5854, lng: 73.7125 },
  "ajmer": { lat: 26.4499, lng: 74.6399 },
  "kota": { lat: 25.2138, lng: 75.8648 },
  "jabalpur": { lat: 23.1815, lng: 79.9864 },
  "ratlam": { lat: 23.3315, lng: 75.0367 },
  "ujjain": { lat: 23.1765, lng: 75.7885 },
  "visakhapatnam": { lat: 17.6868, lng: 83.2185 },
  "vijayawada": { lat: 16.5062, lng: 80.6480 },
  "nellore": { lat: 14.4426, lng: 79.9865 },
  "kurnool": { lat: 15.8281, lng: 78.0373 },
  "anantapur": { lat: 14.6819, lng: 77.6006 },
  "warangal": { lat: 17.9784, lng: 79.5941 },
  "moradabad": { lat: 28.8386, lng: 78.7733 },
  "gorakhpur": { lat: 26.7606, lng: 83.3732 },
  "meerut": { lat: 28.9845, lng: 77.7064 },
  "varanasi": { lat: 25.3176, lng: 82.9739 },
  "allahabad": { lat: 25.4358, lng: 81.8463 },
  "noida": { lat: 28.5355, lng: 77.3910 },
  "mysore": { lat: 12.2958, lng: 76.6394 },
  "hubli": { lat: 15.3647, lng: 75.1240 },
  "mangalore": { lat: 12.9141, lng: 74.8560 },
  "kochi": { lat: 9.9312, lng: 76.2673 },
  "kozhikode": { lat: 11.2588, lng: 75.7804 },
  "thrissur": { lat: 10.5276, lng: 76.2144 },
  "pondicherry": { lat: 11.9416, lng: 79.8083 },
  "tumkur": { lat: 13.3379, lng: 77.1173 },
  "bellary": { lat: 15.1394, lng: 76.9214 },
  "raichur": { lat: 16.2120, lng: 77.3439 },
  "gautam buddha nagar": { lat: 28.5355, lng: 77.3910 },
  "kamrup": { lat: 26.1445, lng: 91.7362 },
};

const DELHI_VARIANTS = [
  "central delhi", "north delhi", "east delhi", "west delhi",
  "south delhi", "south west delhi", "new delhi", "north west delhi", "north east delhi", "shahdara"
];

function getCityCoords(cityRaw) {
  if (!cityRaw) return null;
  const city = cityRaw.trim().toLowerCase();
  if (DELHI_VARIANTS.some(v => city.includes(v) || v.includes(city))) {
    return { lat: 28.6139 + (Math.random() - 0.5) * 0.1, lng: 77.2090 + (Math.random() - 0.5) * 0.1 };
  }
  if (city.includes("k.v.rangareddy") || city.includes("rangareddy") || city.includes("ranga reddy")) {
    return { lat: 17.3850 + (Math.random() - 0.5) * 0.05, lng: 78.4867 + (Math.random() - 0.5) * 0.05 };
  }
  if (city.includes("chattisgarh") || city.includes("chhattisgarh")) return CITY_COORDINATES["raipur"];
  return CITY_COORDINATES[city] || null;
}

// ─── PROJECTION ───────────────────────────────────────────────────────────────
const SVG_W = 800, SVG_H = 900;
function project(lat, lng) {
  const x = ((lng - 68) / (97 - 68)) * SVG_W;
  const y = ((37 - lat) / (37 - 8)) * SVG_H;
  return { x, y };
}

// ─── STATUS LOGIC ─────────────────────────────────────────────────────────────
function getStatus(login, disbursal) {
  const l = (login || "").trim().toLowerCase();
  const d = (disbursal || "").trim().toLowerCase();
  if (l === "active" && d === "active") return "active";
  if (l === "active" && d !== "active") return "moderate";
  if (l !== "active" && d !== "active" && (l || d)) return "inactive";
  return "unknown";
}

const STATUS_COLOR = {
  active: "#22c55e",
  moderate: "#eab308",
  inactive: "#ef4444",
  unknown: "#94a3b8",
};
const STATUS_LABEL = { active: "Active", moderate: "Moderate", inactive: "Inactive", unknown: "Unknown" };

// ─── INDIA SVG PATHS ──────────────────────────────────────────────────────────
// Simplified but geographically representative state paths for 800×900 viewBox
const INDIA_STATES = [
  { id: "JK", name: "Jammu & Kashmir", d: "M220,10 L280,15 L310,30 L320,60 L300,80 L280,90 L260,85 L240,75 L220,60 L210,40 Z" },
  { id: "LA", name: "Ladakh", d: "M280,15 L370,10 L400,25 L410,55 L380,75 L340,80 L310,70 L310,30 Z" },
  { id: "HP", name: "Himachal Pradesh", d: "M240,75 L280,90 L300,100 L295,120 L270,130 L250,125 L235,110 L230,95 Z" },
  { id: "PB", name: "Punjab", d: "M210,75 L240,75 L235,110 L220,120 L200,115 L190,95 L195,80 Z" },
  { id: "UK", name: "Uttarakhand", d: "M295,100 L330,95 L345,110 L340,130 L315,140 L295,135 L285,120 L295,110 Z" },
  { id: "HR", name: "Haryana", d: "M200,110 L235,110 L250,125 L245,145 L225,155 L205,150 L195,135 L195,120 Z" },
  { id: "DL", name: "Delhi", d: "M225,145 L240,140 L245,150 L240,158 L225,155 Z" },
  { id: "RJ", name: "Rajasthan", d: "M130,120 L200,115 L205,150 L225,155 L230,200 L220,250 L200,280 L170,295 L140,290 L115,265 L100,235 L95,200 L105,165 L120,140 Z" },
  { id: "UP", name: "Uttar Pradesh", d: "M245,145 L340,130 L370,140 L390,160 L395,185 L380,210 L355,225 L325,235 L290,240 L260,235 L240,220 L230,200 L225,155 Z" },
  { id: "BR", name: "Bihar", d: "M370,180 L415,175 L440,185 L445,205 L430,220 L405,225 L380,220 L365,205 Z" },
  { id: "SK", name: "Sikkim", d: "M445,175 L455,170 L465,178 L460,190 L448,190 Z" },
  { id: "AR", name: "Arunachal Pradesh", d: "M490,140 L570,135 L590,150 L585,170 L555,175 L520,170 L495,165 L488,155 Z" },
  { id: "NL", name: "Nagaland", d: "M560,175 L585,170 L590,185 L575,200 L555,200 L548,190 Z" },
  { id: "MN", name: "Manipur", d: "M560,200 L580,198 L585,215 L575,228 L558,228 L552,215 Z" },
  { id: "MZ", name: "Mizoram", d: "M550,228 L568,228 L572,245 L560,255 L545,250 L542,238 Z" },
  { id: "TR", name: "Tripura", d: "M535,230 L548,228 L552,245 L542,252 L530,248 L527,238 Z" },
  { id: "ML", name: "Meghalaya", d: "M480,200 L525,196 L530,210 L515,218 L485,215 L475,208 Z" },
  { id: "AS", name: "Assam", d: "M460,190 L555,185 L560,200 L548,210 L525,210 L490,208 L465,205 Z" },
  { id: "WB", name: "West Bengal", d: "M440,200 L465,198 L470,225 L465,255 L450,280 L435,285 L420,265 L415,240 L420,220 Z" },
  { id: "JH", name: "Jharkhand", d: "M390,220 L440,215 L445,240 L435,260 L415,265 L390,260 L375,245 L375,230 Z" },
  { id: "OD", name: "Odisha", d: "M395,260 L440,255 L455,270 L460,300 L450,325 L430,335 L405,330 L385,315 L380,290 L385,268 Z" },
  { id: "CG", name: "Chhattisgarh", d: "M290,255 L375,248 L385,270 L385,310 L365,330 L335,335 L305,325 L285,305 L280,280 L282,260 Z" },
  { id: "MP", name: "Madhya Pradesh", d: "M175,220 L260,215 L285,235 L285,265 L265,280 L230,285 L195,278 L165,260 L155,235 L160,220 Z" },
  { id: "GJ", name: "Gujarat", d: "M85,235 L135,235 L155,250 L160,275 L150,305 L130,325 L105,330 L80,320 L65,300 L60,270 L70,248 Z" },
  { id: "MH", name: "Maharashtra", d: "M150,280 L220,278 L260,280 L285,300 L285,330 L270,355 L245,370 L215,375 L185,368 L160,355 L140,335 L130,310 L135,290 Z" },
  { id: "GA", name: "Goa", d: "M155,370 L170,368 L175,380 L165,388 L153,383 Z" },
  { id: "KA", name: "Karnataka", d: "M150,340 L215,335 L245,355 L255,380 L250,410 L235,435 L210,445 L180,445 L155,430 L138,410 L135,385 L140,360 Z" },
  { id: "TG", name: "Telangana", d: "M260,290 L310,285 L335,300 L340,325 L325,345 L300,355 L275,350 L258,335 L255,315 Z" },
  { id: "AP", name: "Andhra Pradesh", d: "M255,345 L335,340 L365,355 L375,380 L365,410 L345,428 L315,435 L285,430 L260,415 L248,390 L250,365 Z" },
  { id: "TN", name: "Tamil Nadu", d: "M220,450 L290,440 L330,455 L340,480 L330,510 L310,530 L285,540 L260,535 L240,515 L225,490 L218,468 Z" },
  { id: "KL", name: "Kerala", d: "M195,440 L225,445 L228,475 L220,505 L205,525 L188,520 L178,498 L180,470 L190,450 Z" },
  { id: "PY", name: "Puducherry", d: "M296,488 L305,486 L306,496 L296,497 Z" },
  { id: "AN", name: "Andaman & Nicobar Islands", d: "M640,420 L648,418 L650,435 L645,450 L636,445 L634,430 Z" },
  { id: "LD", name: "Lakshadweep", d: "M120,490 L127,488 L128,498 L120,500 Z" },
];

// ─── PARSE EXCEL ─────────────────────────────────────────────────────────────
function parseExcel(file, callback) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets["Sheet1"] || workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    const branches = rows.map((row, idx) => {
      const loginVal = row["MTD Jun'26 Login Active - Andromeda"] || "";
      const disbursalVal = row["MTD Jun'26 Disbursal Active - Andromeda"] || "";
      const status = getStatus(loginVal, disbursalVal);
      const coords = getCityCoords(row["CITY"]);
      return {
        id: row["Branch Code"] || `BR${idx}`,
        branchCode: String(row["Branch Code"] || ""),
        branchName: row["BRANCH_NAME"] || row["DHFL_NAME"] || "",
        city: row["CITY"] || "",
        state: row["STATE"] || "",
        address: row["BRANCH_ADDRESS"] || "",
        pincode: row["PINCODE"] || "",
        zone: row["ZONE_HL"] || "",
        geo: row["GEO_HL"] || "",
        cluster: row["CLUSTER_HL"] || "",
        loginActive: loginVal,
        disbursalActive: disbursalVal,
        activationStatus: status,
        hl: row["HL"] || "",
        msme: row["MSME"] || "",
        sahyog: row["SAHYOG"] || "",
        coords: coords,
        unmapped: !coords,
      };
    });
    callback(branches);
  };
  reader.readAsArrayBuffer(file);
}

// ─── KPI CARD ─────────────────────────────────────────────────────────────────
function KpiCard({ label, value, color, icon: Icon, sub }) {
  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-4 flex items-center gap-3 min-w-0">
      <div className="rounded-lg p-2 flex-shrink-0" style={{ backgroundColor: color + "20" }}>
        <Icon size={20} style={{ color }} />
      </div>
      <div className="min-w-0">
        <div className="text-2xl font-bold text-slate-800 leading-tight">{value}</div>
        <div className="text-xs text-slate-500 font-medium mt-0.5 leading-tight">{label}</div>
        {sub && <div className="text-xs text-slate-400 mt-0.5">{sub}</div>}
      </div>
    </div>
  );
}

// ─── SELECT DROPDOWN ──────────────────────────────────────────────────────────
function Select({ value, onChange, options, label }) {
  return (
    <div className="relative">
      <select
        value={value}
        onChange={e => onChange(e.target.value)}
        className="appearance-none bg-white border border-slate-200 rounded-lg px-3 py-2 pr-8 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-300 cursor-pointer"
      >
        <option value="">{label}</option>
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
      <ChevronDown size={14} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
    </div>
  );
}

// ─── SIDE PANEL ───────────────────────────────────────────────────────────────
function SidePanel({ branch, onClose }) {
  if (!branch) return null;
  const statusColors = { active: "bg-green-100 text-green-700", moderate: "bg-yellow-100 text-yellow-700", inactive: "bg-red-100 text-red-700", unknown: "bg-slate-100 text-slate-600" };
  const sc = statusColors[branch.activationStatus] || statusColors.unknown;
  return (
    <div className="w-80 flex-shrink-0 bg-white border-l border-slate-200 overflow-y-auto flex flex-col">
      <div className="flex items-start justify-between p-4 border-b border-slate-100">
        <div>
          <div className="font-bold text-slate-800 text-base leading-tight">{branch.branchName}</div>
          <div className="text-xs text-slate-400 mt-0.5 font-mono">{branch.branchCode}</div>
        </div>
        <button onClick={onClose} className="p-1 hover:bg-slate-100 rounded-lg ml-2 flex-shrink-0">
          <X size={16} className="text-slate-500" />
        </button>
      </div>

      <div className="p-4 space-y-4 text-sm">
        <section>
          <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">
            <MapPin size={12} /> Location
          </div>
          <div className="text-slate-700">{branch.city}{branch.city && branch.state ? ", " : ""}{branch.state}</div>
          {branch.address && <div className="text-slate-500 text-xs mt-1 leading-relaxed">{branch.address}</div>}
          {branch.pincode && <div className="text-slate-400 text-xs mt-0.5">📮 {branch.pincode}</div>}
        </section>

        <section>
          <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">
            <Building2 size={12} /> Geography
          </div>
          <div className="space-y-1 text-xs text-slate-600">
            <div><span className="text-slate-400">Zone:</span> {branch.zone || "—"}</div>
            <div><span className="text-slate-400">Geo:</span> {branch.geo || "—"}</div>
            <div><span className="text-slate-400">Cluster:</span> {branch.cluster || "—"}</div>
          </div>
        </section>

        <section>
          <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">
            <Activity size={12} /> Activation Status
          </div>
          <div className="space-y-1.5">
            <div className="flex items-center justify-between text-xs">
              <span className="text-slate-500">Login</span>
              <span className={`px-2 py-0.5 rounded-full font-medium ${branch.loginActive.toLowerCase() === "active" ? "bg-green-100 text-green-700" : "bg-red-100 text-red-700"}`}>
                {branch.loginActive || "—"}
              </span>
            </div>
            <div className="flex items-center justify-between text-xs">
              <span className="text-slate-500">Disbursal</span>
              <span className={`px-2 py-0.5 rounded-full font-medium ${branch.disbursalActive.toLowerCase() === "active" ? "bg-green-100 text-green-700" : "bg-red-100 text-red-700"}`}>
                {branch.disbursalActive || "—"}
              </span>
            </div>
            <div className="flex items-center justify-between text-xs mt-1">
              <span className="text-slate-500">Overall</span>
              <span className={`px-2 py-0.5 rounded-full font-semibold capitalize ${sc}`}>
                {STATUS_LABEL[branch.activationStatus]}
              </span>
            </div>
          </div>
        </section>

        <section>
          <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">
            <Layers size={12} /> Products
          </div>
          <div className="flex gap-2 text-xs">
            <span className={`px-2 py-0.5 rounded-full ${branch.hl ? "bg-indigo-100 text-indigo-700" : "bg-slate-100 text-slate-400"}`}>HL</span>
            <span className={`px-2 py-0.5 rounded-full ${branch.msme ? "bg-purple-100 text-purple-700" : "bg-slate-100 text-slate-400"}`}>MSME</span>
            {branch.sahyog && <span className="px-2 py-0.5 rounded-full bg-teal-100 text-teal-700">Sahyog</span>}
          </div>
        </section>

        <section>
          <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">
            🔗 DSA
          </div>
          <div className="text-xs text-indigo-600 font-medium">Andromeda</div>
        </section>
      </div>
    </div>
  );
}

// ─── TOOLTIP ──────────────────────────────────────────────────────────────────
function PinTooltip({ tooltip }) {
  if (!tooltip) return null;
  return (
    <div
      className="absolute z-50 bg-slate-900 text-white text-xs rounded-lg shadow-xl p-2.5 pointer-events-none w-48"
      style={{ left: tooltip.x + 12, top: tooltip.y - 10 }}
    >
      <div className="font-semibold truncate">{tooltip.name}</div>
      <div className="text-slate-300 mt-0.5">{tooltip.city}{tooltip.city && tooltip.state ? ", " : ""}{tooltip.state}</div>
      <div className="mt-1">
        <span
          className="px-1.5 py-0.5 rounded-full text-xs font-medium"
          style={{ backgroundColor: STATUS_COLOR[tooltip.status] + "30", color: STATUS_COLOR[tooltip.status] }}
        >
          {STATUS_LABEL[tooltip.status]}
        </span>
      </div>
    </div>
  );
}

// ─── CLUSTER LIST PANEL ───────────────────────────────────────────────────────
function ClusterPanel({ branches, onSelect, onClose }) {
  if (!branches) return null;
  return (
    <div className="absolute z-40 right-4 top-16 bg-white rounded-xl shadow-2xl border border-slate-200 w-72 max-h-80 overflow-y-auto">
      <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100">
        <span className="font-semibold text-slate-700 text-sm">{branches.length} Branches</span>
        <button onClick={onClose} className="p-1 hover:bg-slate-100 rounded"><X size={14} /></button>
      </div>
      {branches.map(b => (
        <button key={b.id} onClick={() => { onSelect(b); onClose(); }}
          className="w-full text-left px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 border-b border-slate-50">
          <span className="w-2 h-2 rounded-full flex-shrink-0" style={{ backgroundColor: STATUS_COLOR[b.activationStatus] }} />
          <div className="min-w-0">
            <div className="text-xs font-medium text-slate-700 truncate">{b.branchName}</div>
            <div className="text-xs text-slate-400">{b.city}</div>
          </div>
        </button>
      ))}
    </div>
  );
}

// ─── CHARTS ───────────────────────────────────────────────────────────────────
function Charts({ filtered }) {
  const stateData = useMemo(() => {
    const map = {};
    filtered.forEach(b => {
      const s = b.state || "Unknown";
      if (!map[s]) map[s] = { state: s, total: 0, active: 0 };
      map[s].total++;
      if (b.activationStatus === "active") map[s].active++;
    });
    return Object.values(map)
      .map(d => ({ ...d, pct: d.total ? Math.round((d.active / d.total) * 100) : 0 }))
      .sort((a, b) => b.pct - a.pct).slice(0, 15);
  }, [filtered]);

  const zoneData = useMemo(() => {
    const map = {};
    filtered.forEach(b => {
      const z = b.zone || "Unknown";
      if (!map[z]) map[z] = { zone: z, login: 0, disbursal: 0 };
      if (b.loginActive?.toLowerCase() === "active") map[z].login++;
      if (b.disbursalActive?.toLowerCase() === "active") map[z].disbursal++;
    });
    return Object.values(map).sort((a, b) => b.login - a.login);
  }, [filtered]);

  const top10Active = useMemo(() =>
    filtered.filter(b => b.activationStatus === "active")
      .slice(0, 10)
      .map(b => ({ name: (b.branchName || "").slice(0, 22), count: 1 })),
    [filtered]);

  const bottom10Inactive = useMemo(() =>
    filtered.filter(b => b.activationStatus === "inactive")
      .slice(0, 10)
      .map(b => ({ name: (b.branchName || "").slice(0, 22), count: 1 })),
    [filtered]);

  const chartTitle = "text-sm font-semibold text-slate-700 mb-3";
  const cardCls = "bg-white rounded-xl shadow-sm border border-slate-100 p-4";

  return (
    <div className="grid grid-cols-2 gap-4 p-4">
      <div className={cardCls}>
        <div className={chartTitle}>State-wise Activation Rate (%)</div>
        <ResponsiveContainer width="100%" height={320}>
          <BarChart data={stateData} layout="vertical" margin={{ left: 10, right: 30 }}>
            <CartesianGrid strokeDasharray="3 3" horizontal={false} />
            <XAxis type="number" domain={[0, 100]} tick={{ fontSize: 10 }} />
            <YAxis type="category" dataKey="state" tick={{ fontSize: 10 }} width={100} />
            <Tooltip formatter={(v) => [`${v}%`, "Activation"]} />
            <Bar dataKey="pct" radius={[0, 4, 4, 0]}>
              {stateData.map((entry, i) => (
                <Cell key={i} fill={entry.pct >= 60 ? "#22c55e" : entry.pct >= 30 ? "#eab308" : "#ef4444"} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className={cardCls}>
        <div className={chartTitle}>Login vs Disbursal Active by Zone</div>
        <ResponsiveContainer width="100%" height={320}>
          <BarChart data={zoneData} margin={{ left: 5, right: 10 }}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="zone" tick={{ fontSize: 9 }} />
            <YAxis tick={{ fontSize: 10 }} />
            <Tooltip />
            <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
            <Bar dataKey="login" name="Login Active" fill="#6366f1" radius={[4, 4, 0, 0]} />
            <Bar dataKey="disbursal" name="Disbursal Active" fill="#22c55e" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className={cardCls}>
        <div className={chartTitle}>Top 10 Active Branches</div>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={top10Active} layout="vertical" margin={{ left: 10, right: 20 }}>
            <CartesianGrid strokeDasharray="3 3" horizontal={false} />
            <XAxis type="number" tick={{ fontSize: 10 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 9 }} width={130} />
            <Tooltip />
            <Bar dataKey="count" fill="#22c55e" radius={[0, 4, 4, 0]} name="Active" />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className={cardCls}>
        <div className={chartTitle}>Bottom 10 Inactive Branches</div>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={bottom10Inactive} layout="vertical" margin={{ left: 10, right: 20 }}>
            <CartesianGrid strokeDasharray="3 3" horizontal={false} />
            <XAxis type="number" tick={{ fontSize: 10 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 9 }} width={130} />
            <Tooltip formatter={() => ["Needs Activation"]} />
            <Bar dataKey="count" fill="#ef4444" radius={[0, 4, 4, 0]} name="Inactive" />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

// ─── MAIN COMPONENT ───────────────────────────────────────────────────────────
export default function KnowYourDSA() {
  const [branches, setBranches] = useState(null);
  const [dragging, setDragging] = useState(false);
  const [transform, setTransform] = useState({ scale: 1, x: 0, y: 0 });
  const [isPanning, setIsPanning] = useState(false);
  const [panStart, setPanStart] = useState({ x: 0, y: 0, tx: 0, ty: 0 });
  const [hoveredState, setHoveredState] = useState(null);
  const [selectedBranch, setSelectedBranch] = useState(null);
  const [tooltip, setTooltip] = useState(null);
  const [clusterBranches, setClusterBranches] = useState(null);

  // Filters
  const [filterState, setFilterState] = useState("");
  const [filterCity, setFilterCity] = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [filterZone, setFilterZone] = useState("");

  const svgRef = useRef(null);

  // ── Filter options ──────────────────────────────────────────────────────────
  const states = useMemo(() => branches ? [...new Set(branches.map(b => b.state).filter(Boolean))].sort() : [], [branches]);
  const cities = useMemo(() => {
    if (!branches) return [];
    const src = filterState ? branches.filter(b => b.state === filterState) : branches;
    return [...new Set(src.map(b => b.city).filter(Boolean))].sort();
  }, [branches, filterState]);
  const zones = useMemo(() => branches ? [...new Set(branches.map(b => b.zone).filter(Boolean))].sort() : [], [branches]);

  // ── Filtered branches ───────────────────────────────────────────────────────
  const filtered = useMemo(() => {
    if (!branches) return [];
    return branches.filter(b => {
      if (filterState && b.state !== filterState) return false;
      if (filterCity && b.city !== filterCity) return false;
      if (filterStatus && b.activationStatus !== filterStatus) return false;
      if (filterZone && b.zone !== filterZone) return false;
      return true;
    });
  }, [branches, filterState, filterCity, filterStatus, filterZone]);

  const mappedBranches = useMemo(() => filtered.filter(b => b.coords), [filtered]);

  // ── KPI ─────────────────────────────────────────────────────────────────────
  const kpis = useMemo(() => {
    const total = filtered.length;
    const active = filtered.filter(b => b.activationStatus === "active").length;
    const inactive = filtered.filter(b => b.activationStatus === "inactive").length;
    const moderate = filtered.filter(b => b.activationStatus === "moderate").length;
    const loginActive = filtered.filter(b => b.loginActive?.toLowerCase() === "active").length;
    const disbursalActive = filtered.filter(b => b.disbursalActive?.toLowerCase() === "active").length;
    return {
      total, active, inactive, moderate,
      loginPct: total ? Math.round((loginActive / total) * 100) : 0,
      disbursalPct: total ? Math.round((disbursalActive / total) * 100) : 0,
    };
  }, [filtered]);

  // ── Legend counts ────────────────────────────────────────────────────────────
  const legendCounts = useMemo(() => ({
    active: filtered.filter(b => b.activationStatus === "active").length,
    moderate: filtered.filter(b => b.activationStatus === "moderate").length,
    inactive: filtered.filter(b => b.activationStatus === "inactive").length,
    unknown: filtered.filter(b => b.activationStatus === "unknown").length,
  }), [filtered]);

  // ── Clusters vs individual pins ──────────────────────────────────────────────
  const pinGroups = useMemo(() => {
    if (transform.scale >= 2.5) return null;
    const map = {};
    mappedBranches.forEach(b => {
      const key = b.city.toLowerCase().trim() || `${b.coords.lat.toFixed(2)},${b.coords.lng.toFixed(2)}`;
      if (!map[key]) map[key] = [];
      map[key].push(b);
    });
    return map;
  }, [mappedBranches, transform.scale]);

  // ── File handlers ────────────────────────────────────────────────────────────
  const handleFile = useCallback((file) => {
    if (!file) return;
    parseExcel(file, (data) => {
      setBranches(data);
      setSelectedBranch(null);
      setTransform({ scale: 1, x: 0, y: 0 });
    });
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  // ── Zoom / Pan ────────────────────────────────────────────────────────────────
  const handleWheel = useCallback((e) => {
    e.preventDefault();
    const rect = svgRef.current.getBoundingClientRect();
    const cursorX = e.clientX - rect.left;
    const cursorY = e.clientY - rect.top;
    setTransform(prev => {
      const newScale = Math.min(8, Math.max(1, prev.scale * (e.deltaY < 0 ? 1.15 : 0.87)));
      const ratio = newScale / prev.scale;
      return { scale: newScale, x: cursorX - ratio * (cursorX - prev.x), y: cursorY - ratio * (cursorY - prev.y) };
    });
  }, []);

  const handleMouseDown = useCallback((e) => {
    if (e.button !== 0) return;
    setIsPanning(true);
    setPanStart({ x: e.clientX, y: e.clientY, tx: transform.x, ty: transform.y });
  }, [transform]);

  const handleMouseMove = useCallback((e) => {
    if (!isPanning) return;
    setTransform(prev => ({
      ...prev,
      x: panStart.tx + (e.clientX - panStart.x),
      y: panStart.ty + (e.clientY - panStart.y),
    }));
  }, [isPanning, panStart]);

  const handleMouseUp = useCallback(() => setIsPanning(false), []);

  const zoomIn = () => setTransform(prev => ({ ...prev, scale: Math.min(8, prev.scale * 1.3) }));
  const zoomOut = () => setTransform(prev => ({ ...prev, scale: Math.max(1, prev.scale / 1.3) }));
  const resetZoom = () => setTransform({ scale: 1, x: 0, y: 0 });

  const zoomToPoint = (cx, cy, targetScale = 3) => {
    const rect = svgRef.current?.getBoundingClientRect();
    if (!rect) return;
    const vw = rect.width / 2, vh = rect.height / 2;
    setTransform({
      scale: targetScale,
      x: vw - cx * targetScale,
      y: vh - cy * targetScale,
    });
  };

  // ── Upload screen ─────────────────────────────────────────────────────────────
  if (!branches) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-slate-900 to-indigo-950 flex flex-col items-center justify-center">
        <div className="text-center mb-8">
          <div className="text-4xl font-black text-white tracking-tight mb-1">Know Your DSA</div>
          <div className="text-indigo-300 text-sm">Upload Piramal Branch Activation File to get started</div>
        </div>
        <div
          onDragOver={e => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
          onClick={() => document.getElementById("file-input").click()}
          className={`relative cursor-pointer rounded-2xl border-2 border-dashed transition-all duration-200 p-12 w-96 flex flex-col items-center gap-4
            ${dragging ? "border-indigo-400 bg-indigo-900/30 scale-[1.02]" : "border-slate-600 bg-slate-800/40 hover:border-indigo-500 hover:bg-slate-800/60"}`}
        >
          <div className={`rounded-full p-4 ${dragging ? "bg-indigo-500" : "bg-slate-700"}`}>
            <Upload size={28} className={dragging ? "text-white" : "text-slate-300"} />
          </div>
          <div className="text-center">
            <div className="text-white font-semibold">Drop your Excel file here</div>
            <div className="text-slate-400 text-sm mt-1">or click to browse</div>
            <div className="text-slate-500 text-xs mt-2">Branch Mapping &amp; Activation_piramal.xlsx</div>
          </div>
          <input
            id="file-input"
            type="file"
            accept=".xlsx"
            className="hidden"
            onChange={e => handleFile(e.target.files[0])}
          />
        </div>
      </div>
    );
  }

  // ── Dashboard ──────────────────────────────────────────────────────────────
  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      {/* Header */}
      <header className="bg-slate-900 text-white px-5 py-3 flex items-center justify-between shadow-lg flex-shrink-0">
        <div>
          <span className="text-lg font-black tracking-tight">Know Your DSA</span>
          <span className="ml-3 text-xs text-indigo-300 font-medium bg-indigo-900/50 px-2 py-0.5 rounded-full">Andromeda</span>
        </div>
        <button
          onClick={() => { setBranches(null); setFilterState(""); setFilterCity(""); setFilterStatus(""); setFilterZone(""); }}
          className="flex items-center gap-1.5 text-xs text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 px-3 py-1.5 rounded-lg transition"
        >
          <Upload size={13} /> Upload New File
        </button>
      </header>

      {/* KPI Cards */}
      <div className="px-4 pt-4 pb-2 grid grid-cols-6 gap-3 flex-shrink-0">
        <KpiCard label="Total Branches" value={kpis.total} color="#6366f1" icon={Building2} />
        <KpiCard label="Active Branches" value={kpis.active} color="#22c55e" icon={Activity} />
        <KpiCard label="Inactive Branches" value={kpis.inactive} color="#ef4444" icon={TrendingDown} />
        <KpiCard label="Moderate Branches" value={kpis.moderate} color="#eab308" icon={TrendingUp} />
        <KpiCard label="Login Activation" value={`${kpis.loginPct}%`} color="#8b5cf6" icon={Activity} sub={`of ${kpis.total} branches`} />
        <KpiCard label="Disbursal Activation" value={`${kpis.disbursalPct}%`} color="#0ea5e9" icon={Activity} sub={`of ${kpis.total} branches`} />
      </div>

      {/* Filter Bar */}
      <div className="px-4 py-2 bg-white border-b border-slate-100 flex items-center gap-3 flex-shrink-0 shadow-sm">
        <Filter size={14} className="text-slate-400 flex-shrink-0" />
        <Select value={filterState} onChange={v => { setFilterState(v); setFilterCity(""); }} options={states} label="All States" />
        <Select value={filterCity} onChange={setFilterCity} options={cities} label="All Cities" />
        <Select value={filterStatus} onChange={setFilterStatus}
          options={["active", "moderate", "inactive", "unknown"]}
          label="All Statuses" />
        <Select value={filterZone} onChange={setFilterZone} options={zones} label="All Zones" />
        {(filterState || filterCity || filterStatus || filterZone) && (
          <button
            onClick={() => { setFilterState(""); setFilterCity(""); setFilterStatus(""); setFilterZone(""); }}
            className="flex items-center gap-1 text-xs text-red-500 hover:text-red-700 ml-1"
          >
            <RefreshCw size={12} /> Reset
          </button>
        )}
        <span className="ml-auto text-xs text-slate-400">{filtered.length} branches shown</span>
      </div>

      {/* Map + Side Panel */}
      <div className="flex flex-1 relative" style={{ minHeight: 0 }}>
        <div className={`relative overflow-hidden bg-sky-50 flex-1`} style={{ height: 600 }}>
          <svg
            ref={svgRef}
            viewBox={`0 0 ${SVG_W} ${SVG_H}`}
            className="w-full h-full"
            style={{ cursor: isPanning ? "grabbing" : "grab", userSelect: "none" }}
            onWheel={handleWheel}
            onMouseDown={handleMouseDown}
            onMouseMove={handleMouseMove}
            onMouseUp={handleMouseUp}
            onMouseLeave={handleMouseUp}
          >
            {/* Ocean bg */}
            <rect width={SVG_W} height={SVG_H} fill="#EFF6FF" />

            <g transform={`translate(${transform.x},${transform.y}) scale(${transform.scale})`}>
              {/* States */}
              {INDIA_STATES.map(s => (
                <path
                  key={s.id}
                  id={s.id}
                  data-state={s.name}
                  d={s.d}
                  fill={hoveredState === s.id ? "#94a3b8" : "#CBD5E1"}
                  stroke="#FFFFFF"
                  strokeWidth={0.5}
                  onMouseEnter={() => setHoveredState(s.id)}
                  onMouseLeave={() => setHoveredState(null)}
                  onDoubleClick={(e) => {
                    e.stopPropagation();
                    // Compute bounding box centroid of path (approx)
                    const bbox = e.target.getBBox();
                    zoomToPoint(bbox.x + bbox.width / 2, bbox.y + bbox.height / 2, 3);
                  }}
                  style={{ cursor: "pointer" }}
                />
              ))}

              {/* Clusters (scale < 2.5) */}
              {pinGroups && Object.entries(pinGroups).map(([key, grp]) => {
                const coords = grp[0].coords;
                const { x, y } = project(coords.lat, coords.lng);
                const statusCounts = { active: 0, moderate: 0, inactive: 0, unknown: 0 };
                grp.forEach(b => statusCounts[b.activationStatus]++);
                const dominant = Object.entries(statusCounts).sort((a, b) => b[1] - a[1])[0][0];
                const color = STATUS_COLOR[dominant];
                const count = grp.length;
                return (
                  <g
                    key={key}
                    transform={`translate(${x},${y}) scale(${1 / transform.scale})`}
                    style={{ cursor: "pointer" }}
                    onMouseEnter={(e) => {
                      const rect = svgRef.current.getBoundingClientRect();
                      setTooltip({
                        x: e.clientX - rect.left,
                        y: e.clientY - rect.top,
                        name: count === 1 ? grp[0].branchName : `${count} branches`,
                        city: grp[0].city,
                        state: grp[0].state,
                        status: dominant,
                      });
                    }}
                    onMouseLeave={() => setTooltip(null)}
                    onClick={(e) => {
                      e.stopPropagation();
                      if (count === 1) {
                        setSelectedBranch(grp[0]);
                      } else {
                        zoomToPoint(x, y, 3);
                      }
                    }}
                  >
                    <circle r={count > 1 ? 16 : 10} fill={color} stroke="white" strokeWidth={2} opacity={0.9} />
                    {count > 1 && (
                      <text textAnchor="middle" dy="4" fontSize="11" fontWeight="bold" fill="white">{count}</text>
                    )}
                  </g>
                );
              })}

              {/* Individual pins (scale >= 2.5) */}
              {transform.scale >= 2.5 && mappedBranches.map(b => {
                const { x, y } = project(b.coords.lat, b.coords.lng);
                const color = STATUS_COLOR[b.activationStatus];
                return (
                  <g
                    key={b.id}
                    transform={`translate(${x},${y}) scale(${1 / transform.scale})`}
                    style={{ cursor: "pointer" }}
                    onMouseEnter={(e) => {
                      const rect = svgRef.current.getBoundingClientRect();
                      setTooltip({
                        x: e.clientX - rect.left,
                        y: e.clientY - rect.top,
                        name: b.branchName,
                        city: b.city,
                        state: b.state,
                        status: b.activationStatus,
                      });
                    }}
                    onMouseLeave={() => setTooltip(null)}
                    onClick={(e) => { e.stopPropagation(); setSelectedBranch(b); }}
                  >
                    <path
                      d="M0,-14 C5,-14 9,-10 9,-6 C9,0 0,10 0,10 C0,10 -9,0 -9,-6 C-9,-10 -5,-14 0,-14 Z"
                      fill={color} stroke="white" strokeWidth="1.5"
                    />
                    <circle cx="0" cy="-6" r="3" fill="white" />
                  </g>
                );
              })}
            </g>
          </svg>

          {/* Tooltip */}
          <PinTooltip tooltip={tooltip} />

          {/* Legend */}
          <div className="absolute bottom-4 left-4 bg-white/90 backdrop-blur-sm rounded-xl shadow-md px-3 py-2.5 text-xs space-y-1.5 border border-slate-100">
            {[["active", "Active"], ["moderate", "Moderate"], ["inactive", "Inactive"], ["unknown", "Unknown"]].map(([s, l]) => (
              <div key={s} className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full flex-shrink-0" style={{ backgroundColor: STATUS_COLOR[s] }} />
                <span className="text-slate-600">{l}</span>
                <span className="ml-auto font-semibold text-slate-700 pl-2">{legendCounts[s]}</span>
              </div>
            ))}
          </div>

          {/* Map Controls */}
          <div className="absolute bottom-4 right-4 flex flex-col gap-1">
            <button onClick={zoomIn} className="w-8 h-8 bg-white shadow-md rounded-lg flex items-center justify-center hover:bg-slate-50 border border-slate-200">
              <Plus size={14} className="text-slate-700" />
            </button>
            <button onClick={zoomOut} className="w-8 h-8 bg-white shadow-md rounded-lg flex items-center justify-center hover:bg-slate-50 border border-slate-200">
              <Minus size={14} className="text-slate-700" />
            </button>
            <button onClick={resetZoom} className="w-8 h-8 bg-white shadow-md rounded-lg flex items-center justify-center hover:bg-slate-50 border border-slate-200">
              <Home size={13} className="text-slate-700" />
            </button>
            <div className="w-8 h-6 bg-slate-800 rounded text-white text-center leading-6 font-mono" style={{ fontSize: 9 }}>
              {transform.scale.toFixed(1)}x
            </div>
          </div>

          {/* Hovered state label */}
          {hoveredState && (
            <div className="absolute top-3 left-1/2 -translate-x-1/2 bg-slate-800/80 text-white text-xs px-3 py-1 rounded-full pointer-events-none">
              {INDIA_STATES.find(s => s.id === hoveredState)?.name}
            </div>
          )}

          {/* Cluster list panel */}
          <ClusterPanel
            branches={clusterBranches}
            onSelect={setSelectedBranch}
            onClose={() => setClusterBranches(null)}
          />
        </div>

        {/* Side Panel */}
        {selectedBranch && (
          <SidePanel branch={selectedBranch} onClose={() => setSelectedBranch(null)} />
        )}
      </div>

      {/* Charts */}
      <Charts filtered={filtered} />
    </div>
  );
}
