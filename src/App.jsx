import { useState, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────────────────────
// SEED DATA
// ─────────────────────────────────────────────────────────────
const FIXED_COSTS = [
  { name: "ค่าเช่า", amount: 4500 },
  { name: "ค่าพนักงาน", amount: 35000 },
  { name: "ค่าไฟ", amount: 8000 },
  { name: "เครื่องล้างจาน", amount: 4000 },
  { name: "ค่าน้ำ", amount: 1000 },
  { name: "อื่นๆ", amount: 1000 },
];
const TOTAL_FIXED = FIXED_COSTS.reduce((a, b) => a + b.amount, 0);

const SUPPLIERS = [
  { id: 1, name: "ตลาดสดนครชัย", type: "ผัก" },
  { id: 2, name: "ฟาร์มหมูสยาม", type: "หมู/เนื้อ" },
  { id: 3, name: "อาหารทะเลสด", type: "ทะเล" },
  { id: 4, name: "ซอสมาล่าพรีเมียม", type: "ซอส" },
];

// costHistory: [{date, unitCost, qty, total}] — ใช้คำนวณ weighted avg
const INIT_STOCK = [
  { id: 1, name: "หมูสามชั้น", unit: "kg", qty: 8, minQty: 5, dailyUse: 4, supplierId: 2, expiryDays: 3,
    costHistory: [
      { date: "2026-04-25", unitCost: 175, qty: 20, total: 3500 },
      { date: "2026-05-02", unitCost: 178, qty: 20, total: 3560 },
      { date: "2026-05-08", unitCost: 180, qty: 20, total: 3600 },
    ]},
  { id: 2, name: "กุ้งแวนนาไม", unit: "kg", qty: 3, minQty: 4, dailyUse: 2, supplierId: 3, expiryDays: 2,
    costHistory: [
      { date: "2026-04-28", unitCost: 275, qty: 10, total: 2750 },
      { date: "2026-05-05", unitCost: 280, qty: 10, total: 2800 },
    ]},
  { id: 3, name: "หอยแมลงภู่", unit: "kg", qty: 6, minQty: 3, dailyUse: 2, supplierId: 3, expiryDays: 2,
    costHistory: [
      { date: "2026-04-28", unitCost: 118, qty: 15, total: 1770 },
      { date: "2026-05-05", unitCost: 120, qty: 15, total: 1800 },
    ]},
  { id: 4, name: "ผักกาดขาว", unit: "kg", qty: 2, minQty: 3, dailyUse: 3, supplierId: 1, expiryDays: 2,
    costHistory: [
      { date: "2026-05-07", unitCost: 25, qty: 10, total: 250 },
      { date: "2026-05-08", unitCost: 26, qty: 10, total: 260 },
      { date: "2026-05-09", unitCost: 32, qty: 8, total: 256 }, // spike!
    ]},
  { id: 5, name: "เห็ดหอม", unit: "kg", qty: 4, minQty: 2, dailyUse: 1.5, supplierId: 1, expiryDays: 3,
    costHistory: [
      { date: "2026-05-06", unitCost: 90, qty: 5, total: 450 },
      { date: "2026-05-08", unitCost: 92, qty: 5, total: 460 },
    ]},
  { id: 6, name: "เต้าหู้ขาว", unit: "kg", qty: 5, minQty: 3, dailyUse: 2, supplierId: 1, expiryDays: 2,
    costHistory: [
      { date: "2026-05-07", unitCost: 35, qty: 10, total: 350 },
      { date: "2026-05-09", unitCost: 35, qty: 10, total: 350 },
    ]},
  { id: 7, name: "ซอสเบสหม่าล่า", unit: "ถุง", qty: 12, minQty: 6, dailyUse: 2, supplierId: 4, expiryDays: 30,
    costHistory: [
      { date: "2026-04-20", unitCost: 215, qty: 20, total: 4300 },
      { date: "2026-05-06", unitCost: 220, qty: 20, total: 4400 },
    ]},
  { id: 8, name: "น้ำซุปกระดูก", unit: "ลิตร", qty: 20, minQty: 10, dailyUse: 5, supplierId: 4, expiryDays: 5,
    costHistory: [
      { date: "2026-05-07", unitCost: 45, qty: 30, total: 1350 },
    ]},
  { id: 9, name: "กระเทียมสับ", unit: "kg", qty: 3, minQty: 2, dailyUse: 0.8, supplierId: 1, expiryDays: 4,
    costHistory: [
      { date: "2026-05-08", unitCost: 80, qty: 5, total: 400 },
    ]},
  { id: 10, name: "บรรจุภัณฑ์", unit: "ชิ้น", qty: 150, minQty: 50, dailyUse: 30, supplierId: 4, expiryDays: 365,
    costHistory: [
      { date: "2026-05-01", unitCost: 4, qty: 200, total: 800 },
    ]},
];

const INIT_CF = [
  { id: 1, date: "2026-05-09", flow: "in", cat: "ยอดขาย dine-in", amount: 9200, method: "เงินสด", note: "58 ออร์เดอร์" },
  { id: 2, date: "2026-05-09", flow: "in", cat: "ยอดขาย delivery", amount: 3800, method: "โอน", note: "GrabFood" },
  { id: 3, date: "2026-05-09", flow: "out", cat: "วัตถุดิบ/ผัก", amount: 1400, method: "เงินสด", note: "ตลาดนครชัย" },
  { id: 4, date: "2026-05-08", flow: "in", cat: "ยอดขาย dine-in", amount: 8500, method: "เงินสด", note: "52 ออร์เดอร์" },
  { id: 5, date: "2026-05-08", flow: "in", cat: "ยอดขาย delivery", amount: 3200, method: "โอน", note: "" },
  { id: 6, date: "2026-05-08", flow: "out", cat: "วัตถุดิบ/ผัก", amount: 1200, method: "เงินสด", note: "" },
  { id: 7, date: "2026-05-08", flow: "out", cat: "ค่าแรงพนักงาน", amount: 1200, method: "โอน", note: "" },
  { id: 8, date: "2026-05-07", flow: "in", cat: "ยอดขาย dine-in", amount: 7800, method: "เงินสด", note: "48 ออร์เดอร์" },
  { id: 9, date: "2026-05-07", flow: "in", cat: "ยอดขาย delivery", amount: 2900, method: "โอน", note: "" },
  { id: 10, date: "2026-05-07", flow: "out", cat: "วัตถุดิบ/หมู-ทะเล", amount: 4200, method: "โอน", note: "สัปดาห์นี้" },
  { id: 11, date: "2026-05-06", flow: "in", cat: "ยอดขาย dine-in", amount: 9200, method: "เงินสด", note: "60 ออร์เดอร์ วันหยุด" },
  { id: 12, date: "2026-05-06", flow: "out", cat: "ค่าไฟ/น้ำ", amount: 1200, method: "โอน", note: "" },
  { id: 13, date: "2026-05-06", flow: "out", cat: "ซอส/เครื่องปรุง", amount: 1760, method: "โอน", note: "ซอสหม่าล่า 8 ถุง" },
  { id: 14, date: "2026-05-05", flow: "in", cat: "ยอดขาย dine-in", amount: 8800, method: "เงินสด", note: "" },
  { id: 15, date: "2026-05-05", flow: "out", cat: "วัตถุดิบ/ผัก", amount: 1100, method: "เงินสด", note: "" },
  { id: 16, date: "2026-05-05", flow: "out", cat: "ค่าแรงพนักงาน", amount: 1200, method: "โอน", note: "" },
];

const IN_CATS = ["ยอดขาย dine-in", "ยอดขาย delivery", "รับชำระหนี้", "เงินทุนเพิ่ม", "รายได้อื่นๆ"];
const OUT_CATS = [
  // วัตถุดิบ — แยกหมวด
  "วัตถุดิบ/ผัก",
  "วัตถุดิบ/เนื้อสัตว์",
  "วัตถุดิบ/ลูกชิ้น",
  "วัตถุดิบ/เครื่องดื่ม",
  "วัตถุดิบ/ซอส",
  "วัตถุดิบ/บรรจุภัณฑ์",
  // ค่าใช้จ่ายดำเนินการ
  "ค่าแรงพนักงาน","ค่าเช่า","ค่าไฟ/น้ำ","ค่าแก๊ส","ค่าการตลาด","ค่าซ่อมบำรุง","จ่ายชำระหนี้","อื่นๆ"
];
const PAY_METHODS = ["เงินสด","โอนธนาคาร","QR Code","บัตรเครดิต","GrabFood","LINE MAN"];

// หมวดวัตถุดิบและสีประจำหมวด
const INGREDIENT_GROUPS = [
  { key: "วัตถุดิบ/ผัก",          label: "🥬 ผัก",          color: "#4ade80", bg: "rgba(74,222,128,0.10)" },
  { key: "วัตถุดิบ/เนื้อสัตว์",   label: "🥩 เนื้อสัตว์",   color: "#f87171", bg: "rgba(248,113,113,0.10)" },
  { key: "วัตถุดิบ/ลูกชิ้น",      label: "🍢 ลูกชิ้น",      color: "#fb923c", bg: "rgba(251,146,60,0.10)"  },
  { key: "วัตถุดิบ/เครื่องดื่ม",  label: "🧋 เครื่องดื่ม",  color: "#38bdf8", bg: "rgba(56,189,248,0.10)"  },
  { key: "วัตถุดิบ/ซอส",           label: "🌶️ ซอส",          color: "#f43f5e", bg: "rgba(244,63,94,0.10)"   },
  { key: "วัตถุดิบ/บรรจุภัณฑ์",   label: "📦 บรรจุภัณฑ์",   color: "#a78bfa", bg: "rgba(167,139,250,0.10)" },
];

// ค่าใช้จ่ายดำเนินการ
const OPEX_CATS = ["ค่าแรงพนักงาน","ค่าเช่า","ค่าไฟ/น้ำ","ค่าแก๊ส","ค่าการตลาด","ค่าซ่อมบำรุง","จ่ายชำระหนี้","อื่นๆ"];

// ─────────────────────────────────────────────────────────────
// COST UTILS
// ─────────────────────────────────────────────────────────────
function weightedAvgCost(item) {
  const h = item.costHistory || [];
  if (!h.length) return 0;
  const totalQty = h.reduce((a, b) => a + b.qty, 0);
  const totalCost = h.reduce((a, b) => a + b.total, 0);
  return totalQty > 0 ? totalCost / totalQty : 0;
}

function latestCost(item) {
  const h = item.costHistory || [];
  if (!h.length) return 0;
  return [...h].sort((a, b) => b.date.localeCompare(a.date))[0].unitCost;
}

// spike = ราคาล่าสุด > ค่าเฉลี่ย * 1.15 (15%)
function isCostSpike(item) {
  const avg = weightedAvgCost(item);
  const latest = latestCost(item);
  return avg > 0 && latest > avg * 1.15;
}

function spikePercent(item) {
  const avg = weightedAvgCost(item);
  const latest = latestCost(item);
  if (avg <= 0) return 0;
  return ((latest - avg) / avg * 100).toFixed(1);
}

// ─────────────────────────────────────────────────────────────
// STOCK STATUS
// ─────────────────────────────────────────────────────────────
function stockStatus(item) {
  const daysLeft = item.dailyUse > 0 ? item.qty / item.dailyUse : 99;
  if (item.qty <= 0) return "out";
  if (item.qty <= item.minQty || daysLeft <= 1) return "critical";
  if (item.qty <= item.minQty * 1.5 || daysLeft <= 2) return "low";
  return "ok";
}
const SC = { ok: "#4ade80", low: "#fbbf24", critical: "#f87171", out: "#ef4444" };
const SL = { ok: "ปกติ", low: "ใกล้หมด", critical: "น้อยมาก", out: "หมด" };

// ─────────────────────────────────────────────────────────────
// SHARED UI
// ─────────────────────────────────────────────────────────────
const fmt = (n) => Number(n || 0).toLocaleString("th-TH");
const fmtD = (n) => Number(n || 0).toFixed(2);
const todayStr = () => new Date().toISOString().split("T")[0];

const inp = {
  background: "rgba(255,255,255,0.06)", border: "1px solid rgba(249,115,22,0.2)",
  borderRadius: 8, padding: "8px 11px", color: "#fff", fontSize: 13,
  width: "100%", boxSizing: "border-box", outline: "none", fontFamily: "inherit",
};
const btn = (c = "#f97316") => ({
  background: c === "#f97316" ? "linear-gradient(135deg,#f97316,#ea580c)" : `linear-gradient(135deg,${c},${c}cc)`,
  border: "none", borderRadius: 8, padding: "8px 16px",
  color: c === "#f97316" ? "#111111" : "#fff",
  fontWeight: 700, cursor: "pointer", fontSize: 13, fontFamily: "inherit",
});
const ghost = {
  background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)",
  borderRadius: 8, padding: "8px 14px", color: "#888", cursor: "pointer",
  fontSize: 13, fontFamily: "inherit",
};

function Card({ children, style = {} }) {
  return <div style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.09)", borderRadius: 16, padding: "16px 18px", ...style }}>{children}</div>;
}

function Tag({ status }) {
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 4, background: SC[status] + "18", border: `1px solid ${SC[status]}44`, borderRadius: 6, padding: "2px 7px", fontSize: 11, color: SC[status], fontWeight: 700 }}>
      <span style={{ width: 5, height: 5, borderRadius: "50%", background: SC[status], boxShadow: `0 0 4px ${SC[status]}` }} />
      {SL[status]}
    </span>
  );
}

function Hdr({ title, action }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
      <h2 style={{ margin: 0, color: "#f97316", fontSize: 17, fontWeight: 900 }}>{title}</h2>
      {action}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// EXCEL HELPERS
// ─────────────────────────────────────────────────────────────
function exportStockToXLSX(stock) {
  const rows = stock.map(s => ({
    "ID": s.id,
    "ชื่อวัตถุดิบ": s.name,
    "หน่วย": s.unit,
    "จำนวนคงเหลือ": s.qty,
    "จำนวนขั้นต่ำ": s.minQty,
    "ใช้ต่อวัน": s.dailyUse,
    "ต้นทุนเฉลี่ย (฿/หน่วย)": +weightedAvgCost(s).toFixed(2),
    "ราคาล่าสุด (฿/หน่วย)": latestCost(s),
    "มูลค่าสต็อค (฿)": +(s.qty * weightedAvgCost(s)).toFixed(2),
    "สถานะ": SL[stockStatus(s)],
    "แจ้งเตือนราคา": isCostSpike(s) ? `⚠️ ราคาสูงขึ้น ${spikePercent(s)}%` : "-",
    "อายุ (วัน)": s.expiryDays,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [8,20,8,14,14,12,22,22,20,12,22,10].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "สต็อค");

  // Cost history sheet
  const hist = [];
  stock.forEach(s => {
    (s.costHistory || []).forEach(h => {
      hist.push({ "วัตถุดิบ": s.name, "วันที่": h.date, "ราคา/หน่วย": h.unitCost, "จำนวน": h.qty, "รวม": h.total });
    });
  });
  const ws2 = XLSX.utils.json_to_sheet(hist);
  ws2["!cols"] = [20,12,14,10,12].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws2, "ประวัติราคา");

  XLSX.writeFile(wb, `stock_taiguo_${todayStr()}.xlsx`);
}

function exportCFToXLSX(cf) {
  const rows = cf.map(e => ({
    "วันที่": e.date,
    "ประเภท": e.flow === "in" ? "รายรับ" : "รายจ่าย",
    "หมวดหมู่": e.cat,
    "จำนวนเงิน (฿)": e.amount,
    "ช่องทาง": e.method,
    "หมายเหตุ": e.note,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [12,10,24,16,14,20].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Cash Flow");
  XLSX.writeFile(wb, `cashflow_taiguo_${todayStr()}.xlsx`);
}

function exportFullReport(stock, cf) {
  const wb = XLSX.utils.book_new();

  // Summary sheet
  const monthKey = todayStr().slice(0, 7);
  const mCF = cf.filter(e => e.date.startsWith(monthKey));
  const mIn = mCF.filter(e => e.flow === "in").reduce((a, b) => a + b.amount, 0);
  const mOut = mCF.filter(e => e.flow === "out").reduce((a, b) => a + b.amount, 0);
  const summary = [
    ["ไท่กั๋วหม่าล่า — รายงานสรุป", "", todayStr()],
    [],
    ["P&L เดือน " + monthKey],
    ["รายรับรวม", mIn],
    ["รายจ่ายผันแปร", mOut],
    ["ต้นทุนคงที่", TOTAL_FIXED],
    ["กำไรสุทธิโดยประมาณ", mIn - mOut - TOTAL_FIXED],
    ["Gross Margin (%)", mIn > 0 ? +((mIn - mOut) / mIn * 100).toFixed(1) : 0],
    [],
    ["ต้นทุนคงที่รายการ"],
    ...FIXED_COSTS.map(f => [f.name, f.amount]),
    ["รวม", TOTAL_FIXED],
  ];
  const ws0 = XLSX.utils.aoa_to_sheet(summary);
  ws0["!cols"] = [30, 16].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws0, "สรุป");

  // Stock
  const stockRows = stock.map(s => ({
    "ชื่อ": s.name, "หน่วย": s.unit, "คงเหลือ": s.qty,
    "ขั้นต่ำ": s.minQty, "ใช้/วัน": s.dailyUse,
    "ต้นทุนเฉลี่ย": +weightedAvgCost(s).toFixed(2),
    "ราคาล่าสุด": latestCost(s),
    "มูลค่า": +(s.qty * weightedAvgCost(s)).toFixed(2),
    "สถานะ": SL[stockStatus(s)],
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stockRows), "สต็อค");

  // CF
  const cfRows = cf.map(e => ({ "วันที่": e.date, "ประเภท": e.flow === "in" ? "รายรับ" : "รายจ่าย", "หมวด": e.cat, "จำนวน": e.amount, "ช่องทาง": e.method, "หมายเหตุ": e.note }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cfRows), "Cash Flow");

  XLSX.writeFile(wb, `full_report_taiguo_${todayStr()}.xlsx`);
}

// ─────────────────────────────────────────────────────────────
// LOGIN
// ─────────────────────────────────────────────────────────────
function LoginPage({ onLogin }) {
  const [role, setRole] = useState("owner");
  const [pin, setPin] = useState("");
  const [err, setErr] = useState(false);
  const PINS = { owner: "1234", staff: "0000" };
  const go = () => { if (pin === PINS[role]) onLogin(role); else { setErr(true); setPin(""); } };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(160deg,#1a1a1a 0%,#111111 100%)", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{ textAlign: "center", marginBottom: 28 }}>
        <div style={{ fontSize: 52 }}>🫕</div>
        <div style={{ color: "#f97316", fontSize: 22, fontWeight: 900 }}>ไท่กั๋วหม่าล่า</div>
        <div style={{ color: "#444", fontSize: 11, marginTop: 2 }}>ระบบจัดการร้าน Pro</div>
      </div>
      <Card style={{ width: "100%", maxWidth: 340 }}>
        <div style={{ color: "#666", fontSize: 12, marginBottom: 8 }}>บทบาท</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          {[["owner","👑 เจ้าของ"],["staff","👷 พนักงาน"]].map(([v,l]) => (
            <button key={v} onClick={() => { setRole(v); setPin(""); setErr(false); }} style={{ flex:1, padding:10, background: role===v ? "rgba(249,115,22,0.15)" : "rgba(255,255,255,0.04)", border:`1.5px solid ${role===v ? "#f97316" : "rgba(255,255,255,0.1)"}`, borderRadius:10, color: role===v ? "#f97316":"#555", fontWeight:700, cursor:"pointer", fontSize:14, fontFamily:"inherit" }}>{l}</button>
          ))}
        </div>
        <div style={{ color:"#666", fontSize:12, marginBottom:6 }}>PIN {role==="owner" ? "(ทดสอบ: 1234)" : "(ทดสอบ: 0000)"}</div>
        <input type="password" maxLength={4} value={pin} onChange={e=>{setPin(e.target.value);setErr(false);}} onKeyDown={e=>e.key==="Enter"&&go()} style={{...inp, fontSize:22, letterSpacing:8, textAlign:"center", marginBottom:6}} placeholder="••••" />
        {err && <div style={{ color:"#f87171", fontSize:12, marginBottom:6 }}>PIN ไม่ถูกต้อง</div>}
        <button onClick={go} style={{...btn(), width:"100%", padding:12, fontSize:15, marginTop:4}}>เข้าสู่ระบบ</button>
      </Card>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// COST DASHBOARD (main new feature)
// ─────────────────────────────────────────────────────────────
function CostDashboard({ stock, setStock }) {
  const [selItem, setSelItem] = useState(null);
  const [addForm, setAddForm] = useState({ date: todayStr(), unitCost: "", qty: "", note: "" });
  const [importing, setImporting] = useState(false);
  const fileRef = useRef();

  const spikes = stock.filter(isCostSpike);
  const totalStockValue = stock.reduce((a, s) => a + s.qty * weightedAvgCost(s), 0);

  // Import cost history from Excel
  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      // Expected columns: ชื่อวัตถุดิบ, วันที่, ราคา/หน่วย, จำนวน
      let updated = 0;
      setStock(prev => prev.map(item => {
        const newEntries = rows.filter(r => r["ชื่อวัตถุดิบ"] === item.name && r["วันที่"] && r["ราคา/หน่วย"]);
        if (!newEntries.length) return item;
        updated++;
        const toAdd = newEntries.map(r => ({
          date: String(r["วันที่"]).slice(0, 10),
          unitCost: +r["ราคา/หน่วย"],
          qty: +(r["จำนวน"] || 1),
          total: +(r["ราคา/หน่วย"]) * +(r["จำนวน"] || 1),
        }));
        const existDates = new Set((item.costHistory || []).map(h => h.date));
        const merged = [...(item.costHistory || []), ...toAdd.filter(h => !existDates.has(h.date))];
        return { ...item, costHistory: merged.sort((a, b) => a.date.localeCompare(b.date)) };
      }));
      alert(`นำเข้าข้อมูลราคาสำเร็จ ${updated} รายการ`);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const addCostEntry = () => {
    if (!addForm.unitCost || !addForm.qty || !selItem) return;
    const entry = {
      date: addForm.date,
      unitCost: +addForm.unitCost,
      qty: +addForm.qty,
      total: +addForm.unitCost * +addForm.qty,
    };
    setStock(prev => prev.map(s => s.id === selItem.id ? {
      ...s,
      costHistory: [...(s.costHistory || []), entry].sort((a, b) => a.date.localeCompare(b.date))
    } : s));
    setAddForm({ date: todayStr(), unitCost: "", qty: "", note: "" });
    setSelItem(prev => {
      const updated = stock.find(s => s.id === prev.id);
      return updated ? { ...updated, costHistory: [...(updated.costHistory || []), entry] } : prev;
    });
  };

  const delCostEntry = (itemId, idx) => {
    setStock(prev => prev.map(s => s.id === itemId ? {
      ...s,
      costHistory: (s.costHistory || []).filter((_, i) => i !== idx)
    } : s));
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <Hdr title="💹 ต้นทุนวัตถุดิบ"
        action={
          <div style={{ display: "flex", gap: 6 }}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display: "none" }} />
            <button onClick={() => fileRef.current?.click()} style={{ ...ghost, fontSize: 12, padding: "7px 11px" }}>📥 Import Excel</button>
            <button onClick={() => exportStockToXLSX(stock)} style={{ ...btn("#4ade80"), fontSize: 12, padding: "7px 11px" }}>📤 Export</button>
          </div>
        }
      />

      {/* KPI Row */}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        {[
          ["💎", "มูลค่าสต็อครวม", `฿${fmt(Math.round(totalStockValue))}`, "#a78bfa"],
          ["🚨", "ราคาผิดปกติ", `${spikes.length} รายการ`, spikes.length > 0 ? "#f87171" : "#4ade80"],
          ["📦", "วัตถุดิบทั้งหมด", `${stock.length} รายการ`, "#f97316"],
        ].map(([ic, l, v, c]) => (
          <div key={l} style={{ flex: 1, minWidth: 120, background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "12px 14px" }}>
            <div style={{ fontSize: 18, marginBottom: 3 }}>{ic}</div>
            <div style={{ color: "#666", fontSize: 11 }}>{l}</div>
            <div style={{ color: c, fontWeight: 800, fontSize: 18 }}>{v}</div>
          </div>
        ))}
      </div>

      {/* Spike alerts */}
      {spikes.length > 0 && (
        <Card style={{ borderColor: "rgba(248,113,113,0.4)", background: "rgba(248,113,113,0.04)" }}>
          <div style={{ color: "#f87171", fontWeight: 800, marginBottom: 10, fontSize: 13 }}>🚨 แจ้งเตือน: ราคาสูงผิดปกติ ({">"} ค่าเฉลี่ย 15%)</div>
          {spikes.map(s => (
            <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: "1px solid rgba(255,255,255,0.04)" }}>
              <div>
                <div style={{ color: "#fff", fontSize: 13 }}>{s.name}</div>
                <div style={{ color: "#666", fontSize: 11 }}>ราคาล่าสุด ฿{latestCost(s)} vs ค่าเฉลี่ย ฿{weightedAvgCost(s).toFixed(2)}/{s.unit}</div>
              </div>
              <span style={{ color: "#f87171", fontWeight: 800, fontSize: 14 }}>+{spikePercent(s)}%</span>
            </div>
          ))}
        </Card>
      )}

      {/* Stock cost table */}
      <Card>
        <div style={{ color: "#f97316", fontWeight: 800, fontSize: 13, marginBottom: 10 }}>ตารางต้นทุนวัตถุดิบ (คลิกเพื่อดูประวัติ)</div>
        {/* Header */}
        <div style={{ display: "grid", gridTemplateColumns: "1.5fr 0.7fr 1fr 1fr 0.9fr 0.9fr", gap: 4, padding: "6px 4px", borderBottom: "1px solid rgba(249,115,22,0.15)", marginBottom: 4 }}>
          {["ชื่อ","หน่วย","ต้นทุนเฉลี่ย","ราคาล่าสุด","จำนวน","มูลค่า"].map(h => (
            <div key={h} style={{ color: "#f97316", fontSize: 10, fontWeight: 700 }}>{h}</div>
          ))}
        </div>
        {stock.map(s => {
          const avg = weightedAvgCost(s);
          const latest = latestCost(s);
          const spike = isCostSpike(s);
          return (
            <div key={s.id} onClick={() => setSelItem(selItem?.id === s.id ? null : s)}
              style={{ display: "grid", gridTemplateColumns: "1.5fr 0.7fr 1fr 1fr 0.9fr 0.9fr", gap: 4, padding: "8px 4px", borderBottom: "1px solid rgba(255,255,255,0.04)", cursor: "pointer", background: selItem?.id === s.id ? "rgba(249,115,22,0.05)" : "transparent", borderRadius: 6 }}>
              <div style={{ color: "#ddd", fontSize: 12, display: "flex", alignItems: "center", gap: 5 }}>
                {spike && <span title={`ราคาสูงขึ้น ${spikePercent(s)}%`} style={{ color: "#f87171" }}>⚠️</span>}
                {s.name}
              </div>
              <div style={{ color: "#777", fontSize: 12 }}>{s.unit}</div>
              <div style={{ color: "#4ade80", fontSize: 12, fontWeight: 700 }}>฿{avg.toFixed(2)}</div>
              <div style={{ color: spike ? "#f87171" : "#ddd", fontSize: 12, fontWeight: spike ? 700 : 400 }}>฿{latest}</div>
              <div style={{ color: "#aaa", fontSize: 12 }}>{s.qty}</div>
              <div style={{ color: "#a78bfa", fontSize: 12 }}>฿{fmt(Math.round(s.qty * avg))}</div>
            </div>
          );
        })}
      </Card>

      {/* Detail panel */}
      {selItem && (() => {
        const item = stock.find(s => s.id === selItem.id) || selItem;
        const hist = [...(item.costHistory || [])].sort((a, b) => b.date.localeCompare(a.date));
        const avg = weightedAvgCost(item);
        return (
          <Card style={{ borderColor: "rgba(249,115,22,0.3)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div style={{ color: "#f97316", fontWeight: 800, fontSize: 14 }}>📋 {item.name} — ประวัติราคา</div>
              <button onClick={() => setSelItem(null)} style={{ ...ghost, padding: "4px 10px", fontSize: 12 }}>✕</button>
            </div>

            {/* Stats */}
            <div style={{ display: "flex", gap: 12, marginBottom: 14, flexWrap: "wrap" }}>
              {[
                ["ต้นทุนเฉลี่ย (WAC)", `฿${avg.toFixed(2)}`, "#4ade80"],
                ["ราคาล่าสุด", `฿${latestCost(item)}`, isCostSpike(item) ? "#f87171" : "#ddd"],
                ["จำนวนซื้อ", `${hist.length} ครั้ง`, "#a78bfa"],
                ["มูลค่าสต็อค", `฿${fmt(Math.round(item.qty * avg))}`, "#fbbf24"],
              ].map(([l, v, c]) => (
                <div key={l} style={{ background: "rgba(255,255,255,0.03)", borderRadius: 8, padding: "8px 12px" }}>
                  <div style={{ color: "#555", fontSize: 10 }}>{l}</div>
                  <div style={{ color: c, fontWeight: 800, fontSize: 15 }}>{v}</div>
                </div>
              ))}
            </div>

            {/* Mini sparkline */}
            {hist.length > 1 && (
              <div style={{ marginBottom: 12 }}>
                <div style={{ color: "#555", fontSize: 11, marginBottom: 6 }}>แนวโน้มราคา</div>
                <div style={{ display: "flex", alignItems: "flex-end", gap: 3, height: 50 }}>
                  {[...hist].reverse().map((h, i) => {
                    const maxP = Math.max(...hist.map(x => x.unitCost));
                    const minP = Math.min(...hist.map(x => x.unitCost));
                    const range = maxP - minP || 1;
                    const pct = ((h.unitCost - minP) / range * 40 + 10);
                    const isSpike = h.unitCost > avg * 1.15;
                    return (
                      <div key={i} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 2 }}>
                        <div style={{ width: "100%", height: `${pct}px`, background: isSpike ? "#f87171" : "#4ade80", borderRadius: "3px 3px 0 0", opacity: 0.8 }} />
                        <div style={{ color: "#333", fontSize: 8 }}>{h.date.slice(5)}</div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* History table */}
            <div style={{ marginBottom: 12 }}>
              {hist.map((h, i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: "1px solid rgba(255,255,255,0.04)", fontSize: 12 }}>
                  <span style={{ color: "#888" }}>{h.date}</span>
                  <span style={{ color: h.unitCost > avg * 1.15 ? "#f87171" : "#ddd" }}>฿{h.unitCost}/{item.unit}</span>
                  <span style={{ color: "#666" }}>×{h.qty} {item.unit}</span>
                  <span style={{ color: "#a78bfa" }}>฿{fmt(h.total)}</span>
                  <button onClick={() => delCostEntry(item.id, hist.length - 1 - i)} style={{ background: "none", border: "none", color: "#444", cursor: "pointer", fontSize: 14, padding: "0 2px" }}>×</button>
                </div>
              ))}
            </div>

            {/* Add entry */}
            <div style={{ borderTop: "1px solid rgba(255,255,255,0.06)", paddingTop: 12 }}>
              <div style={{ color: "#888", fontSize: 12, marginBottom: 8 }}>➕ บันทึกราคาซื้อครั้งใหม่</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 8 }}>
                {[["วันที่","date","date"],["ราคา/หน่วย","unitCost","number"],["จำนวน","qty","number"]].map(([l,k,t]) => (
                  <div key={k}>
                    <div style={{ color: "#555", fontSize: 10, marginBottom: 3 }}>{l}</div>
                    <input type={t} value={addForm[k]} onChange={e => setAddForm(p => ({...p,[k]:e.target.value}))} style={inp} />
                  </div>
                ))}
              </div>
              <button onClick={addCostEntry} style={{ ...btn(), width: "100%" }}>บันทึกราคา</button>
            </div>
          </Card>
        );
      })()}

      {/* Template download */}
      <Card style={{ background: "rgba(34,197,94,0.03)", borderColor: "rgba(34,197,94,0.15)" }}>
        <div style={{ color: "#22c55e", fontWeight: 800, fontSize: 13, marginBottom: 6 }}>📋 Template สำหรับ Import</div>
        <div style={{ color: "#555", fontSize: 12, marginBottom: 8 }}>ไฟล์ Excel ต้องมีคอลัมน์: ชื่อวัตถุดิบ | วันที่ | ราคา/หน่วย | จำนวน</div>
        <button onClick={() => {
          const template = [{ "ชื่อวัตถุดิบ": "หมูสามชั้น", "วันที่": todayStr(), "ราคา/หน่วย": 180, "จำนวน": 20 }];
          const ws = XLSX.utils.json_to_sheet(template);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, "ราคาวัตถุดิบ");
          XLSX.writeFile(wb, "template_cost_import.xlsx");
        }} style={{ ...btn("#22c55e"), fontSize: 12 }}>⬇️ ดาวน์โหลด Template</button>
      </Card>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// STOCK PAGE (Excel-style editable grid)
// ─────────────────────────────────────────────────────────────
function StockPage({ stock, setStock, role }) {
  const [editCell, setEditCell] = useState(null); // {id, field}
  const [editVal, setEditVal] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [newItem, setNewItem] = useState({ name:"", unit:"kg", qty:0, minQty:3, dailyUse:1, supplierId:1, expiryDays:3 });
  const [sortField, setSortField] = useState("name");
  const [sortDir, setSortDir] = useState(1);
  const [filterStatus, setFilterStatus] = useState("all");
  const fileRef = useRef();

  const EDITABLE_FIELDS = [
    { key: "name", label: "ชื่อ", w: "1.4fr", type: "text" },
    { key: "unit", label: "หน่วย", w: "0.6fr", type: "text" },
    { key: "qty", label: "คงเหลือ", w: "0.8fr", type: "number" },
    { key: "minQty", label: "ขั้นต่ำ", w: "0.8fr", type: "number" },
    { key: "dailyUse", label: "ใช้/วัน", w: "0.8fr", type: "number" },
    { key: "_avgCost", label: "ต้นทุนเฉลี่ย", w: "1fr", type: "readonly" },
    { key: "_value", label: "มูลค่า", w: "1fr", type: "readonly" },
    { key: "_status", label: "สถานะ", w: "0.9fr", type: "readonly" },
    { key: "_actions", label: "", w: "0.5fr", type: "readonly" },
  ];

  const sorted = useMemo(() => {
    let list = filterStatus === "all" ? stock : stock.filter(s => stockStatus(s) === filterStatus);
    return [...list].sort((a, b) => {
      const va = a[sortField] ?? ""; const vb = b[sortField] ?? "";
      return (va < vb ? -1 : va > vb ? 1 : 0) * sortDir;
    });
  }, [stock, sortField, sortDir, filterStatus]);

  const startEdit = (id, field, val) => { setEditCell({ id, field }); setEditVal(String(val)); };
  const commitEdit = () => {
    if (!editCell) return;
    setStock(prev => prev.map(s => s.id === editCell.id ? {
      ...s, [editCell.field]: ["qty","minQty","dailyUse","expiryDays"].includes(editCell.field) ? +editVal : editVal
    } : s));
    setEditCell(null);
  };

  const delItem = (id) => setStock(prev => prev.filter(s => s.id !== id));

  const addItem = () => {
    setStock(prev => [...prev, { ...newItem, id: Date.now(), qty: +newItem.qty, minQty: +newItem.minQty, dailyUse: +newItem.dailyUse, expiryDays: +newItem.expiryDays, costHistory: [] }]);
    setShowAdd(false);
    setNewItem({ name:"",unit:"kg",qty:0,minQty:3,dailyUse:1,supplierId:1,expiryDays:3 });
  };

  // Import stock from Excel
  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      const imported = rows.map((r, i) => ({
        id: Date.now() + i,
        name: r["ชื่อวัตถุดิบ"] || r["name"] || "",
        unit: r["หน่วย"] || r["unit"] || "kg",
        qty: +(r["จำนวนคงเหลือ"] || r["qty"] || 0),
        minQty: +(r["จำนวนขั้นต่ำ"] || r["minQty"] || 3),
        dailyUse: +(r["ใช้ต่อวัน"] || r["dailyUse"] || 1),
        expiryDays: +(r["อายุ (วัน)"] || r["expiryDays"] || 3),
        supplierId: +(r["supplierId"] || 1),
        costHistory: r["ราคาล่าสุด (฿/หน่วย)"] ? [{ date: todayStr(), unitCost: +r["ราคาล่าสุด (฿/หน่วย)"], qty: 1, total: +r["ราคาล่าสุด (฿/หน่วย)"] }] : [],
      })).filter(r => r.name);
      if (window.confirm(`นำเข้าสต็อค ${imported.length} รายการ (จะ replace ข้อมูลเดิม)?`)) setStock(imported);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const colTemplate = EDITABLE_FIELDS.map(f => f.w).join(" ");

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <Hdr title="📦 สต็อค (แก้ไขแบบตาราง)"
        action={
          <div style={{ display: "flex", gap: 6 }}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display: "none" }} />
            <button onClick={() => fileRef.current?.click()} style={{ ...ghost, fontSize: 11, padding: "6px 10px" }}>📥 Import</button>
            <button onClick={() => exportStockToXLSX(stock)} style={{ ...ghost, fontSize: 11, padding: "6px 10px" }}>📤 Export</button>
            {role==="owner" && <button onClick={() => setShowAdd(!showAdd)} style={{ ...btn(), fontSize: 11, padding: "6px 10px" }}>+ เพิ่ม</button>}
          </div>
        }
      />

      {/* Filter */}
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
        {[["all","ทั้งหมด"],["critical","🔴 วิกฤต"],["low","🟡 ต่ำ"],["ok","🟢 ปกติ"]].map(([v,l]) => (
          <button key={v} onClick={() => setFilterStatus(v)} style={{ background: filterStatus===v ? "rgba(249,115,22,0.15)" : "rgba(255,255,255,0.04)", border:`1px solid ${filterStatus===v ? "#f97316aa" : "rgba(255,255,255,0.08)"}`, borderRadius:7, padding:"5px 12px", color:filterStatus===v ? "#f97316":"#666", cursor:"pointer", fontSize:11, fontFamily:"inherit" }}>{l}</button>
        ))}
      </div>

      {/* Add form */}
      {showAdd && (
        <Card style={{ borderColor: "rgba(249,115,22,0.3)" }}>
          <div style={{ color: "#f97316", fontWeight: 800, marginBottom: 10, fontSize: 13 }}>➕ เพิ่มวัตถุดิบ</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8 }}>
            {[["ชื่อ","name","text"],["หน่วย","unit","text"],["จำนวน","qty","number"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"],["อายุ(วัน)","expiryDays","number"]].map(([l,k,t]) => (
              <div key={k}>
                <div style={{ color:"#666", fontSize:10, marginBottom:3 }}>{l}</div>
                <input type={t} value={newItem[k]} onChange={e => setNewItem(p=>({...p,[k]:e.target.value}))} style={inp} />
              </div>
            ))}
          </div>
          <div style={{ display:"flex", gap:8, marginTop:10 }}>
            <button onClick={addItem} style={{...btn(), flex:1}}>บันทึก</button>
            <button onClick={() => setShowAdd(false)} style={ghost}>ยกเลิก</button>
          </div>
        </Card>
      )}

      {/* Excel-style table */}
      <Card style={{ padding: "10px 12px", overflowX: "auto" }}>
        {/* Header */}
        <div style={{ display: "grid", gridTemplateColumns: colTemplate, gap: 2, padding: "6px 4px", borderBottom: "1.5px solid rgba(249,115,22,0.2)", marginBottom: 2 }}>
          {EDITABLE_FIELDS.map(f => (
            <div key={f.key} onClick={() => { if (f.type !== "readonly") { setSortField(f.key); setSortDir(p => p * -1); }}}
              style={{ color: "#f97316", fontSize: 10, fontWeight: 700, cursor: f.type !== "readonly" ? "pointer" : "default", userSelect: "none", display: "flex", alignItems: "center", gap: 3 }}>
              {f.label}
              {sortField === f.key && <span style={{ fontSize: 8 }}>{sortDir > 0 ? "▲" : "▼"}</span>}
            </div>
          ))}
        </div>

        {/* Rows */}
        {sorted.map((item, ri) => {
          const avg = weightedAvgCost(item);
          const spike = isCostSpike(item);
          const st = stockStatus(item);
          return (
            <div key={item.id} style={{ display: "grid", gridTemplateColumns: colTemplate, gap: 2, padding: "3px 4px", borderBottom: "1px solid rgba(255,255,255,0.04)", background: ri % 2 === 0 ? "transparent" : "rgba(255,255,255,0.01)", alignItems: "center" }}>
              {EDITABLE_FIELDS.map(f => {
                const isEditing = editCell?.id === item.id && editCell?.field === f.key;
                if (f.type === "readonly") {
                  let val, color = "#888";
                  if (f.key === "_avgCost") { val = `฿${avg.toFixed(2)}`; color = "#4ade80"; }
                  else if (f.key === "_value") { val = `฿${fmt(Math.round(item.qty * avg))}`; color = "#a78bfa"; }
                  else if (f.key === "_status") return <div key={f.key}><Tag status={st} /></div>;
                  else if (f.key === "_actions") return (
                    <div key={f.key} style={{ display: "flex", gap: 3 }}>
                      {spike && <span title={`ราคาสูง ${spikePercent(item)}%`} style={{ fontSize: 13, cursor: "default" }}>⚠️</span>}
                      <button onClick={() => delItem(item.id)} style={{ background:"none", border:"none", color:"#333", cursor:"pointer", fontSize:14, padding:0 }}>🗑</button>
                    </div>
                  );
                  return <div key={f.key} style={{ color, fontSize: 12 }}>{val}</div>;
                }
                return (
                  <div key={f.key}>
                    {isEditing ? (
                      <input
                        type={f.type} value={editVal} autoFocus
                        onChange={e => setEditVal(e.target.value)}
                        onBlur={commitEdit}
                        onKeyDown={e => { if (e.key === "Enter") commitEdit(); if (e.key === "Escape") setEditCell(null); }}
                        style={{ ...inp, padding: "3px 6px", fontSize: 12, width: "100%", background: "rgba(249,115,22,0.12)", border: "1.5px solid #f97316" }}
                      />
                    ) : (
                      <div
                        onClick={() => role === "owner" ? startEdit(item.id, f.key, item[f.key]) : null}
                        style={{ color: "#ddd", fontSize: 12, padding: "3px 5px", borderRadius: 5, cursor: role==="owner" ? "text" : "default", background: "transparent", minHeight: 22, border: "1px solid transparent", transition: "all .15s" }}
                        onMouseEnter={e => { if(role==="owner") e.target.style.background="rgba(255,255,255,0.05)"; e.target.style.borderColor="rgba(249,115,22,0.2)"; }}
                        onMouseLeave={e => { e.target.style.background="transparent"; e.target.style.borderColor="transparent"; }}
                      >
                        {item[f.key]}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          );
        })}
      </Card>

      <div style={{ color: "#444", fontSize: 11, textAlign: "center" }}>
        💡 คลิกที่ตัวเลขในตารางเพื่อแก้ไขได้เลย • กด Enter เพื่อบันทึก • กด Escape เพื่อยกเลิก
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// CASHFLOW PAGE
// ─────────────────────────────────────────────────────────────
function CashflowPage({ cf, setCF, role }) {
  const [showForm, setShowForm] = useState(false);
  const [viewTab, setViewTab] = useState("list");
  const [filterMonth, setFilterMonth] = useState(todayStr().slice(0,7));
  const [form, setForm] = useState({ date:todayStr(), flow:"in", cat:IN_CATS[0], itemName:"", amount:"", method:"เงินสด", note:"" });
  const [openingBal, setOpeningBal] = useState(15000);
  const fileRef = useRef();

  const months = [...new Set(cf.map(e => e.date.slice(0,7)))].sort((a,b) => b.localeCompare(a));
  const filtered = cf.filter(e => e.date.startsWith(filterMonth));
  const totalIn = filtered.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const totalOut = filtered.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const netFlow = totalIn - totalOut;

  const withBal = (() => {
    let bal = openingBal;
    return [...filtered].sort((a,b)=>a.date.localeCompare(b.date)||a.id-b.id).map(e=>{
      bal += e.flow==="in" ? e.amount : -e.amount;
      return {...e, runBal:bal};
    });
  })();
  const sorted = [...withBal].sort((a,b)=>b.date.localeCompare(a.date)||b.id-a.id);
  const byDate = sorted.reduce((acc,e)=>{ if(!acc[e.date])acc[e.date]=[]; acc[e.date].push(e); return acc; },{});
  const dates = Object.keys(byDate).sort((a,b)=>b.localeCompare(a));
  const closingBal = openingBal + netFlow;

  const varCostCats = ["วัตถุดิบ/ผัก","วัตถุดิบ/หมู-ทะเล","ซอส/เครื่องปรุง","บรรจุภัณฑ์"];
  const cogs = filtered.filter(e=>e.flow==="out"&&varCostCats.includes(e.cat)).reduce((a,b)=>a+b.amount,0);
  const grossP = totalIn - cogs;
  const grossM = totalIn>0 ? ((grossP/totalIn)*100).toFixed(1) : 0;

  const addEntry = () => {
    if (!form.amount) return;
    setCF(prev => [{...form, id:Date.now(), amount:+form.amount}, ...prev]);
    setShowForm(false);
    setForm({date:todayStr(), flow:"in", cat:IN_CATS[0], itemName:"", amount:"", method:"เงินสด", note:""});
  };
  const del = (id) => setCF(prev=>prev.filter(e=>e.id!==id));

  // Import CF from Excel
  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type:"array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      const imported = rows.map((r, i) => ({
        id: Date.now() + i,
        date: String(r["วันที่"] || "").slice(0,10),
        flow: r["ประเภท"]==="รายรับ" ? "in" : "out",
        cat: r["หมวดหมู่"] || r["หมวด"] || "อื่นๆ",
        amount: +(r["จำนวนเงิน (฿)"] || r["จำนวน"] || 0),
        method: r["ช่องทาง"] || "เงินสด",
        note: r["หมายเหตุ"] || "",
      })).filter(r => r.date && r.amount > 0);
      setCF(prev => {
        const existIds = new Set(prev.map(e => e.date + e.cat + e.amount));
        return [...prev, ...imported.filter(r => !existIds.has(r.date + r.cat + r.amount))];
      });
      alert(`นำเข้า Cash Flow ${imported.length} รายการ`);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const catSum = filtered.reduce((acc,e)=>{ const k=`${e.flow}__${e.cat}`; if(!acc[k])acc[k]={flow:e.flow,cat:e.cat,total:0,count:0}; acc[k].total+=e.amount; acc[k].count++; return acc; },{});

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <Hdr title="💵 Cash Flow"
        action={
          <div style={{ display:"flex", gap:6 }}>
            {role==="owner" && <>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display:"none" }} />
              <button onClick={()=>fileRef.current?.click()} style={{...ghost,fontSize:11,padding:"6px 10px"}}>📥 Import</button>
              <button onClick={()=>exportCFToXLSX(cf)} style={{...ghost,fontSize:11,padding:"6px 10px"}}>📤 Export</button>
            </>}
            <button onClick={()=>setShowForm(!showForm)} style={{...btn(),fontSize:12,padding:"7px 12px"}}>+ กรอก</button>
          </div>
        }
      />

      {/* Opening balance — owner only */}
      {role==="owner" && (
        <Card style={{ padding:"12px 16px" }}>
          <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
            <span style={{ color:"#777", fontSize:12 }}>ยอดเปิดต้นงวด ฿</span>
            <input type="number" value={openingBal} onChange={e=>setOpeningBal(+e.target.value||0)} style={{...inp,width:110,textAlign:"right"}} />
            <div style={{ marginLeft:"auto", textAlign:"right" }}>
              <div style={{ color:"#555", fontSize:10 }}>ยอดปิดงวด</div>
              <div style={{ color:closingBal>=0?"#4ade80":"#f87171", fontWeight:800, fontSize:17 }}>฿{fmt(closingBal)}</div>
            </div>
          </div>
        </Card>
      )}

      {/* KPI */}
      <div style={{ display:"flex", gap:8 }}>
        {[["เงินเข้า",`฿${fmt(totalIn)}`,"#4ade80"],["เงินออก",`฿${fmt(totalOut)}`,"#f87171"],["Net",`${netFlow>=0?"+":""}฿${fmt(netFlow)}`,netFlow>=0?"#4ade80":"#f87171"],
          ...(role==="owner"?[["Gross Margin",`${grossM}%`,"#a78bfa"]]:[])]
          .map(([l,v,c])=>(
          <div key={l} style={{ flex:1, background:"rgba(255,255,255,0.03)", border:"1px solid rgba(255,255,255,0.06)", borderRadius:10, padding:"9px 5px", textAlign:"center" }}>
            <div style={{ color:"#555", fontSize:9 }}>{l}</div>
            <div style={{ color:c, fontWeight:800, fontSize:13 }}>{v}</div>
          </div>
        ))}
      </div>

      {/* Tabs */}
      <div style={{ display:"flex", gap:6, justifyContent:"space-between", flexWrap:"wrap" }}>
        <div style={{ display:"flex", gap:5 }}>
          {months.map(m=>(
            <button key={m} onClick={()=>setFilterMonth(m)} style={{ background:filterMonth===m?"rgba(249,115,22,0.15)":"rgba(255,255,255,0.04)", border:`1px solid ${filterMonth===m?"#f97316aa":"rgba(255,255,255,0.08)"}`, borderRadius:7, padding:"4px 10px", color:filterMonth===m?"#f97316":"#666", cursor:"pointer", fontSize:11, fontFamily:"inherit" }}>{m}</button>
          ))}
        </div>
        <div style={{ display:"flex", gap:5 }}>
          {[["list","รายการ"],["summary","สรุป"]].map(([v,l])=>(
            <button key={v} onClick={()=>setViewTab(v)} style={{ background:viewTab===v?"rgba(167,139,250,0.15)":"rgba(255,255,255,0.04)", border:`1px solid ${viewTab===v?"#a78bfa88":"rgba(255,255,255,0.08)"}`, borderRadius:7, padding:"4px 11px", color:viewTab===v?"#a78bfa":"#666", cursor:"pointer", fontSize:11, fontFamily:"inherit" }}>{l}</button>
          ))}
        </div>
      </div>

      {/* Add form */}
      {showForm && (
        <Card style={{ borderColor:"rgba(249,115,22,0.3)" }}>
          <div style={{ color:"#f97316", fontWeight:800, marginBottom:12, fontSize:14 }}>📝 บันทึกรายการ</div>

          {/* flow type */}
          <div style={{ display:"flex", gap:8, marginBottom:12 }}>
            {[["in","💰 รายรับ","#4ade80"],["out","💸 รายจ่าย","#f87171"]].map(([v,l,c])=>(
              <button key={v} onClick={()=>setForm(p=>({...p,flow:v,cat:v==="in"?IN_CATS[0]:OUT_CATS[0],itemName:""}))} style={{ flex:1, padding:10, background:form.flow===v?`${c}18`:"rgba(255,255,255,0.04)", border:`1.5px solid ${form.flow===v?c:"rgba(255,255,255,0.1)"}`, borderRadius:10, color:form.flow===v?c:"#555", fontWeight:700, cursor:"pointer", fontSize:13, fontFamily:"inherit" }}>{l}</button>
            ))}
          </div>

          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
            {/* Date */}
            <div>
              <div style={{ color:"#666",fontSize:11,marginBottom:3 }}>วันที่</div>
              <input type="date" value={form.date} onChange={e=>setForm(p=>({...p,date:e.target.value}))} style={inp} />
            </div>
            {/* Amount */}
            <div>
              <div style={{ color:"#666",fontSize:11,marginBottom:3 }}>จำนวนเงิน (฿)</div>
              <input type="number" value={form.amount} placeholder="0.00" onChange={e=>setForm(p=>({...p,amount:e.target.value}))} style={inp} />
            </div>

            {/* Category (หมวด) */}
            <div>
              <div style={{ color:"#666",fontSize:11,marginBottom:3 }}>
                หมวดหมู่ {form.flow==="out" && <span style={{ color:"#555",fontSize:10 }}>(สำหรับรายงาน)</span>}
              </div>
              <select value={form.cat} onChange={e=>setForm(p=>({...p,cat:e.target.value}))} style={inp}>
                {(form.flow==="in"?IN_CATS:OUT_CATS).map(c=><option key={c}>{c}</option>)}
              </select>
            </div>

            {/* Item name — ชื่อรายการละเอียด */}
            <div>
              <div style={{ color:"#666",fontSize:11,marginBottom:3 }}>
                ชื่อรายการ <span style={{ color:"#555",fontSize:10 }}>(ระบุให้ชัด)</span>
              </div>
              <input type="text" value={form.itemName} onChange={e=>setForm(p=>({...p,itemName:e.target.value}))}
                style={inp}
                placeholder={
                  form.cat==="วัตถุดิบ/เนื้อสัตว์" ? "เช่น หมูสามชั้น, หมูชาบู, กุ้ง" :
                  form.cat==="วัตถุดิบ/ผัก"         ? "เช่น ผักกาด, เห็ดหอม, ข้าวโพด" :
                  form.cat==="วัตถุดิบ/ลูกชิ้น"     ? "เช่น ลูกชิ้นหมู, เกี๊ยว" :
                  form.cat==="วัตถุดิบ/เครื่องดื่ม" ? "เช่น น้ำอัดลม, ชาไทย" :
                  form.cat==="วัตถุดิบ/ซอส"          ? "เช่น ซอสหม่าล่า, น้ำซุป" :
                  form.cat==="วัตถุดิบ/บรรจุภัณฑ์"  ? "เช่น กล่องโฟม, ถุงซิป" :
                  form.flow==="in" ? "เช่น ยอดขาย dine-in 52 ออร์เดอร์" :
                  "ระบุรายการ..."
                }
              />
            </div>

            {/* Payment method */}
            <div>
              <div style={{ color:"#666",fontSize:11,marginBottom:3 }}>ช่องทางชำระ</div>
              <select value={form.method} onChange={e=>setForm(p=>({...p,method:e.target.value}))} style={inp}>{PAY_METHODS.map(m=><option key={m}>{m}</option>)}</select>
            </div>

            {/* Note */}
            <div>
              <div style={{ color:"#666",fontSize:11,marginBottom:3 }}>หมายเหตุเพิ่มเติม</div>
              <input type="text" value={form.note} onChange={e=>setForm(p=>({...p,note:e.target.value}))} style={inp} placeholder="(ไม่บังคับ)" />
            </div>
          </div>

          {/* Preview */}
          {form.itemName && (
            <div style={{ marginTop:10, padding:"8px 12px", background:"rgba(255,255,255,0.04)", borderRadius:8, fontSize:12 }}>
              <span style={{ color:"#555" }}>จะบันทึก: </span>
              <span style={{ color:form.flow==="in"?"#4ade80":"#f87171", fontWeight:700 }}>{form.flow==="in"?"รายรับ":"รายจ่าย"}</span>
              <span style={{ color:"#aaa" }}> • {form.cat} • </span>
              <span style={{ color:"#fff", fontWeight:700 }}>{form.itemName}</span>
              {form.amount && <span style={{ color:"#f97316", fontWeight:700 }}> — ฿{fmt(+form.amount)}</span>}
            </div>
          )}

          <div style={{ display:"flex", gap:8, marginTop:12 }}>
            <button onClick={addEntry} style={{...btn(form.flow==="in"?"#4ade80":"#f87171"),flex:1,padding:10,fontSize:14}}>✅ บันทึก</button>
            <button onClick={()=>setShowForm(false)} style={{...ghost,padding:"10px 14px"}}>ยกเลิก</button>
          </div>
        </Card>
      )}

      {/* List view */}
      {viewTab==="list" && dates.map(date=>{
        const items = byDate[date];
        const dIn = items.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
        const dOut = items.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
        return (
          <Card key={date} style={{ padding:"12px 15px" }}>
            <div style={{ display:"flex", justifyContent:"space-between", paddingBottom:8, marginBottom:8, borderBottom:"1px solid rgba(249,115,22,0.1)" }}>
              <span style={{ color:"#f97316", fontWeight:800, fontSize:13 }}>{date}</span>
              <div style={{ display:"flex", gap:8, fontSize:12 }}>
                <span style={{ color:"#4ade80" }}>+{fmt(dIn)}</span>
                <span style={{ color:"#f87171" }}>-{fmt(dOut)}</span>
                <span style={{ color:dIn-dOut>=0?"#4ade80":"#f87171", fontWeight:700 }}>=฿{fmt(dIn-dOut)}</span>
              </div>
            </div>
            {items.map(e=>(
              <div key={e.id} style={{ display:"flex", alignItems:"center", gap:9, padding:"6px 0", borderBottom:"1px solid rgba(255,255,255,0.03)" }}>
                <div style={{ width:24,height:24,borderRadius:7,flexShrink:0, background:e.flow==="in"?"rgba(74,222,128,0.12)":"rgba(248,113,113,0.12)", display:"flex",alignItems:"center",justifyContent:"center",fontSize:12, color:e.flow==="in"?"#4ade80":"#f87171" }}>{e.flow==="in"?"↓":"↑"}</div>
                <div style={{ flex:1,minWidth:0 }}>
                  <div style={{ color:"#ddd",fontSize:13,fontWeight: e.itemName ? 600 : 400 }}>
                    {e.itemName || e.cat}
                  </div>
                  <div style={{ color:"#555",fontSize:10 }}>
                    {e.itemName ? <span style={{ color:"#666" }}>{e.cat} • </span> : ""}
                    {e.method}{e.note?` • ${e.note}`:""}
                  </div>
                </div>
                <div style={{ textAlign:"right",flexShrink:0 }}>
                  <div style={{ color:e.flow==="in"?"#4ade80":"#f87171",fontWeight:700,fontSize:14 }}>{e.flow==="in"?"+":"-"}฿{fmt(e.amount)}</div>
                  <div style={{ color:"#333",fontSize:9 }}>฿{fmt(e.runBal)}</div>
                </div>
                <button onClick={()=>del(e.id)} style={{ background:"none",border:"none",color:"#333",cursor:"pointer",fontSize:14,padding:"0 2px" }}>×</button>
              </div>
            ))}
          </Card>
        );
      })}

      {/* Summary view */}
      {viewTab==="summary" && (
        <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
          <Card>
            <div style={{ color:"#f97316",fontWeight:800,marginBottom:10,fontSize:13 }}>Gross Profit Analysis</div>
            {[["รายรับรวม",totalIn,"#4ade80"],["− ต้นทุนวัตถุดิบ (COGS)",cogs,"#f87171"],["= กำไรขั้นต้น",grossP,"#4ade80"],["Gross Margin %",`${grossM}%`,"#a78bfa"]].map(([l,v,c])=>(
              <div key={l} style={{ display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid rgba(255,255,255,0.04)",fontSize:13 }}>
                <span style={{ color:"#aaa" }}>{l}</span>
                <span style={{ color:c, fontWeight:700 }}>{typeof v==="string"?v:`฿${fmt(v)}`}</span>
              </div>
            ))}
          </Card>
          <Card>
            <div style={{ color:"#f87171",fontWeight:800,marginBottom:10,fontSize:13 }}>รายจ่ายแยกหมวด + รายละเอียด</div>
            {Object.values(catSum).filter(c=>c.flow==="out").sort((a,b)=>b.total-a.total).map(c=>{
              const pct = totalOut>0?(c.total/totalOut*100).toFixed(0):0;
              // group items within this cat by itemName
              const catEntries = filtered.filter(e=>e.flow==="out"&&e.cat===c.cat);
              const byItem = catEntries.reduce((acc,e)=>{
                const k = e.itemName || e.note || "(ไม่ระบุชื่อ)";
                if(!acc[k]) acc[k]=0; acc[k]+=e.amount; return acc;
              },{});
              const itemList = Object.entries(byItem).sort((a,b)=>b[1]-a[1]);
              const grp = INGREDIENT_GROUPS.find(g=>g.key===c.cat);
              const barColor = grp ? grp.color : "#f87171";
              return (
                <div key={c.cat} style={{ marginBottom:12 }}>
                  <div style={{ display:"flex",justifyContent:"space-between",marginBottom:4 }}>
                    <span style={{ color:grp?grp.color:"#ccc",fontSize:13,fontWeight:700 }}>
                      {grp?grp.label:c.cat} <span style={{ color:"#444",fontSize:10,fontWeight:400 }}>({c.count}ครั้ง)</span>
                    </span>
                    <span style={{ color:barColor,fontSize:13,fontWeight:700 }}>฿{fmt(c.total)} ({pct}%)</span>
                  </div>
                  <div style={{ background:"rgba(255,255,255,0.05)",borderRadius:4,height:6,marginBottom:6 }}>
                    <div style={{ background:`linear-gradient(90deg,${barColor},${barColor}88)`,width:`${pct}%`,height:"100%",borderRadius:4 }} />
                  </div>
                  {/* itemName breakdown */}
                  {itemList.length > 0 && (
                    <div style={{ paddingLeft:8,display:"flex",flexWrap:"wrap",gap:5 }}>
                      {itemList.map(([name,amt])=>(
                        <div key={name} style={{ background:"rgba(255,255,255,0.04)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:6,padding:"3px 8px",fontSize:11,display:"flex",gap:6,alignItems:"center" }}>
                          <span style={{ color:"#aaa" }}>{name}</span>
                          <span style={{ color:barColor,fontWeight:700 }}>฿{fmt(amt)}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </Card>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────
function Dashboard({ stock, cf }) {
  const todayCF = cf.filter(e => e.date===todayStr());
  const todayIn = todayCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const todayOut = todayCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const todayGP = todayIn - todayOut;

  const mk = todayStr().slice(0,7);
  const mCF = cf.filter(e=>e.date.startsWith(mk));
  const mIn = mCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mOut = mCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const mNet = mIn - mOut - TOTAL_FIXED;

  const spikes = stock.filter(isCostSpike);
  const critical = stock.filter(s=>["critical","out"].includes(stockStatus(s)));
  const lowS = stock.filter(s=>stockStatus(s)==="low");
  const totalStockVal = stock.reduce((a,s)=>a+s.qty*weightedAvgCost(s),0);

  const last7 = Array.from({length:7},(_,i)=>{
    const d = new Date(); d.setDate(d.getDate()-(6-i));
    const k = d.toISOString().split("T")[0];
    const dIn = cf.filter(e=>e.date===k&&e.flow==="in").reduce((a,b)=>a+b.amount,0);
    const dOut = cf.filter(e=>e.date===k&&e.flow==="out").reduce((a,b)=>a+b.amount,0);
    return {k, label:`${d.getDate()}/${d.getMonth()+1}`, in:dIn, out:dOut, p:dIn-dOut};
  });
  const maxBar = Math.max(...last7.map(d=>d.in),1);

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:14 }}>
      <div>
        <div style={{ color:"#f97316",fontSize:18,fontWeight:900 }}>ภาพรวมวันนี้</div>
        <div style={{ color:"#444",fontSize:11 }}>{todayStr()}</div>
      </div>

      {/* Today KPI */}
      <div style={{ display:"flex",gap:10,flexWrap:"wrap" }}>
        {[["💰","รายรับวันนี้",`฿${fmt(todayIn)}`,"#4ade80"],["💸","รายจ่าย",`฿${fmt(todayOut)}`,"#f87171"],["📈","กำไรวันนี้",`฿${fmt(todayGP)}`,todayGP>=0?"#4ade80":"#f87171"],["💎","มูลค่าสต็อค",`฿${fmt(Math.round(totalStockVal))}`,"#a78bfa"]].map(([ic,l,v,c])=>(
          <div key={l} style={{ flex:1,minWidth:120, background:"linear-gradient(135deg,rgba(249,115,22,.07),rgba(180,40,40,.04))", border:"1px solid rgba(249,115,22,.15)", borderRadius:12, padding:"13px 14px" }}>
            <div style={{ fontSize:20 }}>{ic}</div>
            <div style={{ color:"#666",fontSize:11,marginTop:2 }}>{l}</div>
            <div style={{ color:c,fontWeight:800,fontSize:18 }}>{v}</div>
          </div>
        ))}
      </div>

      {/* Month P&L */}
      <Card style={{ background:"rgba(249,115,22,0.03)", borderColor:"rgba(249,115,22,0.2)" }}>
        <div style={{ color:"#f97316",fontWeight:800,fontSize:13,marginBottom:10 }}>📅 P&L เดือน {mk}</div>
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
          {[["รายรับ",`฿${fmt(mIn)}`,"#4ade80"],["รายจ่ายผันแปร",`฿${fmt(mOut)}`,"#f87171"],["ต้นทุนคงที่",`฿${fmt(TOTAL_FIXED)}`,"#fbbf24"],["กำไรสุทธิ",`฿${fmt(mNet)}`,mNet>=0?"#4ade80":"#f87171"]].map(([l,v,c])=>(
            <div key={l}><div style={{ color:"#555",fontSize:11 }}>{l}</div><div style={{ color:c,fontWeight:800,fontSize:16 }}>{v}</div></div>
          ))}
        </div>
        <div style={{ marginTop:10,borderTop:"1px solid rgba(255,255,255,0.05)",paddingTop:8,color:"#444",fontSize:10 }}>
          Breakeven/วัน ≈ ฿{fmt(Math.ceil(TOTAL_FIXED/30))} • เป้า/วัน ≈ ฿{fmt(Math.ceil(TOTAL_FIXED/30*2.5))}
        </div>
      </Card>

      {/* Chart */}
      <Card>
        <div style={{ color:"#f97316",fontWeight:800,fontSize:13,marginBottom:10 }}>📊 รายรับ 7 วันล่าสุด</div>
        <div style={{ display:"flex",alignItems:"flex-end",gap:5,height:65 }}>
          {last7.map(d=>(
            <div key={d.k} style={{ flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:3 }}>
              <div style={{ width:"100%",borderRadius:"3px 3px 0 0", background:d.p>=0?"rgba(74,222,128,0.55)":"rgba(248,113,113,0.55)", height:`${Math.max((d.in/maxBar)*55,3)}px`, border:d.k===todayStr()?"1.5px solid #f97316":"none" }} />
              <div style={{ color:"#383838",fontSize:8 }}>{d.label}</div>
            </div>
          ))}
        </div>
      </Card>

      {/* Alerts */}
      {spikes.length>0 && (
        <Card style={{ borderColor:"rgba(248,113,113,0.35)",background:"rgba(248,113,113,0.04)" }}>
          <div style={{ color:"#f87171",fontWeight:800,fontSize:13,marginBottom:8 }}>💹 ราคาวัตถุดิบผิดปกติ ({spikes.length} รายการ)</div>
          {spikes.map(s=>(
            <div key={s.id} style={{ display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid rgba(255,255,255,0.04)",fontSize:12 }}>
              <span style={{ color:"#ddd" }}>{s.name}</span>
              <span style={{ color:"#f87171",fontWeight:700 }}>+{spikePercent(s)}% จากค่าเฉลี่ย</span>
            </div>
          ))}
        </Card>
      )}
      {critical.length>0 && (
        <Card style={{ borderColor:"rgba(248,113,113,0.3)" }}>
          <div style={{ color:"#f87171",fontWeight:800,fontSize:13,marginBottom:8 }}>🚨 สต็อควิกฤต ({critical.length})</div>
          {critical.map(s=>(
            <div key={s.id} style={{ display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid rgba(255,255,255,0.04)",fontSize:12 }}>
              <span style={{ color:"#ddd" }}>{s.name}</span>
              <div style={{ display:"flex",gap:8,alignItems:"center" }}>
                <span style={{ color:"#666" }}>เหลือ {s.qty} {s.unit}</span><Tag status={stockStatus(s)} />
              </div>
            </div>
          ))}
        </Card>
      )}
      {lowS.length>0 && (
        <Card style={{ borderColor:"rgba(251,191,36,0.2)" }}>
          <div style={{ color:"#fbbf24",fontWeight:800,fontSize:13,marginBottom:7 }}>⚠️ ควรสั่งเพิ่ม ({lowS.length})</div>
          <div style={{ display:"flex",flexWrap:"wrap",gap:6 }}>
            {lowS.map(s=><span key={s.id} style={{ background:"rgba(251,191,36,0.1)",border:"1px solid rgba(251,191,36,0.25)",borderRadius:7,padding:"3px 9px",fontSize:11,color:"#fbbf24" }}>{s.name}</span>)}
          </div>
        </Card>
      )}

      {/* Export button */}
      <button onClick={() => exportFullReport(INIT_STOCK, INIT_CF)} style={{...btn("#a78bfa"), width:"100%", padding:12, fontSize:14}}>
        📊 Export รายงานสรุปทั้งหมด (.xlsx)
      </button>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// REPORT PAGE — สรุปต้นทุน-กำไร แยกหมวดวัตถุดิบ
// ─────────────────────────────────────────────────────────────
function ReportPage({ cf, stock }) {
  const months = [...new Set(cf.map(e=>e.date.slice(0,7)))].sort((a,b)=>b.localeCompare(a));
  const [m, setM] = useState(months[0]||todayStr().slice(0,7));
  const [tab, setTab] = useState("pl"); // pl | cogs | opex | daily

  const mCF = cf.filter(e=>e.date.startsWith(m));
  const mIn  = mCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mOut = mCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);

  // COGS แยกแต่ละหมวดวัตถุดิบ
  const cogsMap = {};
  INGREDIENT_GROUPS.forEach(g => {
    cogsMap[g.key] = mCF.filter(e=>e.flow==="out"&&e.cat===g.key).reduce((a,b)=>a+b.amount,0);
  });
  const totalCOGS = Object.values(cogsMap).reduce((a,b)=>a+b,0);

  // OPEX แยกรายการ (จ่ายจริง จาก CF) + ต้นทุนคงที่ที่ยังไม่ได้บันทึก
  const opexActual = {};
  OPEX_CATS.forEach(cat => {
    opexActual[cat] = mCF.filter(e=>e.flow==="out"&&e.cat===cat).reduce((a,b)=>a+b.amount,0);
  });
  const totalOPEX = Object.values(opexActual).reduce((a,b)=>a+b,0);

  // P&L waterfall
  const grossP = mIn - totalCOGS;
  const netP   = grossP - totalOPEX;
  const grossM = mIn>0?((grossP/mIn)*100).toFixed(1):"0";
  const netM   = mIn>0?((netP/mIn)*100).toFixed(1):"0";
  const cogsPct= mIn>0?((totalCOGS/mIn)*100).toFixed(1):"0";
  const opexPct= mIn>0?((totalOPEX/mIn)*100).toFixed(1):"0";

  // Daily data
  const dates = [...new Set(mCF.map(e=>e.date))].sort();
  const dailyData = dates.map(d=>{
    const dIn  = mCF.filter(e=>e.date===d&&e.flow==="in").reduce((a,b)=>a+b.amount,0);
    const dCogs= mCF.filter(e=>e.date===d&&e.flow==="out"&&INGREDIENT_GROUPS.some(g=>g.key===e.cat)).reduce((a,b)=>a+b.amount,0);
    const dOpex= mCF.filter(e=>e.date===d&&e.flow==="out"&&OPEX_CATS.includes(e.cat)).reduce((a,b)=>a+b.amount,0);
    return {d, in:dIn, cogs:dCogs, opex:dOpex, net:dIn-dCogs-dOpex};
  });
  const maxBar = Math.max(...dailyData.map(d=>d.in),1);

  // Donut segments helper (simple CSS trick with conic-gradient)
  const donutData = INGREDIENT_GROUPS.map(g=>({...g, val:cogsMap[g.key]||0})).filter(g=>g.val>0);
  const donutTotal = donutData.reduce((a,b)=>a+b.val,0)||1;
  let acc=0;
  const segments = donutData.map(g=>{
    const start=acc; acc+=g.val/donutTotal*360;
    return {...g, start, end:acc};
  });
  const conicGrad = segments.map(s=>`${s.color} ${s.start.toFixed(1)}deg ${s.end.toFixed(1)}deg`).join(", ");

  const tabBtn = (id,label) => (
    <button onClick={()=>setTab(id)} style={{ background:tab===id?"rgba(249,115,22,0.15)":"rgba(255,255,255,0.04)", border:`1px solid ${tab===id?"#f97316aa":"rgba(255,255,255,0.08)"}`, borderRadius:7, padding:"5px 13px", color:tab===id?"#f97316":"#666", cursor:"pointer", fontSize:11, fontFamily:"inherit", fontWeight:tab===id?800:400 }}>{label}</button>
  );

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:14 }}>
      {/* Header */}
      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center" }}>
        <h2 style={{ margin:0,color:"#f97316",fontSize:17,fontWeight:900 }}>📊 รายงานต้นทุน-กำไร</h2>
        <button onClick={()=>exportFullReport(stock,cf)} style={{...btn("#a78bfa"),fontSize:11,padding:"6px 11px"}}>📤 Export</button>
      </div>

      {/* Month picker */}
      <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
        {months.map(mo=>(
          <button key={mo} onClick={()=>setM(mo)} style={{ background:m===mo?"rgba(249,115,22,0.15)":"rgba(255,255,255,0.04)", border:`1px solid ${m===mo?"#f97316aa":"rgba(255,255,255,0.08)"}`, borderRadius:7, padding:"5px 12px", color:m===mo?"#f97316":"#666", cursor:"pointer", fontSize:12, fontFamily:"inherit" }}>{mo}</button>
        ))}
      </div>

      {/* KPI strip */}
      <div style={{ display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8 }}>
        {[
          ["💰","รายรับ",`฿${fmt(mIn)}`,"#4ade80"],
          ["🧾","COGS",`฿${fmt(totalCOGS)}`,"#f87171"],
          ["📈","กำไรขั้นต้น",`฿${fmt(grossP)}`,grossP>=0?"#4ade80":"#f87171"],
          ["🏁","กำไรสุทธิ",`฿${fmt(netP)}`,netP>=0?"#4ade80":"#f87171"],
        ].map(([ic,l,v,c])=>(
          <div key={l} style={{ background:"rgba(255,255,255,0.03)",border:"1px solid rgba(255,255,255,0.06)",borderRadius:11,padding:"10px 10px" }}>
            <div style={{ fontSize:16 }}>{ic}</div>
            <div style={{ color:"#555",fontSize:9,marginTop:2 }}>{l}</div>
            <div style={{ color:c,fontWeight:800,fontSize:14,lineHeight:1.2 }}>{v}</div>
          </div>
        ))}
      </div>

      {/* Tab nav */}
      <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
        {tabBtn("pl","📑 P&L")}
        {tabBtn("cogs","🥩 ต้นทุนวัตถุดิบ")}
        {tabBtn("opex","🏢 ค่าใช้จ่าย")}
        {tabBtn("daily","📅 รายวัน")}
      </div>

      {/* ── TAB: P&L Waterfall ── */}
      {tab==="pl" && (
        <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
          <Card style={{ background:"rgba(249,115,22,0.04)",borderColor:"rgba(249,115,22,0.2)" }}>
            <div style={{ color:"#f97316",fontWeight:800,fontSize:14,marginBottom:12 }}>📑 งบกำไร-ขาดทุน เดือน {m}</div>

            {/* Revenue bar */}
            <div style={{ marginBottom:14 }}>
              <div style={{ display:"flex",justifyContent:"space-between",marginBottom:4 }}>
                <span style={{ color:"#4ade80",fontWeight:700,fontSize:13 }}>💰 รายรับรวม</span>
                <span style={{ color:"#4ade80",fontWeight:800,fontSize:15 }}>฿{fmt(mIn)}</span>
              </div>
              <div style={{ background:"rgba(255,255,255,0.06)",borderRadius:6,height:12,overflow:"hidden" }}>
                <div style={{ background:"linear-gradient(90deg,#4ade80,#16a34a)",width:"100%",height:"100%",borderRadius:6 }}/>
              </div>
            </div>

            {/* COGS block */}
            <div style={{ background:"rgba(248,113,113,0.06)",border:"1px solid rgba(248,113,113,0.15)",borderRadius:10,padding:"10px 12px",marginBottom:10 }}>
              <div style={{ display:"flex",justifyContent:"space-between",marginBottom:8 }}>
                <span style={{ color:"#f87171",fontWeight:700,fontSize:13 }}>🧾 ต้นทุนวัตถุดิบ (COGS)</span>
                <span style={{ color:"#f87171",fontWeight:800 }}>฿{fmt(totalCOGS)} ({cogsPct}%)</span>
              </div>
              {INGREDIENT_GROUPS.map(g=>{
                const val = cogsMap[g.key]||0;
                const pct = mIn>0?(val/mIn*100).toFixed(1):0;
                const barPct = totalCOGS>0?(val/totalCOGS*100):0;
                return (
                  <div key={g.key} style={{ marginBottom:7 }}>
                    <div style={{ display:"flex",justifyContent:"space-between",marginBottom:3 }}>
                      <span style={{ color:"#ccc",fontSize:12 }}>{g.label}</span>
                      <div style={{ display:"flex",gap:8 }}>
                        <span style={{ color:"#777",fontSize:11 }}>{pct}% ของรายรับ</span>
                        <span style={{ color:g.color,fontWeight:700,fontSize:12 }}>฿{fmt(val)}</span>
                      </div>
                    </div>
                    <div style={{ background:"rgba(255,255,255,0.05)",borderRadius:4,height:6 }}>
                      <div style={{ background:g.color,width:`${barPct}%`,height:"100%",borderRadius:4,opacity:0.8 }}/>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Gross profit line */}
            <div style={{ display:"flex",justifyContent:"space-between",padding:"8px 0",borderTop:"1.5px solid rgba(249,115,22,0.2)",borderBottom:"1.5px solid rgba(249,115,22,0.2)",marginBottom:10 }}>
              <span style={{ color:"#fff",fontWeight:800,fontSize:14 }}>= กำไรขั้นต้น</span>
              <div style={{ textAlign:"right" }}>
                <span style={{ color:grossP>=0?"#4ade80":"#f87171",fontWeight:800,fontSize:16 }}>฿{fmt(grossP)}</span>
                <span style={{ color:"#666",fontSize:11,marginLeft:8 }}>Gross {grossM}%</span>
              </div>
            </div>

            {/* OPEX block */}
            <div style={{ background:"rgba(251,191,36,0.05)",border:"1px solid rgba(251,191,36,0.12)",borderRadius:10,padding:"10px 12px",marginBottom:10 }}>
              <div style={{ display:"flex",justifyContent:"space-between",marginBottom:6 }}>
                <span style={{ color:"#fbbf24",fontWeight:700,fontSize:13 }}>🏢 ค่าใช้จ่ายดำเนินการ (OPEX)</span>
                <span style={{ color:"#fbbf24",fontWeight:800 }}>฿{fmt(totalOPEX)} ({opexPct}%)</span>
              </div>
              {OPEX_CATS.filter(cat=>(opexActual[cat]||0)>0).map(cat=>(
                <div key={cat} style={{ display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid rgba(255,255,255,0.04)",fontSize:12 }}>
                  <span style={{ color:"#bbb" }}>{cat}</span>
                  <span style={{ color:"#fbbf24" }}>฿{fmt(opexActual[cat])}</span>
                </div>
              ))}
              {/* Fixed costs not yet recorded */}
              {FIXED_COSTS.filter(f=>!OPEX_CATS.includes(f.name)||(opexActual[f.name]||0)===0).map(f=>(
                <div key={f.name} style={{ display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid rgba(255,255,255,0.04)",fontSize:12 }}>
                  <span style={{ color:"#555" }}>{f.name} <span style={{ color:"#383838",fontSize:10 }}>(คงที่)</span></span>
                  <span style={{ color:"#555" }}>฿{fmt(f.amount)}</span>
                </div>
              ))}
            </div>

            {/* Net profit */}
            <div style={{ background:netP>=0?"rgba(74,222,128,0.08)":"rgba(248,113,113,0.08)", border:`1.5px solid ${netP>=0?"rgba(74,222,128,0.3)":"rgba(248,113,113,0.3)"}`, borderRadius:12,padding:"12px 14px" }}>
              <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                <div>
                  <div style={{ color:"#fff",fontWeight:900,fontSize:15 }}>= กำไรสุทธิ</div>
                  <div style={{ color:"#555",fontSize:11 }}>หลังหักทุกค่าใช้จ่าย</div>
                </div>
                <div style={{ textAlign:"right" }}>
                  <div style={{ color:netP>=0?"#4ade80":"#f87171",fontWeight:900,fontSize:22 }}>฿{fmt(netP)}</div>
                  <div style={{ color:"#666",fontSize:11 }}>Net Margin {netM}%</div>
                </div>
              </div>
              {/* Visual profit bar */}
              <div style={{ marginTop:10,background:"rgba(255,255,255,0.05)",borderRadius:6,height:8,overflow:"hidden" }}>
                <div style={{ background:netP>=0?"linear-gradient(90deg,#4ade80,#16a34a)":"linear-gradient(90deg,#f87171,#dc2626)", width:`${Math.min(Math.abs(netP)/mIn*100,100)}%`, height:"100%",borderRadius:6 }}/>
              </div>
            </div>

            {/* Metrics row */}
            <div style={{ display:"flex",gap:0,marginTop:12,borderTop:"1px solid rgba(255,255,255,0.05)",paddingTop:10 }}>
              {[["Gross Margin",`${grossM}%`,"#4ade80"],["Net Margin",`${netM}%`,netP>=0?"#4ade80":"#f87171"],["COGS Ratio",`${cogsPct}%`,"#f87171"],["OPEX Ratio",`${opexPct}%`,"#fbbf24"],["BEP/วัน",`฿${fmt(Math.ceil((totalOPEX||TOTAL_FIXED)/Math.max(dates.length,1)))}`,"#a78bfa"]].map(([l,v,c])=>(
                <div key={l} style={{ flex:1,textAlign:"center",padding:"4px 2px",borderRight:"1px solid rgba(255,255,255,0.05)" }}>
                  <div style={{ color:"#444",fontSize:9 }}>{l}</div>
                  <div style={{ color:c,fontWeight:800,fontSize:12 }}>{v}</div>
                </div>
              ))}
            </div>
          </Card>
        </div>
      )}

      {/* ── TAB: COGS detail ── */}
      {tab==="cogs" && (
        <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
          {/* Donut chart */}
          <Card style={{ alignItems:"center" }}>
            <div style={{ color:"#f97316",fontWeight:800,fontSize:13,marginBottom:12 }}>🥧 สัดส่วนต้นทุนวัตถุดิบ</div>
            <div style={{ display:"flex",gap:16,alignItems:"center",flexWrap:"wrap" }}>
              {/* Donut */}
              <div style={{ position:"relative",width:110,height:110,flexShrink:0 }}>
                <div style={{ width:110,height:110,borderRadius:"50%", background:conicGrad||"rgba(255,255,255,0.05)" }}/>
                <div style={{ position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)", width:58,height:58,borderRadius:"50%",background:"#111111", display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center" }}>
                  <div style={{ color:"#f97316",fontSize:9,fontWeight:700 }}>COGS</div>
                  <div style={{ color:"#fff",fontSize:12,fontWeight:800 }}>฿{fmt(totalCOGS)}</div>
                </div>
              </div>
              {/* Legend */}
              <div style={{ flex:1,display:"flex",flexDirection:"column",gap:5 }}>
                {donutData.map(g=>(
                  <div key={g.key} style={{ display:"flex",alignItems:"center",gap:7 }}>
                    <div style={{ width:10,height:10,borderRadius:2,background:g.color,flexShrink:0 }}/>
                    <span style={{ color:"#ccc",fontSize:12,flex:1 }}>{g.label}</span>
                    <span style={{ color:g.color,fontWeight:700,fontSize:12 }}>฿{fmt(g.val)}</span>
                    <span style={{ color:"#555",fontSize:10 }}>({(g.val/donutTotal*100).toFixed(0)}%)</span>
                  </div>
                ))}
              </div>
            </div>
          </Card>

          {/* Per-category deep dive */}
          {INGREDIENT_GROUPS.map(g=>{
            const val = cogsMap[g.key]||0;
            if (!val) return null;
            const pctRev = mIn>0?(val/mIn*100).toFixed(1):0;
            const pctCOGS = totalCOGS>0?(val/totalCOGS*100).toFixed(0):0;
            // items in this category from stock
            const catItems = stock.filter(s=>{
              if (g.key==="วัตถุดิบ/ผัก") return s.supplierId===1;
              if (g.key==="วัตถุดิบ/เนื้อสัตว์") return s.supplierId===2||s.supplierId===3;
              if (g.key==="วัตถุดิบ/ซอส") return s.supplierId===4;
              return false;
            });
            return (
              <Card key={g.key} style={{ borderColor:`${g.color}33`,background:`${g.bg}` }}>
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8 }}>
                  <div style={{ color:g.color,fontWeight:800,fontSize:13 }}>{g.label}</div>
                  <div style={{ textAlign:"right" }}>
                    <div style={{ color:g.color,fontWeight:800,fontSize:16 }}>฿{fmt(val)}</div>
                    <div style={{ color:"#555",fontSize:10 }}>{pctRev}% ของรายรับ • {pctCOGS}% ของ COGS</div>
                  </div>
                </div>
                {/* Progress bar */}
                <div style={{ background:"rgba(255,255,255,0.06)",borderRadius:5,height:7,marginBottom:8 }}>
                  <div style={{ background:g.color,width:`${pctCOGS}%`,height:"100%",borderRadius:5,opacity:0.7 }}/>
                </div>
                {/* CF entries grouped by itemName */}
                {(() => {
                  const entries = mCF.filter(e=>e.flow==="out"&&e.cat===g.key);
                  // group by itemName
                  const byItem = entries.reduce((acc,e)=>{
                    const k = e.itemName || e.note || "(ไม่ระบุชื่อ)";
                    if(!acc[k]) acc[k]={name:k,total:0,entries:[]};
                    acc[k].total+=e.amount; acc[k].entries.push(e);
                    return acc;
                  },{});
                  const items = Object.values(byItem).sort((a,b)=>b.total-a.total);
                  return items.map(item=>(
                    <div key={item.name} style={{ marginBottom:8 }}>
                      <div style={{ display:"flex",justifyContent:"space-between",padding:"5px 8px",background:"rgba(255,255,255,0.04)",borderRadius:7,marginTop:5 }}>
                        <span style={{ color:"#ccc",fontSize:12,fontWeight:600 }}>📌 {item.name}</span>
                        <span style={{ color:g.color,fontWeight:700,fontSize:12 }}>฿{fmt(item.total)}</span>
                      </div>
                      {item.entries.map(e=>(
                        <div key={e.id} style={{ display:"flex",justifyContent:"space-between",padding:"3px 8px 3px 20px",fontSize:11 }}>
                          <span style={{ color:"#555" }}>{e.date}</span>
                          <span style={{ color:"#666",flex:1,marginLeft:6 }}>{e.method}{e.note?` • ${e.note}`:""}</span>
                          <span style={{ color:"#888" }}>฿{fmt(e.amount)}</span>
                        </div>
                      ))}
                    </div>
                  ));
                })()}
              </Card>
            );
          })}
        </div>
      )}

      {/* ── TAB: OPEX ── */}
      {tab==="opex" && (
        <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
          <Card style={{ background:"rgba(251,191,36,0.04)",borderColor:"rgba(251,191,36,0.2)" }}>
            <div style={{ color:"#fbbf24",fontWeight:800,fontSize:14,marginBottom:12 }}>🏢 ค่าใช้จ่ายดำเนินการ</div>

            {/* Recorded OPEX */}
            <div style={{ color:"#888",fontSize:11,marginBottom:6,fontWeight:600 }}>บันทึกแล้วใน Cash Flow</div>
            {OPEX_CATS.map(cat=>{
              const val = opexActual[cat]||0;
              if (!val) return null;
              const pct = totalOPEX>0?(val/totalOPEX*100).toFixed(0):0;
              return (
                <div key={cat} style={{ marginBottom:8 }}>
                  <div style={{ display:"flex",justifyContent:"space-between",marginBottom:3 }}>
                    <span style={{ color:"#ccc",fontSize:12 }}>{cat}</span>
                    <span style={{ color:"#fbbf24",fontWeight:700,fontSize:12 }}>฿{fmt(val)} ({pct}%)</span>
                  </div>
                  <div style={{ background:"rgba(255,255,255,0.05)",borderRadius:4,height:5 }}>
                    <div style={{ background:"linear-gradient(90deg,#fbbf24,#d97706)",width:`${pct}%`,height:"100%",borderRadius:4 }}/>
                  </div>
                </div>
              );
            })}

            <div style={{ borderTop:"1px solid rgba(255,255,255,0.06)",paddingTop:10,marginTop:6 }}>
              <div style={{ color:"#444",fontSize:11,marginBottom:6,fontWeight:600 }}>ต้นทุนคงที่ประจำเดือน (รวมอยู่ในการคำนวณ)</div>
              {FIXED_COSTS.map(f=>(
                <div key={f.name} style={{ display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid rgba(255,255,255,0.03)",fontSize:12 }}>
                  <span style={{ color:"#555" }}>{f.name}</span>
                  <span style={{ color:"#555" }}>฿{fmt(f.amount)}</span>
                </div>
              ))}
              <div style={{ display:"flex",justifyContent:"space-between",paddingTop:8,fontWeight:800 }}>
                <span style={{ color:"#fbbf24" }}>รวมคงที่/เดือน</span>
                <span style={{ color:"#fbbf24",fontSize:16 }}>฿{fmt(TOTAL_FIXED)}</span>
              </div>
            </div>

            <div style={{ display:"flex",justifyContent:"space-between",marginTop:12,paddingTop:10,borderTop:"1.5px solid rgba(251,191,36,0.2)",fontWeight:900 }}>
              <span style={{ color:"#fff",fontSize:14 }}>OPEX รวม (บันทึกแล้ว)</span>
              <span style={{ color:"#fbbf24",fontSize:18 }}>฿{fmt(totalOPEX)}</span>
            </div>
          </Card>
        </div>
      )}

      {/* ── TAB: Daily breakdown ── */}
      {tab==="daily" && (
        <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
          {/* Stacked bar chart */}
          <Card>
            <div style={{ color:"#f97316",fontWeight:800,fontSize:13,marginBottom:8 }}>รายรับ vs ต้นทุน รายวัน</div>
            <div style={{ display:"flex",gap:3,marginBottom:6 }}>
              {[["#4ade80","รายรับ"],["#f87171","COGS"],["#fbbf24","OPEX"]].map(([c,l])=>(
                <div key={l} style={{ display:"flex",alignItems:"center",gap:4,marginRight:8 }}>
                  <div style={{ width:8,height:8,borderRadius:2,background:c }}/>
                  <span style={{ color:"#666",fontSize:10 }}>{l}</span>
                </div>
              ))}
            </div>
            <div style={{ display:"flex",alignItems:"flex-end",gap:4,height:90 }}>
              {dailyData.map(d=>(
                <div key={d.d} style={{ flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:1 }}>
                  <div style={{ width:"100%",display:"flex",flexDirection:"column-reverse",gap:1 }}>
                    <div style={{ width:"100%",borderRadius:"0 0 2px 2px",background:"rgba(74,222,128,0.6)", height:`${Math.max((d.in/maxBar)*55,2)}px` }}/>
                    <div style={{ width:"100%",background:"rgba(248,113,113,0.7)",height:`${Math.max((d.cogs/maxBar)*55,d.cogs?2:0)}px` }}/>
                    <div style={{ width:"100%",borderRadius:"2px 2px 0 0",background:"rgba(251,191,36,0.7)",height:`${Math.max((d.opex/maxBar)*55,d.opex?2:0)}px` }}/>
                  </div>
                  <div style={{ color:"#333",fontSize:8 }}>{d.d.slice(8)}</div>
                </div>
              ))}
            </div>
          </Card>

          {/* Daily table */}
          <Card style={{ padding:"12px 14px" }}>
            <div style={{ color:"#f97316",fontWeight:800,fontSize:13,marginBottom:8 }}>ตารางรายวัน</div>
            {/* header */}
            <div style={{ display:"grid",gridTemplateColumns:"0.8fr 1fr 1fr 1fr 1fr",gap:4,padding:"5px 0",borderBottom:"1.5px solid rgba(249,115,22,0.15)",marginBottom:4 }}>
              {["วันที่","รายรับ","COGS","OPEX","กำไร"].map(h=>(
                <div key={h} style={{ color:"#f97316",fontSize:10,fontWeight:700 }}>{h}</div>
              ))}
            </div>
            {[...dailyData].reverse().map(d=>(
              <div key={d.d} style={{ display:"grid",gridTemplateColumns:"0.8fr 1fr 1fr 1fr 1fr",gap:4,padding:"6px 0",borderBottom:"1px solid rgba(255,255,255,0.04)",fontSize:12,alignItems:"center" }}>
                <div style={{ color:"#888" }}>{d.d.slice(5)}</div>
                <div style={{ color:"#4ade80" }}>฿{fmt(d.in)}</div>
                <div style={{ color:"#f87171" }}>฿{fmt(d.cogs)}</div>
                <div style={{ color:"#fbbf24" }}>฿{fmt(d.opex)}</div>
                <div style={{ color:d.net>=0?"#4ade80":"#f87171",fontWeight:700 }}>฿{fmt(d.net)}</div>
              </div>
            ))}
            {/* Total row */}
            <div style={{ display:"grid",gridTemplateColumns:"0.8fr 1fr 1fr 1fr 1fr",gap:4,padding:"8px 0 0",marginTop:4,borderTop:"1.5px solid rgba(249,115,22,0.15)",fontSize:12 }}>
              <div style={{ color:"#f97316",fontWeight:800 }}>รวม</div>
              <div style={{ color:"#4ade80",fontWeight:800 }}>฿{fmt(mIn)}</div>
              <div style={{ color:"#f87171",fontWeight:800 }}>฿{fmt(totalCOGS)}</div>
              <div style={{ color:"#fbbf24",fontWeight:800 }}>฿{fmt(totalOPEX)}</div>
              <div style={{ color:netP>=0?"#4ade80":"#f87171",fontWeight:800 }}>฿{fmt(netP)}</div>
            </div>
          </Card>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// PURCHASE
// ─────────────────────────────────────────────────────────────
function PurchasePage({ stock }) {
  const [sel, setSel] = useState({});
  const [sent, setSent] = useState(false);
  const need = stock.filter(s=>["critical","out","low"].includes(stockStatus(s)));
  const selItems = need.filter(s=>sel[s.id]);
  const bySupp = selItems.reduce((acc,s)=>{ const su=SUPPLIERS.find(x=>x.id===s.supplierId); if(!acc[su.id])acc[su.id]={su,items:[]}; acc[su.id].items.push(s); return acc; },{});
  const totalCost = selItems.reduce((a,s)=>a+Math.max(s.minQty*2-s.qty,s.minQty)*weightedAvgCost(s),0);

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:14 }}>
      <Hdr title="🛒 ใบสั่งซื้อ" />
      <Card>
        <div style={{ color:"#f97316",fontWeight:800,fontSize:13,marginBottom:10 }}>รายการที่ต้องสั่ง ({need.length})</div>
        {need.length===0 ? <div style={{ color:"#555",textAlign:"center",padding:20 }}>✅ สต็อคปกติทั้งหมด</div> : need.map(s=>{
          const su = SUPPLIERS.find(x=>x.id===s.supplierId);
          const qty = Math.max(s.minQty*2-s.qty,s.minQty);
          const avgC = weightedAvgCost(s);
          return (
            <div key={s.id} style={{ display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:"1px solid rgba(255,255,255,0.04)" }}>
              <input type="checkbox" checked={!!sel[s.id]} onChange={()=>setSel(p=>({...p,[s.id]:!p[s.id]}))} style={{ width:16,height:16,accentColor:"#f97316",cursor:"pointer",flexShrink:0 }} />
              <div style={{ flex:1 }}>
                <div style={{ color:"#ddd",fontSize:13,display:"flex",alignItems:"center",gap:6 }}>
                  {s.name} {isCostSpike(s)&&<span title="ราคาสูงผิดปกติ" style={{ color:"#f87171",fontSize:11 }}>⚠️ ราคาผิดปกติ</span>}
                </div>
                <div style={{ color:"#555",fontSize:10 }}>เหลือ {s.qty} {s.unit} • สั่ง {qty} {s.unit} • ≈฿{fmt(Math.round(qty*avgC))} • {su?.name}</div>
              </div>
              <Tag status={stockStatus(s)} />
            </div>
          );
        })}
      </Card>
      {selItems.length>0 && (
        <Card style={{ borderColor:"rgba(74,222,128,0.2)" }}>
          <div style={{ color:"#4ade80",fontWeight:800,fontSize:13,marginBottom:10 }}>📋 ใบสั่งซื้อ — {selItems.length} รายการ • ≈฿{fmt(Math.round(totalCost))}</div>
          {Object.values(bySupp).map(({su,items})=>(
            <div key={su.id} style={{ marginBottom:10 }}>
              <div style={{ color:"#f97316",fontWeight:700,fontSize:12,marginBottom:4 }}>{su.name}</div>
              {items.map(i=>{const q=Math.max(i.minQty*2-i.qty,i.minQty); return <div key={i.id} style={{ color:"#bbb",fontSize:12,padding:"2px 0 2px 10px" }}>• {i.name} {q} {i.unit} (฿{fmt(Math.round(q*weightedAvgCost(i)))})</div>;})}
            </div>
          ))}
        </Card>
      )}
      <button onClick={()=>{setSent(true);setTimeout(()=>setSent(false),3000);}} disabled={selItems.length===0} style={{...btn(selItems.length>0?"#22c55e":"#333"),width:"100%",padding:13,fontSize:15,opacity:selItems.length===0?0.4:1}}>
        {sent?"✅ ส่ง LINE เรียบร้อย!":`📲 ส่งใบสั่งซื้อผ่าน LINE (${selItems.length} รายการ)`}
      </button>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// ADMIN PANEL — แก้ไขข้อมูลหลังบ้าน
// ─────────────────────────────────────────────────────────────
function AdminPage({ stock, setStock, cf, setCF }) {
  const [tab, setTab] = useState("cf"); // cf | stock | fixed | line
  const [editId, setEditId] = useState(null);
  const [editData, setEditData] = useState({});

  // ── CF editor ──
  const startEditCF = (e) => { setEditId(e.id); setEditData({...e}); };
  const saveCF = () => {
    setCF(prev => prev.map(e => e.id===editId ? {...editData, amount:+editData.amount} : e));
    setEditId(null);
  };
  const delCF = (id) => { if(window.confirm("ลบรายการนี้?")) setCF(prev=>prev.filter(e=>e.id!==id)); };

  // ── Stock editor ──
  const startEditStock = (s) => { setEditId("s_"+s.id); setEditData({...s}); };
  const saveStock = () => {
    setStock(prev => prev.map(s => s.id===editData.id ? {...editData, qty:+editData.qty, minQty:+editData.minQty, dailyUse:+editData.dailyUse, expiryDays:+editData.expiryDays} : s));
    setEditId(null);
  };
  const delStock = (id) => { if(window.confirm("ลบวัตถุดิบนี้?")) setStock(prev=>prev.filter(s=>s.id!==id)); };

  const tabBtn = (id, label) => (
    <button onClick={()=>{setTab(id);setEditId(null);}} style={{
      background: tab===id ? "rgba(249,115,22,0.18)" : "rgba(255,255,255,0.04)",
      border: `1px solid ${tab===id ? "#f97316aa" : "rgba(255,255,255,0.08)"}`,
      borderRadius:7, padding:"6px 13px", color:tab===id ? "#f97316":"#777",
      cursor:"pointer", fontSize:12, fontFamily:"inherit", fontWeight: tab===id?700:400
    }}>{label}</button>
  );

  const rowStyle = { display:"flex", alignItems:"flex-start", gap:8, padding:"8px 0", borderBottom:"1px solid rgba(255,255,255,0.05)", fontSize:12 };
  const cellStyle = (w) => ({ width:w, flexShrink:0, color:"#ccc" });
  const editInp = { ...inp, padding:"4px 8px", fontSize:12 };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <Hdr title="⚙️ Admin — แก้ไขข้อมูล" />

      <div style={{ background:"rgba(249,115,22,0.06)", border:"1px solid rgba(249,115,22,0.2)", borderRadius:10, padding:"10px 14px" }}>
        <div style={{ color:"#f97316", fontWeight:700, fontSize:12 }}>🔐 เฉพาะเจ้าของเท่านั้น</div>
        <div style={{ color:"#666", fontSize:11, marginTop:2 }}>แก้ไข/ลบข้อมูลทุกรายการ รวมถึงตั้งค่า LINE Token</div>
      </div>

      <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
        {tabBtn("cf","💵 Cash Flow")}
        {tabBtn("stock","📦 สต็อค")}
        {tabBtn("line","📲 LINE Setup")}
      </div>

      {/* ── Cash Flow Editor ── */}
      {tab==="cf" && (
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          <div style={{ color:"#888", fontSize:11 }}>รายการทั้งหมด {cf.length} รายการ • คลิก ✏️ เพื่อแก้ไข</div>
          {[...cf].sort((a,b)=>b.date.localeCompare(a.date)).map(e => {
            const isEdit = editId===e.id;
            return (
              <Card key={e.id} style={{ padding:"10px 14px", borderColor: isEdit?"rgba(249,115,22,0.4)":"rgba(255,255,255,0.06)" }}>
                {isEdit ? (
                  <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                    <div style={{ color:"#f97316", fontWeight:700, fontSize:12, marginBottom:4 }}>✏️ แก้ไขรายการ</div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
                      {[["วันที่","date","date"],["จำนวน (฿)","amount","number"],["ชื่อรายการ","itemName","text"],["หมายเหตุ","note","text"]].map(([l,k,t])=>(
                        <div key={k}>
                          <div style={{ color:"#666", fontSize:10, marginBottom:2 }}>{l}</div>
                          <input type={t} value={editData[k]||""} onChange={e2=>setEditData(p=>({...p,[k]:e2.target.value}))} style={editInp} />
                        </div>
                      ))}
                      <div>
                        <div style={{ color:"#666", fontSize:10, marginBottom:2 }}>หมวดหมู่</div>
                        <select value={editData.cat||""} onChange={e2=>setEditData(p=>({...p,cat:e2.target.value}))} style={editInp}>
                          {[...IN_CATS,...OUT_CATS].map(c=><option key={c}>{c}</option>)}
                        </select>
                      </div>
                      <div>
                        <div style={{ color:"#666", fontSize:10, marginBottom:2 }}>ประเภท</div>
                        <select value={editData.flow||"in"} onChange={e2=>setEditData(p=>({...p,flow:e2.target.value}))} style={editInp}>
                          <option value="in">รายรับ</option>
                          <option value="out">รายจ่าย</option>
                        </select>
                      </div>
                    </div>
                    <div style={{ display:"flex", gap:8 }}>
                      <button onClick={saveCF} style={{...btn("#f97316"), padding:"7px 16px"}}>💾 บันทึก</button>
                      <button onClick={()=>setEditId(null)} style={{...ghost, padding:"7px 12px"}}>ยกเลิก</button>
                    </div>
                  </div>
                ) : (
                  <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                    <div style={{ width:6, height:6, borderRadius:"50%", background:e.flow==="in"?"#4ade80":"#f87171", flexShrink:0 }}/>
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ color:"#ddd", fontSize:12 }}>{e.itemName||e.cat} <span style={{ color:"#555", fontSize:10 }}>• {e.cat} • {e.date}</span></div>
                      <div style={{ color:e.flow==="in"?"#4ade80":"#f87171", fontWeight:700, fontSize:13 }}>{e.flow==="in"?"+":"-"}฿{fmt(e.amount)}</div>
                    </div>
                    <button onClick={()=>startEditCF(e)} style={{ background:"rgba(249,115,22,0.12)", border:"1px solid rgba(249,115,22,0.3)", borderRadius:6, padding:"3px 9px", color:"#f97316", cursor:"pointer", fontSize:11, fontFamily:"inherit" }}>✏️</button>
                    <button onClick={()=>delCF(e.id)} style={{ background:"rgba(248,113,113,0.1)", border:"1px solid rgba(248,113,113,0.25)", borderRadius:6, padding:"3px 9px", color:"#f87171", cursor:"pointer", fontSize:11, fontFamily:"inherit" }}>🗑</button>
                  </div>
                )}
              </Card>
            );
          })}
        </div>
      )}

      {/* ── Stock Editor ── */}
      {tab==="stock" && (
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          <div style={{ color:"#888", fontSize:11 }}>วัตถุดิบทั้งหมด {stock.length} รายการ</div>
          {stock.map(s => {
            const isEdit = editId==="s_"+s.id;
            return (
              <Card key={s.id} style={{ padding:"10px 14px", borderColor: isEdit?"rgba(249,115,22,0.4)":"rgba(255,255,255,0.06)" }}>
                {isEdit ? (
                  <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                    <div style={{ color:"#f97316", fontWeight:700, fontSize:12, marginBottom:4 }}>✏️ แก้ไข: {editData.name}</div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
                      {[["ชื่อ","name","text"],["หน่วย","unit","text"],["คงเหลือ","qty","number"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"],["อายุ(วัน)","expiryDays","number"]].map(([l,k,t])=>(
                        <div key={k}>
                          <div style={{ color:"#666", fontSize:10, marginBottom:2 }}>{l}</div>
                          <input type={t} value={editData[k]||""} onChange={e2=>setEditData(p=>({...p,[k]:e2.target.value}))} style={editInp} />
                        </div>
                      ))}
                    </div>
                    <div style={{ display:"flex", gap:8 }}>
                      <button onClick={saveStock} style={{...btn("#f97316"), padding:"7px 16px"}}>💾 บันทึก</button>
                      <button onClick={()=>setEditId(null)} style={{...ghost, padding:"7px 12px"}}>ยกเลิก</button>
                    </div>
                  </div>
                ) : (
                  <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                    <Tag status={stockStatus(s)} />
                    <div style={{ flex:1 }}>
                      <div style={{ color:"#ddd", fontSize:12, fontWeight:600 }}>{s.name}</div>
                      <div style={{ color:"#666", fontSize:10 }}>{s.qty} {s.unit} • ขั้นต่ำ {s.minQty} • ใช้ {s.dailyUse}/วัน</div>
                    </div>
                    <button onClick={()=>startEditStock(s)} style={{ background:"rgba(249,115,22,0.12)", border:"1px solid rgba(249,115,22,0.3)", borderRadius:6, padding:"3px 9px", color:"#f97316", cursor:"pointer", fontSize:11, fontFamily:"inherit" }}>✏️</button>
                    <button onClick={()=>delStock(s.id)} style={{ background:"rgba(248,113,113,0.1)", border:"1px solid rgba(248,113,113,0.25)", borderRadius:6, padding:"3px 9px", color:"#f87171", cursor:"pointer", fontSize:11, fontFamily:"inherit" }}>🗑</button>
                  </div>
                )}
              </Card>
            );
          })}
        </div>
      )}

      {/* ── LINE Setup ── */}
      {tab==="line" && (
        <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
          <Card style={{ background:"rgba(34,197,94,0.05)", borderColor:"rgba(34,197,94,0.2)" }}>
            <div style={{ color:"#22c55e", fontWeight:800, fontSize:14, marginBottom:10 }}>📲 วิธีเชื่อม LINE จริง</div>
            <div style={{ color:"#888", fontSize:12, lineHeight:1.9 }}>
              ปัจจุบันปุ่ม "ส่ง LINE" เป็นการจำลองเท่านั้น เพราะ LINE Notify ต้องส่งผ่าน server จริงถึงจะทำงานได้ (browser ไม่อนุญาตโดยตรง)
            </div>
            <div style={{ marginTop:12, display:"flex", flexDirection:"column", gap:10 }}>
              {[
                ["1️⃣ สมัคร LINE Notify", "เปิด notify-bot.line.me → Login → \"Generate token\" → ตั้งชื่อ เช่น 'ร้านไท่กั๋ว' → Copy Token", "#f97316"],
                ["2️⃣ สร้าง Backend (ฟรี)", "ใช้ Make.com หรือ n8n.io รับ webhook จากแอป แล้วส่งต่อไป LINE Notify ด้วย Token ที่ได้", "#a78bfa"],
                ["3️⃣ ใส่ Webhook URL", "นำ URL ที่ได้จาก Make/n8n มาใส่ในช่องด้านล่าง แอปจะ POST ข้อมูลไปให้อัตโนมัติ", "#38bdf8"],
                ["4️⃣ ทดสอบ", "กดปุ่ม 'ทดสอบส่ง LINE' เพื่อตรวจสอบว่าข้อความถึง LINE ของคุณหรือไม่", "#4ade80"],
              ].map(([step, desc, c]) => (
                <div key={step} style={{ background:"rgba(255,255,255,0.03)", border:`1px solid ${c}22`, borderRadius:9, padding:"10px 12px" }}>
                  <div style={{ color:c, fontWeight:700, fontSize:12, marginBottom:3 }}>{step}</div>
                  <div style={{ color:"#777", fontSize:11, lineHeight:1.6 }}>{desc}</div>
                </div>
              ))}
            </div>
          </Card>

          <Card>
            <div style={{ color:"#f97316", fontWeight:700, fontSize:13, marginBottom:10 }}>🔧 ตั้งค่า Webhook URL</div>
            <div style={{ color:"#666", fontSize:11, marginBottom:6 }}>URL จาก Make.com / n8n (สำหรับ deploy จริง)</div>
            <input type="url" style={inp} placeholder="https://hook.make.com/xxxxx หรือ https://your-n8n.com/webhook/xxxx" />
            <div style={{ color:"#666", fontSize:11, marginTop:8, marginBottom:6 }}>LINE Notify Token (ใส่ใน server ห้ามใส่ใน browser)</div>
            <input type="password" style={inp} placeholder="LINE Notify Token (เก็บเป็นความลับ)" />
            <div style={{ display:"flex", gap:8, marginTop:12 }}>
              <button style={{...btn(), flex:1}}>💾 บันทึก URL</button>
              <button style={{...ghost}}>🧪 ทดสอบส่ง</button>
            </div>
          </Card>

          <Card style={{ background:"rgba(249,115,22,0.04)", borderColor:"rgba(249,115,22,0.15)" }}>
            <div style={{ color:"#f97316", fontWeight:700, fontSize:13, marginBottom:8 }}>🔔 กำหนดการแจ้งเตือน LINE</div>
            {[
              ["สต็อคต่ำกว่าขั้นต่ำ", "แจ้งทันที", true],
              ["สรุปยอดขายประจำวัน", "23:00 น.", true],
              ["ราคาวัตถุดิบผิดปกติ", "แจ้งทันที", true],
              ["ใบสั่งซื้อ", "เมื่อกดส่ง", true],
              ["รายงานกำไรรายสัปดาห์", "จันทร์ 08:00", false],
            ].map(([name, time, on]) => (
              <div key={name} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"7px 0", borderBottom:"1px solid rgba(255,255,255,0.04)" }}>
                <div>
                  <div style={{ color:"#ddd", fontSize:12 }}>{name}</div>
                  <div style={{ color:"#555", fontSize:10 }}>{time}</div>
                </div>
                <div style={{ width:36, height:20, borderRadius:10, background:on?"rgba(34,197,94,0.3)":"rgba(255,255,255,0.08)", border:`1px solid ${on?"#22c55e44":"rgba(255,255,255,0.1)"}`, display:"flex", alignItems:"center", padding:"2px", cursor:"pointer" }}>
                  <div style={{ width:14, height:14, borderRadius:"50%", background:on?"#22c55e":"#555", transform:on?"translateX(16px)":"translateX(0)", transition:"transform .2s" }}/>
                </div>
              </div>
            ))}
          </Card>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────
export default function App() {
  const [role, setRole] = useState(null);
  const [page, setPage] = useState("dashboard");
  const [stock, setStock] = useState(INIT_STOCK);
  const [cf, setCF] = useState(INIT_CF);

  if (!role) return <LoginPage onLogin={r=>{ setRole(r); setPage("dashboard"); }} />;

  const ownerNav = [
    { id:"dashboard", icon:"🏠", label:"หลัก" },
    { id:"cashflow", icon:"💵", label:"Cash Flow" },
    { id:"stock", icon:"📦", label:"สต็อค" },
    { id:"report", icon:"📊", label:"รายงาน" },
    { id:"purchase", icon:"🛒", label:"สั่งซื้อ" },
    { id:"admin", icon:"⚙️", label:"Admin" },
  ];
  const staffNav = [
    { id:"dashboard", icon:"🏠", label:"หลัก" },
    { id:"cashflow", icon:"💵", label:"Cash Flow" },
    { id:"stock", icon:"📦", label:"สต็อค" },
    { id:"purchase", icon:"🛒", label:"สั่งซื้อ" },
  ];
  const nav = role==="owner" ? ownerNav : staffNav;

  const pages = {
    dashboard: <Dashboard stock={stock} cf={cf} />,
    cost: <CostDashboard stock={stock} setStock={setStock} />,
    cashflow: <CashflowPage cf={cf} setCF={setCF} role={role} />,
    stock: <StockPage stock={stock} setStock={setStock} role={role} />,
    purchase: <PurchasePage stock={stock} />,
    report: <ReportPage cf={cf} stock={stock} />,
    admin: <AdminPage stock={stock} setStock={setStock} cf={cf} setCF={setCF} />,
  };

  const spikes = stock.filter(isCostSpike).length;
  const critical = stock.filter(s=>["critical","out"].includes(stockStatus(s))).length;

  return (
    <div style={{ minHeight:"100vh", background:"#111111", fontFamily:"'Noto Sans Thai','Noto Sans',sans-serif", color:"#fff", fontSize:14 }}>
      {/* Header */}
      <div style={{
        background:"rgba(20,20,20,0.95)", borderBottom:"1px solid rgba(249,115,22,0.2)",
        padding:"10px 14px", display:"flex", alignItems:"center", gap:10,
        position:"sticky", top:0, zIndex:50, backdropFilter:"blur(12px)"
      }}>
        <div style={{ width:32, height:32, borderRadius:9, background:"linear-gradient(135deg,#f97316,#ea580c)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:16 }}>🫕</div>
        <div>
          <div style={{ color:"#f97316", fontWeight:900, fontSize:14, letterSpacing:.3 }}>ไท่กั๋วหม่าล่า</div>
          <div style={{ color:"#555", fontSize:9 }}>{role==="owner"?"👑 เจ้าของ":"👷 พนักงาน"} • {todayStr()}</div>
        </div>
        <div style={{ display:"flex", gap:5, marginLeft:8 }}>
          {spikes>0 && (
            <span style={{ background:"rgba(248,113,113,0.15)", border:"1px solid rgba(248,113,113,0.3)", borderRadius:10, padding:"2px 7px", fontSize:10, color:"#f87171", fontWeight:700 }}>
              ⚠️ {spikes}
            </span>
          )}
          {critical>0 && (
            <span style={{ background:"rgba(239,68,68,0.15)", border:"1px solid rgba(239,68,68,0.3)", borderRadius:10, padding:"2px 7px", fontSize:10, color:"#ef4444", fontWeight:700 }}>
              🚨 {critical}
            </span>
          )}
        </div>
        <button onClick={()=>setRole(null)} style={{ marginLeft:"auto", background:"rgba(255,255,255,0.05)", border:"1px solid rgba(255,255,255,0.1)", borderRadius:7, padding:"4px 10px", color:"#666", cursor:"pointer", fontSize:10, fontFamily:"inherit" }}>
          ออกจากระบบ
        </button>
      </div>

      {/* Content */}
      <div style={{ padding:"16px 13px 90px", maxWidth:680, margin:"0 auto" }}>
        {pages[page] || pages["dashboard"]}
      </div>

      {/* Bottom nav */}
      <div style={{
        position:"fixed", bottom:0, left:0, right:0,
        background:"rgba(15,15,15,0.97)", backdropFilter:"blur(20px)",
        borderTop:"1px solid rgba(249,115,22,0.15)",
        display:"flex", justifyContent:"space-around",
        padding:"8px 0 14px", zIndex:100
      }}>
        {nav.map(n=>(
          <button key={n.id} onClick={()=>setPage(n.id)} style={{
            background:"none", border:"none", cursor:"pointer",
            display:"flex", flexDirection:"column", alignItems:"center", gap:2,
            padding:"2px 6px", minWidth:44,
            opacity: page===n.id ? 1 : 0.35,
            transition:"opacity .2s",
          }}>
            <span style={{ fontSize:19 }}>{n.icon}</span>
            <span style={{ fontSize:9, color:page===n.id?"#f97316":"#777", fontWeight:page===n.id?800:400 }}>{n.label}</span>
            {page===n.id && <div style={{ width:14, height:2, borderRadius:2, background:"#f97316" }} />}
          </button>
        ))}
      </div>
    </div>
  );
}
