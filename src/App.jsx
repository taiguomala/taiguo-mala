import { useState, useRef, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────
// THEME — Light / Orange / Gray
// ─────────────────────────────────────────────
const T = {
  bg:       "#f4f4f5",
  card:     "#ffffff",
  border:   "#e4e4e7",
  borderOr: "#fed7aa",
  orange:   "#f97316",
  orangeDk: "#ea580c",
  orangeLt: "#fff7ed",
  text:     "#18181b",
  textMd:   "#3f3f46",
  textSm:   "#71717a",
  textXs:   "#a1a1aa",
  green:    "#16a34a",
  greenLt:  "#f0fdf4",
  red:      "#dc2626",
  redLt:    "#fef2f2",
  yellow:   "#d97706",
  yellowLt: "#fffbeb",
  blue:     "#2563eb",
  blueLt:   "#eff6ff",
  shadow:   "0 1px 3px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.04)",
  shadowMd: "0 4px 6px rgba(0,0,0,0.07), 0 2px 4px rgba(0,0,0,0.05)",
};

// ─────────────────────────────────────────────
// STAFF ACCOUNTS
// ─────────────────────────────────────────────
const INIT_STAFF = [
  { id: "owner", name: "เจ้าของ (DR.Fresh)", pin: "1234", role: "owner", branch: "all", active: true,
    perms: { cashflow: true, stock: true, purchase: true, report: true, cost: true, admin: true } },
  { id: "s1", name: "น้องมิ้ว (สาขา 1)", pin: "1111", role: "staff", branch: "1", active: true,
    perms: { cashflow: true, stock: true, purchase: true, report: false, cost: false, admin: false } },
  { id: "s2", name: "น้องปาล์ม (สาขา 2)", pin: "2222", role: "staff", branch: "2", active: true,
    perms: { cashflow: true, stock: true, purchase: false, report: false, cost: false, admin: false } },
  { id: "s3", name: "น้องเจ (สาขา 3)", pin: "3333", role: "staff", branch: "3", active: true,
    perms: { cashflow: false, stock: true, purchase: false, report: false, cost: false, admin: false } },
  { id: "s4", name: "น้องบิ๊ก (สาขา 4)", pin: "4444", role: "staff", branch: "4", active: false,
    perms: { cashflow: true, stock: true, purchase: true, report: false, cost: false, admin: false } },
];

const BRANCHES = [
  { id: "1", name: "สาขา 1 — ใจกลางเมือง" },
  { id: "2", name: "สาขา 2 — ห้างเซ็นทรัล" },
  { id: "3", name: "สาขา 3 — ตลาดสด" },
  { id: "4", name: "สาขา 4 — ออนไลน์" },
];

// ─────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────
const FIXED_COSTS = [
  { name: "ค่าเช่า", amount: 4500 },
  { name: "ค่าพนักงาน", amount: 35000 },
  { name: "ค่าไฟ", amount: 8000 },
  { name: "เครื่องล้างจาน", amount: 4000 },
  { name: "ค่าน้ำ", amount: 1000 },
  { name: "อื่นๆ", amount: 1000 },
];
const TOTAL_FIXED = FIXED_COSTS.reduce((a, b) => a + b.amount, 0);

const INGREDIENT_GROUPS = [
  { key: "วัตถุดิบ/ผัก",         label: "🥬 ผัก",         color: T.green,  bg: T.greenLt  },
  { key: "วัตถุดิบ/เนื้อสัตว์",  label: "🥩 เนื้อสัตว์",  color: T.red,    bg: T.redLt    },
  { key: "วัตถุดิบ/ลูกชิ้น",     label: "🍢 ลูกชิ้น",     color: "#ea580c",bg: "#fff7ed"  },
  { key: "วัตถุดิบ/เครื่องดื่ม", label: "🧋 เครื่องดื่ม", color: T.blue,   bg: T.blueLt   },
  { key: "วัตถุดิบ/ซอส",          label: "🌶️ ซอส",         color: "#be123c",bg: "#fff1f2"  },
  { key: "วัตถุดิบ/บรรจุภัณฑ์",  label: "📦 บรรจุภัณฑ์",  color: "#7c3aed",bg: "#f5f3ff"  },
];
const OPEX_CATS = ["ค่าแรงพนักงาน","ค่าเช่า","ค่าไฟ/น้ำ","ค่าแก๊ส","ค่าการตลาด","ค่าซ่อมบำรุง","จ่ายชำระหนี้","อื่นๆ"];
const IN_CATS  = ["ยอดขาย dine-in","ยอดขาย delivery","รับชำระหนี้","เงินทุนเพิ่ม","รายได้อื่นๆ"];
const OUT_CATS = [...INGREDIENT_GROUPS.map(g=>g.key), ...OPEX_CATS];
const PAY_METHODS = ["เงินสด","โอนธนาคาร","QR Code","บัตรเครดิต","GrabFood","LINE MAN"];

// ─────────────────────────────────────────────
// SEED STOCK + MOVEMENTS
// ─────────────────────────────────────────────
const INIT_STOCK = [
  { id:1, name:"หมูสามชั้น",   unit:"kg",    qty:8,   minQty:5,  dailyUse:4,   supplierId:2, expiryDays:3,
    costHistory:[{date:"2026-05-08",unitCost:180,qty:20,total:3600}] },
  { id:2, name:"กุ้งแวนนาไม",  unit:"kg",    qty:3,   minQty:4,  dailyUse:2,   supplierId:3, expiryDays:2,
    costHistory:[{date:"2026-05-05",unitCost:280,qty:10,total:2800}] },
  { id:3, name:"ผักกาดขาว",    unit:"kg",    qty:2,   minQty:3,  dailyUse:3,   supplierId:1, expiryDays:2,
    costHistory:[{date:"2026-05-09",unitCost:32, qty:8, total:256}] },
  { id:4, name:"เห็ดหอม",      unit:"kg",    qty:4,   minQty:2,  dailyUse:1.5, supplierId:1, expiryDays:3,
    costHistory:[{date:"2026-05-08",unitCost:92, qty:5, total:460}] },
  { id:5, name:"ซอสเบสหม่าล่า",unit:"ถุง",   qty:12,  minQty:6,  dailyUse:2,   supplierId:4, expiryDays:30,
    costHistory:[{date:"2026-05-06",unitCost:220,qty:20,total:4400}] },
  { id:6, name:"ลูกชิ้นหมู",   unit:"kg",    qty:5,   minQty:3,  dailyUse:2,   supplierId:2, expiryDays:4,
    costHistory:[{date:"2026-05-07",unitCost:95, qty:10,total:950}] },
  { id:7, name:"น้ำอัดลม",     unit:"ขวด",   qty:24,  minQty:12, dailyUse:8,   supplierId:1, expiryDays:90,
    costHistory:[{date:"2026-05-01",unitCost:15, qty:48,total:720}] },
  { id:8, name:"บรรจุภัณฑ์",   unit:"ชิ้น",  qty:150, minQty:50, dailyUse:30,  supplierId:4, expiryDays:365,
    costHistory:[{date:"2026-05-01",unitCost:4,  qty:200,total:800}] },
];

// Stock movements log
const INIT_MOVEMENTS = [
  { id:1, itemId:1, type:"in",  qty:20,  date:"2026-05-08", staffId:"s1", note:"รับจากซัพฯ", branch:"1" },
  { id:2, itemId:3, type:"out", qty:3,   date:"2026-05-09", staffId:"s2", note:"ใช้วันนี้",   branch:"2" },
  { id:3, itemId:5, type:"in",  qty:8,   date:"2026-05-06", staffId:"s1", note:"สั่งเพิ่ม",   branch:"1" },
  { id:4, itemId:2, type:"out", qty:2,   date:"2026-05-09", staffId:"s3", note:"ใช้บุฟเฟต์",  branch:"3" },
];

const INIT_CF = [
  { id:1, date:"2026-05-09", flow:"in",  cat:"ยอดขาย dine-in",     itemName:"58 ออร์เดอร์", amount:9200,  method:"เงินสด", note:"", branch:"1", staffId:"s1" },
  { id:2, date:"2026-05-09", flow:"in",  cat:"ยอดขาย delivery",     itemName:"GrabFood",      amount:3800,  method:"โอน",    note:"", branch:"1", staffId:"s1" },
  { id:3, date:"2026-05-09", flow:"out", cat:"วัตถุดิบ/ผัก",        itemName:"ผักกาดขาว",    amount:1400,  method:"เงินสด", note:"", branch:"1", staffId:"s1" },
  { id:4, date:"2026-05-08", flow:"in",  cat:"ยอดขาย dine-in",     itemName:"52 ออร์เดอร์", amount:8500,  method:"เงินสด", note:"", branch:"2", staffId:"s2" },
  { id:5, date:"2026-05-08", flow:"out", cat:"วัตถุดิบ/เนื้อสัตว์", itemName:"หมูสามชั้น",   amount:3600,  method:"โอน",    note:"", branch:"2", staffId:"s2" },
  { id:6, date:"2026-05-07", flow:"in",  cat:"ยอดขาย dine-in",     itemName:"48 ออร์เดอร์", amount:7800,  method:"เงินสด", note:"", branch:"1", staffId:"s1" },
  { id:7, date:"2026-05-07", flow:"out", cat:"วัตถุดิบ/ลูกชิ้น",   itemName:"ลูกชิ้นหมู",   amount:950,   method:"เงินสด", note:"", branch:"3", staffId:"s3" },
  { id:8, date:"2026-05-01", flow:"out", cat:"ค่าแรงพนักงาน",       itemName:"เงินเดือน",     amount:35000, method:"โอน",    note:"", branch:"1", staffId:"owner" },
  { id:9, date:"2026-05-01", flow:"out", cat:"ค่าเช่า",             itemName:"ค่าเช่ารายเดือน",amount:4500,method:"โอน",    note:"", branch:"1", staffId:"owner" },
];

// ─────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────
const fmt   = n => Number(n||0).toLocaleString("th-TH");
const today = () => new Date().toISOString().split("T")[0];

function wac(item) {
  const h = item.costHistory||[];
  const tq = h.reduce((a,b)=>a+b.qty,0);
  const tc = h.reduce((a,b)=>a+b.total,0);
  return tq>0 ? tc/tq : 0;
}
function latestCost(item) {
  const h = item.costHistory||[];
  if (!h.length) return 0;
  return [...h].sort((a,b)=>b.date.localeCompare(a.date))[0].unitCost;
}
function isSpike(item) {
  const avg=wac(item), latest=latestCost(item);
  return avg>0 && latest>avg*1.15;
}
function stockSt(item) {
  const d = item.dailyUse>0 ? item.qty/item.dailyUse : 99;
  if (item.qty<=0) return "out";
  if (item.qty<=item.minQty||d<=1) return "critical";
  if (item.qty<=item.minQty*1.5||d<=2) return "low";
  return "ok";
}
const ST_COLOR = {ok:T.green, low:T.yellow, critical:T.red, out:T.red};
const ST_BG    = {ok:T.greenLt, low:T.yellowLt, critical:T.redLt, out:T.redLt};
const ST_LABEL = {ok:"ปกติ", low:"ใกล้หมด", critical:"น้อยมาก", out:"หมด"};

// ─────────────────────────────────────────────
// SHARED UI
// ─────────────────────────────────────────────
const S = {
  card: { background:T.card, border:`1px solid ${T.border}`, borderRadius:16, padding:"20px 22px", boxShadow:T.shadow },
  inp:  { background:T.card, border:`1px solid ${T.border}`, borderRadius:10, padding:"10px 14px", color:T.text, fontSize:17, width:"100%", boxSizing:"border-box", outline:"none", fontFamily:"inherit" },
  btn:  (bg=T.orange,fg="#fff") => ({ background:bg, border:"none", borderRadius:10, padding:"10px 20px", color:fg, fontWeight:700, cursor:"pointer", fontSize:16, fontFamily:"inherit" }),
  ghost:{ background:"transparent", border:`1px solid ${T.border}`, borderRadius:10, padding:"10px 16px", color:T.textMd, cursor:"pointer", fontSize:16, fontFamily:"inherit" },
};

function Card({ children, style={} }) {
  return <div style={{...S.card,...style}}>{children}</div>;
}

function Badge({ status }) {
  return (
    <span style={{ display:"inline-flex",alignItems:"center",gap:5, background:ST_BG[status], color:ST_COLOR[status], border:`1px solid ${ST_COLOR[status]}44`, borderRadius:8, padding:"3px 10px", fontSize:13, fontWeight:700 }}>
      <span style={{ width:7,height:7,borderRadius:"50%",background:ST_COLOR[status] }}/>
      {ST_LABEL[status]}
    </span>
  );
}

function SectionTitle({ title, action }) {
  return (
    <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18 }}>
      <h2 style={{ margin:0,color:T.text,fontSize:22,fontWeight:800 }}>{title}</h2>
      {action}
    </div>
  );
}

function TabBar({ tabs, active, onChange }) {
  return (
    <div style={{ display:"flex",gap:6,flexWrap:"wrap",marginBottom:16 }}>
      {tabs.map(([id,label])=>(
        <button key={id} onClick={()=>onChange(id)} style={{
          background:active===id?T.orange:"transparent",
          border:`1px solid ${active===id?T.orange:T.border}`,
          borderRadius:8, padding:"8px 16px",
          color:active===id?"#fff":T.textMd,
          cursor:"pointer", fontSize:15, fontFamily:"inherit", fontWeight:active===id?700:400,
          transition:"all .15s"
        }}>{label}</button>
      ))}
    </div>
  );
}

// ─────────────────────────────────────────────
// NOTIFICATIONS (in-app bot)
// ─────────────────────────────────────────────
function buildNotifications(stock, cf, movements, staff) {
  const notes = [];
  // Stock alerts
  stock.filter(s=>["critical","out"].includes(stockSt(s))).forEach(s=>{
    notes.push({ id:`st_${s.id}`, type:"danger", icon:"🚨", title:`สต็อควิกฤต: ${s.name}`, body:`เหลือ ${s.qty} ${s.unit} ต่ำกว่าขั้นต่ำ ${s.minQty} ${s.unit}`, time:"ตอนนี้" });
  });
  stock.filter(s=>stockSt(s)==="low").forEach(s=>{
    notes.push({ id:`sl_${s.id}`, type:"warn", icon:"⚠️", title:`ควรสั่งเพิ่ม: ${s.name}`, body:`เหลือ ${s.qty} ${s.unit} ใช้ได้อีก ${(s.qty/s.dailyUse).toFixed(1)} วัน`, time:"วันนี้" });
  });
  // Price spike
  stock.filter(isSpike).forEach(s=>{
    notes.push({ id:`sp_${s.id}`, type:"warn", icon:"📈", title:`ราคาผิดปกติ: ${s.name}`, body:`ราคาล่าสุด ฿${latestCost(s)} สูงกว่าค่าเฉลี่ย ${((latestCost(s)-wac(s))/wac(s)*100).toFixed(0)}%`, time:"วันนี้" });
  });
  // Staff activity today
  const todayMvs = movements.filter(m=>m.date===today());
  if (todayMvs.length>0) {
    notes.push({ id:"mv_today", type:"info", icon:"📦", title:`กิจกรรมสต็อควันนี้ ${todayMvs.length} รายการ`, body:`พนักงาน ${[...new Set(todayMvs.map(m=>m.staffId))].map(id=>staff.find(s=>s.id===id)?.name||id).join(", ")} บันทึกข้อมูล`, time:"วันนี้" });
  }
  // Daily summary
  const todayCF = cf.filter(e=>e.date===today());
  const todayIn = todayCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  if (todayIn>0) {
    notes.push({ id:"cf_today", type:"success", icon:"💰", title:`ยอดขายวันนี้ ฿${fmt(todayIn)}`, body:`บันทึกแล้ว ${todayCF.length} รายการ`, time:"วันนี้" });
  }
  return notes;
}

// ─────────────────────────────────────────────
// LOGIN
// ─────────────────────────────────────────────
function LoginPage({ staff, onLogin }) {
  const [pin, setPin] = useState("");
  const [err, setErr] = useState(false);
  const [shake, setShake] = useState(false);

  const go = () => {
    const found = staff.find(s=>s.pin===pin && s.active);
    if (found) { onLogin(found); }
    else { setErr(true); setShake(true); setPin(""); setTimeout(()=>setShake(false),500); }
  };

  return (
    <div style={{ minHeight:"100vh", background:`linear-gradient(135deg,${T.orangeLt} 0%,#fff 60%,${T.bg} 100%)`, display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24 }}>
      <div style={{ marginBottom:32,textAlign:"center" }}>
        <div style={{ fontSize:56,marginBottom:8 }}>🫕</div>
        <div style={{ color:T.orange,fontSize:26,fontWeight:900,letterSpacing:.5 }}>ไท่กั๋วหม่าล่า</div>
        <div style={{ color:T.textSm,fontSize:15,marginTop:4 }}>ระบบจัดการร้าน</div>
      </div>

      <div style={{ ...S.card, width:"100%", maxWidth:380, boxShadow:T.shadowMd }}>
        <div style={{ color:T.textMd, fontSize:15, marginBottom:10 }}>กรอก PIN เพื่อเข้าใช้งาน</div>
        <input
          type="password" maxLength={4} value={pin}
          onChange={e=>{setPin(e.target.value);setErr(false);}}
          onKeyDown={e=>e.key==="Enter"&&go()}
          style={{ ...S.inp, fontSize:28, letterSpacing:10, textAlign:"center", marginBottom:8,
            border:`2px solid ${err?T.red:T.border}`,
            animation: shake?"shake .3s ease-in-out":undefined }}
          placeholder="••••"
        />
        {err && <div style={{ color:T.red,fontSize:14,marginBottom:8 }}>PIN ไม่ถูกต้องหรือบัญชีถูกระงับ</div>}
        <button onClick={go} style={{ ...S.btn(), width:"100%", padding:14, fontSize:17, marginTop:4 }}>
          เข้าสู่ระบบ
        </button>

        <div style={{ marginTop:20, padding:"14px", background:T.bg, borderRadius:10 }}>
          <div style={{ color:T.textSm,fontSize:13,marginBottom:8,fontWeight:600 }}>PIN ทดสอบ</div>
          {staff.filter(s=>s.active).map(s=>(
            <div key={s.id} style={{ display:"flex",justifyContent:"space-between",fontSize:13,padding:"3px 0",color:T.textMd }}>
              <span>{s.name}</span>
              <span style={{ fontFamily:"monospace",color:T.orange,fontWeight:700 }}>{s.pin}</span>
            </div>
          ))}
        </div>
      </div>
      <style>{`@keyframes shake{0%,100%{transform:translateX(0)}25%{transform:translateX(-8px)}75%{transform:translateX(8px)}}`}</style>
    </div>
  );
}

// ─────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
// DASHBOARD — Real-time with comparison charts
// ─────────────────────────────────────────────
function Dashboard({ stock, cf, movements, user, staff, notifications, lineToken }) {
  const [tick, setTick] = useState(0);
  const [branchFilter, setBranchFilter] = useState("all");

  // Auto-refresh every 30s to simulate real-time
  useEffect(() => {
    const t = setInterval(() => setTick(p => p + 1), 30000);
    return () => clearInterval(t);
  }, []);

  const branchCF = user.role === "owner"
    ? (branchFilter === "all" ? cf : cf.filter(e => e.branch === branchFilter))
    : cf.filter(e => e.branch === user.branch || e.staffId === user.id);

  // Date helpers
  const todayStr = today();
  const yesterdayStr = (() => { const d = new Date(); d.setDate(d.getDate() - 1); return d.toISOString().split("T")[0]; })();
  const lastWeekStr  = (() => { const d = new Date(); d.setDate(d.getDate() - 7); return d.toISOString().split("T")[0]; })();

  const dayIncome = (dateStr) => branchCF.filter(e => e.date === dateStr && e.flow === "in").reduce((a, b) => a + b.amount, 0);
  const dayExpense= (dateStr) => branchCF.filter(e => e.date === dateStr && e.flow === "out").reduce((a, b) => a + b.amount, 0);

  const todayIn    = dayIncome(todayStr);
  const todayOut   = dayExpense(todayStr);
  const yestIn     = dayIncome(yesterdayStr);
  const lastWkIn   = dayIncome(lastWeekStr);

  const pctVsYest  = yestIn > 0 ? (((todayIn - yestIn) / yestIn) * 100).toFixed(1) : null;
  const pctVsLastWk= lastWkIn > 0 ? (((todayIn - lastWkIn) / lastWkIn) * 100).toFixed(1) : null;

  // Last 14 days chart
  const last14 = Array.from({ length: 14 }, (_, i) => {
    const d = new Date(); d.setDate(d.getDate() - (13 - i));
    const k = d.toISOString().split("T")[0];
    const inc = branchCF.filter(e => e.date === k && e.flow === "in").reduce((a, b) => a + b.amount, 0);
    const exp = branchCF.filter(e => e.date === k && e.flow === "out").reduce((a, b) => a + b.amount, 0);
    return { k, day: d.getDate(), inc, exp, profit: inc - exp, isToday: k === todayStr, isYest: k === yesterdayStr };
  });
  const maxVal = Math.max(...last14.map(d => d.inc), 1);

  // Expiry alerts
  const expiringToday    = stock.filter(s => s.expiryDays <= 1 && s.qty > 0);
  const expiringSoon     = stock.filter(s => s.expiryDays > 1 && s.expiryDays <= 3 && s.qty > 0);
  const criticalStock    = stock.filter(s => ["critical", "out"].includes(stockSt(s)));
  const warn = notifications.filter(n => n.type === "danger" || n.type === "warn");

  // Branch comparison (owner only)
  const branchStats = BRANCHES.map(b => ({
    ...b,
    inc: cf.filter(e => e.date === todayStr && e.flow === "in" && e.branch === b.id).reduce((a, c) => a + c.amount, 0),
  }));

  const mk = todayStr.slice(0, 7);
  const mIn  = branchCF.filter(e => e.date.startsWith(mk) && e.flow === "in").reduce((a, b) => a + b.amount, 0);
  const mOut = branchCF.filter(e => e.date.startsWith(mk) && e.flow === "out").reduce((a, b) => a + b.amount, 0);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      {/* Header row */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <div>
          <div style={{ color: T.text, fontSize: 22, fontWeight: 900 }}>สวัสดี, {user.name.split(" ")[0]} 👋</div>
          <div style={{ color: T.textSm, fontSize: 14 }}>{todayStr} • อัพเดทอัตโนมัติทุก 30 วินาที</div>
        </div>
        {/* Branch filter (owner only) */}
        {user.role === "owner" && (
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
            {[["all", "ทุกสาขา"], ...BRANCHES.map(b => [b.id, b.name.split("—")[0].trim()])].map(([v, l]) => (
              <button key={v} onClick={() => setBranchFilter(v)} style={{
                background: branchFilter === v ? T.orange : "transparent",
                border: `1px solid ${branchFilter === v ? T.orange : T.border}`,
                borderRadius: 8, padding: "6px 12px", color: branchFilter === v ? "#fff" : T.textMd,
                cursor: "pointer", fontSize: 13, fontFamily: "inherit"
              }}>{l}</button>
            ))}
          </div>
        )}
      </div>

      {/* Today KPI with comparison */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 12 }}>
        {/* Revenue today */}
        <Card style={{ padding: "16px 18px", borderLeft: `4px solid ${T.orange}` }}>
          <div style={{ fontSize: 22, marginBottom: 4 }}>💰</div>
          <div style={{ color: T.textSm, fontSize: 14 }}>รายรับวันนี้</div>
          <div style={{ color: T.text, fontWeight: 900, fontSize: 24, margin: "4px 0" }}>฿{fmt(todayIn)}</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            {pctVsYest !== null && (
              <span style={{ fontSize: 12, color: +pctVsYest >= 0 ? T.green : T.red, fontWeight: 700 }}>
                {+pctVsYest >= 0 ? "▲" : "▼"} {Math.abs(pctVsYest)}% vs เมื่อวาน
              </span>
            )}
            {pctVsLastWk !== null && (
              <span style={{ fontSize: 12, color: +pctVsLastWk >= 0 ? T.green : T.red, fontWeight: 700 }}>
                {+pctVsLastWk >= 0 ? "▲" : "▼"} {Math.abs(pctVsLastWk)}% vs สัปดาห์ก่อน
              </span>
            )}
          </div>
        </Card>

        <Card style={{ padding: "16px 18px", borderLeft: `4px solid ${T.red}` }}>
          <div style={{ fontSize: 22, marginBottom: 4 }}>💸</div>
          <div style={{ color: T.textSm, fontSize: 14 }}>รายจ่ายวันนี้</div>
          <div style={{ color: T.red, fontWeight: 900, fontSize: 24, margin: "4px 0" }}>฿{fmt(todayOut)}</div>
          <div style={{ color: T.textSm, fontSize: 12 }}>เมื่อวาน ฿{fmt(dayExpense(yesterdayStr))}</div>
        </Card>

        <Card style={{ padding: "16px 18px", borderLeft: `4px solid ${todayIn - todayOut >= 0 ? T.green : T.red}` }}>
          <div style={{ fontSize: 22, marginBottom: 4 }}>📈</div>
          <div style={{ color: T.textSm, fontSize: 14 }}>กำไรวันนี้</div>
          <div style={{ color: todayIn - todayOut >= 0 ? T.green : T.red, fontWeight: 900, fontSize: 24, margin: "4px 0" }}>฿{fmt(todayIn - todayOut)}</div>
          <div style={{ color: T.textSm, fontSize: 12 }}>เมื่อวาน ฿{fmt(yestIn - dayExpense(yesterdayStr))}</div>
        </Card>

        <Card style={{ padding: "16px 18px", borderLeft: `4px solid ${T.blue}` }}>
          <div style={{ fontSize: 22, marginBottom: 4 }}>📅</div>
          <div style={{ color: T.textSm, fontSize: 14 }}>กำไรเดือน {mk}</div>
          <div style={{ color: T.blue, fontWeight: 900, fontSize: 24, margin: "4px 0" }}>฿{fmt(mIn - mOut - TOTAL_FIXED)}</div>
          <div style={{ color: T.textSm, fontSize: 12 }}>รายรับ ฿{fmt(mIn)}</div>
        </Card>
      </div>

      {/* 14-day chart */}
      <Card>
        <div style={{ color: T.text, fontWeight: 800, fontSize: 17, marginBottom: 4 }}>📊 ยอดขาย 14 วันล่าสุด</div>
        <div style={{ display: "flex", gap: 12, marginBottom: 12 }}>
          {[["#f97316", "วันนี้"], ["#fed7aa", "วันอื่น"], ["#fef3c7", "เมื่อวาน"]].map(([c, l]) => (
            <div key={l} style={{ display: "flex", alignItems: "center", gap: 5 }}>
              <div style={{ width: 10, height: 10, borderRadius: 2, background: c }} />
              <span style={{ color: T.textSm, fontSize: 12 }}>{l}</span>
            </div>
          ))}
        </div>
        <div style={{ display: "flex", alignItems: "flex-end", gap: 4, height: 100 }}>
          {last14.map(d => (
            <div key={d.k} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 3 }}>
              <div style={{ width: "100%", borderRadius: "4px 4px 0 0",
                background: d.isToday ? T.orange : d.isYest ? "#fef3c7" : T.borderOr,
                height: `${Math.max((d.inc / maxVal) * 85, 3)}px`,
                border: d.isToday ? `2px solid ${T.orangeDk}` : "none",
                transition: "height .3s"
              }} />
              <div style={{ color: T.textXs, fontSize: 9 }}>{d.day}</div>
            </div>
          ))}
        </div>

        {/* Comparison row */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, marginTop: 14, paddingTop: 14, borderTop: `1px solid ${T.bg}` }}>
          {[
            ["วันนี้", todayIn, T.orange],
            ["เมื่อวาน", yestIn, T.textMd],
            ["สัปดาห์ก่อน", lastWkIn, T.textSm],
          ].map(([l, v, c]) => (
            <div key={l} style={{ textAlign: "center" }}>
              <div style={{ color: T.textSm, fontSize: 12 }}>{l}</div>
              <div style={{ color: c, fontWeight: 800, fontSize: 17 }}>฿{fmt(v)}</div>
            </div>
          ))}
        </div>
      </Card>

      {/* Branch comparison (owner only) */}
      {user.role === "owner" && branchFilter === "all" && (
        <Card>
          <div style={{ color: T.text, fontWeight: 800, fontSize: 17, marginBottom: 12 }}>🏪 เปรียบเทียบสาขาวันนี้</div>
          {branchStats.map((b, i) => {
            const maxBranch = Math.max(...branchStats.map(x => x.inc), 1);
            const pct = (b.inc / maxBranch) * 100;
            return (
              <div key={b.id} style={{ marginBottom: 10 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <span style={{ color: T.textMd, fontSize: 15 }}>{b.name.split("—")[0].trim()}</span>
                  <span style={{ color: b.inc > 0 ? T.orange : T.textXs, fontWeight: 700, fontSize: 15 }}>฿{fmt(b.inc)}</span>
                </div>
                <div style={{ background: T.bg, borderRadius: 6, height: 10 }}>
                  <div style={{ background: `linear-gradient(90deg,${T.orange},${T.orangeDk})`, width: `${pct}%`, height: "100%", borderRadius: 6, transition: "width .5s" }} />
                </div>
              </div>
            );
          })}
        </Card>
      )}

      {/* Expiry alerts */}
      {expiringToday.length > 0 && (
        <Card style={{ borderColor: T.red + "55", background: T.redLt }}>
          <div style={{ color: T.red, fontWeight: 800, fontSize: 17, marginBottom: 10 }}>🗓️ หมดอายุวันนี้! ({expiringToday.length} รายการ) — ใช้ให้หมดก่อน</div>
          {expiringToday.map(s => (
            <div key={s.id} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: `1px solid ${T.redLt}`, fontSize: 15 }}>
              <span style={{ color: T.text, fontWeight: 600 }}>⚠️ {s.name}</span>
              <span style={{ color: T.red }}>เหลือ {s.qty} {s.unit}</span>
            </div>
          ))}
        </Card>
      )}

      {expiringSoon.length > 0 && (
        <Card style={{ borderColor: T.yellow + "55", background: T.yellowLt }}>
          <div style={{ color: T.yellow, fontWeight: 800, fontSize: 16, marginBottom: 8 }}>⏰ ใกล้หมดอายุ 1-3 วัน ({expiringSoon.length} รายการ)</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {expiringSoon.map(s => (
              <span key={s.id} style={{ background: "#fff", border: `1px solid ${T.yellow}44`, borderRadius: 8, padding: "4px 12px", fontSize: 14, color: T.yellow, fontWeight: 600 }}>
                {s.name} ({s.expiryDays} วัน)
              </span>
            ))}
          </div>
        </Card>
      )}

      {/* Stock alerts */}
      {criticalStock.length > 0 && (
        <Card style={{ borderColor: T.red + "44" }}>
          <div style={{ color: T.red, fontWeight: 800, fontSize: 16, marginBottom: 8 }}>🚨 สต็อควิกฤต ({criticalStock.length})</div>
          {criticalStock.map(s => (
            <div key={s.id} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: `1px solid ${T.bg}`, fontSize: 15 }}>
              <span style={{ color: T.text }}>{s.name}</span>
              <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                <span style={{ color: T.textSm }}>เหลือ {s.qty} {s.unit}</span>
                <Badge status={stockSt(s)} />
              </div>
            </div>
          ))}
        </Card>
      )}

      {/* LINE notify status */}
      <Card style={{ padding: "12px 16px", background: lineToken ? T.greenLt : T.bg, borderColor: lineToken ? T.green + "44" : T.border }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontSize: 14, color: lineToken ? T.green : T.textSm, fontWeight: 600 }}>
            📲 LINE Notify: {lineToken ? "✅ เชื่อมแล้ว — แจ้งเตือนอัตโนมัติ" : "❌ ยังไม่ได้เชื่อม — ตั้งค่าใน Settings"}
          </div>
        </div>
      </Card>
    </div>
  );
}

// ─────────────────────────────────────────────
// PURCHASE PAGE — Auto-generate + LINE send
// ─────────────────────────────────────────────
const SUPPLIERS_FULL = [
  { id: 1, name: "ตลาดสดนครชัย",     type: "ผัก",          line: "@vegmarket",  phone: "081-234-5678" },
  { id: 2, name: "ฟาร์มหมูสยาม",     type: "หมู/เนื้อ",    line: "@siampork",   phone: "082-345-6789" },
  { id: 3, name: "อาหารทะเลสด",       type: "ทะเล",         line: "@freshsea",   phone: "083-456-7890" },
  { id: 4, name: "ซอสมาล่าพรีเมียม",  type: "ซอส/บรรจุ",    line: "@malabase",   phone: "084-567-8901" },
];

function PurchasePage({ stock, lineToken }) {
  const [selected, setSelected] = useState({});
  const [orderQtys, setOrderQtys] = useState({});
  const [sent, setSent] = useState({});
  const [note, setNote] = useState("");
  const [preview, setPreview] = useState(false);

  const needOrder = stock.filter(s => ["critical", "out", "low"].includes(stockSt(s)));

  // Auto-calculate suggested order qty
  const suggestQty = (s) => Math.max(s.minQty * 2 - s.qty, s.minQty);

  const toggle = (id) => {
    setSelected(p => ({ ...p, [id]: !p[id] }));
    if (!orderQtys[id]) setOrderQtys(p => ({ ...p, [id]: suggestQty(stock.find(s => s.id === id)) }));
  };

  const selItems = needOrder.filter(s => selected[s.id]);

  const bySupplier = selItems.reduce((acc, s) => {
    const sup = SUPPLIERS_FULL.find(x => x.id === s.supplierId) || SUPPLIERS_FULL[0];
    if (!acc[sup.id]) acc[sup.id] = { sup, items: [] };
    acc[sup.id].items.push(s);
    return acc;
  }, {});

  const totalCost = selItems.reduce((a, s) => a + (orderQtys[s.id] || suggestQty(s)) * (s.costHistory?.slice(-1)[0]?.unitCost || 0), 0);

  const buildMessage = (sup, items) => {
    const lines = items.map(i => `• ${i.name} ${orderQtys[i.id] || suggestQty(i)} ${i.unit}`).join("\n");
    return `🫕 ใบสั่งซื้อ ไท่กั๋วหม่าล่า\n📅 วันที่: ${today()}\n\n${lines}${note ? `\n\nหมายเหตุ: ${note}` : ""}\n\nขอบคุณครับ/ค่ะ`;
  };

  const sendLine = async (supId) => {
    const { sup, items } = bySupplier[supId];
    const msg = buildMessage(sup, items);
    if (lineToken) {
      try {
        await fetch("https://notify-api.line.me/api/notify", {
          method: "POST",
          headers: { "Authorization": `Bearer ${lineToken}`, "Content-Type": "application/x-www-form-urlencoded" },
          body: `message=${encodeURIComponent(msg)}`
        });
      } catch (e) { /* CORS in browser — needs backend */ }
    }
    setSent(p => ({ ...p, [supId]: true }));
    setTimeout(() => setSent(p => ({ ...p, [supId]: false })), 4000);
  };

  const sendAll = () => Object.keys(bySupplier).forEach(id => sendLine(+id));

  const exportPO = () => {
    const rows = selItems.map(s => {
      const sup = SUPPLIERS_FULL.find(x => x.id === s.supplierId);
      return {
        "ซัพพลายเออร์": sup?.name, "วัตถุดิบ": s.name, "หน่วย": s.unit,
        "จำนวนสั่ง": orderQtys[s.id] || suggestQty(s),
        "ราคา/หน่วย": s.costHistory?.slice(-1)[0]?.unitCost || 0,
        "รวม": (orderQtys[s.id] || suggestQty(s)) * (s.costHistory?.slice(-1)[0]?.unitCost || 0),
        "วันที่สั่ง": today(),
      };
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "ใบสั่งซื้อ");
    XLSX.writeFile(wb, `PO_${today()}.xlsx`);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <SectionTitle title="🛒 สั่งซื้อวัตถุดิบ"
        action={
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={exportPO} style={{ ...S.ghost, fontSize: 14, padding: "8px 14px" }}>📤 Export PO</button>
            {selItems.length > 0 && <button onClick={sendAll} style={{ ...S.btn("#22c55e"), fontSize: 14, padding: "8px 16px" }}>📲 ส่ง LINE ทุกเจ้า</button>}
          </div>
        }
      />

      {/* LINE status */}
      <Card style={{ background: lineToken ? T.greenLt : T.yellowLt, borderColor: lineToken ? T.green + "44" : T.yellow + "44", padding: "12px 16px" }}>
        <div style={{ color: lineToken ? T.green : T.yellow, fontWeight: 700, fontSize: 15 }}>
          {lineToken ? "✅ เชื่อม LINE แล้ว — กดส่งจะส่งจริงทันที" : "⚠️ ยังไม่ได้เชื่อม LINE — ตั้งค่า Token ใน Settings ก่อน"}
        </div>
      </Card>

      {/* Select all */}
      {needOrder.length > 0 && (
        <button onClick={() => {
          const allSel = needOrder.every(s => selected[s.id]);
          const newSel = {};
          needOrder.forEach(s => { newSel[s.id] = !allSel; });
          setSelected(newSel);
        }} style={{ ...S.ghost, width: "100%", padding: 12, fontSize: 15 }}>
          {needOrder.every(s => selected[s.id]) ? "✕ ยกเลิกทั้งหมด" : "✓ เลือกทั้งหมด"} ({needOrder.length} รายการ)
        </button>
      )}

      {/* Items to order */}
      <Card>
        <div style={{ color: T.text, fontWeight: 800, fontSize: 17, marginBottom: 12 }}>รายการที่ต้องสั่ง ({needOrder.length})</div>
        {needOrder.length === 0
          ? <div style={{ color: T.textSm, textAlign: "center", padding: 24, fontSize: 16 }}>✅ สต็อคทุกอย่างปกติดี ไม่มีรายการที่ต้องสั่ง</div>
          : needOrder.map(s => {
            const sup = SUPPLIERS_FULL.find(x => x.id === s.supplierId) || SUPPLIERS_FULL[0];
            const suggested = suggestQty(s);
            const cost = s.costHistory?.slice(-1)[0]?.unitCost || 0;
            const qty = orderQtys[s.id] || suggested;
            return (
              <div key={s.id} style={{ display: "flex", alignItems: "flex-start", gap: 12, padding: "12px 0", borderBottom: `1px solid ${T.bg}` }}>
                <input type="checkbox" checked={!!selected[s.id]} onChange={() => toggle(s.id)}
                  style={{ width: 18, height: 18, accentColor: T.orange, cursor: "pointer", marginTop: 3, flexShrink: 0 }} />
                <div style={{ flex: 1 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    <span style={{ color: T.text, fontWeight: 700, fontSize: 16 }}>{s.name}</span>
                    <Badge status={stockSt(s)} />
                  </div>
                  <div style={{ color: T.textSm, fontSize: 13, marginTop: 2 }}>
                    เหลือ {s.qty} {s.unit} • ซัพฯ: {sup.name} ({sup.line})
                  </div>
                  {selected[s.id] && (
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8 }}>
                      <span style={{ color: T.textSm, fontSize: 14 }}>สั่ง:</span>
                      <input type="number" value={orderQtys[s.id] || suggested}
                        onChange={e => setOrderQtys(p => ({ ...p, [s.id]: +e.target.value }))}
                        style={{ ...S.inp, width: 80, padding: "6px 10px", fontSize: 15 }} />
                      <span style={{ color: T.textSm, fontSize: 14 }}>{s.unit}</span>
                      {cost > 0 && <span style={{ color: T.orange, fontSize: 14, fontWeight: 700 }}>≈ ฿{fmt(qty * cost)}</span>}
                    </div>
                  )}
                </div>
              </div>
            );
          })
        }
      </Card>

      {/* Note */}
      <div>
        <div style={{ color: T.textSm, fontSize: 14, marginBottom: 6 }}>หมายเหตุในใบสั่งซื้อ</div>
        <textarea value={note} onChange={e => setNote(e.target.value)}
          style={{ ...S.inp, height: 70, resize: "vertical" }}
          placeholder="เช่น ส่งก่อน 10 โมง, จ่ายเงินสด..." />
      </div>

      {/* Order preview by supplier */}
      {selItems.length > 0 && (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ color: T.text, fontWeight: 800, fontSize: 17 }}>📋 ใบสั่งซื้อ — รวม ≈ ฿{fmt(Math.round(totalCost))}</div>
          {Object.values(bySupplier).map(({ sup, items }) => (
            <Card key={sup.id} style={{ borderColor: T.green + "44" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <div>
                  <div style={{ color: T.text, fontWeight: 700, fontSize: 16 }}>{sup.name}</div>
                  <div style={{ color: T.textSm, fontSize: 13 }}>LINE: {sup.line} • โทร: {sup.phone}</div>
                </div>
                <button onClick={() => sendLine(sup.id)} style={{
                  ...S.btn(sent[sup.id] ? T.green : "#22c55e"),
                  fontSize: 14, padding: "9px 16px"
                }}>
                  {sent[sup.id] ? "✅ ส่งแล้ว!" : "📲 ส่ง LINE"}
                </button>
              </div>
              <div style={{ background: T.bg, borderRadius: 10, padding: "10px 14px", fontFamily: "monospace", fontSize: 14, whiteSpace: "pre-wrap", color: T.textMd }}>
                {buildMessage(sup, items)}
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// VAT / TAX REPORT PAGE
// ─────────────────────────────────────────────
const VAT_RATE = 0.07;

function VATReportPage({ cf }) {
  const months = [...new Set(cf.map(e => e.date.slice(0, 7)))].sort((a, b) => b.localeCompare(a));
  const [m, setM] = useState(months[0] || today().slice(0, 7));
  const [vatType, setVatType] = useState("included"); // included | excluded

  const mCF = cf.filter(e => e.date.startsWith(m));
  const mIn  = mCF.filter(e => e.flow === "in").reduce((a, b) => a + b.amount, 0);
  const mOut = mCF.filter(e => e.flow === "out").reduce((a, b) => a + b.amount, 0);

  const vatSales = vatType === "included" ? mIn / 1.07 * 0.07 : mIn * 0.07;
  const salesExVAT = vatType === "included" ? mIn / 1.07 : mIn;
  const vatPurchase = vatType === "included" ? mOut / 1.07 * 0.07 : mOut * 0.07;
  const purchaseExVAT = vatType === "included" ? mOut / 1.07 : mOut;
  const vatPayable = vatSales - vatPurchase;

  const dailySales = [...new Set(mCF.map(e => e.date))].sort().map(d => ({
    date: d,
    inc: mCF.filter(e => e.date === d && e.flow === "in").reduce((a, b) => a + b.amount, 0),
    exp: mCF.filter(e => e.date === d && e.flow === "out").reduce((a, b) => a + b.amount, 0),
  }));

  const exportVAT = () => {
    const wb = XLSX.utils.book_new();

    // Summary sheet
    const summary = [
      ["รายงาน VAT", "", `ไท่กั๋วหม่าล่า — เดือน ${m}`],
      [],
      ["รายการ", "ยอดรวม (฿)", "ฐานภาษี (฿)", "VAT 7% (฿)"],
      ["รายรับ (ขาย)", mIn.toFixed(2), salesExVAT.toFixed(2), vatSales.toFixed(2)],
      ["รายจ่าย (ซื้อ)", mOut.toFixed(2), purchaseExVAT.toFixed(2), vatPurchase.toFixed(2)],
      [],
      ["VAT ที่ต้องนำส่ง", "", "", vatPayable.toFixed(2)],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), "สรุป VAT");

    // Daily detail
    const daily = dailySales.map(d => ({
      "วันที่": d.date,
      "รายรับ": d.inc,
      "ฐานภาษีขาย": (d.inc / 1.07).toFixed(2),
      "VAT ขาย": (d.inc / 1.07 * 0.07).toFixed(2),
      "รายจ่าย": d.exp,
      "ฐานภาษีซื้อ": (d.exp / 1.07).toFixed(2),
      "VAT ซื้อ": (d.exp / 1.07 * 0.07).toFixed(2),
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(daily), "รายวัน");

    XLSX.writeFile(wb, `VAT_${m}.xlsx`);
  };

  const row = (label, val1, val2, val3, bold = false) => (
    <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr 1fr 1fr", gap: 8, padding: "8px 0", borderBottom: `1px solid ${T.bg}`, fontSize: 15, fontWeight: bold ? 800 : 400 }}>
      <span style={{ color: bold ? T.text : T.textMd }}>{label}</span>
      <span style={{ color: T.text, textAlign: "right" }}>฿{fmt(Math.round(val1))}</span>
      <span style={{ color: T.blue, textAlign: "right" }}>฿{fmt(Math.round(val2))}</span>
      <span style={{ color: T.orange, textAlign: "right" }}>฿{fmt(Math.round(val3))}</span>
    </div>
  );

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <SectionTitle title="🧾 รายงาน VAT & ภาษี"
        action={<button onClick={exportVAT} style={{ ...S.btn("#7c3aed"), fontSize: 14, padding: "9px 16px" }}>📤 Export สำหรับนักบัญชี</button>}
      />

      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
        {months.map(mo => (
          <button key={mo} onClick={() => setM(mo)} style={{
            background: m === mo ? T.orange : "transparent", border: `1px solid ${m === mo ? T.orange : T.border}`,
            borderRadius: 8, padding: "7px 13px", color: m === mo ? "#fff" : T.textMd,
            cursor: "pointer", fontSize: 14, fontFamily: "inherit"
          }}>{mo}</button>
        ))}
      </div>

      {/* VAT type toggle */}
      <Card style={{ padding: "12px 16px" }}>
        <div style={{ color: T.textSm, fontSize: 14, marginBottom: 8 }}>ราคาที่บันทึกใน Cash Flow</div>
        <div style={{ display: "flex", gap: 8 }}>
          {[["included", "ราคารวม VAT แล้ว"], ["excluded", "ราคาไม่รวม VAT"]].map(([v, l]) => (
            <button key={v} onClick={() => setVatType(v)} style={{
              flex: 1, padding: "10px", fontSize: 15,
              background: vatType === v ? T.orange : "transparent",
              border: `1px solid ${vatType === v ? T.orange : T.border}`,
              borderRadius: 10, color: vatType === v ? "#fff" : T.textMd,
              cursor: "pointer", fontFamily: "inherit", fontWeight: vatType === v ? 700 : 400
            }}>{l}</button>
          ))}
        </div>
      </Card>

      {/* VAT Summary table */}
      <Card>
        <div style={{ color: T.text, fontWeight: 800, fontSize: 17, marginBottom: 12 }}>สรุป VAT เดือน {m}</div>
        <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr 1fr 1fr", gap: 8, padding: "6px 0", borderBottom: `2px solid ${T.border}`, fontSize: 13, color: T.textSm, fontWeight: 700 }}>
          <span>รายการ</span>
          <span style={{ textAlign: "right" }}>ยอดรวม</span>
          <span style={{ textAlign: "right", color: T.blue }}>ฐานภาษี</span>
          <span style={{ textAlign: "right", color: T.orange }}>VAT 7%</span>
        </div>
        {row("รายรับ (ภาษีขาย)", mIn, salesExVAT, vatSales)}
        {row("รายจ่าย (ภาษีซื้อ)", mOut, purchaseExVAT, vatPurchase)}

        {/* VAT payable */}
        <div style={{ background: vatPayable >= 0 ? T.orangeLt : T.greenLt, border: `1px solid ${vatPayable >= 0 ? T.borderOr : T.green + "44"}`, borderRadius: 10, padding: "14px 16px", marginTop: 12 }}>
          <div style={{ color: T.textSm, fontSize: 14 }}>VAT ที่ต้องนำส่งสรรพากร</div>
          <div style={{ color: vatPayable >= 0 ? T.orange : T.green, fontWeight: 900, fontSize: 26, marginTop: 4 }}>฿{fmt(Math.round(Math.abs(vatPayable)))}</div>
          <div style={{ color: T.textSm, fontSize: 13, marginTop: 2 }}>
            {vatPayable >= 0 ? "ต้องจ่าย VAT เพิ่ม" : "VAT ซื้อ > ขาย — ขอคืนภาษีได้"} • ยื่นภาษีภายในวันที่ 15 ของเดือนถัดไป
          </div>
        </div>
      </Card>

      {/* Daily VAT table */}
      <Card>
        <div style={{ color: T.text, fontWeight: 800, fontSize: 17, marginBottom: 12 }}>รายงานรายวัน</div>
        <div style={{ overflowX: "auto" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", gap: 4, padding: "6px 0", borderBottom: `2px solid ${T.border}`, fontSize: 13, color: T.textSm, fontWeight: 700, minWidth: 500 }}>
            {["วันที่", "รายรับ", "ฐานภาษีขาย", "VAT ขาย", "สะสม VAT"].map(h => <span key={h} style={{ textAlign: "right" }}>{h}</span>)}
          </div>
          {(() => {
            let cumVAT = 0;
            return dailySales.map(d => {
              const vat = d.inc / 1.07 * 0.07;
              cumVAT += vat;
              return (
                <div key={d.date} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", gap: 4, padding: "7px 0", borderBottom: `1px solid ${T.bg}`, fontSize: 14, minWidth: 500 }}>
                  <span style={{ color: T.textMd }}>{d.date.slice(5)}</span>
                  <span style={{ color: T.green, textAlign: "right" }}>฿{fmt(d.inc)}</span>
                  <span style={{ textAlign: "right" }}>฿{fmt(Math.round(d.inc / 1.07))}</span>
                  <span style={{ color: T.orange, textAlign: "right" }}>฿{fmt(Math.round(vat))}</span>
                  <span style={{ color: T.blue, textAlign: "right", fontWeight: 700 }}>฿{fmt(Math.round(cumVAT))}</span>
                </div>
              );
            });
          })()}
        </div>
      </Card>

      {/* Tips for accountant */}
      <Card style={{ background: T.blueLt, borderColor: T.blue + "33" }}>
        <div style={{ color: T.blue, fontWeight: 800, fontSize: 16, marginBottom: 8 }}>💡 สำหรับนักบัญชี</div>
        {["กด Export เพื่อได้ไฟล์ Excel สรุป VAT รายเดือน",
          "Sheet 1: สรุปภาษีขาย-ซื้อ | Sheet 2: รายวันครบทุกรายการ",
          "ยื่น ภ.พ.30 ทุกเดือน ภายในวันที่ 15 หรือ 23 (ออนไลน์)",
          "เก็บใบกำกับภาษีซื้อเพื่อใช้ขอคืน VAT"
        ].map(t => (
          <div key={t} style={{ color: T.textMd, fontSize: 14, padding: "4px 0" }}>• {t}</div>
        ))}
      </Card>
    </div>
  );
}

function StockPage({ stock, setStock, movements, setMovements, user }) {
  const [tab, setTab] = useState("stock");
  const [selItem, setSelItem] = useState(null);
  const [mvForm, setMvForm] = useState({ type:"in", qty:"", note:"" });
  const [showAdd, setShowAdd] = useState(false);
  const [newItem, setNewItem] = useState({ name:"",unit:"kg",qty:0,minQty:3,dailyUse:1,expiryDays:3 });
  const [filter, setFilter] = useState("all");
  const [importMsg, setImportMsg] = useState("");
  const fileRef = useRef();

  const sorted = useMemo(()=>{
    const ord = {out:0,critical:1,low:2,ok:3};
    let list = filter==="all"?stock:stock.filter(s=>stockSt(s)===filter);
    return [...list].sort((a,b)=>ord[stockSt(a)]-ord[stockSt(b)]);
  },[stock,filter]);

  const addMovement = () => {
    if (!mvForm.qty||!selItem) return;
    const qty = parseFloat(mvForm.qty);
    setMovements(prev=>[...prev,{ id:Date.now(), itemId:selItem.id, type:mvForm.type, qty, date:today(), staffId:user.id, note:mvForm.note, branch:user.branch||"1" }]);
    setStock(prev=>prev.map(s=>s.id===selItem.id?{...s, qty:mvForm.type==="in"?s.qty+qty:Math.max(0,s.qty-qty)}:s));
    setMvForm({type:"in",qty:"",note:""});
    setSelItem(null);
  };

  const itemMovements = selItem ? movements.filter(m=>m.itemId===selItem.id).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,10) : [];

  // Excel Import
  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type:"array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        let added=0, updated=0;
        const newStock = [...stock];
        rows.forEach(r => {
          const name = r["ชื่อวัตถุดิบ"]||r["name"]||"";
          if (!name) return;
          const idx = newStock.findIndex(s=>s.name===name);
          if (idx>=0) {
            if (r["จำนวนคงเหลือ"]!==undefined) newStock[idx] = {...newStock[idx], qty:+r["จำนวนคงเหลือ"]};
            if (r["จำนวนขั้นต่ำ"]!==undefined)  newStock[idx] = {...newStock[idx], minQty:+r["จำนวนขั้นต่ำ"]};
            if (r["ใช้ต่อวัน"]!==undefined)      newStock[idx] = {...newStock[idx], dailyUse:+r["ใช้ต่อวัน"]};
            updated++;
          } else {
            newStock.push({
              id: Date.now()+Math.random(),
              name,
              unit:       r["หน่วย"]||"kg",
              qty:        +(r["จำนวนคงเหลือ"]||0),
              minQty:     +(r["จำนวนขั้นต่ำ"]||3),
              dailyUse:   +(r["ใช้ต่อวัน"]||1),
              expiryDays: +(r["อายุ (วัน)"]||3),
              supplierId: 1,
              costHistory: r["ราคาต่อหน่วย"] ? [{date:today(),unitCost:+r["ราคาต่อหน่วย"],qty:+(r["จำนวนคงเหลือ"]||1),total:+r["ราคาต่อหน่วย"]*(+(r["จำนวนคงเหลือ"]||1))}] : [],
            });
            added++;
          }
        });
        setStock(newStock);
        setImportMsg(`เพิ่ม ${added} รายการ อัพเดท ${updated} รายการ`);
        setTimeout(()=>setImportMsg(""),4000);
      } catch(err) {
        setImportMsg("error");
        setTimeout(()=>setImportMsg(""),4000);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value="";
  };

  // Excel Export
  const handleExport = () => {
    const rows = stock.map(s=>({
      "ชื่อวัตถุดิบ": s.name, "หน่วย": s.unit,
      "จำนวนคงเหลือ": s.qty, "จำนวนขั้นต่ำ": s.minQty,
      "ใช้ต่อวัน": s.dailyUse, "อายุ (วัน)": s.expiryDays,
      "ราคาต่อหน่วย": s.costHistory?.slice(-1)[0]?.unitCost||0,
      "สถานะ": ST_LABEL[stockSt(s)],
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [20,8,14,14,12,10,16,12].map(w=>({wch:w}));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "stock");
    XLSX.writeFile(wb, `stock_${today()}.xlsx`);
  };

  // Template Download
  const downloadTemplate = () => {
    const tpl = [
      {"ชื่อวัตถุดิบ":"หมูสามชั้น","หน่วย":"kg","จำนวนคงเหลือ":10,"จำนวนขั้นต่ำ":5,"ใช้ต่อวัน":4,"อายุ (วัน)":3,"ราคาต่อหน่วย":180},
      {"ชื่อวัตถุดิบ":"ผักกาดขาว","หน่วย":"kg","จำนวนคงเหลือ":5,"จำนวนขั้นต่ำ":3,"ใช้ต่อวัน":3,"อายุ (วัน)":2,"ราคาต่อหน่วย":30},
      {"ชื่อวัตถุดิบ":"ลูกชิ้นหมู","หน่วย":"kg","จำนวนคงเหลือ":8,"จำนวนขั้นต่ำ":3,"ใช้ต่อวัน":2,"อายุ (วัน)":4,"ราคาต่อหน่วย":95},
    ];
    const ws = XLSX.utils.json_to_sheet(tpl);
    ws["!cols"] = [20,8,14,14,12,10,16].map(w=>({wch:w}));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "template");
    XLSX.writeFile(wb, "template_stock.xlsx");
  };

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <SectionTitle title="📦 สต็อควัตถุดิบ"
        action={
          <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display:"none" }} />
            <button onClick={()=>fileRef.current?.click()} style={{...S.ghost,fontSize:14,padding:"8px 14px"}}>📥 Import Excel</button>
            <button onClick={handleExport} style={{...S.ghost,fontSize:14,padding:"8px 14px"}}>📤 Export Excel</button>
            {user.perms?.admin && <button onClick={()=>setShowAdd(!showAdd)} style={S.btn()}>+ เพิ่ม</button>}
          </div>
        }
      />

      {/* Import result message */}
      {importMsg && (
        <div style={{ background:importMsg==="error"?T.redLt:T.greenLt, border:`1px solid ${importMsg==="error"?T.red:T.green}`, borderRadius:10, padding:"12px 16px", color:importMsg==="error"?T.red:T.green, fontWeight:700, fontSize:16 }}>
          {importMsg==="error" ? "❌ ไฟล์ผิดรูปแบบ กรุณาดาวน์โหลด Template ก่อน" : `✅ นำเข้าสำเร็จ! ${importMsg}`}
        </div>
      )}

      {/* Import guide */}
      <Card style={{ background:T.orangeLt,borderColor:T.borderOr,padding:"14px 18px" }}>
        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10 }}>
          <div style={{ flex:1 }}>
            <div style={{ color:T.orange,fontWeight:700,fontSize:16 }}>📋 นำเข้าจาก Excel</div>
            <div style={{ color:T.textMd,fontSize:14,marginTop:2 }}>
              คอลัมน์ที่ต้องมี: <b>ชื่อวัตถุดิบ | หน่วย | จำนวนคงเหลือ | จำนวนขั้นต่ำ | ใช้ต่อวัน</b>
            </div>
            <div style={{ color:T.textSm,fontSize:13,marginTop:2 }}>
              ชื่อซ้ำ = อัพเดทข้อมูลเดิม • ชื่อใหม่ = เพิ่มรายการใหม่
            </div>
          </div>
          <button onClick={downloadTemplate} style={{...S.btn(T.orange),fontSize:14,padding:"10px 16px",whiteSpace:"nowrap"}}>
            ⬇️ ดาวน์โหลด Template
          </button>
        </div>
      </Card>

      <TabBar tabs={[["stock","รายการสต็อค"],["movements","ประวัติเคลื่อนไหว"]]} active={tab} onChange={setTab} />

      {tab==="stock" && (
        <>
          {/* Filter */}
          <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
            {[["all","ทั้งหมด"],["critical","🔴 วิกฤต"],["low","🟡 ต่ำ"],["ok","🟢 ปกติ"]].map(([v,l])=>(
              <button key={v} onClick={()=>setFilter(v)} style={{ background:filter===v?T.orange:"transparent", border:`1px solid ${filter===v?T.orange:T.border}`, borderRadius:8, padding:"7px 14px", color:filter===v?"#fff":T.textMd, cursor:"pointer", fontSize:14, fontFamily:"inherit" }}>{l}</button>
            ))}
          </div>

          {/* Add form */}
          {showAdd && (
            <Card style={{ borderColor:T.orangeLt }}>
              <div style={{ color:T.orange,fontWeight:800,fontSize:17,marginBottom:12 }}>➕ เพิ่มวัตถุดิบ</div>
              <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
                {[["ชื่อ","name","text"],["หน่วย","unit","text"],["จำนวน","qty","number"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"],["อายุ(วัน)","expiryDays","number"]].map(([l,k,t])=>(
                  <div key={k}>
                    <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>{l}</div>
                    <input type={t} value={newItem[k]} onChange={e=>setNewItem(p=>({...p,[k]:e.target.value}))} style={S.inp} />
                  </div>
                ))}
              </div>
              <div style={{ display:"flex",gap:8,marginTop:12 }}>
                <button onClick={()=>{setStock(prev=>[...prev,{...newItem,id:Date.now(),qty:+newItem.qty,minQty:+newItem.minQty,dailyUse:+newItem.dailyUse,expiryDays:+newItem.expiryDays,costHistory:[]}]);setShowAdd(false);}} style={{...S.btn(),flex:1}}>บันทึก</button>
                <button onClick={()=>setShowAdd(false)} style={S.ghost}>ยกเลิก</button>
              </div>
            </Card>
          )}

          {/* Stock list */}
          {sorted.map(item=>{
            const st=stockSt(item);
            const daysLeft=item.dailyUse>0?(item.qty/item.dailyUse).toFixed(1):"∞";
            const isOpen=selItem?.id===item.id;
            return (
              <Card key={item.id} style={{ borderColor:st!=="ok"?ST_COLOR[st]+"33":T.border, padding:"16px 18px" }}>
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",cursor:"pointer" }} onClick={()=>setSelItem(isOpen?null:item)}>
                  <div style={{ flex:1 }}>
                    <div style={{ display:"flex",alignItems:"center",gap:10,flexWrap:"wrap" }}>
                      <span style={{ color:T.text,fontWeight:700,fontSize:18 }}>{item.name}</span>
                      <Badge status={st} />
                    </div>
                    <div style={{ color:T.textSm,fontSize:14,marginTop:4 }}>
                      ขั้นต่ำ {item.minQty} {item.unit} • ใช้ {item.dailyUse}/วัน • เหลืออีก {daysLeft} วัน
                    </div>
                  </div>
                  <div style={{ textAlign:"right",marginLeft:12 }}>
                    <div style={{ color:ST_COLOR[st],fontWeight:900,fontSize:26 }}>{item.qty}</div>
                    <div style={{ color:T.textSm,fontSize:14 }}>{item.unit}</div>
                  </div>
                </div>

                {/* Movement form */}
                {isOpen && (
                  <div style={{ marginTop:14,paddingTop:14,borderTop:`1px solid ${T.bg}` }}>
                    <div style={{ color:T.text,fontWeight:700,fontSize:16,marginBottom:10 }}>บันทึกการเคลื่อนไหว</div>
                    <div style={{ display:"flex",gap:8,marginBottom:10 }}>
                      {[["in","📥 รับเข้า",T.green],["out","📤 จ่ายออก",T.red]].map(([v,l,c])=>(
                        <button key={v} onClick={()=>setMvForm(p=>({...p,type:v}))} style={{ flex:1,padding:10,
                          background:mvForm.type===v?c+"15":"transparent",
                          border:`1.5px solid ${mvForm.type===v?c:T.border}`,
                          borderRadius:10,color:mvForm.type===v?c:T.textMd,fontWeight:700,cursor:"pointer",fontSize:16,fontFamily:"inherit" }}>{l}</button>
                      ))}
                    </div>
                    <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8 }}>
                      <div>
                        <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>จำนวน ({item.unit})</div>
                        <input type="number" value={mvForm.qty} onChange={e=>setMvForm(p=>({...p,qty:e.target.value}))} style={S.inp} placeholder="0" />
                      </div>
                      <div>
                        <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>หมายเหตุ</div>
                        <input type="text" value={mvForm.note} onChange={e=>setMvForm(p=>({...p,note:e.target.value}))} style={S.inp} placeholder="เช่น รับจากซัพฯ" />
                      </div>
                    </div>
                    <div style={{ display:"flex",gap:8 }}>
                      <button onClick={addMovement} style={{...S.btn(mvForm.type==="in"?T.green:T.red),flex:1,padding:10}}>
                        ✅ บันทึก
                      </button>
                      <button onClick={()=>setSelItem(null)} style={{...S.ghost,padding:"10px 14px"}}>ยกเลิก</button>
                    </div>

                    {/* Recent movements */}
                    {itemMovements.length>0 && (
                      <div style={{ marginTop:12 }}>
                        <div style={{ color:T.textSm,fontSize:13,marginBottom:6,fontWeight:600 }}>ประวัติล่าสุด</div>
                        {itemMovements.map(m=>{
                          const who=movements&&INIT_STAFF.find(s=>s.id===m.staffId);
                          return (
                            <div key={m.id} style={{ display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:`1px solid ${T.bg}`,fontSize:14 }}>
                              <span style={{ color:T.textMd }}>{m.date} • {m.note||"-"}</span>
                              <div style={{ display:"flex",gap:8,alignItems:"center" }}>
                                <span style={{ color:T.textXs,fontSize:12 }}>{who?.name||m.staffId}</span>
                                <span style={{ color:m.type==="in"?T.green:T.red,fontWeight:700 }}>{m.type==="in"?"+":"-"}{m.qty}</span>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                )}
              </Card>
            );
          })}
        </>
      )}

      {tab==="movements" && (
        <Card>
          <div style={{ color:T.text,fontWeight:700,fontSize:17,marginBottom:12 }}>ประวัติการเคลื่อนไหวทั้งหมด</div>
          {[...movements].sort((a,b)=>b.date.localeCompare(a.date)).map(m=>{
            const item=stock.find(s=>s.id===m.itemId);
            const who=INIT_STAFF.find(s=>s.id===m.staffId);
            return (
              <div key={m.id} style={{ display:"flex",alignItems:"center",gap:12,padding:"10px 0",borderBottom:`1px solid ${T.bg}` }}>
                <div style={{ width:36,height:36,borderRadius:10,background:m.type==="in"?T.greenLt:T.redLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0 }}>
                  {m.type==="in"?"📥":"📤"}
                </div>
                <div style={{ flex:1 }}>
                  <div style={{ color:T.text,fontSize:15,fontWeight:600 }}>{item?.name}</div>
                  <div style={{ color:T.textSm,fontSize:13 }}>{m.date} • {who?.name||m.staffId} • {m.note||"-"}</div>
                </div>
                <span style={{ color:m.type==="in"?T.green:T.red,fontWeight:700,fontSize:16 }}>
                  {m.type==="in"?"+":"-"}{m.qty} {item?.unit}
                </span>
              </div>
            );
          })}
        </Card>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// CASHFLOW PAGE
// ─────────────────────────────────────────────
function CashflowPage({ cf, setCF, user }) {
  const [showForm, setShowForm] = useState(false);
  const [viewTab, setViewTab] = useState("list");
  const [filterMonth, setFilterMonth] = useState(today().slice(0,7));
  const [form, setForm] = useState({ date:today(),flow:"in",cat:IN_CATS[0],itemName:"",amount:"",method:"เงินสด",note:"" });

  const myCF = user.role==="owner" ? cf : cf.filter(e=>e.branch===user.branch||e.staffId===user.id);
  const months = [...new Set(myCF.map(e=>e.date.slice(0,7)))].sort((a,b)=>b.localeCompare(a));
  const filtered = myCF.filter(e=>e.date.startsWith(filterMonth));
  const totalIn  = filtered.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const totalOut = filtered.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);

  const sorted = [...filtered].sort((a,b)=>b.date.localeCompare(a.date)||b.id-a.id);
  const byDate = sorted.reduce((acc,e)=>{ if(!acc[e.date])acc[e.date]=[]; acc[e.date].push(e); return acc; },{});
  const dates = Object.keys(byDate).sort((a,b)=>b.localeCompare(a));

  const addEntry = () => {
    if (!form.amount) return;
    setCF(prev=>[{...form,id:Date.now(),amount:+form.amount,branch:user.branch||"1",staffId:user.id},...prev]);
    setShowForm(false);
    setForm({date:today(),flow:"in",cat:IN_CATS[0],itemName:"",amount:"",method:"เงินสด",note:""});
  };

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <SectionTitle title="💵 Cash Flow"
        action={
          <div style={{ display:"flex",gap:8 }}>
            <button onClick={()=>setShowForm(!showForm)} style={S.btn()}>+ กรอก</button>
          </div>
        }
      />

      {/* KPI */}
      <div style={{ display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10 }}>
        {[["เงินเข้า",`฿${fmt(totalIn)}`,T.green],["เงินออก",`฿${fmt(totalOut)}`,T.red],["คงเหลือ",`฿${fmt(totalIn-totalOut)}`,totalIn-totalOut>=0?T.green:T.red]].map(([l,v,c])=>(
          <Card key={l} style={{ padding:"14px 16px",textAlign:"center" }}>
            <div style={{ color:T.textSm,fontSize:14 }}>{l}</div>
            <div style={{ color:c,fontWeight:800,fontSize:20,marginTop:2 }}>{v}</div>
          </Card>
        ))}
      </div>

      {/* Month + View tabs */}
      <div style={{ display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8 }}>
        <div style={{ display:"flex",gap:5 }}>
          {months.map(m=>(
            <button key={m} onClick={()=>setFilterMonth(m)} style={{ background:filterMonth===m?T.orange:"transparent", border:`1px solid ${filterMonth===m?T.orange:T.border}`, borderRadius:8, padding:"7px 13px", color:filterMonth===m?"#fff":T.textMd, cursor:"pointer", fontSize:14, fontFamily:"inherit" }}>{m}</button>
          ))}
        </div>
        <div style={{ display:"flex",gap:5 }}>
          {[["list","รายการ"],["summary","สรุป"]].map(([v,l])=>(
            <button key={v} onClick={()=>setViewTab(v)} style={{ background:viewTab===v?T.blue:"transparent", border:`1px solid ${viewTab===v?T.blue:T.border}`, borderRadius:8, padding:"7px 13px", color:viewTab===v?"#fff":T.textMd, cursor:"pointer", fontSize:14, fontFamily:"inherit" }}>{l}</button>
          ))}
        </div>
      </div>

      {/* Add form */}
      {showForm && (
        <Card style={{ borderColor:T.orangeLt }}>
          <div style={{ color:T.orange,fontWeight:800,fontSize:18,marginBottom:12 }}>📝 บันทึกรายการ</div>
          <div style={{ display:"flex",gap:8,marginBottom:12 }}>
            {[["in","💰 รายรับ",T.green],["out","💸 รายจ่าย",T.red]].map(([v,l,c])=>(
              <button key={v} onClick={()=>setForm(p=>({...p,flow:v,cat:v==="in"?IN_CATS[0]:OUT_CATS[0],itemName:""}))} style={{ flex:1,padding:12,background:form.flow===v?c+"15":"transparent",border:`1.5px solid ${form.flow===v?c:T.border}`,borderRadius:10,color:form.flow===v?c:T.textMd,fontWeight:700,cursor:"pointer",fontSize:16,fontFamily:"inherit" }}>{l}</button>
            ))}
          </div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
            <div><div style={{ color:T.textSm,fontSize:14,marginBottom:4 }}>วันที่</div><input type="date" value={form.date} onChange={e=>setForm(p=>({...p,date:e.target.value}))} style={S.inp}/></div>
            <div><div style={{ color:T.textSm,fontSize:14,marginBottom:4 }}>จำนวน (฿)</div><input type="number" value={form.amount} onChange={e=>setForm(p=>({...p,amount:e.target.value}))} style={S.inp} placeholder="0"/></div>
            <div><div style={{ color:T.textSm,fontSize:14,marginBottom:4 }}>หมวดหมู่</div><select value={form.cat} onChange={e=>setForm(p=>({...p,cat:e.target.value}))} style={S.inp}>{(form.flow==="in"?IN_CATS:OUT_CATS).map(c=><option key={c}>{c}</option>)}</select></div>
            <div><div style={{ color:T.textSm,fontSize:14,marginBottom:4 }}>ชื่อรายการ</div><input type="text" value={form.itemName} onChange={e=>setForm(p=>({...p,itemName:e.target.value}))} style={S.inp} placeholder="เช่น หมูสามชั้น"/></div>
            <div><div style={{ color:T.textSm,fontSize:14,marginBottom:4 }}>ช่องทาง</div><select value={form.method} onChange={e=>setForm(p=>({...p,method:e.target.value}))} style={S.inp}>{PAY_METHODS.map(m=><option key={m}>{m}</option>)}</select></div>
            <div><div style={{ color:T.textSm,fontSize:14,marginBottom:4 }}>หมายเหตุ</div><input type="text" value={form.note} onChange={e=>setForm(p=>({...p,note:e.target.value}))} style={S.inp} placeholder="(ไม่บังคับ)"/></div>
          </div>
          <div style={{ display:"flex",gap:8,marginTop:14 }}>
            <button onClick={addEntry} style={{...S.btn(form.flow==="in"?T.green:T.red),flex:1,padding:12,fontSize:17}}>✅ บันทึก</button>
            <button onClick={()=>setShowForm(false)} style={{...S.ghost,padding:"12px 16px"}}>ยกเลิก</button>
          </div>
        </Card>
      )}

      {/* List view */}
      {viewTab==="list" && dates.map(date=>{
        const items=byDate[date];
        const dIn=items.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
        const dOut=items.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
        return (
          <Card key={date} style={{ padding:"14px 18px" }}>
            <div style={{ display:"flex",justifyContent:"space-between",paddingBottom:10,marginBottom:10,borderBottom:`1px solid ${T.bg}` }}>
              <span style={{ color:T.text,fontWeight:800,fontSize:16 }}>{date}</span>
              <div style={{ display:"flex",gap:12,fontSize:15 }}>
                <span style={{ color:T.green,fontWeight:600 }}>+฿{fmt(dIn)}</span>
                <span style={{ color:T.red,fontWeight:600 }}>-฿{fmt(dOut)}</span>
                <span style={{ color:dIn-dOut>=0?T.green:T.red,fontWeight:800 }}>= ฿{fmt(dIn-dOut)}</span>
              </div>
            </div>
            {items.map(e=>(
              <div key={e.id} style={{ display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:`1px solid ${T.bg}` }}>
                <div style={{ width:32,height:32,borderRadius:9,background:e.flow==="in"?T.greenLt:T.redLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0 }}>
                  {e.flow==="in"?"↓":"↑"}
                </div>
                <div style={{ flex:1 }}>
                  <div style={{ color:T.text,fontSize:16,fontWeight:e.itemName?600:400 }}>{e.itemName||e.cat}</div>
                  <div style={{ color:T.textSm,fontSize:13 }}>{e.itemName?`${e.cat} • `:"" }{e.method}{e.note?` • ${e.note}`:""}</div>
                </div>
                <span style={{ color:e.flow==="in"?T.green:T.red,fontWeight:700,fontSize:17 }}>
                  {e.flow==="in"?"+":"-"}฿{fmt(e.amount)}
                </span>
              </div>
            ))}
          </Card>
        );
      })}

      {/* Summary view */}
      {viewTab==="summary" && (
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          {INGREDIENT_GROUPS.map(g=>{
            const val=filtered.filter(e=>e.flow==="out"&&e.cat===g.key).reduce((a,b)=>a+b.amount,0);
            if (!val) return null;
            const pct=totalOut>0?(val/totalOut*100).toFixed(0):0;
            return (
              <Card key={g.key} style={{ borderColor:g.color+"33",background:g.bg }}>
                <div style={{ display:"flex",justifyContent:"space-between",marginBottom:6 }}>
                  <span style={{ color:g.color,fontWeight:700,fontSize:16 }}>{g.label}</span>
                  <span style={{ color:g.color,fontWeight:800,fontSize:17 }}>฿{fmt(val)} ({pct}%)</span>
                </div>
                <div style={{ background:"rgba(255,255,255,0.5)",borderRadius:5,height:8 }}>
                  <div style={{ background:g.color,width:`${pct}%`,height:"100%",borderRadius:5 }}/>
                </div>
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// REPORT PAGE
// ─────────────────────────────────────────────
function ReportPage({ cf, stock, movements, staff, user }) {
  const months = [...new Set(cf.map(e=>e.date.slice(0,7)))].sort((a,b)=>b.localeCompare(a));
  const [m, setM] = useState(months[0]||today().slice(0,7));

  const mCF = cf.filter(e=>e.date.startsWith(m));
  const mIn  = mCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mOut = mCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const cogs = mCF.filter(e=>e.flow==="out"&&INGREDIENT_GROUPS.some(g=>g.key===e.cat)).reduce((a,b)=>a+b.amount,0);
  const grossP = mIn-cogs;
  const netP   = grossP-mOut+cogs-TOTAL_FIXED;
  const grossM = mIn>0?((grossP/mIn)*100).toFixed(1):0;

  // Staff performance
  const staffPerf = staff.filter(s=>s.role!=="owner").map(s=>{
    const sCF = mCF.filter(e=>e.staffId===s.id);
    const sMv = movements.filter(mv=>mv.staffId===s.id&&mv.date.startsWith(m));
    const sIn = sCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
    return { ...s, cfCount:sCF.length, mvCount:sMv.length, revenue:sIn };
  });

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <SectionTitle title="📊 รายงาน"
        action={<button onClick={()=>{
          const wb=XLSX.utils.book_new();
          const rows=mCF.map(e=>({วันที่:e.date,ประเภท:e.flow==="in"?"รายรับ":"รายจ่าย",หมวด:e.cat,รายการ:e.itemName,จำนวน:e.amount,พนักงาน:staff.find(s=>s.id===e.staffId)?.name||e.staffId}));
          XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"รายงาน");
          XLSX.writeFile(wb,`report_${m}.xlsx`);
        }} style={{...S.btn("#7c3aed"),fontSize:14}}>📤 Export</button>}
      />

      <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
        {months.map(mo=>(
          <button key={mo} onClick={()=>setM(mo)} style={{ background:m===mo?T.orange:"transparent", border:`1px solid ${m===mo?T.orange:T.border}`, borderRadius:8, padding:"7px 13px", color:m===mo?"#fff":T.textMd, cursor:"pointer", fontSize:14, fontFamily:"inherit" }}>{mo}</button>
        ))}
      </div>

      {/* P&L */}
      <Card style={{ background:T.orangeLt,borderColor:T.borderOr }}>
        <div style={{ color:T.orange,fontWeight:800,fontSize:18,marginBottom:14 }}>📑 P&L เดือน {m}</div>
        {[["💰 รายรับรวม",mIn,T.green,false],["− ต้นทุนวัตถุดิบ",cogs,T.red,true],["= กำไรขั้นต้น",grossP,T.green,false],["− ค่าใช้จ่ายดำเนินการ",mOut-cogs+TOTAL_FIXED,T.yellow,true],["= กำไรสุทธิ",netP,netP>=0?T.green:T.red,false]].map(([l,v,c,indent])=>(
          <div key={l} style={{ display:"flex",justifyContent:"space-between",padding:"8px 0",paddingLeft:indent?16:0,borderBottom:`1px solid ${T.borderOr}`,fontSize:16 }}>
            <span style={{ color:l.startsWith("=")?T.text:T.textMd,fontWeight:l.startsWith("=")?800:400 }}>{l}</span>
            <span style={{ color:c,fontWeight:l.startsWith("=")?900:600,fontSize:l.startsWith("=")?19:16 }}>฿{fmt(Math.abs(v))}</span>
          </div>
        ))}
        <div style={{ display:"flex",gap:16,marginTop:12,paddingTop:10,borderTop:`1px solid ${T.borderOr}` }}>
          {[["Gross Margin",`${grossM}%`],["BEP/วัน",`฿${fmt(Math.ceil(TOTAL_FIXED/30))}`]].map(([l,v])=>(
            <div key={l}><div style={{ color:T.textSm,fontSize:13 }}>{l}</div><div style={{ color:T.orange,fontWeight:800,fontSize:18 }}>{v}</div></div>
          ))}
        </div>
      </Card>

      {/* Staff performance — owner only */}
      {user.role==="owner" && (
        <Card>
          <div style={{ color:T.text,fontWeight:800,fontSize:17,marginBottom:12 }}>👷 รายงานพนักงาน เดือน {m}</div>
          <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:10 }}>
            {staffPerf.map(s=>(
              <div key={s.id} style={{ background:s.active?T.bg:T.redLt, border:`1px solid ${T.border}`, borderRadius:12, padding:"14px 16px" }}>
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8 }}>
                  <span style={{ color:T.text,fontWeight:700,fontSize:16 }}>{s.name.split("(")[0].trim()}</span>
                  <span style={{ background:s.active?T.greenLt:T.redLt,color:s.active?T.green:T.red,borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700 }}>{s.active?"ใช้งาน":"ระงับ"}</span>
                </div>
                <div style={{ fontSize:14,color:T.textMd }}>
                  <div>📝 กรอก CF: <b>{s.cfCount}</b> ครั้ง</div>
                  <div>📦 บันทึกสต็อค: <b>{s.mvCount}</b> ครั้ง</div>
                  <div>💰 ยอดขาย: <b style={{ color:T.green }}>฿{fmt(s.revenue)}</b></div>
                </div>
              </div>
            ))}
          </div>
        </Card>
      )}

      {/* COGS breakdown */}
      <Card>
        <div style={{ color:T.text,fontWeight:800,fontSize:17,marginBottom:12 }}>🧾 ต้นทุนวัตถุดิบแยกหมวด</div>
        {INGREDIENT_GROUPS.map(g=>{
          const val=mCF.filter(e=>e.flow==="out"&&e.cat===g.key).reduce((a,b)=>a+b.amount,0);
          if (!val) return null;
          const pct=cogs>0?(val/cogs*100).toFixed(0):0;
          return (
            <div key={g.key} style={{ marginBottom:10 }}>
              <div style={{ display:"flex",justifyContent:"space-between",marginBottom:4,fontSize:15 }}>
                <span style={{ color:g.color,fontWeight:600 }}>{g.label}</span>
                <span style={{ color:g.color,fontWeight:700 }}>฿{fmt(val)} ({pct}%)</span>
              </div>
              <div style={{ background:T.bg,borderRadius:5,height:8 }}>
                <div style={{ background:g.color,width:`${pct}%`,height:"100%",borderRadius:5 }}/>
              </div>
            </div>
          );
        })}
      </Card>
    </div>
  );
}

// ─────────────────────────────────────────────
// SETTINGS — Permission Management
// ─────────────────────────────────────────────
function SettingsPage({ staff, setStaff, notifications, lineToken, setLineToken }) {
  const [tab, setTab] = useState("staff");
  const [editId, setEditId] = useState(null);

  const PERM_LABELS = { cashflow:"💵 Cash Flow", stock:"📦 สต็อค", purchase:"🛒 สั่งซื้อ", report:"📊 รายงาน", cost:"💹 ต้นทุน", admin:"⚙️ Admin" };

  const togglePerm = (staffId, perm) => {
    setStaff(prev=>prev.map(s=>s.id===staffId?{...s,perms:{...s.perms,[perm]:!s.perms[perm]}}:s));
  };
  const toggleActive = (staffId) => {
    setStaff(prev=>prev.map(s=>s.id===staffId?{...s,active:!s.active}:s));
  };
  const updatePin = (staffId, pin) => {
    setStaff(prev=>prev.map(s=>s.id===staffId?{...s,pin}:s));
  };
  const updateName = (staffId, name) => {
    setStaff(prev=>prev.map(s=>s.id===staffId?{...s,name}:s));
  };

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <SectionTitle title="⚙️ ตั้งค่าระบบ" />
      <TabBar tabs={[["staff","👷 พนักงาน"],["notif","🔔 แจ้งเตือน"],["line","📲 LINE"]]} active={tab} onChange={setTab} />

      {/* Staff management */}
      {tab==="staff" && (
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          {staff.filter(s=>s.role!=="owner").map(s=>(
            <Card key={s.id} style={{ borderColor:s.active?T.border:T.redLt }}>
              {/* Header */}
              <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12 }}>
                <div>
                  <div style={{ color:T.text,fontWeight:700,fontSize:17 }}>{s.name}</div>
                  <div style={{ color:T.textSm,fontSize:14 }}>{BRANCHES.find(b=>b.id===s.branch)?.name}</div>
                </div>
                <div style={{ display:"flex",gap:8,alignItems:"center" }}>
                  {/* Active toggle */}
                  <button onClick={()=>toggleActive(s.id)} style={{
                    background:s.active?T.greenLt:T.redLt, border:`1px solid ${s.active?T.green:T.red}`,
                    borderRadius:8, padding:"6px 14px", color:s.active?T.green:T.red,
                    cursor:"pointer", fontSize:14, fontWeight:700, fontFamily:"inherit"
                  }}>{s.active?"ใช้งาน ✓":"ระงับ ✗"}</button>
                  <button onClick={()=>setEditId(editId===s.id?null:s.id)} style={{...S.ghost,padding:"6px 12px",fontSize:14}}>✏️</button>
                </div>
              </div>

              {/* Edit mode */}
              {editId===s.id && (
                <div style={{ background:T.bg,borderRadius:10,padding:"12px 14px",marginBottom:12 }}>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
                    <div>
                      <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>ชื่อพนักงาน</div>
                      <input value={s.name} onChange={e=>updateName(s.id,e.target.value)} style={{...S.inp,fontSize:15}} />
                    </div>
                    <div>
                      <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>PIN (4 หลัก)</div>
                      <input maxLength={4} value={s.pin} onChange={e=>updatePin(s.id,e.target.value)} style={{...S.inp,fontSize:15,letterSpacing:4}} placeholder="••••" />
                    </div>
                  </div>
                </div>
              )}

              {/* Permissions grid */}
              <div style={{ color:T.textSm,fontSize:14,marginBottom:8,fontWeight:600 }}>สิทธิ์การใช้งาน</div>
              <div style={{ display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8 }}>
                {Object.entries(PERM_LABELS).map(([perm,label])=>(
                  <button key={perm} onClick={()=>togglePerm(s.id,perm)} style={{
                    background:s.perms[perm]?T.orange:T.bg,
                    border:`1px solid ${s.perms[perm]?T.orange:T.border}`,
                    borderRadius:9, padding:"8px 6px",
                    color:s.perms[perm]?"#fff":T.textMd,
                    cursor:"pointer", fontSize:13, fontFamily:"inherit",
                    fontWeight:s.perms[perm]?700:400, textAlign:"center",
                    transition:"all .15s"
                  }}>{label}</button>
                ))}
              </div>
            </Card>
          ))}
        </div>
      )}

      {/* Notifications */}
      {tab==="notif" && (
        <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
          <Card style={{ background:T.orangeLt,borderColor:T.borderOr }}>
            <div style={{ color:T.orange,fontWeight:800,fontSize:17,marginBottom:12 }}>🔔 การแจ้งเตือนในแอป</div>
            {notifications.length===0 ? (
              <div style={{ color:T.textSm,textAlign:"center",padding:20,fontSize:16 }}>✅ ไม่มีการแจ้งเตือนตอนนี้</div>
            ) : notifications.map(n=>(
              <div key={n.id} style={{ display:"flex",gap:12,padding:"10px 0",borderBottom:`1px solid ${T.borderOr}` }}>
                <div style={{ fontSize:22,flexShrink:0 }}>{n.icon}</div>
                <div>
                  <div style={{ color:T.text,fontWeight:600,fontSize:16 }}>{n.title}</div>
                  <div style={{ color:T.textSm,fontSize:14 }}>{n.body}</div>
                  <div style={{ color:T.textXs,fontSize:12,marginTop:2 }}>{n.time}</div>
                </div>
              </div>
            ))}
          </Card>

          <Card>
            <div style={{ color:T.text,fontWeight:800,fontSize:17,marginBottom:12 }}>⏰ ตั้งค่าการแจ้งเตือน</div>
            {[["สต็อคต่ำกว่าขั้นต่ำ","แจ้งทันที",true],["สรุปยอดขายประจำวัน","23:00 น.",true],["ราคาวัตถุดิบผิดปกติ","แจ้งทันที",true],["รายงานรายสัปดาห์","จันทร์ 08:00",false]].map(([name,time,on])=>(
              <div key={name} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:`1px solid ${T.bg}` }}>
                <div>
                  <div style={{ color:T.text,fontSize:16 }}>{name}</div>
                  <div style={{ color:T.textSm,fontSize:13 }}>{time}</div>
                </div>
                <div style={{ width:44,height:24,borderRadius:12,background:on?T.orange:T.border,position:"relative",cursor:"pointer" }}>
                  <div style={{ position:"absolute",top:2,width:20,height:20,borderRadius:"50%",background:"#fff",boxShadow:T.shadow,left:on?22:2,transition:"left .2s" }}/>
                </div>
              </div>
            ))}
          </Card>
        </div>
      )}

      {/* LINE Setup */}
      {tab==="line" && (
        <Card>
          <div style={{ color:T.orange,fontWeight:800,fontSize:17,marginBottom:14 }}>📲 ตั้งค่า LINE Notify</div>
          {[["1️⃣ สมัคร LINE Notify Token","เปิด notify-bot.line.me → Login → Generate token → Copy Token"],["2️⃣ สร้าง Webhook","ใช้ Make.com (ฟรี) รับ webhook แล้วส่งต่อ LINE"],["3️⃣ ใส่ URL ด้านล่าง","แอปจะส่งข้อมูลไปให้อัตโนมัติ"]].map(([s,d])=>(
            <div key={s} style={{ background:T.bg,borderRadius:10,padding:"12px 14px",marginBottom:8 }}>
              <div style={{ color:T.orange,fontWeight:700,fontSize:15 }}>{s}</div>
              <div style={{ color:T.textMd,fontSize:14,marginTop:2 }}>{d}</div>
            </div>
          ))}
          <div style={{ marginTop:4 }}>
            <div style={{ color:T.textSm,fontSize:14,marginBottom:6 }}>Webhook URL</div>
            <div style={{ marginBottom:10 }}>
            <div style={{ color:T.textSm,fontSize:14,marginBottom:6 }}>LINE Notify Token (รับจาก notify-bot.line.me)</div>
            <input type="password" value={lineToken} onChange={e=>setLineToken(e.target.value)} style={S.inp} placeholder="ใส่ Token ที่นี่..." />
            {lineToken && <div style={{ color:T.green,fontSize:13,marginTop:4,fontWeight:600 }}>✅ Token บันทึกแล้ว — LINE พร้อมใช้งาน</div>}
          </div>
          <div><div style={{ color:T.textSm,fontSize:14,marginBottom:6 }}>Webhook URL (Make.com / n8n)</div>
            <input type="url" style={S.inp} placeholder="https://hook.make.com/xxxxx" />
            <button style={{...S.btn(),width:"100%",marginTop:10,padding:12,fontSize:16}}>💾 บันทึก</button>
          </div>
        </Card>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
// MAIN APP — v6 with all features
// ─────────────────────────────────────────────
export default function App() {
  const [user, setUser]           = useState(null);
  const [page, setPage]           = useState("dashboard");
  const [stock, setStock]         = useState(INIT_STOCK);
  const [cf, setCF]               = useState(INIT_CF);
  const [movements, setMovements] = useState(INIT_MOVEMENTS);
  const [staff, setStaff]         = useState(INIT_STAFF);
  const [notifOpen, setNotifOpen] = useState(false);
  const [lineToken, setLineToken] = useState("");

  const notifications = useMemo(() => buildNotifications(stock, cf, movements, staff), [stock, cf, movements]);

  if (!user) return <LoginPage staff={staff} onLogin={u => { setUser(u); setPage("dashboard"); }} />;

  const p = user.perms;

  const ownerNav = [
    { id: "dashboard", icon: "🏠", label: "หลัก"     },
    { id: "cashflow",  icon: "💵", label: "Cash Flow" },
    { id: "stock",     icon: "📦", label: "สต็อค"    },
    { id: "purchase",  icon: "🛒", label: "สั่งซื้อ"  },
    { id: "vat",       icon: "🧾", label: "ภาษี"     },
    { id: "settings",  icon: "⚙️", label: "ตั้งค่า"  },
  ];

  const staffNav = [
    { id: "dashboard", icon: "🏠", label: "หลัก"     },
    ...(p.cashflow ? [{ id: "cashflow", icon: "💵", label: "Cash Flow" }] : []),
    ...(p.stock    ? [{ id: "stock",    icon: "📦", label: "สต็อค"    }] : []),
    ...(p.purchase ? [{ id: "purchase", icon: "🛒", label: "สั่งซื้อ"  }] : []),
    ...(p.report   ? [{ id: "report",   icon: "📊", label: "รายงาน"   }] : []),
  ];

  const nav = user.role === "owner" ? ownerNav : staffNav;

  const pages = {
    dashboard: <Dashboard stock={stock} cf={cf} movements={movements} user={user} staff={staff} notifications={notifications} lineToken={lineToken} />,
    cashflow:  <CashflowPage cf={cf} setCF={setCF} user={user} />,
    stock:     <StockPage stock={stock} setStock={setStock} movements={movements} setMovements={setMovements} user={user} />,
    purchase:  <PurchasePage stock={stock} lineToken={lineToken} />,
    report:    <ReportPage cf={cf} stock={stock} movements={movements} staff={staff} user={user} />,
    vat:       <VATReportPage cf={cf} />,
    settings:  <SettingsPage staff={staff} setStaff={setStaff} notifications={notifications} lineToken={lineToken} setLineToken={setLineToken} />,
  };

  const alertCount = notifications.filter(n => n.type === "danger" || n.type === "warn").length;

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Noto Sans Thai','Noto Sans',sans-serif", color: T.text, fontSize: 18 }}>

      {/* Header */}
      <div style={{ background: T.card, borderBottom: `1px solid ${T.border}`, padding: "12px 20px", display: "flex", alignItems: "center", gap: 12, position: "sticky", top: 0, zIndex: 50, boxShadow: T.shadow }}>
        <div style={{ width: 36, height: 36, borderRadius: 10, background: `linear-gradient(135deg,${T.orange},${T.orangeDk})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>🫕</div>
        <div>
          <div style={{ color: T.orange, fontWeight: 900, fontSize: 16 }}>ไท่กั๋วหม่าล่า</div>
          <div style={{ color: T.textXs, fontSize: 12 }}>{user.name}</div>
        </div>

        {/* Notification bell */}
        <button onClick={() => setNotifOpen(!notifOpen)} style={{ marginLeft: "auto", position: "relative", background: "transparent", border: "none", cursor: "pointer", padding: 8, borderRadius: 10 }}>
          <span style={{ fontSize: 22 }}>🔔</span>
          {alertCount > 0 && (
            <span style={{ position: "absolute", top: 4, right: 4, background: T.red, color: "#fff", borderRadius: "50%", width: 18, height: 18, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 800 }}>
              {alertCount}
            </span>
          )}
        </button>

        <button onClick={() => setUser(null)} style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 8, padding: "6px 12px", color: T.textMd, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>
          ออก
        </button>
      </div>

      {/* Notification dropdown */}
      {notifOpen && (
        <div onClick={() => setNotifOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 190 }} />
      )}
      {notifOpen && (
        <div style={{ position: "fixed", top: 64, right: 12, zIndex: 200, width: 330, background: T.card, border: `1px solid ${T.border}`, borderRadius: 14, boxShadow: T.shadowMd, padding: "14px 16px", maxHeight: "65vh", overflowY: "auto" }}>
          <div style={{ color: T.text, fontWeight: 800, fontSize: 16, marginBottom: 10 }}>🔔 การแจ้งเตือน</div>
          {notifications.length === 0
            ? <div style={{ color: T.textSm, textAlign: "center", padding: 16 }}>ไม่มีการแจ้งเตือน ✅</div>
            : notifications.map(n => (
              <div key={n.id} style={{ display: "flex", gap: 10, padding: "9px 0", borderBottom: `1px solid ${T.bg}` }}>
                <span style={{ fontSize: 20, flexShrink: 0 }}>{n.icon}</span>
                <div>
                  <div style={{ color: T.text, fontSize: 14, fontWeight: 600 }}>{n.title}</div>
                  <div style={{ color: T.textSm, fontSize: 12 }}>{n.body}</div>
                </div>
              </div>
            ))
          }
        </div>
      )}

      {/* Content */}
      <div style={{ padding: "20px 16px 100px", maxWidth: 900, margin: "0 auto" }}>
        {pages[page] || pages["dashboard"]}
      </div>

      {/* Bottom nav */}
      <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: T.card, borderTop: `1px solid ${T.border}`, display: "flex", justifyContent: "space-around", padding: "10px 0 16px", zIndex: 100, boxShadow: "0 -2px 10px rgba(0,0,0,0.06)" }}>
        {nav.map(n => (
          <button key={n.id} onClick={() => setPage(n.id)} style={{ background: "none", border: "none", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, padding: "2px 6px", minWidth: 48 }}>
            <span style={{ fontSize: 20 }}>{n.icon}</span>
            <span style={{ fontSize: 11, color: page === n.id ? T.orange : T.textXs, fontWeight: page === n.id ? 800 : 400 }}>{n.label}</span>
            {page === n.id && <div style={{ width: 20, height: 3, borderRadius: 2, background: T.orange }} />}
          </button>
        ))}
      </div>
    </div>
  );
}
