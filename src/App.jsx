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
// ROLES: owner | staff | franchise
const INIT_STAFF = [
  { id:"owner", name:"DR.Fresh (เจ้าของ)", pin:"1234", role:"owner", franchiseId:null, active:true,
    perms:{ cashflow:true, stock:true, purchase:true, report:true, ai:true, admin:true } },
  { id:"s1", name:"มิ้ว", pin:"1111", role:"staff", franchiseId:null, active:true,
    perms:{ cashflow:true, stock:true, purchase:true, report:false, ai:false, admin:false } },
  { id:"s2", name:"ปาล์ม", pin:"2222", role:"staff", franchiseId:null, active:true,
    perms:{ cashflow:true, stock:true, purchase:false, report:false, ai:false, admin:false } },
  { id:"s3", name:"เจ", pin:"3333", role:"staff", franchiseId:null, active:true,
    perms:{ cashflow:false, stock:true, purchase:false, report:false, ai:false, admin:false } },
  { id:"fr1", name:"แฟรนไชส์ เชียงใหม่", pin:"5555", role:"franchise", franchiseId:"fr1", active:true,
    perms:{ cashflow:true, stock:true, purchase:true, report:true, ai:false, admin:false } },
  { id:"fr2", name:"แฟรนไชส์ ขอนแก่น", pin:"6666", role:"franchise", franchiseId:"fr2", active:true,
    perms:{ cashflow:true, stock:true, purchase:true, report:true, ai:false, admin:false } },
  { id:"fr3", name:"แฟรนไชส์ ภูเก็ต", pin:"7777", role:"franchise", franchiseId:"fr3", active:false,
    perms:{ cashflow:true, stock:true, purchase:false, report:true, ai:false, admin:false } },
];

const INIT_FRANCHISES = [
  { id:"fr1", name:"ไท่กั๋วหม่าล่า เชียงใหม่", owner:"คุณสมชาย",  phone:"091-111-2222", openDate:"2025-01-15", royaltyPct:5, active:true,  monthlyTarget:150000 },
  { id:"fr2", name:"ไท่กั๋วหม่าล่า ขอนแก่น",  owner:"คุณสมหญิง", phone:"092-222-3333", openDate:"2025-03-01", royaltyPct:5, active:true,  monthlyTarget:120000 },
  { id:"fr3", name:"ไท่กั๋วหม่าล่า ภูเก็ต",   owner:"คุณสมศรี",  phone:"093-333-4444", openDate:"2025-06-01", royaltyPct:5, active:false, monthlyTarget:180000 },
];

const BRANCHES = [
  { id:"main", name:"สาขาหลัก" },
  { id:"fr1",  name:"เชียงใหม่" },
  { id:"fr2",  name:"ขอนแก่น"  },
  { id:"fr3",  name:"ภูเก็ต"   },
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
  { id:1, itemId:1, type:"in",  qty:20,  date:"2026-05-08", staffId:"s1", note:"รับจากซัพฯ", branch:"main" },
  { id:2, itemId:3, type:"out", qty:3,   date:"2026-05-09", staffId:"s2", note:"ใช้วันนี้",   branch:"main" },
  { id:3, itemId:5, type:"in",  qty:8,   date:"2026-05-06", staffId:"s1", note:"สั่งเพิ่ม",   branch:"main" },
  { id:4, itemId:2, type:"out", qty:2,   date:"2026-05-09", staffId:"s3", note:"ใช้บุฟเฟต์",  branch:"main" },
];

const INIT_CF = [
  { id:1, date:"2026-05-09", flow:"in",  cat:"ยอดขาย dine-in",     itemName:"58 ออร์เดอร์", amount:9200,  method:"เงินสด", note:"", branch:"main", staffId:"s1" },
  { id:2, date:"2026-05-09", flow:"in",  cat:"ยอดขาย delivery",     itemName:"GrabFood",      amount:3800,  method:"โอน",    note:"", branch:"main", staffId:"s1" },
  { id:3, date:"2026-05-09", flow:"out", cat:"วัตถุดิบ/ผัก",        itemName:"ผักกาดขาว",    amount:1400,  method:"เงินสด", note:"", branch:"main", staffId:"s1" },
  { id:4, date:"2026-05-08", flow:"in",  cat:"ยอดขาย dine-in",     itemName:"52 ออร์เดอร์", amount:8500,  method:"เงินสด", note:"", branch:"main", staffId:"s2" },
  { id:5, date:"2026-05-08", flow:"out", cat:"วัตถุดิบ/เนื้อสัตว์", itemName:"หมูสามชั้น",   amount:3600,  method:"โอน",    note:"", branch:"main", staffId:"s2" },
  { id:6, date:"2026-05-07", flow:"in",  cat:"ยอดขาย dine-in",     itemName:"48 ออร์เดอร์", amount:7800,  method:"เงินสด", note:"", branch:"main", staffId:"s1" },
  { id:7, date:"2026-05-07", flow:"out", cat:"วัตถุดิบ/ลูกชิ้น",   itemName:"ลูกชิ้นหมู",   amount:950,   method:"เงินสด", note:"", branch:"main", staffId:"s3" },
  { id:8, date:"2026-05-01", flow:"out", cat:"ค่าแรงพนักงาน",       itemName:"เงินเดือน",     amount:35000, method:"โอน",    note:"", branch:"main", staffId:"owner" },
  { id:9,  date:"2026-05-01", flow:"out", cat:"ค่าเช่า",         itemName:"ค่าเช่ารายเดือน",amount:4500,  method:"โอน",    note:"", branch:"main", staffId:"owner" },
  // แฟรนไชส์ เชียงใหม่
  { id:10, date:"2026-05-09", flow:"in",  cat:"ยอดขาย dine-in",  itemName:"65 ออร์เดอร์",  amount:12500, method:"เงินสด", note:"", branch:"fr1",  staffId:"fr1" },
  { id:11, date:"2026-05-09", flow:"in",  cat:"ยอดขาย delivery", itemName:"GrabFood",       amount:4200,  method:"โอน",    note:"", branch:"fr1",  staffId:"fr1" },
  { id:12, date:"2026-05-09", flow:"out", cat:"วัตถุดิบ/ผัก",    itemName:"ผักรวม",         amount:2100,  method:"เงินสด", note:"", branch:"fr1",  staffId:"fr1" },
  { id:13, date:"2026-05-08", flow:"in",  cat:"ยอดขาย dine-in",  itemName:"72 ออร์เดอร์",  amount:14000, method:"เงินสด", note:"", branch:"fr1",  staffId:"fr1" },
  // แฟรนไชส์ ขอนแก่น
  { id:14, date:"2026-05-09", flow:"in",  cat:"ยอดขาย dine-in",  itemName:"45 ออร์เดอร์",  amount:8500,  method:"เงินสด", note:"", branch:"fr2",  staffId:"fr2" },
  { id:15, date:"2026-05-09", flow:"in",  cat:"ยอดขาย delivery", itemName:"LINE MAN",       amount:3200,  method:"โอน",    note:"", branch:"fr2",  staffId:"fr2" },
  { id:16, date:"2026-05-08", flow:"in",  cat:"ยอดขาย dine-in",  itemName:"50 ออร์เดอร์",  amount:9500,  method:"เงินสด", note:"", branch:"fr2",  staffId:"fr2" },
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
            <div key={s.id} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:13,padding:"3px 0",color:T.textMd }}>
              <div style={{ display:"flex",alignItems:"center",gap:6 }}>
                <span style={{ fontSize:12 }}>{s.role==="owner"?"👑":s.role==="franchise"?"🏪":"👷"}</span>
                <span>{s.name}</span>
              </div>
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
    : user.role === "franchise"
    ? cf.filter(e => e.branch === user.franchiseId)
    : cf.filter(e => e.staffId === user.id || e.branch === "main");

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
function PurchasePage({ stock, lineToken, suppliers }) {
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
    const sup = suppliers.find(x => x.id === s.supplierId) || suppliers[0];
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
      const sup = suppliers.find(x => x.id === s.supplierId);
      return {
        "ซัพพลายเออร์": sup?.name, "วัตถุดิบ": s.name, "หน่วย": s.unit,
        "จำนวนสั่ง": orderQtys[s.id] || suggestQty(s),
        "ราคา/หน่วย": s.costHistory?.slice(-1)[0]?.unitCost || 0,
        "รวม": (orderQtys[s.id] || suggestQty(s)) * (s.costHistory?.slice(-1)[0]?.unitCost || 0),
        "order_date": today(),
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
            const sup = suppliers.find(x => x.id === s.supplierId) || suppliers[0];
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
          const rows=mCF.map(e=>({"date":e.date,"type":e.flow==="in"?"income":"expense","category":e.cat,"item":e.itemName||"","amount":e.amount,"staff":(staff.find(s=>s.id===e.staffId)||{}).name||e.staffId}));
          XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"report");
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
// ─────────────────────────────────────────────
// SUPPLIERS — editable from Settings
// ─────────────────────────────────────────────
const INIT_SUPPLIERS = [
  { id:1, name:"ตลาดสดนครชัย",    type:"ผัก",         line:"@vegmarket", phone:"081-234-5678", active:true },
  { id:2, name:"ฟาร์มหมูสยาม",    type:"หมู/เนื้อ",   line:"@siampork",  phone:"082-345-6789", active:true },
  { id:3, name:"อาหารทะเลสด",      type:"ทะเล",        line:"@freshsea",  phone:"083-456-7890", active:true },
  { id:4, name:"ซอสมาล่าพรีเมียม", type:"ซอส/บรรจุ",   line:"@malabase",  phone:"084-567-8901", active:true },
];

// ─────────────────────────────────────────────
// ENHANCED STOCK PAGE — with price tracking
// ─────────────────────────────────────────────
function StockPage({ stock, setStock, movements, setMovements, user, suppliers }) {
  const [tab, setTab] = useState("stock");
  const [selItem, setSelItem] = useState(null);
  const [mvForm, setMvForm] = useState({ type:"in", qty:"", unitCost:"", note:"" });
  const [showAdd, setShowAdd] = useState(false);
  const [editMinId, setEditMinId] = useState(null);
  const [editMinVal, setEditMinVal] = useState("");
  const [newItem, setNewItem] = useState({ name:"",unit:"kg",qty:0,minQty:3,dailyUse:1,expiryDays:3,supplierId:1 });
  const [filter, setFilter] = useState("all");
  const [importMsg, setImportMsg] = useState("");
  const fileRef = useRef();

  const sorted = useMemo(()=>{
    const ord = {out:0,critical:1,low:2,ok:3};
    let list = filter==="all"?stock:stock.filter(s=>stockSt(s)===filter);
    return [...list].sort((a,b)=>ord[stockSt(a)]-ord[stockSt(b)]);
  },[stock,filter]);

  // Add movement with price tracking
  const addMovement = () => {
    if (!mvForm.qty || !selItem) return;
    const qty = parseFloat(mvForm.qty);
    const unitCost = parseFloat(mvForm.unitCost) || 0;

    // Update stock quantity
    setStock(prev => prev.map(s => {
      if (s.id !== selItem.id) return s;
      const newQty = mvForm.type==="in" ? s.qty+qty : Math.max(0, s.qty-qty);
      // Add cost history if price provided on "in"
      const newHistory = mvForm.type==="in" && unitCost>0
        ? [...(s.costHistory||[]), { date:today(), unitCost, qty, total:unitCost*qty }]
        : s.costHistory;
      return { ...s, qty:newQty, costHistory:newHistory };
    }));

    // Log movement
    setMovements(prev => [...prev, {
      id:Date.now(), itemId:selItem.id, type:mvForm.type, qty,
      unitCost: mvForm.type==="in" ? unitCost : 0,
      date:today(), staffId:user.id, note:mvForm.note, branch:user.branch||"1"
    }]);

    setMvForm({type:"in",qty:"",unitCost:"",note:""});
    setSelItem(null);
  };

  // Save min qty edit
  const saveMinQty = (id) => {
    setStock(prev=>prev.map(s=>s.id===id?{...s,minQty:+editMinVal}:s));
    setEditMinId(null);
  };

  // Import Excel
  const handleImport = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, {type:"array"});
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        let added=0, updated=0;
        const newStock = [...stock];
        rows.forEach(r => {
          const name = r["ชื่อวัตถุดิบ"]||r["name"]||""; if (!name) return;
          const idx = newStock.findIndex(s=>s.name===name);
          if (idx>=0) {
            if (r["จำนวนคงเหลือ"]!==undefined) newStock[idx]={...newStock[idx],qty:+r["จำนวนคงเหลือ"]};
            if (r["จำนวนขั้นต่ำ"]!==undefined)  newStock[idx]={...newStock[idx],minQty:+r["จำนวนขั้นต่ำ"]};
            if (r["ใช้ต่อวัน"]!==undefined)      newStock[idx]={...newStock[idx],dailyUse:+r["ใช้ต่อวัน"]};
            updated++;
          } else {
            newStock.push({ id:Date.now()+Math.random(), name, unit:r["หน่วย"]||"kg",
              qty:+(r["จำนวนคงเหลือ"]||0), minQty:+(r["จำนวนขั้นต่ำ"]||3),
              dailyUse:+(r["ใช้ต่อวัน"]||1), expiryDays:+(r["อายุ (วัน)"]||3),
              supplierId:1, costHistory:r["ราคาต่อหน่วย"]?[{date:today(),unitCost:+r["ราคาต่อหน่วย"],qty:+(r["จำนวนคงเหลือ"]||1),total:+r["ราคาต่อหน่วย"]*(+(r["จำนวนคงเหลือ"]||1))}]:[],
            }); added++;
          }
        });
        setStock(newStock);
        setImportMsg(`ok:เพิ่ม ${added} อัพเดท ${updated} รายการ`);
        setTimeout(()=>setImportMsg(""),4000);
      } catch { setImportMsg("err"); setTimeout(()=>setImportMsg(""),4000); }
    };
    reader.readAsArrayBuffer(file); e.target.value="";
  };

  const handleExport = () => {
    const rows = stock.map(s=>({
      "ชื่อวัตถุดิบ":s.name,"หน่วย":s.unit,"จำนวนคงเหลือ":s.qty,
      "จำนวนขั้นต่ำ":s.minQty,"ใช้ต่อวัน":s.dailyUse,"อายุ (วัน)":s.expiryDays,
      "ราคาเฉลี่ย":+wac(s).toFixed(2),"ราคาล่าสุด":latestCost(s),
      "มูลค่าสต็อค":+(s.qty*wac(s)).toFixed(2),"สถานะ":ST_LABEL[stockSt(s)],
    }));
    const ws=XLSX.utils.json_to_sheet(rows); ws["!cols"]=[20,8,14,14,12,10,14,14,16,10].map(w=>({wch:w}));
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"stock");
    XLSX.writeFile(wb,`stock_${today()}.xlsx`);
  };

  const itemMovements = selItem ? [...movements].filter(m=>m.itemId===selItem.id).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,8) : [];

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <SectionTitle title="📦 สต็อควัตถุดิบ"
        action={
          <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display:"none" }} />
            <button onClick={()=>fileRef.current?.click()} style={{...S.ghost,fontSize:13,padding:"7px 12px"}}>📥 Import</button>
            <button onClick={handleExport} style={{...S.ghost,fontSize:13,padding:"7px 12px"}}>📤 Export</button>
            {user.perms?.admin && <button onClick={()=>setShowAdd(!showAdd)} style={S.btn()}>+ เพิ่ม</button>}
          </div>
        }
      />

      {importMsg && (
        <div style={{ background:importMsg.startsWith("ok")?T.greenLt:T.redLt, border:`1px solid ${importMsg.startsWith("ok")?T.green:T.red}44`, borderRadius:10, padding:"10px 16px", color:importMsg.startsWith("ok")?T.green:T.red, fontWeight:700, fontSize:15 }}>
          {importMsg.startsWith("ok") ? `✅ ${importMsg.slice(3)}` : "❌ ไฟล์ผิดรูปแบบ"}
        </div>
      )}

      <TabBar tabs={[["stock","รายการสต็อค"],["price","ราคาเฉลี่ย"],["movements","ประวัติ"]]} active={tab} onChange={setTab} />

      {/* Filter */}
      {tab==="stock" && (
        <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
          {[["all","ทั้งหมด"],["critical","🔴 วิกฤต"],["low","🟡 ต่ำ"],["ok","🟢 ปกติ"]].map(([v,l])=>(
            <button key={v} onClick={()=>setFilter(v)} style={{ background:filter===v?T.orange:"transparent", border:`1px solid ${filter===v?T.orange:T.border}`, borderRadius:8, padding:"7px 14px", color:filter===v?"#fff":T.textMd, cursor:"pointer", fontSize:14, fontFamily:"inherit" }}>{l}</button>
          ))}
        </div>
      )}

      {/* Add form */}
      {showAdd && tab==="stock" && (
        <Card style={{ borderColor:T.borderOr }}>
          <div style={{ color:T.orange,fontWeight:800,fontSize:17,marginBottom:12 }}>➕ เพิ่มวัตถุดิบ</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
            {[["ชื่อ","name","text"],["หน่วย","unit","text"],["จำนวน","qty","number"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"],["อายุ(วัน)","expiryDays","number"]].map(([l,k,t])=>(
              <div key={k}>
                <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>{l}</div>
                <input type={t} value={newItem[k]} onChange={e=>setNewItem(p=>({...p,[k]:e.target.value}))} style={S.inp} />
              </div>
            ))}
            <div style={{ gridColumn:"1/-1" }}>
              <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>ซัพพลายเออร์</div>
              <select value={newItem.supplierId} onChange={e=>setNewItem(p=>({...p,supplierId:+e.target.value}))} style={S.inp}>
                {suppliers.map(s=><option key={s.id} value={s.id}>{s.name} ({s.type})</option>)}
              </select>
            </div>
          </div>
          <div style={{ display:"flex",gap:8,marginTop:12 }}>
            <button onClick={()=>{setStock(prev=>[...prev,{...newItem,id:Date.now(),qty:+newItem.qty,minQty:+newItem.minQty,dailyUse:+newItem.dailyUse,expiryDays:+newItem.expiryDays,costHistory:[]}]);setShowAdd(false);}} style={{...S.btn(),flex:1}}>บันทึก</button>
            <button onClick={()=>setShowAdd(false)} style={S.ghost}>ยกเลิก</button>
          </div>
        </Card>
      )}

      {/* STOCK LIST TAB */}
      {tab==="stock" && sorted.map(item=>{
        const st=stockSt(item);
        const daysLeft=item.dailyUse>0?(item.qty/item.dailyUse).toFixed(1):"∞";
        const isOpen=selItem?.id===item.id;
        const avgCost=wac(item);
        const lastCost=latestCost(item);
        const spike=isSpike(item);
        const sup=suppliers.find(x=>x.id===item.supplierId);

        return (
          <Card key={item.id} style={{ borderColor:st!=="ok"?ST_COLOR[st]+"33":spike?"#f97316"+33:T.border }}>
            <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",cursor:"pointer" }} onClick={()=>setSelItem(isOpen?null:item)}>
              <div style={{ flex:1 }}>
                <div style={{ display:"flex",alignItems:"center",gap:8,flexWrap:"wrap" }}>
                  <span style={{ color:T.text,fontWeight:700,fontSize:17 }}>{item.name}</span>
                  <Badge status={st} />
                  {spike && <span style={{ background:"#fff7ed",border:"1px solid #fed7aa",borderRadius:6,padding:"2px 8px",fontSize:12,color:T.orange,fontWeight:700 }}>⚠️ ราคาผิดปกติ</span>}
                </div>
                <div style={{ color:T.textSm,fontSize:13,marginTop:4 }}>
                  ซัพฯ: {sup?.name||"-"} • เหลืออีก {daysLeft} วัน • ราคาเฉลี่ย ฿{avgCost.toFixed(2)}/{item.unit}
                </div>
                {/* Min qty edit */}
                <div style={{ display:"flex",alignItems:"center",gap:8,marginTop:4 }}>
                  <span style={{ color:T.textXs,fontSize:12 }}>ขั้นต่ำ:</span>
                  {editMinId===item.id ? (
                    <>
                      <input type="number" value={editMinVal} onChange={e=>setEditMinVal(e.target.value)} autoFocus
                        style={{...S.inp,width:60,padding:"3px 8px",fontSize:13}} onKeyDown={e=>e.key==="Enter"&&saveMinQty(item.id)} />
                      <button onClick={()=>saveMinQty(item.id)} style={{...S.btn(T.green),padding:"3px 10px",fontSize:12}}>✓</button>
                      <button onClick={()=>setEditMinId(null)} style={{...S.ghost,padding:"3px 8px",fontSize:12}}>✕</button>
                    </>
                  ) : (
                    <span onClick={e=>{e.stopPropagation();setEditMinId(item.id);setEditMinVal(item.minQty);}}
                      style={{ color:T.orange,fontSize:13,fontWeight:700,cursor:"pointer",textDecoration:"underline dotted" }}>
                      {item.minQty} {item.unit} ✏️
                    </span>
                  )}
                </div>
              </div>
              <div style={{ textAlign:"right",marginLeft:12,flexShrink:0 }}>
                <div style={{ color:ST_COLOR[st],fontWeight:900,fontSize:28 }}>{item.qty}</div>
                <div style={{ color:T.textSm,fontSize:14 }}>{item.unit}</div>
              </div>
            </div>

            {/* Movement form */}
            {isOpen && (
              <div style={{ marginTop:14,paddingTop:14,borderTop:`1px solid ${T.bg}` }}>
                <div style={{ color:T.text,fontWeight:700,fontSize:16,marginBottom:10 }}>บันทึกการเคลื่อนไหว</div>
                <div style={{ display:"flex",gap:8,marginBottom:10 }}>
                  {[["in","📥 รับเข้า",T.green],["out","📤 จ่ายออก",T.red]].map(([v,l,c])=>(
                    <button key={v} onClick={()=>setMvForm(p=>({...p,type:v}))} style={{ flex:1,padding:10,background:mvForm.type===v?c+"15":"transparent",border:`1.5px solid ${mvForm.type===v?c:T.border}`,borderRadius:10,color:mvForm.type===v?c:T.textMd,fontWeight:700,cursor:"pointer",fontSize:15,fontFamily:"inherit" }}>{l}</button>
                  ))}
                </div>
                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8 }}>
                  <div>
                    <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>จำนวน ({item.unit})</div>
                    <input type="number" value={mvForm.qty} onChange={e=>setMvForm(p=>({...p,qty:e.target.value}))} style={S.inp} placeholder="0" />
                  </div>
                  {mvForm.type==="in" && (
                    <div>
                      <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>ราคา/หน่วย (฿) <span style={{ color:T.orange,fontSize:11 }}>*สำคัญ</span></div>
                      <input type="number" value={mvForm.unitCost} onChange={e=>setMvForm(p=>({...p,unitCost:e.target.value}))} style={{...S.inp,borderColor:T.orange}} placeholder={`ล่าสุด ฿${lastCost||"-"}`} />
                    </div>
                  )}
                  <div style={{ gridColumn: mvForm.type==="out"?"1/-1":"auto" }}>
                    <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>หมายเหตุ</div>
                    <input type="text" value={mvForm.note} onChange={e=>setMvForm(p=>({...p,note:e.target.value}))} style={S.inp} placeholder="เช่น รับจากซัพฯ" />
                  </div>
                </div>

                {/* Price spike warning preview */}
                {mvForm.type==="in" && mvForm.unitCost && avgCost>0 && (+mvForm.unitCost > avgCost*1.15) && (
                  <div style={{ background:T.yellowLt,border:`1px solid ${T.yellow}44`,borderRadius:8,padding:"8px 12px",marginBottom:8,fontSize:13,color:T.yellow,fontWeight:600 }}>
                    ⚠️ ราคานี้ ({mvForm.unitCost}) สูงกว่าค่าเฉลี่ย ({avgCost.toFixed(2)}) {((+mvForm.unitCost-avgCost)/avgCost*100).toFixed(0)}% — ยืนยันได้เลยถ้าถูกต้อง
                  </div>
                )}

                {mvForm.type==="in" && mvForm.qty && mvForm.unitCost && (
                  <div style={{ background:T.greenLt,borderRadius:8,padding:"8px 12px",marginBottom:8,fontSize:14,color:T.green,fontWeight:600 }}>
                    รวมจ่าย: ฿{(+mvForm.qty * +mvForm.unitCost).toLocaleString("th-TH")} • ราคาเฉลี่ยใหม่จะเป็น: ฿{(()=>{
                      const totalOld=(item.costHistory||[]).reduce((a,b)=>a+b.total,0);
                      const qtyOld=(item.costHistory||[]).reduce((a,b)=>a+b.qty,0);
                      const newTotal=totalOld+(+mvForm.qty*+mvForm.unitCost);
                      const newQty=qtyOld+(+mvForm.qty);
                      return newQty>0?(newTotal/newQty).toFixed(2):"0";
                    })()}/{item.unit}
                  </div>
                )}

                <div style={{ display:"flex",gap:8 }}>
                  <button onClick={addMovement} style={{...S.btn(mvForm.type==="in"?T.green:T.red),flex:1,padding:10,fontSize:15}}>✅ บันทึก</button>
                  <button onClick={()=>setSelItem(null)} style={{...S.ghost,padding:"10px 14px"}}>ยกเลิก</button>
                </div>

                {/* Movement history */}
                {itemMovements.length>0 && (
                  <div style={{ marginTop:12 }}>
                    <div style={{ color:T.textSm,fontSize:13,marginBottom:6,fontWeight:600 }}>ประวัติล่าสุด</div>
                    {itemMovements.map(m=>(
                      <div key={m.id} style={{ display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:`1px solid ${T.bg}`,fontSize:13 }}>
                        <span style={{ color:T.textMd }}>{m.date} • {m.note||"-"}</span>
                        <div style={{ display:"flex",gap:10 }}>
                          {m.unitCost>0 && <span style={{ color:T.textXs }}>฿{m.unitCost}/{item.unit}</span>}
                          <span style={{ color:m.type==="in"?T.green:T.red,fontWeight:700 }}>{m.type==="in"?"+":"-"}{m.qty} {item.unit}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </Card>
        );
      })}

      {/* PRICE TRACKING TAB */}
      {tab==="price" && (
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          <Card style={{ background:T.orangeLt,borderColor:T.borderOr }}>
            <div style={{ color:T.orange,fontWeight:800,fontSize:17,marginBottom:4 }}>💹 ราคาเฉลี่ยถ่วงน้ำหนัก (WAC)</div>
            <div style={{ color:T.textSm,fontSize:13 }}>คำนวณจากทุกครั้งที่รับเข้า • แจ้งเตือนเมื่อราคาสูงกว่าค่าเฉลี่ย {">"} 15%</div>
          </Card>

          {/* Spike alerts */}
          {stock.filter(isSpike).length>0 && (
            <Card style={{ borderColor:T.red+"44",background:T.redLt }}>
              <div style={{ color:T.red,fontWeight:800,fontSize:16,marginBottom:10 }}>🚨 ราคาผิดปกติ ({stock.filter(isSpike).length} รายการ)</div>
              {stock.filter(isSpike).map(s=>(
                <div key={s.id} style={{ display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.border}`,fontSize:15 }}>
                  <span style={{ color:T.text,fontWeight:600 }}>{s.name}</span>
                  <div style={{ textAlign:"right" }}>
                    <div style={{ color:T.red,fontWeight:700 }}>ล่าสุด ฿{latestCost(s)}/{s.unit}</div>
                    <div style={{ color:T.textSm,fontSize:12 }}>เฉลี่ย ฿{wac(s).toFixed(2)} • สูงขึ้น {(((latestCost(s)-wac(s))/wac(s))*100).toFixed(1)}%</div>
                  </div>
                </div>
              ))}
            </Card>
          )}

          {/* All items price table */}
          <Card>
            <div style={{ display:"grid",gridTemplateColumns:"1.4fr 1fr 1fr 1fr 0.8fr",gap:4,padding:"6px 0",borderBottom:`2px solid ${T.border}`,fontSize:12,color:T.textSm,fontWeight:700 }}>
              {["วัตถุดิบ","ราคาล่าสุด","เฉลี่ย WAC","ซื้อ","สถานะ"].map(h=><span key={h}>{h}</span>)}
            </div>
            {[...stock].sort((a,b)=>a.name.localeCompare(b.name)).map(s=>{
              const spike=isSpike(s);
              const avg=wac(s);
              const last=latestCost(s);
              const hist=s.costHistory||[];
              return (
                <div key={s.id} style={{ display:"grid",gridTemplateColumns:"1.4fr 1fr 1fr 1fr 0.8fr",gap:4,padding:"8px 0",borderBottom:`1px solid ${T.bg}`,fontSize:14,alignItems:"center" }}>
                  <span style={{ color:T.text,fontWeight:600 }}>{s.name}</span>
                  <span style={{ color:spike?T.red:T.text,fontWeight:spike?700:400 }}>฿{last||"-"}</span>
                  <span style={{ color:T.green,fontWeight:600 }}>฿{avg>0?avg.toFixed(2):"-"}</span>
                  <span style={{ color:T.textSm }}>{hist.length} ครั้ง</span>
                  <span style={{ color:spike?T.red:T.textXs,fontSize:11,fontWeight:spike?700:400 }}>{spike?`+${(((last-avg)/avg)*100).toFixed(0)}%`:"ปกติ"}</span>
                </div>
              );
            })}
          </Card>
        </div>
      )}

      {/* MOVEMENTS TAB */}
      {tab==="movements" && (
        <Card>
          <div style={{ color:T.text,fontWeight:700,fontSize:17,marginBottom:12 }}>ประวัติการเคลื่อนไหวทั้งหมด</div>
          {[...movements].sort((a,b)=>b.date.localeCompare(a.date)).map(m=>{
            const item=stock.find(s=>s.id===m.itemId);
            return (
              <div key={m.id} style={{ display:"flex",alignItems:"center",gap:12,padding:"10px 0",borderBottom:`1px solid ${T.bg}` }}>
                <div style={{ width:32,height:32,borderRadius:9,background:m.type==="in"?T.greenLt:T.redLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0 }}>
                  {m.type==="in"?"📥":"📤"}
                </div>
                <div style={{ flex:1 }}>
                  <div style={{ color:T.text,fontSize:15,fontWeight:600 }}>{item?.name||"?"}</div>
                  <div style={{ color:T.textSm,fontSize:12 }}>{m.date}{m.unitCost>0?` • ฿${m.unitCost}/${item?.unit}`:""} • {m.note||"-"}</div>
                </div>
                <span style={{ color:m.type==="in"?T.green:T.red,fontWeight:700,fontSize:16 }}>
                  {m.type==="in"?"+":"-"}{m.qty} {item?.unit}
                </span>
              </div>
            );
          })}
          {movements.length===0 && <div style={{ color:T.textSm,textAlign:"center",padding:24 }}>ยังไม่มีประวัติ</div>}
        </Card>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// AI FORECAST PAGE
// ─────────────────────────────────────────────
function AIForecastPage({ stock, cf, movements, user }) {
  const [messages, setMessages] = useState([{
    role:"assistant",
    text:"สวัสดีครับ! ผมคือ AI ผู้ช่วยของไท่กั๋วหม่าล่า 🫕\n\nผมวิเคราะห์ข้อมูลร้านแบบ Real-time ได้ครับ ลองถามได้เลย:\n\n• \"พรุ่งนี้ควรเตรียมวัตถุดิบอะไรบ้าง?\"\n• \"เดือนนี้ต้นทุนสูงขึ้นไหม?\"\n• \"ราคาวัตถุดิบไหนผิดปกติ?\"\n• \"สต็อคอะไรที่ควรสั่งเพิ่ม?\"\n• \"วิเคราะห์กำไรขาดทุนให้หน่อย\""
  }]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const bottomRef = useRef();

  useEffect(()=>{ bottomRef.current?.scrollIntoView({behavior:"smooth"}); },[messages]);

  const quickQ = [
    "พรุ่งนี้ควรเตรียมวัตถุดิบอะไร?",
    "สต็อคอะไรควรสั่งด่วน?",
    "ราคาวัตถุดิบผิดปกติไหม?",
    "วิเคราะห์กำไรเดือนนี้",
    "แนะนำลดต้นทุนให้หน่อย",
  ];

  const buildContext = () => {
    const mk = today().slice(0,7);
    const mCF = cf.filter(e=>e.date.startsWith(mk));
    const mIn  = mCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
    const mOut = mCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
    const criticalStock = stock.filter(s=>["critical","out"].includes(stockSt(s))).map(s=>`${s.name}(เหลือ ${s.qty}${s.unit})`).join(", ");
    const spikes = stock.filter(isSpike).map(s=>`${s.name}(ราคาสูงขึ้น ${(((latestCost(s)-wac(s))/wac(s))*100).toFixed(0)}%)`).join(", ");
    const stockSummary = stock.map(s=>`${s.name}: ${s.qty}${s.unit} ใช้${s.dailyUse}/วัน เหลือ${s.dailyUse>0?(s.qty/s.dailyUse).toFixed(1):"∞"}วัน ราคาเฉลี่ย฿${wac(s).toFixed(2)}`).join("\n");
    const last7days = Array.from({length:7},(_,i)=>{ const d=new Date(); d.setDate(d.getDate()-i); const k=d.toISOString().split("T")[0]; return {date:k,income:cf.filter(e=>e.date===k&&e.flow==="in").reduce((a,b)=>a+b.amount,0)}; });

    return `คุณคือ AI ผู้ช่วยวิเคราะห์ธุรกิจร้านอาหาร "ไท่กั๋วหม่าล่า" ร้านหม่าล่าสั่งจานเดี่ยว ตั้งอยู่ที่ไทย

ข้อมูลปัจจุบัน (${today()}):
- รายรับเดือน ${mk}: ฿${mIn.toLocaleString()}
- รายจ่ายเดือน ${mk}: ฿${mOut.toLocaleString()}
- กำไรสุทธิ: ฿${(mIn-mOut-53500).toLocaleString()} (หักต้นทุนคงที่ 53,500)
- สต็อควิกฤต: ${criticalStock||"ไม่มี"}
- ราคาผิดปกติ: ${spikes||"ไม่มี"}
- ยอดขาย 7 วันล่าสุด: ${last7days.map(d=>`${d.date.slice(5)}:฿${d.income.toLocaleString()}`).join(", ")}

รายละเอียดสต็อค:
${stockSummary}

ตอบเป็นภาษาไทย กระชับ ใช้ตัวเลขจริง ให้คำแนะนำที่ปฏิบัติได้จริง มีการพยากรณ์ล่วงหน้า`;
  };

  const send = async () => {
    if (!input.trim()||loading) return;
    const userMsg=input.trim(); setInput("");
    setMessages(prev=>[...prev,{role:"user",text:userMsg}]);
    setLoading(true);
    try {
      const history=messages.map(m=>({role:m.role==="assistant"?"assistant":"user",content:m.text}));
      const res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,system:buildContext(),messages:[...history,{role:"user",content:userMsg}]})
      });
      const data=await res.json();
      const reply=data.content?.find(b=>b.type==="text")?.text||"ขอโทษครับ ไม่สามารถตอบได้ในขณะนี้";
      setMessages(prev=>[...prev,{role:"assistant",text:reply}]);
    } catch { setMessages(prev=>[...prev,{role:"assistant",text:"⚠️ เกิดข้อผิดพลาด กรุณาลองใหม่"}]); }
    setLoading(false);
  };

  return (
    <div style={{ display:"flex",flexDirection:"column",height:"calc(100vh-130px)" }}>
      <SectionTitle title="🤖 AI พยากรณ์ & วิเคราะห์" />

      {/* Stock forecast cards */}
      <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10,marginBottom:14 }}>
        {stock.filter(s=>s.dailyUse>0&&s.qty>0).slice(0,4).map(s=>{
          const daysLeft=s.qty/s.dailyUse;
          const urgency=daysLeft<=1?"🔴":daysLeft<=3?"🟡":"🟢";
          return (
            <Card key={s.id} style={{ padding:"12px 14px",borderColor:daysLeft<=1?T.red+"44":daysLeft<=3?T.yellow+"44":T.border }}>
              <div style={{ fontSize:18,marginBottom:2 }}>{urgency}</div>
              <div style={{ color:T.text,fontWeight:700,fontSize:14 }}>{s.name}</div>
              <div style={{ color:T.textSm,fontSize:12 }}>เหลือ {daysLeft.toFixed(1)} วัน</div>
              <div style={{ color:daysLeft<=2?T.red:T.textXs,fontSize:11,fontWeight:600 }}>
                {daysLeft<=1?"ต้องสั่งวันนี้!":daysLeft<=3?"ควรสั่งเร็วๆ นี้":"ปกติ"}
              </div>
            </Card>
          );
        })}
      </div>

      {/* Quick questions */}
      <div style={{ display:"flex",gap:6,flexWrap:"wrap",marginBottom:12 }}>
        {quickQ.map(q=>(
          <button key={q} onClick={()=>setInput(q)} style={{ background:T.orangeLt,border:`1px solid ${T.borderOr}`,borderRadius:20,padding:"6px 12px",color:T.orange,cursor:"pointer",fontSize:13,fontFamily:"inherit" }}>{q}</button>
        ))}
      </div>

      {/* Chat messages */}
      <div style={{ flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:10,marginBottom:12,maxHeight:380,minHeight:200 }}>
        {messages.map((m,i)=>(
          <div key={i} style={{ display:"flex",justifyContent:m.role==="user"?"flex-end":"flex-start" }}>
            <div style={{ maxWidth:"88%",background:m.role==="user"?T.orange:T.card,border:m.role==="assistant"?`1px solid ${T.border}`:"none",borderRadius:m.role==="user"?"14px 14px 4px 14px":"14px 14px 14px 4px",padding:"10px 14px",color:m.role==="user"?"#fff":T.text,fontSize:15,lineHeight:1.6,whiteSpace:"pre-wrap",boxShadow:T.shadow }}>
              {m.role==="assistant"&&<span style={{ marginRight:5 }}>🤖</span>}
              {m.text}
            </div>
          </div>
        ))}
        {loading && (
          <div style={{ display:"flex",justifyContent:"flex-start" }}>
            <div style={{ background:T.card,border:`1px solid ${T.border}`,borderRadius:"14px 14px 14px 4px",padding:"10px 16px",color:T.textSm,fontSize:15,boxShadow:T.shadow }}>
              ⏳ กำลังวิเคราะห์...
            </div>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      {/* Input */}
      <div style={{ display:"flex",gap:8 }}>
        <input value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&send()}
          style={{...S.inp,flex:1}} placeholder="ถามเกี่ยวกับสต็อค กำไร พยากรณ์..." disabled={loading} />
        <button onClick={send} disabled={!input.trim()||loading} style={{...S.btn(),padding:"9px 18px",opacity:!input.trim()||loading?0.5:1}}>ส่ง</button>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// SETTINGS PAGE — with supplier management
// ─────────────────────────────────────────────
function SettingsPage({ staff, setStaff, notifications, lineToken, setLineToken, suppliers, setSuppliers }) {
  const [tab, setTab] = useState("staff");
  const [editStaffId, setEditStaffId] = useState(null);
  const [editSupId, setEditSupId] = useState(null);
  const [editSupData, setEditSupData] = useState({});
  const [newSup, setNewSup] = useState({ name:"",type:"",line:"",phone:"",active:true });
  const [showAddSup, setShowAddSup] = useState(false);

  const PERM_LABELS = { cashflow:"💵 Cash Flow", stock:"📦 สต็อค", purchase:"🛒 สั่งซื้อ", report:"📊 รายงาน", ai:"🤖 AI", admin:"⚙️ Admin" };

  const togglePerm=(staffId,perm)=>setStaff(prev=>prev.map(s=>s.id===staffId?{...s,perms:{...s.perms,[perm]:!s.perms[perm]}}:s));
  const toggleActive=staffId=>setStaff(prev=>prev.map(s=>s.id===staffId?{...s,active:!s.active}:s));
  const updatePin=(staffId,pin)=>setStaff(prev=>prev.map(s=>s.id===staffId?{...s,pin}:s));
  const updateStaffName=(staffId,name)=>setStaff(prev=>prev.map(s=>s.id===staffId?{...s,name}:s));

  const saveSup=()=>{ setSuppliers(prev=>prev.map(s=>s.id===editSupId?{...s,...editSupData}:s)); setEditSupId(null); };
  const addSup=()=>{ setSuppliers(prev=>[...prev,{...newSup,id:Date.now()}]); setNewSup({name:"",type:"",line:"",phone:"",active:true}); setShowAddSup(false); };
  const toggleSupActive=id=>setSuppliers(prev=>prev.map(s=>s.id===id?{...s,active:!s.active}:s));

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <SectionTitle title="⚙️ ตั้งค่าระบบ" />
      <TabBar tabs={[["staff","👷 พนักงาน"],["supplier","🏪 ซัพพลายเออร์"],["notif","🔔 แจ้งเตือน"],["line","📲 LINE"]]} active={tab} onChange={setTab} />

      {/* STAFF TAB */}
      {tab==="staff" && (
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          {staff.filter(s=>s.role!=="owner").map(s=>(
            <Card key={s.id} style={{ borderColor:s.active?T.border:T.red+"33" }}>
              <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10 }}>
                <div>
                  <div style={{ color:T.text,fontWeight:700,fontSize:16 }}>{s.name}</div>
                  <div style={{ color:T.textSm,fontSize:13 }}>{BRANCHES.find(b=>b.id===s.branch)?.name}</div>
                </div>
                <div style={{ display:"flex",gap:8 }}>
                  <button onClick={()=>toggleActive(s.id)} style={{ background:s.active?T.greenLt:T.redLt,border:`1px solid ${s.active?T.green:T.red}44`,borderRadius:8,padding:"6px 12px",color:s.active?T.green:T.red,cursor:"pointer",fontSize:13,fontWeight:700,fontFamily:"inherit" }}>{s.active?"ใช้งาน":"ระงับ"}</button>
                  <button onClick={()=>setEditStaffId(editStaffId===s.id?null:s.id)} style={{...S.ghost,padding:"6px 10px",fontSize:13}}>✏️</button>
                </div>
              </div>
              {editStaffId===s.id && (
                <div style={{ background:T.bg,borderRadius:10,padding:"12px",marginBottom:10 }}>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
                    <div>
                      <div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>ชื่อ</div>
                      <input value={s.name} onChange={e=>updateStaffName(s.id,e.target.value)} style={{...S.inp,fontSize:14}} />
                    </div>
                    <div>
                      <div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>PIN</div>
                      <input maxLength={4} value={s.pin} onChange={e=>updatePin(s.id,e.target.value)} style={{...S.inp,fontSize:16,letterSpacing:4}} />
                    </div>
                  </div>
                </div>
              )}
              <div style={{ color:T.textSm,fontSize:13,marginBottom:8,fontWeight:600 }}>สิทธิ์การใช้งาน</div>
              <div style={{ display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:7 }}>
                {Object.entries(PERM_LABELS).map(([perm,label])=>(
                  <button key={perm} onClick={()=>togglePerm(s.id,perm)} style={{ background:s.perms[perm]?T.orange:"transparent",border:`1px solid ${s.perms[perm]?T.orange:T.border}`,borderRadius:8,padding:"7px 4px",color:s.perms[perm]?"#fff":T.textMd,cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:s.perms[perm]?700:400,textAlign:"center",transition:"all .15s" }}>{label}</button>
                ))}
              </div>
            </Card>
          ))}
        </div>
      )}

      {/* SUPPLIER TAB */}
      {tab==="supplier" && (
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          <button onClick={()=>setShowAddSup(!showAddSup)} style={{...S.btn(),width:"100%",padding:12,fontSize:15}}>+ เพิ่มซัพพลายเออร์</button>

          {showAddSup && (
            <Card style={{ borderColor:T.borderOr }}>
              <div style={{ color:T.orange,fontWeight:800,fontSize:16,marginBottom:10 }}>➕ เพิ่มซัพพลายเออร์ใหม่</div>
              <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
                {[["ชื่อ","name"],["ประเภท","type"],["LINE ID","line"],["เบอร์โทร","phone"]].map(([l,k])=>(
                  <div key={k}>
                    <div style={{ color:T.textSm,fontSize:13,marginBottom:3 }}>{l}</div>
                    <input value={newSup[k]} onChange={e=>setNewSup(p=>({...p,[k]:e.target.value}))} style={S.inp} placeholder={l} />
                  </div>
                ))}
              </div>
              <div style={{ display:"flex",gap:8,marginTop:10 }}>
                <button onClick={addSup} style={{...S.btn(),flex:1}}>บันทึก</button>
                <button onClick={()=>setShowAddSup(false)} style={S.ghost}>ยกเลิก</button>
              </div>
            </Card>
          )}

          {suppliers.map(s=>(
            <Card key={s.id} style={{ borderColor:s.active?T.border:T.red+"33" }}>
              {editSupId===s.id ? (
                <div>
                  <div style={{ color:T.orange,fontWeight:700,fontSize:15,marginBottom:10 }}>✏️ แก้ไข: {s.name}</div>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10 }}>
                    {[["ชื่อ","name"],["ประเภท","type"],["LINE ID","line"],["เบอร์โทร","phone"]].map(([l,k])=>(
                      <div key={k}>
                        <div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>{l}</div>
                        <input value={editSupData[k]??s[k]} onChange={e=>setEditSupData(p=>({...p,[k]:e.target.value}))} style={S.inp} />
                      </div>
                    ))}
                  </div>
                  <div style={{ display:"flex",gap:8 }}>
                    <button onClick={saveSup} style={{...S.btn(),flex:1}}>บันทึก</button>
                    <button onClick={()=>setEditSupId(null)} style={S.ghost}>ยกเลิก</button>
                  </div>
                </div>
              ) : (
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                  <div>
                    <div style={{ color:T.text,fontWeight:700,fontSize:16 }}>{s.name}</div>
                    <div style={{ color:T.textSm,fontSize:13 }}>{s.type} • {s.line} • {s.phone}</div>
                  </div>
                  <div style={{ display:"flex",gap:8 }}>
                    <button onClick={()=>toggleSupActive(s.id)} style={{ background:s.active?T.greenLt:T.redLt,border:`1px solid ${s.active?T.green:T.red}44`,borderRadius:7,padding:"5px 10px",color:s.active?T.green:T.red,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit" }}>{s.active?"เปิด":"ปิด"}</button>
                    <button onClick={()=>{setEditSupId(s.id);setEditSupData({});}} style={{...S.ghost,padding:"5px 10px",fontSize:12}}>✏️</button>
                  </div>
                </div>
              )}
            </Card>
          ))}
        </div>
      )}

      {/* NOTIFICATIONS TAB */}
      {tab==="notif" && (
        <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
          <Card style={{ background:T.orangeLt,borderColor:T.borderOr }}>
            <div style={{ color:T.orange,fontWeight:800,fontSize:16,marginBottom:10 }}>🔔 การแจ้งเตือนปัจจุบัน</div>
            {notifications.length===0
              ? <div style={{ color:T.textSm,textAlign:"center",padding:20,fontSize:15 }}>ไม่มีการแจ้งเตือน ✅</div>
              : notifications.map(n=>(
                <div key={n.id} style={{ display:"flex",gap:12,padding:"10px 0",borderBottom:`1px solid ${T.borderOr}` }}>
                  <span style={{ fontSize:22,flexShrink:0 }}>{n.icon}</span>
                  <div>
                    <div style={{ color:T.text,fontWeight:600,fontSize:15 }}>{n.title}</div>
                    <div style={{ color:T.textSm,fontSize:13 }}>{n.body}</div>
                  </div>
                </div>
              ))
            }
          </Card>
          <Card>
            <div style={{ color:T.text,fontWeight:800,fontSize:16,marginBottom:10 }}>ตั้งค่าการแจ้งเตือน</div>
            {[["สต็อคต่ำกว่าขั้นต่ำ","แจ้งทันที",true],["ราคาผิดปกติ (>15%)","แจ้งทันที",true],["สรุปยอดขายประจำวัน","23:00 น.",true],["รายงานรายสัปดาห์","จันทร์ 08:00",false]].map(([name,time,on])=>(
              <div key={name} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:`1px solid ${T.bg}` }}>
                <div>
                  <div style={{ color:T.text,fontSize:15 }}>{name}</div>
                  <div style={{ color:T.textSm,fontSize:12 }}>{time}</div>
                </div>
                <div style={{ width:44,height:24,borderRadius:12,background:on?T.orange:T.border,position:"relative",cursor:"pointer" }}>
                  <div style={{ position:"absolute",top:2,width:20,height:20,borderRadius:"50%",background:"#fff",boxShadow:T.shadow,left:on?22:2,transition:"left .2s" }}/>
                </div>
              </div>
            ))}
          </Card>
        </div>
      )}

      {/* LINE TAB */}
      {tab==="line" && (
        <Card>
          <div style={{ color:T.orange,fontWeight:800,fontSize:17,marginBottom:14 }}>📲 ตั้งค่า LINE Notify</div>
          {[["1️⃣ สมัคร Token","เปิด notify-bot.line.me → Login → Generate token → Copy"],["2️⃣ สร้าง Webhook","ใช้ Make.com (ฟรี) รับ webhook แล้วส่งต่อ LINE"],["3️⃣ ใส่ Token ด้านล่าง","ระบบจะส่งแจ้งเตือนผ่าน LINE อัตโนมัติ"]].map(([s,d])=>(
            <div key={s} style={{ background:T.bg,borderRadius:10,padding:"12px 14px",marginBottom:8 }}>
              <div style={{ color:T.orange,fontWeight:700,fontSize:14 }}>{s}</div>
              <div style={{ color:T.textMd,fontSize:13,marginTop:2 }}>{d}</div>
            </div>
          ))}
          <div style={{ marginTop:8 }}>
            <div style={{ color:T.textSm,fontSize:14,marginBottom:6 }}>LINE Notify Token</div>
            <input type="password" value={lineToken} onChange={e=>setLineToken(e.target.value)} style={S.inp} placeholder="ใส่ Token ที่นี่..." />
            {lineToken && <div style={{ color:T.green,fontSize:13,marginTop:6,fontWeight:600 }}>Token บันทึกแล้ว LINE พร้อมใช้งาน</div>}
            <div style={{ color:T.textSm,fontSize:14,marginTop:10,marginBottom:6 }}>Webhook URL (Make.com)</div>
            <input type="url" style={S.inp} placeholder="https://hook.make.com/xxxxx" />
            <button style={{...S.btn(),width:"100%",marginTop:10,padding:12,fontSize:15}}>บันทึก</button>
          </div>
        </Card>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// MAIN APP v7
// ─────────────────────────────────────────────

// ─────────────────────────────────────────────
// FRANCHISE DASHBOARD — เจ้าของดูภาพรวม / แฟรนไชส์ดูสาขาตัวเอง
// ─────────────────────────────────────────────
function FranchisePage({ cf, stock, user, franchises, setFranchises, staff, setStaff }) {
  const [tab, setTab] = useState(user.role==="franchise"?"mypage":"overview");
  const [editFrId, setEditFrId] = useState(null);
  const [editFrData, setEditFrData] = useState({});
  const [showAddFr, setShowAddFr] = useState(false);
  const [newFr, setNewFr] = useState({ name:"", owner:"", phone:"", royaltyPct:5, monthlyTarget:100000, active:true });

  const mk = today().slice(0,7);

  // ── เจ้าของ: ภาพรวมทุกแฟรนไชส์ ──
  const frStats = franchises.map(fr => {
    const frCF = cf.filter(e => e.branch===fr.id && e.date.startsWith(mk));
    const frIn  = frCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
    const frOut = frCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
    const royalty = frIn * (fr.royaltyPct/100);
    const pctTarget = fr.monthlyTarget>0 ? (frIn/fr.monthlyTarget*100).toFixed(0) : 0;
    return { ...fr, frIn, frOut, royalty, pctTarget };
  });

  const totalRoyalty = frStats.reduce((a,b)=>a+b.royalty,0);
  const totalFrSales = frStats.reduce((a,b)=>a+b.frIn,0);

  // ── แฟรนไชส์: ดูข้อมูลสาขาตัวเอง ──
  const myFr = franchises.find(f=>f.id===user.franchiseId);
  const myCF = cf.filter(e=>e.branch===user.franchiseId && e.date.startsWith(mk));
  const myIn  = myCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const myOut = myCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const myRoyalty = myIn * ((myFr?.royaltyPct||5)/100);

  const saveFr = () => {
    setFranchises(prev=>prev.map(f=>f.id===editFrId?{...f,...editFrData,royaltyPct:+editFrData.royaltyPct||f.royaltyPct,monthlyTarget:+editFrData.monthlyTarget||f.monthlyTarget}:f));
    setEditFrId(null);
  };
  const addFr = () => {
    const newId = "fr"+(Date.now());
    setFranchises(prev=>[...prev,{...newFr,id:newId,royaltyPct:+newFr.royaltyPct,monthlyTarget:+newFr.monthlyTarget}]);
    // Add franchise account to staff
    setStaff(prev=>[...prev,{id:newId,name:`แฟรนไชส์ ${newFr.name.split(" ").slice(-1)[0]||"ใหม่"}`,pin:"0000",role:"franchise",franchiseId:newId,active:true,perms:{cashflow:true,stock:true,purchase:true,report:true,ai:false,admin:false}}]);
    setNewFr({name:"",owner:"",phone:"",royaltyPct:5,monthlyTarget:100000,active:true});
    setShowAddFr(false);
  };

  // Franchise-only view
  if (user.role==="franchise") {
    return (
      <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
        <div>
          <div style={{ color:T.text,fontSize:22,fontWeight:900 }}>🏪 {myFr?.name||"สาขาของฉัน"}</div>
          <div style={{ color:T.textSm,fontSize:14 }}>เจ้าของ: {myFr?.owner} • เปิด: {myFr?.openDate}</div>
        </div>

        {/* My KPI */}
        <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12 }}>
          {[["💰","รายรับเดือนนี้",`฿${fmt(myIn)}`,T.green],["💸","รายจ่าย",`฿${fmt(myOut)}`,T.red],["📈","กำไร",`฿${fmt(myIn-myOut)}`,myIn-myOut>=0?T.green:T.red],["🤝","Royalty Fee",`฿${fmt(Math.round(myRoyalty))}`,T.orange]].map(([ic,l,v,c])=>(
            <Card key={l} style={{ padding:"14px 16px" }}>
              <div style={{ fontSize:22,marginBottom:3 }}>{ic}</div>
              <div style={{ color:T.textSm,fontSize:13 }}>{l}</div>
              <div style={{ color:c,fontWeight:800,fontSize:20 }}>{v}</div>
            </Card>
          ))}
        </div>

        {/* Target progress */}
        <Card style={{ background:T.orangeLt,borderColor:T.borderOr }}>
          <div style={{ display:"flex",justifyContent:"space-between",marginBottom:8 }}>
            <div style={{ color:T.orange,fontWeight:700,fontSize:15 }}>เป้ายอดขายเดือนนี้</div>
            <div style={{ textAlign:"right" }}>
              <div style={{ color:T.orange,fontWeight:800,fontSize:17 }}>฿{fmt(myIn)} / ฿{fmt(myFr?.monthlyTarget||0)}</div>
              <div style={{ color:T.textSm,fontSize:12 }}>{(myFr?.monthlyTarget>0?(myIn/myFr.monthlyTarget*100):0).toFixed(0)}% ของเป้า</div>
            </div>
          </div>
          <div style={{ background:"rgba(255,255,255,0.5)",borderRadius:8,height:12 }}>
            <div style={{ background:T.orange,height:"100%",borderRadius:8,width:`${Math.min(myFr?.monthlyTarget>0?(myIn/myFr.monthlyTarget*100):0,100)}%`,transition:"width .5s" }}/>
          </div>
        </Card>

        {/* Royalty info */}
        <Card>
          <div style={{ color:T.text,fontWeight:700,fontSize:16,marginBottom:10 }}>🤝 ค่า Royalty Fee</div>
          <div style={{ display:"flex",flexDirection:"column",gap:8 }}>
            {[["อัตรา Royalty",`${myFr?.royaltyPct||5}% ของรายรับ`],["รายรับเดือนนี้",`฿${fmt(myIn)}`],["Royalty ที่ต้องจ่าย",`฿${fmt(Math.round(myRoyalty))}`],["กำหนดจ่าย","ภายในวันที่ 5 ของเดือนถัดไป"]].map(([l,v])=>(
              <div key={l} style={{ display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${T.bg}`,fontSize:15 }}>
                <span style={{ color:T.textMd }}>{l}</span>
                <span style={{ color:T.text,fontWeight:600 }}>{v}</span>
              </div>
            ))}
          </div>
          <div style={{ marginTop:12,background:T.yellowLt,border:`1px solid ${T.yellow}44`,borderRadius:10,padding:"10px 14px",fontSize:14,color:T.yellow,fontWeight:600 }}>
            ⏰ กรุณาโอน ฿{fmt(Math.round(myRoyalty))} ภายในวันที่ 5/{+(mk.split("-")[1])+1||1}
          </div>
        </Card>
      </div>
    );
  }

  // Owner view
  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <SectionTitle title="🏪 จัดการแฟรนไชส์"
        action={<button onClick={()=>setShowAddFr(!showAddFr)} style={S.btn()}>+ เพิ่มแฟรนไชส์</button>}
      />

      {/* Add franchise form */}
      {showAddFr && (
        <Card style={{ borderColor:T.borderOr }}>
          <div style={{ color:T.orange,fontWeight:800,fontSize:16,marginBottom:12 }}>➕ เพิ่มแฟรนไชส์ใหม่</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
            {[["ชื่อสาขา","name"],["ชื่อเจ้าของ","owner"],["เบอร์โทร","phone"]].map(([l,k])=>(
              <div key={k} style={{ gridColumn:k==="name"?"1/-1":"auto" }}>
                <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>{l}</div>
                <input value={newFr[k]} onChange={e=>setNewFr(p=>({...p,[k]:e.target.value}))} style={S.inp} placeholder={l} />
              </div>
            ))}
            <div>
              <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>Royalty (%)</div>
              <input type="number" value={newFr.royaltyPct} onChange={e=>setNewFr(p=>({...p,royaltyPct:e.target.value}))} style={S.inp} />
            </div>
            <div>
              <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>เป้ายอดขาย/เดือน</div>
              <input type="number" value={newFr.monthlyTarget} onChange={e=>setNewFr(p=>({...p,monthlyTarget:e.target.value}))} style={S.inp} />
            </div>
          </div>
          <div style={{ background:T.blueLt,borderRadius:8,padding:"8px 12px",marginTop:8,fontSize:13,color:T.blue }}>
            ระบบจะสร้าง account แฟรนไชส์ให้อัตโนมัติ PIN เริ่มต้น: 0000 (แก้ได้ใน Settings)
          </div>
          <div style={{ display:"flex",gap:8,marginTop:12 }}>
            <button onClick={addFr} style={{...S.btn(),flex:1}}>สร้างแฟรนไชส์</button>
            <button onClick={()=>setShowAddFr(false)} style={S.ghost}>ยกเลิก</button>
          </div>
        </Card>
      )}

      {/* Summary KPI */}
      <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12 }}>
        {[["🏪","แฟรนไชส์ทั้งหมด",`${franchises.filter(f=>f.active).length} สาขา`,T.blue],["💰","ยอดขายรวม",`฿${fmt(totalFrSales)}`,T.green],["🤝","Royalty รวม",`฿${fmt(Math.round(totalRoyalty))}`,T.orange]].map(([ic,l,v,c])=>(
          <Card key={l} style={{ padding:"14px 16px" }}>
            <div style={{ fontSize:22,marginBottom:3 }}>{ic}</div>
            <div style={{ color:T.textSm,fontSize:13 }}>{l}</div>
            <div style={{ color:c,fontWeight:800,fontSize:20 }}>{v}</div>
          </Card>
        ))}
      </div>

      {/* Franchise list */}
      {frStats.map(fr=>(
        <Card key={fr.id} style={{ borderColor:fr.active?T.border:T.red+"33" }}>
          {editFrId===fr.id ? (
            <div>
              <div style={{ color:T.orange,fontWeight:700,fontSize:15,marginBottom:10 }}>✏️ แก้ไข: {fr.name}</div>
              <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10 }}>
                {[["ชื่อสาขา","name"],["เจ้าของ","owner"],["เบอร์โทร","phone"],["Royalty (%)","royaltyPct"],["เป้า/เดือน","monthlyTarget"]].map(([l,k])=>(
                  <div key={k}>
                    <div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>{l}</div>
                    <input value={editFrData[k]??fr[k]} onChange={e=>setEditFrData(p=>({...p,[k]:e.target.value}))} style={{...S.inp,fontSize:14}} />
                  </div>
                ))}
              </div>
              <div style={{ display:"flex",gap:8 }}>
                <button onClick={saveFr} style={{...S.btn(),flex:1}}>บันทึก</button>
                <button onClick={()=>setEditFrId(null)} style={S.ghost}>ยกเลิก</button>
              </div>
            </div>
          ) : (
            <>
              <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10 }}>
                <div>
                  <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                    <span style={{ color:T.text,fontWeight:800,fontSize:17 }}>{fr.name}</span>
                    <span style={{ background:fr.active?T.greenLt:T.redLt,color:fr.active?T.green:T.red,borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700 }}>{fr.active?"เปิด":"ปิด"}</span>
                  </div>
                  <div style={{ color:T.textSm,fontSize:13,marginTop:2 }}>เจ้าของ: {fr.owner} • {fr.phone} • Royalty {fr.royaltyPct}%</div>
                </div>
                <button onClick={()=>{setEditFrId(fr.id);setEditFrData({});}} style={{...S.ghost,padding:"6px 10px",fontSize:13}}>✏️</button>
              </div>

              {/* Sales vs target */}
              <div style={{ marginBottom:8 }}>
                <div style={{ display:"flex",justifyContent:"space-between",marginBottom:4,fontSize:14 }}>
                  <span style={{ color:T.textMd }}>ยอดขาย {mk}</span>
                  <span style={{ color:T.orange,fontWeight:700 }}>฿{fmt(fr.frIn)} / ฿{fmt(fr.monthlyTarget)} ({fr.pctTarget}%)</span>
                </div>
                <div style={{ background:T.bg,borderRadius:6,height:10 }}>
                  <div style={{ background:`linear-gradient(90deg,${T.orange},${T.orangeDk})`,width:`${Math.min(+fr.pctTarget,100)}%`,height:"100%",borderRadius:6,transition:"width .5s" }}/>
                </div>
              </div>

              {/* Stats row */}
              <div style={{ display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,paddingTop:8,borderTop:`1px solid ${T.bg}` }}>
                {[["รายรับ",`฿${fmt(fr.frIn)}`,T.green],["รายจ่าย",`฿${fmt(fr.frOut)}`,T.red],["กำไร",`฿${fmt(fr.frIn-fr.frOut)}`,fr.frIn-fr.frOut>=0?T.green:T.red],["Royalty",`฿${fmt(Math.round(fr.royalty))}`,T.orange]].map(([l,v,c])=>(
                  <div key={l} style={{ textAlign:"center" }}>
                    <div style={{ color:T.textXs,fontSize:11 }}>{l}</div>
                    <div style={{ color:c,fontWeight:700,fontSize:14 }}>{v}</div>
                  </div>
                ))}
              </div>
            </>
          )}
        </Card>
      ))}
    </div>
  );
}

export default function App() {
  const [user, setUser]             = useState(null);
  const [page, setPage]             = useState("dashboard");
  const [stock, setStock]           = useState(INIT_STOCK);
  const [cf, setCF]                 = useState(INIT_CF);
  const [movements, setMovements]   = useState(INIT_MOVEMENTS);
  const [staff, setStaff]           = useState(INIT_STAFF);
  const [suppliers, setSuppliers]   = useState(INIT_SUPPLIERS);
  const [franchises, setFranchises] = useState(INIT_FRANCHISES);
  const [notifOpen, setNotifOpen]   = useState(false);
  const [lineToken, setLineToken]   = useState("");

  const notifications = useMemo(()=>buildNotifications(stock,cf,movements,staff),[stock,cf,movements]);

  if (!user) return <LoginPage staff={staff} onLogin={u=>{setUser(u);setPage("dashboard");}} />;

  const p = user.perms;
  const ownerNav = [
    {id:"dashboard", icon:"🏠",label:"หลัก"},
    {id:"cashflow",  icon:"💵",label:"Cash Flow"},
    {id:"stock",     icon:"📦",label:"สต็อค"},
    {id:"purchase",  icon:"🛒",label:"สั่งซื้อ"},
    {id:"report",    icon:"📊",label:"รายงาน"},
    {id:"franchise", icon:"🏪",label:"แฟรนไชส์"},
    {id:"ai",        icon:"🤖",label:"AI"},
    {id:"settings",  icon:"⚙️",label:"ตั้งค่า"},
  ];
  const franchiseNav = [
    {id:"dashboard", icon:"🏠",label:"หลัก"},
    {id:"franchise", icon:"🏪",label:"สาขาของฉัน"},
    ...(p.cashflow ?[{id:"cashflow",icon:"💵",label:"Cash Flow"}]:[]),
    ...(p.stock    ?[{id:"stock",   icon:"📦",label:"สต็อค"}]:[]),
    ...(p.report   ?[{id:"report",  icon:"📊",label:"รายงาน"}]:[]),
  ];
  const staffNav = [
    {id:"dashboard",icon:"🏠",label:"หลัก"},
    ...(p.cashflow ?[{id:"cashflow",icon:"💵",label:"Cash Flow"}]:[]),
    ...(p.stock    ?[{id:"stock",   icon:"📦",label:"สต็อค"}]:[]),
    ...(p.purchase ?[{id:"purchase",icon:"🛒",label:"สั่งซื้อ"}]:[]),
    ...(p.ai       ?[{id:"ai",      icon:"🤖",label:"AI"}]:[]),
  ];
  const nav = user.role==="owner" ? ownerNav : user.role==="franchise" ? franchiseNav : staffNav;

  const pages = {
    dashboard: <Dashboard stock={stock} cf={cf} movements={movements} user={user} staff={staff} notifications={notifications} lineToken={lineToken} />,
    cashflow:  <CashflowPage cf={cf} setCF={setCF} user={user} />,
    stock:     <StockPage stock={stock} setStock={setStock} movements={movements} setMovements={setMovements} user={user} suppliers={suppliers} />,
    purchase:  <PurchasePage stock={stock} lineToken={lineToken} suppliers={suppliers} />,
    report:    <ReportPage cf={cf} stock={stock} movements={movements} staff={staff} user={user} />,
    franchise: <FranchisePage cf={cf} stock={stock} user={user} franchises={franchises} setFranchises={setFranchises} staff={staff} setStaff={setStaff} />,
    ai:        <AIForecastPage stock={stock} cf={cf} movements={movements} user={user} />,
    settings:  <SettingsPage staff={staff} setStaff={setStaff} notifications={notifications} lineToken={lineToken} setLineToken={setLineToken} suppliers={suppliers} setSuppliers={setSuppliers} />,
  };

  const alertCount = notifications.filter(n=>n.type==="danger"||n.type==="warn").length;

  return (
    <div style={{ minHeight:"100vh",background:T.bg,fontFamily:"'Noto Sans Thai','Noto Sans',sans-serif",color:T.text,fontSize:17 }}>
      {/* Header */}
      <div style={{ background:T.card,borderBottom:`1px solid ${T.border}`,padding:"11px 18px",display:"flex",alignItems:"center",gap:12,position:"sticky",top:0,zIndex:50,boxShadow:T.shadow }}>
        <div style={{ width:34,height:34,borderRadius:9,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:17 }}>🫕</div>
        <div>
          <div style={{ color:T.orange,fontWeight:900,fontSize:15 }}>ไท่กั๋วหม่าล่า</div>
          <div style={{ color:T.textXs,fontSize:11 }}>
            {user.role==="owner"?"👑 เจ้าของ":user.role==="franchise"?"🏪 แฟรนไชส์":"👷 พนักงาน"}
          </div>
        </div>
        <button onClick={()=>setNotifOpen(!notifOpen)} style={{ marginLeft:"auto",position:"relative",background:"transparent",border:"none",cursor:"pointer",padding:8,borderRadius:10 }}>
          <span style={{ fontSize:22 }}>🔔</span>
          {alertCount>0 && <span style={{ position:"absolute",top:4,right:4,background:T.red,color:"#fff",borderRadius:"50%",width:17,height:17,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800 }}>{alertCount}</span>}
        </button>
        <button onClick={()=>setUser(null)} style={{ background:T.bg,border:`1px solid ${T.border}`,borderRadius:8,padding:"5px 11px",color:T.textMd,cursor:"pointer",fontSize:13,fontFamily:"inherit" }}>ออก</button>
      </div>

      {/* Notification dropdown */}
      {notifOpen && <div onClick={()=>setNotifOpen(false)} style={{ position:"fixed",inset:0,zIndex:190 }}/>}
      {notifOpen && (
        <div style={{ position:"fixed",top:62,right:12,zIndex:200,width:320,background:T.card,border:`1px solid ${T.border}`,borderRadius:14,boxShadow:T.shadowMd,padding:"14px 16px",maxHeight:"60vh",overflowY:"auto" }}>
          <div style={{ color:T.text,fontWeight:800,fontSize:15,marginBottom:10 }}>🔔 การแจ้งเตือน</div>
          {notifications.length===0
            ? <div style={{ color:T.textSm,textAlign:"center",padding:16 }}>ไม่มีการแจ้งเตือน ✅</div>
            : notifications.map(n=>(
              <div key={n.id} style={{ display:"flex",gap:10,padding:"8px 0",borderBottom:`1px solid ${T.bg}` }}>
                <span style={{ fontSize:20,flexShrink:0 }}>{n.icon}</span>
                <div>
                  <div style={{ color:T.text,fontSize:14,fontWeight:600 }}>{n.title}</div>
                  <div style={{ color:T.textSm,fontSize:12 }}>{n.body}</div>
                </div>
              </div>
            ))
          }
        </div>
      )}

      {/* Content */}
      <div style={{ padding:"18px 14px 95px",maxWidth:900,margin:"0 auto" }}>
        {pages[page]||pages["dashboard"]}
      </div>

      {/* Bottom nav */}
      <div style={{ position:"fixed",bottom:0,left:0,right:0,background:T.card,borderTop:`1px solid ${T.border}`,display:"flex",justifyContent:"space-around",padding:"9px 0 15px",zIndex:100,boxShadow:"0 -2px 8px rgba(0,0,0,0.06)" }}>
        {nav.map(n=>(
          <button key={n.id} onClick={()=>setPage(n.id)} style={{ background:"none",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2,padding:"2px 5px",minWidth:44 }}>
            <span style={{ fontSize:20 }}>{n.icon}</span>
            <span style={{ fontSize:10,color:page===n.id?T.orange:T.textXs,fontWeight:page===n.id?800:400 }}>{n.label}</span>
            {page===n.id && <div style={{ width:18,height:3,borderRadius:2,background:T.orange }}/>}
          </button>
        ))}
      </div>
    </div>
  );
}
