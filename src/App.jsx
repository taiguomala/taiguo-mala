import { useState, useRef, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── SUPABASE ───────────────────────────────
const SUPA_URL = "https://klmowpluuvjmbvvmqzep.supabase.co";
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtsbW93cGx1dXZqbWJ2dm1xemVwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzgzMTIzMTMsImV4cCI6MjA5Mzg4ODMxM30.aXQz6WBqE8US5_-ij6GvvY0XaCykMag8x6W2a6uAwMU";
async function supaFetch(path, opts={}) {
  try {
    const res = await fetch(`${SUPA_URL}/rest/v1/${path}`, {
      ...opts,
      headers: { "apikey":SUPA_KEY, "Authorization":`Bearer ${SUPA_KEY}`,
        "Content-Type":"application/json", "Prefer":opts.prefer||"return=representation", ...opts.headers }
    });
    if (!res.ok) return null;
    const t = await res.text(); return t ? JSON.parse(t) : [];
  } catch { return null; }
}
const db = {
  getCF:       ()    => supaFetch("cashflow?order=date.desc,id.desc"),
  addCF:       (r)   => supaFetch("cashflow", { method:"POST", body:JSON.stringify(r) }),
  delCF:       (id)  => supaFetch(`cashflow?id=eq.${id}`, { method:"DELETE", prefer:"" }),
  clearCF:     ()    => supaFetch("cashflow?id=gt.0", { method:"DELETE", prefer:"" }),
  getStock:    ()    => supaFetch("stock?order=name.asc"),
  upsertStock: (rows)=> supaFetch("stock", { method:"POST", body:JSON.stringify(rows), headers:{"Prefer":"resolution=merge-duplicates,return=representation"} }),
  getMvs:      ()    => supaFetch("movements?order=date.desc,id.desc"),
  addMv:       (r)   => supaFetch("movements", { method:"POST", body:JSON.stringify(r) }),
  clearMvs:    ()    => supaFetch("movements?id=gt.0", { method:"DELETE", prefer:"" }),
};

// ─── THEME ──────────────────────────────────
const T = {
  bg:"#f4f4f5", card:"#ffffff", border:"#e4e4e7", text:"#18181b",
  textMd:"#52525b", textSm:"#71717a", textXs:"#a1a1aa",
  orange:"#f97316", orangeDk:"#ea580c", orangeLt:"#fff7ed", borderOr:"#fed7aa",
  green:"#16a34a", greenLt:"#f0fdf4", red:"#dc2626", redLt:"#fef2f2",
  yellow:"#d97706", yellowLt:"#fffbeb", blue:"#2563eb", blueLt:"#eff6ff",
  shadow:"0 1px 3px rgba(0,0,0,0.08)", shadowMd:"0 4px 12px rgba(0,0,0,0.12)",
};
const S = {
  card: { background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:"16px 18px", boxShadow:T.shadow },
  inp:  { width:"100%", padding:"10px 12px", border:`1px solid ${T.border}`, borderRadius:10,
          fontSize:15, fontFamily:"'Noto Sans Thai','Noto Sans',sans-serif", background:T.card,
          outline:"none", boxSizing:"border-box", color:T.text },
  btn:  (bg=T.orange) => ({ background:bg, color:"#fff", border:"none", borderRadius:10,
          padding:"9px 18px", cursor:"pointer", fontWeight:700, fontSize:15,
          fontFamily:"'Noto Sans Thai','Noto Sans',sans-serif" }),
  ghost: { background:"transparent", border:`1px solid ${T.border}`, borderRadius:10,
           padding:"9px 14px", cursor:"pointer", fontSize:14, color:T.textMd,
           fontFamily:"'Noto Sans Thai','Noto Sans',sans-serif" },
};

// ─── CONSTANTS ──────────────────────────────
const today = () => new Date().toISOString().split("T")[0];
const fmt = n => Math.round(n).toLocaleString("th-TH");
const IN_CATS  = ["ยอดขาย dine-in","ยอดขาย delivery","เงินโอนเข้า","อื่นๆ รับเข้า"];
const OUT_CATS = ["วัตถุดิบ/ผัก","หมู/เนื้อ/ทะเล","ค่าเช่า","ค่าพนักงาน","ค่าไฟ/น้ำ","บรรจุภัณฑ์","ค่าขนส่ง","อื่นๆ จ่ายออก"];
const PAY_METHODS = ["เงินสด","โอนธนาคาร","QR Code","GrabFood","LINE MAN","Shopee Food","บัตรเครดิต"];

const INIT_STOCK = [
  { id:1,  name:"หมูสามชั้น",    unit:"kg",  qty:15, minQty:5,  dailyUse:3, supplierId:2, costHistory:[] },
  { id:2,  name:"กุ้งแวนนาไม",   unit:"kg",  qty:8,  minQty:4,  dailyUse:2, supplierId:3, costHistory:[] },
  { id:3,  name:"เต้าหู้ขาว",    unit:"kg",  qty:10, minQty:5,  dailyUse:2, supplierId:1, costHistory:[] },
  { id:4,  name:"ผักกาดขาว",     unit:"kg",  qty:6,  minQty:3,  dailyUse:2, supplierId:1, costHistory:[] },
  { id:5,  name:"เห็ดรวม",       unit:"kg",  qty:4,  minQty:3,  dailyUse:1, supplierId:1, costHistory:[] },
  { id:6,  name:"น้ำซุปมาล่า",   unit:"ถุง", qty:20, minQty:8,  dailyUse:4, supplierId:4, costHistory:[] },
  { id:7,  name:"วุ้นเส้น",      unit:"kg",  qty:5,  minQty:3,  dailyUse:1, supplierId:1, costHistory:[] },
  { id:8,  name:"ลูกชิ้นรวม",    unit:"kg",  qty:10, minQty:4,  dailyUse:2, supplierId:2, costHistory:[] },
];

const INIT_SUPPLIERS = [
  { id:1, name:"ตลาดสดนครชัย",    type:"ผัก",       phone:"081-234-5678", active:true },
  { id:2, name:"ฟาร์มหมูสยาม",    type:"หมู/เนื้อ", phone:"082-345-6789", active:true },
  { id:3, name:"อาหารทะเลสด",      type:"ทะเล",      phone:"083-456-7890", active:true },
  { id:4, name:"ซอสมาล่าพรีเมียม", type:"ซอส/บรรจุ", phone:"084-567-8901", active:true },
];

const INIT_STAFF = [
  { id:"owner",  name:"DR.Fresh (เจ้าของ)", pin:"1234", role:"owner",     franchiseId:null, active:true,
    perms:{ cashflow:true, stock:true, purchase:true, report:true, ai:true, admin:true, viewPrice:true } },
  { id:"s1",     name:"มิ้ว",              pin:"1111", role:"staff",     franchiseId:null, active:true,
    perms:{ cashflow:true, stock:true, purchase:false, report:false, ai:false, admin:false, viewPrice:false } },
  { id:"s2",     name:"ปาล์ม",             pin:"2222", role:"staff",     franchiseId:null, active:true,
    perms:{ cashflow:true, stock:true, purchase:false, report:false, ai:false, admin:false, viewPrice:false } },
  { id:"s3",     name:"เจ",               pin:"3333", role:"staff",     franchiseId:null, active:true,
    perms:{ cashflow:false, stock:true, purchase:false, report:false, ai:false, admin:false, viewPrice:false } },
  { id:"fr1",    name:"แฟรนไชส์ เชียงใหม่", pin:"5555", role:"franchise", franchiseId:"fr1", active:true,
    perms:{ cashflow:true, stock:true, purchase:false, report:true, ai:false, admin:false, viewPrice:true } },
  { id:"fr2",    name:"แฟรนไชส์ ขอนแก่น",  pin:"6666", role:"franchise", franchiseId:"fr2", active:true,
    perms:{ cashflow:true, stock:true, purchase:false, report:true, ai:false, admin:false, viewPrice:true } },
  { id:"emergency", name:"Emergency Reset", pin:"0000", role:"owner", franchiseId:null, active:true,
    perms:{ cashflow:false, stock:false, purchase:false, report:false, ai:false, admin:true, viewPrice:false } },
];

const INIT_FRANCHISES = [
  { id:"fr1", name:"ไท่กั๋วหม่าล่า เชียงใหม่", owner:"คุณสมชาย",  phone:"091-111-2222", openDate:"2025-01-15", royaltyPct:5, active:true,  monthlyTarget:150000 },
  { id:"fr2", name:"ไท่กั๋วหม่าล่า ขอนแก่น",  owner:"คุณสมหญิง", phone:"092-222-3333", openDate:"2025-03-01", royaltyPct:5, active:true,  monthlyTarget:120000 },
];

const INIT_FIXED = [
  { name:"ค่าเช่า",        amount:4500  },
  { name:"ค่าพนักงาน",      amount:35000 },
  { name:"ค่าไฟ",          amount:8000  },
  { name:"เครื่องล้างจาน", amount:4000  },
  { name:"ค่าน้ำ",         amount:1000  },
  { name:"อื่นๆ",          amount:1000  },
];

// ─── STOCK HELPERS ──────────────────────────
const stockSt = s => s.qty <= 0 ? "out" : s.qty < s.minQty * 0.5 ? "critical" : s.qty < s.minQty ? "low" : "ok";
const ST_COLOR = { ok:T.green, low:T.yellow, critical:T.red, out:T.red };
const ST_LABEL = { ok:"ปกติ", low:"ใกล้หมด", critical:"น้อยมาก", out:"หมด" };
const wac = s => { const h=s.costHistory||[]; const t=h.reduce((a,b)=>a+b.total,0); const q=h.reduce((a,b)=>a+b.qty,0); return q>0?t/q:0; };

// ─── UI COMPONENTS ──────────────────────────
function Card({ children, style={}, onClick }) {
  return <div style={{...S.card,...style}} onClick={onClick}>{children}</div>;
}
function Badge({ status }) {
  const colors = { ok:[T.green,T.greenLt], low:[T.yellow,T.yellowLt], critical:[T.red,T.redLt], out:[T.red,T.redLt] };
  const [c,bg] = colors[status]||colors.ok;
  return <span style={{ background:bg,color:c,borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700 }}>{ST_LABEL[status]}</span>;
}
function Tabs({ tabs, active, onChange }) {
  return (
    <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
      {tabs.map(([v,l])=>(
        <button key={v} onClick={()=>onChange(v)} style={{
          padding:"8px 14px", borderRadius:8, border:`1px solid ${active===v?T.orange:T.border}`,
          background:active===v?T.orange:"transparent", color:active===v?"#fff":T.textMd,
          cursor:"pointer", fontSize:13, fontFamily:"inherit", fontWeight:active===v?700:400
        }}>{l}</button>
      ))}
    </div>
  );
}
function Hdr({ title, action }) {
  return (
    <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10 }}>
      <div style={{ color:T.text,fontSize:20,fontWeight:900 }}>{title}</div>
      {action && <div>{action}</div>}
    </div>
  );
}
function Msg({ text, ok }) {
  if (!text) return null;
  return (
    <div style={{ background:ok?T.greenLt:T.yellowLt, border:`1px solid ${ok?T.green:T.yellow}44`,
      borderRadius:8, padding:"10px 14px", fontSize:14, fontWeight:600, color:ok?T.green:T.yellow }}>
      {text}
    </div>
  );
}

// ─── LOGIN PAGE ─────────────────────────────
function LoginPage({ staff, onLogin }) {
  const [pin, setPin] = useState("");
  const [err, setErr] = useState(false);
  const go = () => {
    const u = staff.find(s => s.pin === pin && s.active);
    if (u) { onLogin(u); }
    else { setErr(true); setTimeout(()=>setErr(false),1500); }
  };
  return (
    <div style={{ minHeight:"100vh",background:`linear-gradient(135deg,#fff7ed,#fff,#fff7ed)`,
      display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:20,gap:20,
      fontFamily:"'Noto Sans Thai','Noto Sans',sans-serif" }}>
      <div style={{ textAlign:"center" }}>
        <div style={{ width:110,height:110,borderRadius:28,
          background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,
          display:"flex",alignItems:"center",justifyContent:"center",
          fontSize:60,margin:"0 auto 14px",boxShadow:"0 8px 28px rgba(249,115,22,0.35)" }}>🫕</div>
        <div style={{ color:T.orange,fontWeight:900,fontSize:26 }}>ไท่กั๋วหม่าล่า</div>
        <div style={{ color:T.textSm,fontSize:14,marginTop:4 }}>TAI GUO MALA • ระบบจัดการร้าน</div>
      </div>
      <Card style={{ width:"100%",maxWidth:360,boxShadow:T.shadowMd }}>
        <div style={{ color:T.textMd,fontSize:15,marginBottom:12,textAlign:"center" }}>กรอก PIN เพื่อเข้าใช้งาน</div>
        <input type="password" maxLength={4} value={pin}
          onChange={e=>{setPin(e.target.value);setErr(false);}}
          onKeyDown={e=>e.key==="Enter"&&go()}
          style={{...S.inp,fontSize:28,letterSpacing:10,textAlign:"center",marginBottom:10,
            border:`2px solid ${err?T.red:T.border}`,animation:err?"shake .3s ease-in-out":undefined}}
          placeholder="••••" autoFocus />
        <button onClick={go} style={{...S.btn(),width:"100%",padding:13,fontSize:17}}>เข้าสู่ระบบ</button>
        <div style={{ marginTop:14,textAlign:"center",color:T.textXs,fontSize:13 }}>
          กรอก PIN ที่ได้รับจากเจ้าของร้าน
        </div>
      </Card>
      <style>{`@keyframes shake{0%,100%{transform:translateX(0)}25%{transform:translateX(-8px)}75%{transform:translateX(8px)}}`}</style>
    </div>
  );
}

// ─── CASHFLOW PAGE ───────────────────────────
function CashflowPage({ cf, setCF, user, dbReady }) {
  const [showForm, setShowForm]       = useState(false);
  const [viewTab, setViewTab]         = useState("daily");
  const [filterMode, setFilterMode]   = useState("month");
  const [filterMonth, setFilterMonth] = useState(today().slice(0,7));
  const [filterDay, setFilterDay]     = useState(today());
  const [rangeFrom, setRangeFrom]     = useState(today().slice(0,8)+"01");
  const [rangeTo, setRangeTo]         = useState(today());
  const [openCash, setOpenCash]       = useState(0);
  const [openBank, setOpenBank]       = useState(0);
  const [showSetup, setShowSetup]     = useState(false);
  const [msg, setMsg]                 = useState("");
  const [form, setForm] = useState({ date:today(),flow:"in",cat:IN_CATS[0],itemName:"",amount:"",method:"เงินสด",note:"" });

  const myCF = user.role==="owner" ? cf
    : user.role==="franchise" ? cf.filter(e=>e.branch===user.franchiseId)
    : cf.filter(e=>e.staffId===user.id||e.branch==="main");

  const months = [...new Set(myCF.map(e=>e.date.slice(0,7)))].sort((a,b)=>b.localeCompare(a));
  const filtered = filterMode==="day" ? myCF.filter(e=>e.date===filterDay)
    : filterMode==="range" ? myCF.filter(e=>e.date>=rangeFrom&&e.date<=rangeTo)
    : myCF.filter(e=>e.date.startsWith(filterMonth));

  const isCash = e => e.method==="เงินสด";
  const sum = (arr, flow, fn) => arr.filter(e=>e.flow===flow&&(fn?fn(e):true)).reduce((a,b)=>a+b.amount,0);
  const totalIn  = sum(filtered,"in");
  const totalOut = sum(filtered,"out");
  const cashIn   = sum(filtered,"in",isCash);
  const cashOut  = sum(filtered,"out",isCash);
  const bankIn   = sum(filtered,"in",e=>!isCash(e));
  const bankOut  = sum(filtered,"out",e=>!isCash(e));
  const cashBal  = openCash + cashIn  - cashOut;
  const bankBal  = openBank + bankIn  - bankOut;

  // Running daily balance
  const allDates = [...new Set(myCF.map(e=>e.date))].sort();
  let rCash = openCash, rBank = openBank;
  const dailyRows = allDates.map(d => {
    const ent = myCF.filter(e=>e.date===d);
    const dCI=sum(ent,"in",isCash), dCO=sum(ent,"out",isCash);
    const dBI=sum(ent,"in",e=>!isCash(e)), dBO=sum(ent,"out",e=>!isCash(e));
    rCash+=dCI-dCO; rBank+=dBI-dBO;
    return { d,ent,dCI,dCO,dBI,dBO,balCash:rCash,balBank:rBank };
  });
  const dispDays = [...(filterMode==="month" ? dailyRows.filter(r=>r.d.startsWith(filterMonth))
    : filterMode==="day" ? dailyRows.filter(r=>r.d===filterDay)
    : dailyRows.filter(r=>r.d>=rangeFrom&&r.d<=rangeTo))].reverse();

  const addEntry = async () => {
    if (!form.amount||!+form.amount) return;
    const entry = {...form,id:Date.now(),amount:+form.amount,branch:user.franchiseId||"main",staffId:user.id};
    setCF(prev=>[entry,...prev]);
    setShowForm(false);
    setForm({date:today(),flow:"in",cat:IN_CATS[0],itemName:"",amount:"",method:"เงินสด",note:""});
    if (dbReady) db.addCF({id:entry.id,date:entry.date,flow:entry.flow,cat:entry.cat,
      item_name:entry.itemName,amount:entry.amount,method:entry.method,note:entry.note,
      branch:entry.branch,staff_id:entry.staffId}).catch(()=>{});
  };
  const delEntry = async (id) => {
    setCF(prev=>prev.filter(e=>e.id!==id));
    if (dbReady) db.delCF(id).catch(()=>{});
  };
  const clearAll = async () => {
    if (!window.confirm("ล้าง Cash Flow ทั้งหมด?\nข้อมูลจะหายถาวร")) return;
    setCF([]);
    if (dbReady) db.clearCF().catch(()=>{});
    setMsg("ล้าง Cash Flow เรียบร้อย");
    setTimeout(()=>setMsg(""),3000);
  };

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <Hdr title="💵 Cash Flow" action={
        <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
          <button onClick={()=>setShowSetup(!showSetup)} style={{...S.ghost,fontSize:13,padding:"7px 12px"}}>⚙️ ยอดเปิด</button>
          <button onClick={clearAll} style={{...S.btn(T.red),fontSize:13,padding:"7px 12px"}}>🗑 ล้างข้อมูล</button>
          <button onClick={()=>setShowForm(!showForm)} style={S.btn()}>+ กรอก</button>
        </div>
      } />

      <Msg text={msg} ok={true} />

      {/* ตั้งยอดเปิด */}
      {showSetup && (
        <Card style={{ borderColor:T.borderOr,background:T.orangeLt }}>
          <div style={{ color:T.orange,fontWeight:800,fontSize:16,marginBottom:8 }}>⚙️ ตั้งยอดเปิดบัญชี</div>
          <div style={{ color:T.textMd,fontSize:13,marginBottom:12 }}>ใส่ยอดเงินที่มีจริง ณ วันที่เริ่มใช้งาน</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
            <div>
              <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>💵 เงินสดในมือ (฿)</div>
              <input type="number" value={openCash} onChange={e=>setOpenCash(+e.target.value||0)} style={S.inp} placeholder="0" />
            </div>
            <div>
              <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>🏦 เงินในธนาคาร (฿)</div>
              <input type="number" value={openBank} onChange={e=>setOpenBank(+e.target.value||0)} style={S.inp} placeholder="0" />
            </div>
          </div>
          <button onClick={()=>setShowSetup(false)} style={{...S.btn(),width:"100%",marginTop:12,padding:10}}>✅ บันทึก</button>
        </Card>
      )}

      {/* ยอดคงเหลือ */}
      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:12 }}>
        <Card style={{ padding:"16px 18px",borderLeft:`4px solid #78716c`,background:"#fafaf9" }}>
          <div style={{ display:"flex",alignItems:"center",gap:6,marginBottom:6 }}>
            <span style={{ fontSize:20 }}>💵</span>
            <span style={{ color:"#78716c",fontWeight:700,fontSize:14 }}>เงินสดในมือ</span>
          </div>
          <div style={{ color:cashBal>=0?T.green:T.red,fontWeight:900,fontSize:24 }}>฿{fmt(cashBal)}</div>
          <div style={{ fontSize:12,color:T.textSm,marginTop:4,lineHeight:1.8 }}>
            <div>เปิด ฿{fmt(openCash)}</div>
            <div style={{ color:T.green }}>+ เข้า ฿{fmt(cashIn)}</div>
            <div style={{ color:T.red }}>- ออก ฿{fmt(cashOut)}</div>
          </div>
        </Card>
        <Card style={{ padding:"16px 18px",borderLeft:`4px solid ${T.blue}`,background:T.blueLt }}>
          <div style={{ display:"flex",alignItems:"center",gap:6,marginBottom:6 }}>
            <span style={{ fontSize:20 }}>🏦</span>
            <span style={{ color:T.blue,fontWeight:700,fontSize:14 }}>เงินในธนาคาร</span>
          </div>
          <div style={{ color:bankBal>=0?T.green:T.red,fontWeight:900,fontSize:24 }}>฿{fmt(bankBal)}</div>
          <div style={{ fontSize:12,color:T.textSm,marginTop:4,lineHeight:1.8 }}>
            <div>เปิด ฿{fmt(openBank)}</div>
            <div style={{ color:T.green }}>+ เข้า ฿{fmt(bankIn)}</div>
            <div style={{ color:T.red }}>- ออก ฿{fmt(bankOut)}</div>
          </div>
        </Card>
      </div>
      <Card style={{ padding:"12px 18px",background:T.orangeLt,borderColor:T.borderOr }}>
        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center" }}>
          <div>
            <div style={{ color:T.textSm,fontSize:13 }}>ยอดรวมทั้งหมด</div>
            <div style={{ color:T.orange,fontWeight:900,fontSize:22 }}>฿{fmt(cashBal+bankBal)}</div>
          </div>
          <div style={{ textAlign:"right",fontSize:13 }}>
            <div style={{ color:T.green }}>เข้า ฿{fmt(totalIn)}</div>
            <div style={{ color:T.red }}>ออก ฿{fmt(totalOut)}</div>
          </div>
        </div>
      </Card>

      {/* Filter */}
      <Card style={{ padding:"14px 16px" }}>
        <div style={{ display:"flex",gap:6,marginBottom:10 }}>
          {[["month","📅 เดือน"],["day","📆 วัน"],["range","📊 ช่วง"]].map(([v,l])=>(
            <button key={v} onClick={()=>setFilterMode(v)} style={{
              flex:1,padding:"8px 4px",borderRadius:8,cursor:"pointer",fontSize:13,fontFamily:"inherit",
              background:filterMode===v?T.orange:"transparent",
              border:`1px solid ${filterMode===v?T.orange:T.border}`,
              color:filterMode===v?"#fff":T.textMd,fontWeight:filterMode===v?700:400 }}>{l}</button>
          ))}
        </div>
        {filterMode==="month" && (
          <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
            {(months.length?months:[today().slice(0,7)]).map(m=>(
              <button key={m} onClick={()=>setFilterMonth(m)} style={{
                background:filterMonth===m?T.orange:"transparent",border:`1px solid ${filterMonth===m?T.orange:T.border}`,
                borderRadius:8,padding:"6px 12px",color:filterMonth===m?"#fff":T.textMd,
                cursor:"pointer",fontSize:13,fontFamily:"inherit" }}>{m}</button>
            ))}
          </div>
        )}
        {filterMode==="day" && (
          <div>
            <input type="date" value={filterDay} max={today()} onChange={e=>setFilterDay(e.target.value)} style={{...S.inp,fontSize:16}} />
            <div style={{ display:"flex",gap:6,marginTop:8,flexWrap:"wrap" }}>
              {(()=>{ const g=n=>{const d=new Date();d.setDate(d.getDate()-n);return d.toISOString().split("T")[0];};
                return [["วันนี้",g(0)],["เมื่อวาน",g(1)],["3 วันก่อน",g(3)],["7 วันก่อน",g(7)]].map(([l,v])=>(
                  <button key={l} onClick={()=>setFilterDay(v)} style={{
                    background:filterDay===v?T.blue:"transparent",border:`1px solid ${filterDay===v?T.blue:T.border}`,
                    borderRadius:7,padding:"5px 10px",color:filterDay===v?"#fff":T.textMd,
                    cursor:"pointer",fontSize:12,fontFamily:"inherit" }}>{l}</button>
                ));
              })()}
            </div>
          </div>
        )}
        {filterMode==="range" && (
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>จาก</div>
              <input type="date" value={rangeFrom} max={today()} onChange={e=>setRangeFrom(e.target.value)} style={S.inp} /></div>
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>ถึง</div>
              <input type="date" value={rangeTo} max={today()} onChange={e=>setRangeTo(e.target.value)} style={S.inp} /></div>
            <div style={{ gridColumn:"1/-1",display:"flex",gap:6 }}>
              {[["7 วัน",7],["14 วัน",14],["30 วัน",30]].map(([l,d])=>(
                <button key={l} onClick={()=>{const nd=new Date(),nf=new Date();nf.setDate(nd.getDate()-(d-1));setRangeTo(nd.toISOString().split("T")[0]);setRangeFrom(nf.toISOString().split("T")[0]);}}
                  style={{...S.btn(T.orange),fontSize:12,padding:"5px 10px"}}>{l}</button>
              ))}
            </div>
          </div>
        )}
      </Card>

      {/* View tabs */}
      <Tabs tabs={[["daily","📆 รายวัน+ยอดยก"],["list","📋 รายการ"],["summary","📊 สรุป"]]} active={viewTab} onChange={setViewTab} />

      {/* Add form */}
      {showForm && (
        <Card style={{ borderColor:T.borderOr }}>
          <div style={{ color:T.orange,fontWeight:800,fontSize:17,marginBottom:12 }}>📝 บันทึกรายการ</div>
          <div style={{ display:"flex",gap:8,marginBottom:12 }}>
            {[["in","💰 รายรับ",T.green],["out","💸 รายจ่าย",T.red]].map(([v,l,c])=>(
              <button key={v} onClick={()=>setForm(p=>({...p,flow:v,cat:v==="in"?IN_CATS[0]:OUT_CATS[0],itemName:""}))}
                style={{flex:1,padding:12,background:form.flow===v?c+"18":"transparent",
                  border:`1.5px solid ${form.flow===v?c:T.border}`,borderRadius:10,
                  color:form.flow===v?c:T.textMd,fontWeight:700,cursor:"pointer",fontSize:16,fontFamily:"inherit"}}>{l}</button>
            ))}
          </div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>วันที่</div>
              <input type="date" value={form.date} max={today()} onChange={e=>setForm(p=>({...p,date:e.target.value}))} style={S.inp} /></div>
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>จำนวน (฿)</div>
              <input type="number" value={form.amount} onChange={e=>setForm(p=>({...p,amount:e.target.value}))} style={S.inp} placeholder="0" /></div>
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>หมวดหมู่</div>
              <select value={form.cat} onChange={e=>setForm(p=>({...p,cat:e.target.value}))} style={S.inp}>
                {(form.flow==="in"?IN_CATS:OUT_CATS).map(c=><option key={c}>{c}</option>)}</select></div>
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>ช่องทาง</div>
              <select value={form.method} onChange={e=>setForm(p=>({...p,method:e.target.value}))} style={S.inp}>
                {PAY_METHODS.map(m=><option key={m}>{m}</option>)}</select></div>
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>ชื่อรายการ</div>
              <input type="text" value={form.itemName} onChange={e=>setForm(p=>({...p,itemName:e.target.value}))} style={S.inp} placeholder="เช่น หมูสามชั้น" /></div>
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>หมายเหตุ</div>
              <input type="text" value={form.note} onChange={e=>setForm(p=>({...p,note:e.target.value}))} style={S.inp} placeholder="(ไม่บังคับ)" /></div>
          </div>
          <div style={{ display:"flex",gap:8,marginTop:12 }}>
            <button onClick={addEntry} style={{...S.btn(form.flow==="in"?T.green:T.red),flex:1,padding:12,fontSize:16}}>✅ บันทึก</button>
            <button onClick={()=>setShowForm(false)} style={{...S.ghost,padding:"12px 14px"}}>ยกเลิก</button>
          </div>
        </Card>
      )}

      {/* รายวัน + ยอดยก */}
      {viewTab==="daily" && (
        <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
          {dispDays.length===0 && <div style={{ color:T.textSm,textAlign:"center",padding:32 }}>ไม่มีรายการในช่วงที่เลือก</div>}
          {dispDays.map((row,i)=>{
            const prev = dispDays[i+1];
            return (
              <Card key={row.d} style={{ padding:"14px 18px" }}>
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",paddingBottom:10,marginBottom:10,borderBottom:`1px solid ${T.bg}` }}>
                  <span style={{ fontWeight:800,fontSize:15 }}>{row.d}</span>
                  <div style={{ display:"flex",gap:10,fontSize:13 }}>
                    <span style={{ color:T.green }}>+฿{fmt(row.dCI+row.dBI)}</span>
                    <span style={{ color:T.red }}>-฿{fmt(row.dCO+row.dBO)}</span>
                  </div>
                </div>
                {prev && (
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8,
                    padding:"6px 10px",background:T.bg,borderRadius:8,fontSize:12 }}>
                    <div style={{ color:T.textSm }}>💵 ยกมา: <b style={{ color:T.text }}>฿{fmt(prev.balCash)}</b></div>
                    <div style={{ color:T.textSm }}>🏦 ยกมา: <b style={{ color:T.text }}>฿{fmt(prev.balBank)}</b></div>
                  </div>
                )}
                {row.ent.map(e=>(
                  <div key={e.id} style={{ display:"flex",alignItems:"center",gap:10,padding:"7px 0",borderBottom:`1px solid ${T.bg}` }}>
                    <div style={{ width:28,height:28,borderRadius:8,flexShrink:0,
                      background:e.flow==="in"?T.greenLt:T.redLt,
                      display:"flex",alignItems:"center",justifyContent:"center",
                      fontSize:14,color:e.flow==="in"?T.green:T.red }}>{e.flow==="in"?"↓":"↑"}</div>
                    <div style={{ flex:1,minWidth:0 }}>
                      <div style={{ color:T.text,fontSize:14,fontWeight:600 }}>{e.itemName||e.cat}</div>
                      <div style={{ color:T.textSm,fontSize:11 }}>
                        {e.itemName?`${e.cat} • `:""}{isCash(e)?"💵":"🏦"} {e.method}{e.note?` • ${e.note}`:""}
                      </div>
                    </div>
                    <span style={{ color:e.flow==="in"?T.green:T.red,fontWeight:700,fontSize:15,flexShrink:0 }}>
                      {e.flow==="in"?"+":"-"}฿{fmt(e.amount)}
                    </span>
                    <button onClick={()=>delEntry(e.id)}
                      style={{ background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:18,padding:"0 4px",flexShrink:0 }}>×</button>
                  </div>
                ))}
                <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:10,
                  padding:"8px 10px",background:T.orangeLt,borderRadius:8,fontSize:13 }}>
                  <div>
                    <div style={{ color:T.textXs,fontSize:11 }}>💵 เงินสดสิ้นวัน</div>
                    <div style={{ color:row.balCash>=0?T.green:T.red,fontWeight:800,fontSize:16 }}>฿{fmt(row.balCash)}</div>
                  </div>
                  <div>
                    <div style={{ color:T.textXs,fontSize:11 }}>🏦 ธนาคารสิ้นวัน</div>
                    <div style={{ color:row.balBank>=0?T.green:T.red,fontWeight:800,fontSize:16 }}>฿{fmt(row.balBank)}</div>
                  </div>
                </div>
              </Card>
            );
          })}
        </div>
      )}

      {/* รายการ */}
      {viewTab==="list" && (
        <div style={{ display:"flex",flexDirection:"column",gap:8 }}>
          {filtered.length===0 && <div style={{ color:T.textSm,textAlign:"center",padding:32 }}>ไม่มีรายการ</div>}
          {[...filtered].sort((a,b)=>b.date.localeCompare(a.date)||b.id-a.id).map(e=>(
            <Card key={e.id} style={{ padding:"12px 16px" }}>
              <div style={{ display:"flex",alignItems:"center",gap:10 }}>
                <div style={{ width:32,height:32,borderRadius:9,flexShrink:0,
                  background:e.flow==="in"?T.greenLt:T.redLt,
                  display:"flex",alignItems:"center",justifyContent:"center",
                  fontSize:16,color:e.flow==="in"?T.green:T.red }}>{e.flow==="in"?"↓":"↑"}</div>
                <div style={{ flex:1,minWidth:0 }}>
                  <div style={{ color:T.text,fontSize:15,fontWeight:600 }}>{e.itemName||e.cat}</div>
                  <div style={{ color:T.textSm,fontSize:12 }}>{e.date} • {isCash(e)?"💵":"🏦"} {e.method}{e.note?` • ${e.note}`:""}</div>
                </div>
                <span style={{ color:e.flow==="in"?T.green:T.red,fontWeight:700,fontSize:16,flexShrink:0 }}>
                  {e.flow==="in"?"+":"-"}฿{fmt(e.amount)}
                </span>
                <button onClick={()=>delEntry(e.id)}
                  style={{ background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:18,padding:"0 4px" }}>×</button>
              </div>
            </Card>
          ))}
        </div>
      )}

      {/* สรุป */}
      {viewTab==="summary" && (
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          <Card>
            <div style={{ fontWeight:800,fontSize:16,marginBottom:10 }}>💳 แยกตามช่องทาง</div>
            {Object.entries(filtered.reduce((acc,e)=>{
              if(!acc[e.method])acc[e.method]={in:0,out:0};
              acc[e.method][e.flow]+=e.amount; return acc;
            },{})).sort((a,b)=>(b[1].in+b[1].out)-(a[1].in+a[1].out)).map(([m,{in:mI,out:mO}])=>(
              <div key={m} style={{ padding:"9px 0",borderBottom:`1px solid ${T.bg}` }}>
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4 }}>
                  <span style={{ fontWeight:700,fontSize:15 }}>{m==="เงินสด"?"💵":"🏦"} {m}</span>
                  <span style={{ color:mI-mO>=0?T.green:T.red,fontWeight:800,fontSize:16 }}>฿{fmt(mI-mO)}</span>
                </div>
                <div style={{ display:"flex",gap:12,fontSize:12,color:T.textSm }}>
                  {mI>0&&<span style={{ color:T.green }}>เข้า ฿{fmt(mI)}</span>}
                  {mO>0&&<span style={{ color:T.red }}>ออก ฿{fmt(mO)}</span>}
                </div>
              </div>
            ))}
          </Card>
          <Card>
            <div style={{ fontWeight:800,fontSize:16,marginBottom:10 }}>🧾 รายจ่ายแยกหมวด</div>
            {OUT_CATS.map(cat=>{
              const val=filtered.filter(e=>e.flow==="out"&&e.cat===cat).reduce((a,b)=>a+b.amount,0);
              if(!val)return null;
              const pct=totalOut>0?(val/totalOut*100).toFixed(0):0;
              return (
                <div key={cat} style={{ marginBottom:10 }}>
                  <div style={{ display:"flex",justifyContent:"space-between",marginBottom:4 }}>
                    <span style={{ fontSize:14 }}>{cat}</span>
                    <span style={{ fontWeight:700,fontSize:14 }}>฿{fmt(val)} ({pct}%)</span>
                  </div>
                  <div style={{ background:T.bg,borderRadius:4,height:6 }}>
                    <div style={{ background:T.orange,width:`${pct}%`,height:"100%",borderRadius:4 }}/>
                  </div>
                </div>
              );
            })}
          </Card>
        </div>
      )}
    </div>
  );
}

// ─── STOCK PAGE (เจ้าของ) ───────────────────
function StockPage({ stock, setStock, movements, setMovements, user, suppliers }) {
  const [tab, setTab]       = useState("list");
  const [selId, setSelId]   = useState("");
  const [mvType, setMvType] = useState("in");
  const [qty, setQty]       = useState("");
  const [cost, setCost]     = useState("");
  const [note, setNote]     = useState("");
  const [msg, setMsg]       = useState({ text:"", ok:false });
  const [editMinId, setEditMinId] = useState(null);
  const [editMinVal, setEditMinVal] = useState("");
  const [editItemId, setEditItemId] = useState(null);
  const [editItemData, setEditItemData] = useState({});
  const [showAdd, setShowAdd] = useState(false);
  const [filter, setFilter] = useState("all");
  const [newItem, setNewItem] = useState({ name:"",unit:"kg",qty:0,minQty:3,dailyUse:1,supplierId:1 });
  const fileRef = useRef();

  const canPrice = user.perms?.viewPrice===true;
  const selectedItem = selId ? stock.find(s=>String(s.id)===String(selId)) : null;

  const sorted = useMemo(()=>{
    const ord={out:0,critical:1,low:2,ok:3};
    const list=filter==="all"?stock:stock.filter(s=>stockSt(s)===filter);
    return [...list].sort((a,b)=>ord[stockSt(a)]-ord[stockSt(b)]);
  },[stock,filter]);

  const showMsg = (text,ok) => { setMsg({text,ok}); setTimeout(()=>setMsg({text:"",ok:false}),3000); };

  const save = () => {
    const q=parseFloat(qty);
    if (!selId||!q||q<=0) { showMsg("กรุณาเลือกรายการและกรอกจำนวน",false); return; }
    const item=stock.find(s=>String(s.id)===String(selId));
    if (!item) return;
    if (mvType==="out"&&q>item.qty) { showMsg(`มีแค่ ${item.qty} ${item.unit} ไม่พอ`,false); return; }
    const unitCost=parseFloat(cost)||0;
    const newQty=mvType==="in"?item.qty+q:item.qty-q;
    const newHist=mvType==="in"&&unitCost>0?[...(item.costHistory||[]),{date:today(),unitCost,qty:q,total:unitCost*q}]:item.costHistory;
    setStock(stock.map(s=>String(s.id)===String(selId)?{...s,qty:newQty,costHistory:newHist}:s));
    setMovements(prev=>[...prev,{id:Date.now(),itemId:item.id,type:mvType,qty:q,unitCost,
      date:today(),staffId:user.id,note:note||(mvType==="in"?"รับเข้า":"จ่ายออก"),branch:"main"}]);
    showMsg(`✅ ${mvType==="in"?"รับเข้า":"จ่ายออก"} ${item.name} ${q} ${item.unit} เรียบร้อย`,true);
    setQty(""); setCost(""); setNote("");
  };

  const saveMinQty = id => { setStock(stock.map(s=>s.id===id?{...s,minQty:+editMinVal}:s)); setEditMinId(null); };

  const handleImport = e => {
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=ev=>{
      try {
        const wb=XLSX.read(ev.target.result,{type:"array"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        let added=0,updated=0;
        const newStock=[...stock];
        rows.forEach(r=>{
          const name=r["ชื่อวัตถุดิบ"]||r["name"]||""; if(!name) return;
          const idx=newStock.findIndex(s=>s.name===name);
          if(idx>=0){if(r["จำนวนคงเหลือ"]!==undefined)newStock[idx]={...newStock[idx],qty:+r["จำนวนคงเหลือ"]};updated++;}
          else{newStock.push({id:Date.now()+Math.random(),name,unit:r["หน่วย"]||"kg",qty:+(r["จำนวนคงเหลือ"]||0),minQty:+(r["จำนวนขั้นต่ำ"]||3),dailyUse:+(r["ใช้ต่อวัน"]||1),supplierId:1,costHistory:[]});added++;}
        });
        setStock(newStock);
        showMsg(`✅ เพิ่ม ${added} อัพเดท ${updated} รายการ`,true);
      } catch { showMsg("❌ ไฟล์ผิดรูปแบบ",false); }
    };
    reader.readAsArrayBuffer(file); e.target.value="";
  };

  const handleExport = () => {
    const rows=stock.map(s=>({
      "ชื่อวัตถุดิบ":s.name,"หน่วย":s.unit,"จำนวนคงเหลือ":s.qty,
      "จำนวนขั้นต่ำ":s.minQty,"ใช้ต่อวัน":s.dailyUse,
      "ราคาเฉลี่ย":canPrice?+wac(s).toFixed(2):"-","สถานะ":ST_LABEL[stockSt(s)]
    }));
    const ws=XLSX.utils.json_to_sheet(rows);
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"stock");
    XLSX.writeFile(wb,`stock_${today()}.xlsx`);
  };

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <Hdr title="📦 สต็อควัตถุดิบ" action={
        <div style={{ display:"flex",gap:8,flexWrap:"wrap" }}>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display:"none" }} />
          <button onClick={()=>fileRef.current?.click()} style={{...S.ghost,fontSize:13,padding:"7px 12px"}}>📥 Import</button>
          <button onClick={handleExport} style={{...S.ghost,fontSize:13,padding:"7px 12px"}}>📤 Export</button>
          <button onClick={()=>setShowAdd(!showAdd)} style={S.btn()}>+ เพิ่ม</button>
        </div>
      } />

      <Msg text={msg.text} ok={msg.ok} />

      {showAdd && (
        <Card style={{ borderColor:T.borderOr }}>
          <div style={{ color:T.orange,fontWeight:800,fontSize:16,marginBottom:12 }}>➕ เพิ่มวัตถุดิบ</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
            {[["ชื่อ","name","text"],["หน่วย","unit","text"],["จำนวน","qty","number"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"]].map(([l,k,t])=>(
              <div key={k}><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>{l}</div>
                <input type={t} value={newItem[k]} onChange={e=>setNewItem(p=>({...p,[k]:e.target.value}))} style={S.inp} /></div>
            ))}
          </div>
          <div style={{ display:"flex",gap:8,marginTop:12 }}>
            <button onClick={()=>{setStock([...stock,{...newItem,id:Date.now(),qty:+newItem.qty,minQty:+newItem.minQty,dailyUse:+newItem.dailyUse,costHistory:[]}]);setShowAdd(false);}} style={{...S.btn(),flex:1}}>บันทึก</button>
            <button onClick={()=>setShowAdd(false)} style={S.ghost}>ยกเลิก</button>
          </div>
        </Card>
      )}

      <Tabs tabs={[["list","📋 รายการสต็อค"],["move","📥📤 รับ/จ่าย"],["history","📊 ประวัติ"]]} active={tab} onChange={t=>{setTab(t);setSelId("");setQty("");setMsg({text:"",ok:false});}} />

      {tab==="list" && (
        <>
          <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
            {[["all","ทั้งหมด"],["critical","🔴 น้อยมาก"],["low","🟡 ใกล้หมด"],["ok","🟢 ปกติ"]].map(([v,l])=>(
              <button key={v} onClick={()=>setFilter(v)} style={{
                background:filter===v?T.orange:"transparent",border:`1px solid ${filter===v?T.orange:T.border}`,
                borderRadius:8,padding:"7px 14px",color:filter===v?"#fff":T.textMd,
                cursor:"pointer",fontSize:14,fontFamily:"inherit" }}>{l}</button>
            ))}
          </div>
          {sorted.map(item=>{
            const st=stockSt(item);
            const avgCost=wac(item);
            const sup=suppliers?.find(x=>x.id===item.supplierId);
            const daysLeft=item.dailyUse>0?(item.qty/item.dailyUse).toFixed(1):"∞";
            return (
              <Card key={item.id} style={{ borderColor:st!=="ok"?ST_COLOR[st]+"44":T.border }}>
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start" }}>
                  <div style={{ flex:1 }}>
                    <div style={{ display:"flex",alignItems:"center",gap:8,flexWrap:"wrap" }}>
                      <span style={{ fontWeight:700,fontSize:17 }}>{item.name}</span>
                      <Badge status={st} />
                    </div>
                    <div style={{ color:T.textSm,fontSize:13,marginTop:4 }}>
                      ซัพฯ: {sup?.name||"-"} • เหลือ {daysLeft} วัน
                      {canPrice&&avgCost>0&&<span> • เฉลี่ย ฿{avgCost.toFixed(2)}/{item.unit}</span>}
                    </div>
                    <div style={{ display:"flex",alignItems:"center",gap:6,marginTop:4 }}>
                      <span style={{ color:T.textXs,fontSize:12 }}>ขั้นต่ำ:</span>
                      {editMinId===item.id?(
                        <>
                          <input type="number" value={editMinVal} onChange={e=>setEditMinVal(e.target.value)} autoFocus
                            style={{...S.inp,width:60,padding:"3px 8px",fontSize:13}} onKeyDown={e=>e.key==="Enter"&&saveMinQty(item.id)} />
                          <button onClick={()=>saveMinQty(item.id)} style={{...S.btn(T.green),padding:"3px 10px",fontSize:12}}>✓</button>
                          <button onClick={()=>setEditMinId(null)} style={{...S.ghost,padding:"3px 8px",fontSize:12}}>✕</button>
                        </>
                      ):(
                        <span onClick={()=>{setEditMinId(item.id);setEditMinVal(item.minQty);}}
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
                {editItemId===item.id ? (
                  <div style={{ marginTop:12,paddingTop:12,borderTop:`1px solid ${T.bg}` }}>
                    <div style={{ color:T.orange,fontWeight:700,fontSize:14,marginBottom:10 }}>✏️ แก้ไขรายละเอียด</div>
                    <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10 }}>
                      {[["ชื่อ","name","text"],["หน่วย","unit","text"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"]].map(([l,k,t])=>(
                        <div key={k}><div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>{l}</div>
                          <input type={t} value={editItemData[k]??item[k]}
                            onChange={e=>setEditItemData(p=>({...p,[k]:e.target.value}))}
                            style={S.inp} /></div>
                      ))}
                    </div>
                    <div style={{ display:"flex",gap:8 }}>
                      <button onClick={()=>{
                        setStock(prev=>prev.map(s=>s.id===item.id?{...s,...editItemData,minQty:+editItemData.minQty||s.minQty,dailyUse:+editItemData.dailyUse||s.dailyUse}:s));
                        setEditItemId(null); setEditItemData({});
                      }} style={{...S.btn(),flex:1}}>บันทึก</button>
                      <button onClick={()=>{setEditItemId(null);setEditItemData({});}} style={S.ghost}>ยกเลิก</button>
                    </div>
                  </div>
                ) : (
                  <div style={{ borderTop:`1px solid ${T.bg}`,marginTop:10,paddingTop:8,display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                    <button onClick={()=>{setEditItemId(item.id);setEditItemData({name:item.name,unit:item.unit,minQty:item.minQty,dailyUse:item.dailyUse});}}
                      style={{ background:"none",border:"none",color:T.orange,cursor:"pointer",fontSize:12,fontFamily:"inherit",fontWeight:600 }}>
                      ✏️ แก้ไขรายละเอียด
                    </button>
                    <button onClick={()=>{if(window.confirm(`ลบ "${item.name}" ออกจากสต็อค?`))setStock(prev=>prev.filter(s=>s.id!==item.id));}}
                      style={{ background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:12,fontFamily:"inherit" }}>
                      🗑 ลบ
                    </button>
                  </div>
                )}
              </Card>
            );
          })}
        </>
      )}

      {tab==="move" && (
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          <div style={{ display:"flex",gap:8 }}>
            {[["in","📥 รับเข้า",T.green],["out","📤 จ่ายออก",T.red]].map(([v,l,c])=>(
              <button key={v} onClick={()=>{setMvType(v);setMsg({text:"",ok:false});}}
                style={{ flex:1,padding:12,fontWeight:800,fontSize:16,cursor:"pointer",fontFamily:"inherit",borderRadius:10,
                  background:mvType===v?c+"18":T.card,border:`2px solid ${mvType===v?c:T.border}`,color:mvType===v?c:T.textMd }}>{l}</button>
            ))}
          </div>
          <div>
            <div style={{ color:T.textSm,fontSize:13,marginBottom:6,fontWeight:600 }}>เลือกรายการ</div>
            <select value={selId} onChange={e=>{setSelId(e.target.value);setQty("");setCost("");setMsg({text:"",ok:false});}}
              style={{...S.inp,fontSize:16,height:48}}>
              <option value="">— กรุณาเลือกรายการ —</option>
              {[...stock].sort((a,b)=>({out:0,critical:1,low:2,ok:3}[stockSt(a)]-{out:0,critical:1,low:2,ok:3}[stockSt(b)])).map(s=>{
                const st=stockSt(s);
                const badge=st==="out"?"🔴":st==="critical"?"🟠":st==="low"?"🟡":"🟢";
                return <option key={s.id} value={String(s.id)}>{badge} {s.name} (คงเหลือ {s.qty} {s.unit})</option>;
              })}
            </select>
          </div>
          {selectedItem && (
            <div style={{ background:T.bg,borderRadius:10,padding:"10px 14px",display:"flex",justifyContent:"space-between",fontSize:14 }}>
              <span style={{ color:T.textMd }}>คงเหลือ</span>
              <span style={{ color:ST_COLOR[stockSt(selectedItem)],fontWeight:800,fontSize:18 }}>
                {selectedItem.qty} {selectedItem.unit}
              </span>
            </div>
          )}
          <div style={{ display:"grid",gridTemplateColumns:mvType==="in"&&canPrice?"1fr 1fr":"1fr",gap:10 }}>
            <div>
              <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>จำนวน ({selectedItem?.unit||"หน่วย"})</div>
              <input type="number" inputMode="numeric" value={qty}
                onChange={e=>{setQty(e.target.value);setMsg({text:"",ok:false});}}
                style={{...S.inp,fontSize:22,fontWeight:800,textAlign:"center"}} placeholder="0" />
            </div>
            {mvType==="in"&&canPrice&&(
              <div>
                <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>ราคา/หน่วย (฿)</div>
                <input type="number" inputMode="numeric" value={cost} onChange={e=>setCost(e.target.value)}
                  style={{...S.inp,fontSize:18,borderColor:T.orange}} placeholder="0" />
              </div>
            )}
          </div>
          <div>
            <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>หมายเหตุ</div>
            <input type="text" value={note} onChange={e=>setNote(e.target.value)} style={S.inp}
              placeholder={mvType==="in"?"เช่น รับจากซัพฯ":"เช่น ใช้วันนี้"} />
          </div>
          {selectedItem&&qty&&+qty>0&&(
            <div style={{ background:mvType==="in"?T.greenLt:T.redLt,borderRadius:8,padding:"10px 14px",fontSize:14 }}>
              หลังบันทึก: <b style={{ color:mvType==="in"?T.green:T.red,fontSize:18 }}>
                {mvType==="in"?selectedItem.qty+(+qty):Math.max(0,selectedItem.qty-(+qty))} {selectedItem.unit}
              </b>
              {mvType==="in"&&cost&&canPrice&&<span style={{ color:T.textSm,fontSize:12 }}> • รวมจ่าย ฿{fmt(+qty*+cost)}</span>}
            </div>
          )}
          <Msg text={msg.text} ok={msg.ok} />
          <button onClick={save}
            style={{...S.btn(mvType==="in"?T.green:T.red),width:"100%",padding:14,fontSize:17}}>
            ✅ บันทึก {mvType==="in"?"รับเข้า":"จ่ายออก"}
          </button>
        </div>
      )}

      {tab==="history" && (
        <Card>
          <div style={{ fontWeight:800,fontSize:16,marginBottom:12 }}>ประวัติการเคลื่อนไหว</div>
          {[...movements].sort((a,b)=>b.date.localeCompare(a.date)).slice(0,40).map(m=>{
            const item=stock.find(s=>s.id==m.itemId);
            const col=m.type==="in"?T.green:m.type==="check"?T.blue:T.red;
            return (
              <div key={m.id} style={{ display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:`1px solid ${T.bg}` }}>
                <div style={{ width:32,height:32,borderRadius:8,background:col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,flexShrink:0 }}>
                  {m.type==="in"?"📥":m.type==="check"?"✅":"📤"}
                </div>
                <div style={{ flex:1 }}>
                  <div style={{ fontWeight:600,fontSize:14 }}>{item?.name||"?"}</div>
                  <div style={{ color:T.textSm,fontSize:12 }}>{m.date} • {m.note}{canPrice&&m.unitCost>0?` • ฿${m.unitCost}/${item?.unit}`:""}</div>
                </div>
                <span style={{ color:col,fontWeight:800,fontSize:15,flexShrink:0 }}>
                  {m.type==="check"?"":m.type==="in"?"+":"-"}{m.qty} {item?.unit}
                </span>
                <button onClick={()=>setMovements(prev=>prev.filter(x=>x.id!==m.id))}
                  style={{ background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:18,padding:"0 4px" }}>×</button>
              </div>
            );
          })}
          {movements.length===0&&<div style={{ color:T.textSm,textAlign:"center",padding:24 }}>ยังไม่มีประวัติ</div>}
        </Card>
      )}
    </div>
  );
}

// ─── STAFF STOCK PAGE (พนักงาน) ─────────────
function StaffStockPage({ stock, setStock, movements, setMovements, user }) {
  const [tab, setTab]     = useState("quick");
  const [selId, setSelId] = useState("");
  const [mvType, setMvType] = useState("in");
  const [qty, setQty]     = useState("");
  const [note, setNote]   = useState("");
  const [msg, setMsg]     = useState({ text:"", ok:false });
  const [checked, setChecked] = useState({});
  const [submitted, setSubmitted] = useState(false);

  const urgent = stock.filter(s=>["critical","out"].includes(stockSt(s)));
  const selectedItem = selId ? stock.find(s=>String(s.id)===String(selId)) : null;
  const myMvs = [...movements].filter(m=>m.staffId===user.id).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,20);

  const showMsg = (text,ok) => { setMsg({text,ok}); setTimeout(()=>setMsg({text:"",ok:false}),3000); };

  const save = () => {
    const q=parseFloat(qty);
    if (!selId||!q||q<=0) { showMsg("กรุณาเลือกรายการและกรอกจำนวน",false); return; }
    const item=stock.find(s=>String(s.id)===String(selId));
    if (!item) { showMsg("ไม่พบรายการ",false); return; }
    if (mvType==="out"&&q>item.qty) { showMsg(`มีแค่ ${item.qty} ${item.unit} ไม่พอจ่าย`,false); return; }
    const newQty=mvType==="in"?item.qty+q:item.qty-q;
    setStock(stock.map(s=>String(s.id)===String(selId)?{...s,qty:newQty}:s));
    setMovements(prev=>[...prev,{id:Date.now(),itemId:item.id,type:mvType,qty:q,unitCost:0,
      date:today(),staffId:user.id,note:note||(mvType==="in"?"รับเข้า":"จ่ายออก"),branch:"main"}]);
    showMsg(`✅ ${mvType==="in"?"รับเข้า":"จ่ายออก"} ${item.name} ${q} ${item.unit} เรียบร้อย`,true);
    setQty(""); setNote("");
  };

  const saveChecklist = () => {
    const entries=Object.entries(checked).filter(([,v])=>v!==""&&!isNaN(+v));
    if(!entries.length) return;
    setStock(stock.map(s=>{const v=checked[String(s.id)];if(!v||isNaN(+v))return s;return{...s,qty:+v};}));
    entries.forEach(([id,v])=>{
      const item=stock.find(s=>String(s.id)===id); if(!item) return;
      setMovements(prev=>[...prev,{id:Date.now()+Math.random(),itemId:item.id,type:"check",qty:+v,unitCost:0,
        date:today(),staffId:user.id,note:"นับสต็อครายวัน",branch:"main"}]);
    });
    setChecked({}); setSubmitted(true); setTimeout(()=>setSubmitted(false),3000);
  };

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:14 }}>
      <div>
        <div style={{ color:T.text,fontSize:22,fontWeight:900 }}>📦 บันทึกสต็อค</div>
        <div style={{ color:T.textSm,fontSize:14 }}>สวัสดี {user.name} • {today()}</div>
      </div>
      {urgent.length>0&&(
        <Card style={{ background:T.redLt,borderColor:T.red+"44" }}>
          <div style={{ color:T.red,fontWeight:800,fontSize:15,marginBottom:6 }}>🚨 แจ้งเจ้าของด่วน!</div>
          {urgent.map(s=>(
            <div key={s.id} style={{ display:"flex",justifyContent:"space-between",padding:"4px 0",fontSize:14 }}>
              <span style={{ fontWeight:600 }}>{s.name}</span>
              <span style={{ color:T.red,fontWeight:700 }}>เหลือ {s.qty} {s.unit}</span>
            </div>
          ))}
        </Card>
      )}
      <Tabs tabs={[["quick","⚡ รับเข้า/จ่ายออก"],["checklist","✅ นับรายวัน"],["history","📋 ประวัติ"]]}
        active={tab} onChange={t=>{setTab(t);setSelId("");setQty("");setMsg({text:"",ok:false});}} />

      {tab==="quick"&&(
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          <div style={{ display:"flex",gap:8 }}>
            {[["in","📥 รับเข้า",T.green],["out","📤 จ่ายออก",T.red]].map(([v,l,c])=>(
              <button key={v} onClick={()=>{setMvType(v);setMsg({text:"",ok:false});}}
                style={{ flex:1,padding:12,fontWeight:800,fontSize:16,cursor:"pointer",fontFamily:"inherit",borderRadius:10,
                  background:mvType===v?c+"18":T.card,border:`2px solid ${mvType===v?c:T.border}`,color:mvType===v?c:T.textMd }}>{l}</button>
            ))}
          </div>
          <div>
            <div style={{ color:T.textSm,fontSize:13,marginBottom:6,fontWeight:600 }}>เลือกรายการ</div>
            <select value={selId} onChange={e=>{setSelId(e.target.value);setQty("");setMsg({text:"",ok:false});}}
              style={{...S.inp,fontSize:16,height:48}}>
              <option value="">— กรุณาเลือกรายการ —</option>
              {[...stock].sort((a,b)=>({out:0,critical:1,low:2,ok:3}[stockSt(a)]-{out:0,critical:1,low:2,ok:3}[stockSt(b)])).map(s=>{
                const st=stockSt(s);
                const badge=st==="out"?"🔴":st==="critical"?"🟠":st==="low"?"🟡":"🟢";
                return <option key={s.id} value={String(s.id)}>{badge} {s.name} (คงเหลือ {s.qty} {s.unit})</option>;
              })}
            </select>
          </div>
          {selectedItem&&(
            <div style={{ background:T.bg,borderRadius:10,padding:"10px 14px",display:"flex",justifyContent:"space-between",fontSize:14 }}>
              <span style={{ color:T.textMd }}>คงเหลือปัจจุบัน</span>
              <span style={{ color:ST_COLOR[stockSt(selectedItem)],fontWeight:800,fontSize:18 }}>
                {selectedItem.qty} {selectedItem.unit}
              </span>
            </div>
          )}
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
            <div>
              <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>จำนวน ({selectedItem?.unit||"หน่วย"})</div>
              <input type="number" inputMode="numeric" value={qty}
                onChange={e=>{setQty(e.target.value);setMsg({text:"",ok:false});}}
                style={{...S.inp,fontSize:22,fontWeight:800,textAlign:"center"}} placeholder="0" />
            </div>
            <div>
              <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>หมายเหตุ</div>
              <input type="text" value={note} onChange={e=>setNote(e.target.value)} style={S.inp}
                placeholder={mvType==="in"?"รับจากซัพฯ":"ใช้วันนี้"} />
            </div>
          </div>
          {selectedItem&&qty&&+qty>0&&(
            <div style={{ background:mvType==="in"?T.greenLt:T.redLt,borderRadius:8,padding:"10px 14px",fontSize:14 }}>
              หลังบันทึก: <b style={{ color:mvType==="in"?T.green:T.red,fontSize:18 }}>
                {mvType==="in"?selectedItem.qty+(+qty):Math.max(0,selectedItem.qty-(+qty))} {selectedItem.unit}
              </b>
            </div>
          )}
          <Msg text={msg.text} ok={msg.ok} />
          <button onClick={save} style={{...S.btn(mvType==="in"?T.green:T.red),width:"100%",padding:14,fontSize:17}}>
            ✅ บันทึก {mvType==="in"?"รับเข้า":"จ่ายออก"}
          </button>
        </div>
      )}

      {tab==="checklist"&&(
        <div style={{ display:"flex",flexDirection:"column",gap:10 }}>
          {submitted&&<div style={{ background:T.greenLt,borderRadius:10,padding:"12px 16px",color:T.green,fontWeight:800,fontSize:15 }}>✅ บันทึกสำเร็จ!</div>}
          {stock.map(item=>{
            const st=stockSt(item);
            const val=checked[String(item.id)]??"";
            const diff=val!==""&&!isNaN(+val)?+val-item.qty:null;
            return (
              <Card key={item.id} style={{ padding:"12px 16px",borderColor:st!=="ok"?ST_COLOR[st]+"44":T.border }}>
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                  <div>
                    <div style={{ display:"flex",alignItems:"center",gap:7 }}>
                      <span style={{ fontWeight:700,fontSize:15 }}>{item.name}</span>
                      <Badge status={st} />
                    </div>
                    <div style={{ color:T.textSm,fontSize:12 }}>ระบบ: {item.qty} {item.unit}</div>
                    {diff!==null&&<div style={{ fontSize:12,fontWeight:600,marginTop:2,
                      color:diff<0?T.yellow:diff>0?T.green:T.textSm }}>
                      {diff<0?`⚠️ ขาด ${Math.abs(diff)}`:diff>0?`+เกิน ${diff}`:"✓ ตรง"}
                    </div>}
                  </div>
                  <div style={{ display:"flex",alignItems:"center",gap:6 }}>
                    <input type="number" inputMode="numeric" value={val}
                      onChange={e=>setChecked(p=>({...p,[String(item.id)]:e.target.value}))}
                      style={{...S.inp,width:80,fontSize:18,fontWeight:700,textAlign:"center"}} placeholder="นับได้" />
                    <span style={{ color:T.textSm,fontSize:12 }}>{item.unit}</span>
                  </div>
                </div>
              </Card>
            );
          })}
          <button onClick={saveChecklist}
            style={{...S.btn(T.green),width:"100%",padding:13,fontSize:16}}>
            ✅ ส่งรายงานสต็อค ({Object.entries(checked).filter(([,v])=>v!=="").length} รายการ)
          </button>
        </div>
      )}

      {tab==="history"&&(
        <Card>
          <div style={{ fontWeight:800,fontSize:16,marginBottom:12 }}>ประวัติที่ฉันบันทึก</div>
          {myMvs.length===0?<div style={{ color:T.textSm,textAlign:"center",padding:24 }}>ยังไม่มีประวัติ</div>
          :myMvs.map(m=>{
            const item=stock.find(s=>s.id==m.itemId);
            const col=m.type==="in"?T.green:m.type==="check"?T.blue:T.red;
            return (
              <div key={m.id} style={{ display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:`1px solid ${T.bg}` }}>
                <div style={{ width:32,height:32,borderRadius:8,background:col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,flexShrink:0 }}>
                  {m.type==="in"?"📥":m.type==="check"?"✅":"📤"}
                </div>
                <div style={{ flex:1 }}>
                  <div style={{ fontWeight:600,fontSize:14 }}>{item?.name||"?"}</div>
                  <div style={{ color:T.textSm,fontSize:12 }}>{m.date} • {m.note}</div>
                </div>
                <span style={{ color:col,fontWeight:800,fontSize:15 }}>
                  {m.type==="check"?"":m.type==="in"?"+":"-"}{m.qty} {item?.unit}
                </span>
              </div>
            );
          })}
        </Card>
      )}
    </div>
  );
}

// ─── SETTINGS PAGE ───────────────────────────
function SettingsPage({ staff, setStaff, lineToken, setLineToken, suppliers, setSuppliers, fixedCosts, setFixedCosts }) {
  const [tab, setTab] = useState("staff");
  const [editId, setEditId] = useState(null);
  const [editSup, setEditSup] = useState(null);
  const [editSupData, setEditSupData] = useState({});
  const [showAddSup, setShowAddSup] = useState(false);
  const [newSup, setNewSup] = useState({ name:"",type:"",phone:"",active:true });

  const PERMS = { cashflow:"💵 Cash Flow", stock:"📦 สต็อค", purchase:"🛒 สั่งซื้อ", report:"📊 รายงาน", ai:"🤖 AI", admin:"⚙️ Admin", viewPrice:"💰 ดูราคา" };

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <Hdr title="⚙️ ตั้งค่าระบบ" />
      <Tabs tabs={[["staff","👷 พนักงาน"],["supplier","🏪 ซัพพลายเออร์"],["fixed","🔒 ต้นทุนคงที่"],["line","📲 LINE"]]} active={tab} onChange={setTab} />

      {tab==="staff"&&(
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          {staff.filter(s=>s.role!=="owner"&&s.id!=="emergency").map(s=>(
            <Card key={s.id} style={{ borderColor:s.active?T.border:T.red+"33" }}>
              <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10 }}>
                <div>
                  <div style={{ fontWeight:700,fontSize:16 }}>{s.name}</div>
                  <div style={{ color:T.textSm,fontSize:13 }}>{s.role==="franchise"?"🏪 แฟรนไชส์":"👷 พนักงาน"}</div>
                </div>
                <div style={{ display:"flex",gap:8 }}>
                  <button onClick={()=>setStaff(prev=>prev.map(x=>x.id===s.id?{...x,active:!x.active}:x))}
                    style={{ background:s.active?T.greenLt:T.redLt,border:`1px solid ${s.active?T.green:T.red}44`,borderRadius:8,padding:"6px 12px",color:s.active?T.green:T.red,cursor:"pointer",fontSize:13,fontWeight:700,fontFamily:"inherit" }}>{s.active?"ใช้งาน":"ระงับ"}</button>
                  <button onClick={()=>setEditId(editId===s.id?null:s.id)} style={{...S.ghost,padding:"6px 10px",fontSize:13}}>✏️</button>
                </div>
              </div>
              {editId===s.id&&(
                <div style={{ background:T.bg,borderRadius:10,padding:12,marginBottom:10 }}>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
                    <div><div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>ชื่อ</div>
                      <input value={s.name} onChange={e=>setStaff(prev=>prev.map(x=>x.id===s.id?{...x,name:e.target.value}:x))} style={{...S.inp,fontSize:14}} /></div>
                    <div><div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>PIN</div>
                      <input maxLength={4} value={s.pin} onChange={e=>setStaff(prev=>prev.map(x=>x.id===s.id?{...x,pin:e.target.value}:x))} style={{...S.inp,fontSize:16,letterSpacing:4}} /></div>
                  </div>
                </div>
              )}
              <div style={{ color:T.textSm,fontSize:13,marginBottom:8,fontWeight:600 }}>สิทธิ์การใช้งาน</div>
              <div style={{ display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6 }}>
                {Object.entries(PERMS).map(([perm,label])=>(
                  <button key={perm} onClick={()=>setStaff(prev=>prev.map(x=>x.id===s.id?{...x,perms:{...x.perms,[perm]:!x.perms[perm]}}:x))}
                    style={{ background:s.perms[perm]?T.orange:"transparent",border:`1px solid ${s.perms[perm]?T.orange:T.border}`,borderRadius:8,padding:"7px 4px",color:s.perms[perm]?"#fff":T.textMd,cursor:"pointer",fontSize:11,fontFamily:"inherit",fontWeight:s.perms[perm]?700:400,textAlign:"center" }}>{label}</button>
                ))}
              </div>
            </Card>
          ))}
        </div>
      )}

      {tab==="supplier"&&(
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          <button onClick={()=>setShowAddSup(!showAddSup)} style={{...S.btn(),width:"100%",padding:12,fontSize:15}}>+ เพิ่มซัพพลายเออร์</button>
          {showAddSup&&(
            <Card style={{ borderColor:T.borderOr }}>
              <div style={{ color:T.orange,fontWeight:800,fontSize:16,marginBottom:10 }}>➕ เพิ่มซัพพลายเออร์</div>
              <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8 }}>
                {[["ชื่อ","name"],["ประเภท","type"],["เบอร์โทร","phone"]].map(([l,k])=>(
                  <div key={k} style={{ gridColumn:k==="name"?"1/-1":"auto" }}>
                    <div style={{ color:T.textSm,fontSize:13,marginBottom:3 }}>{l}</div>
                    <input value={newSup[k]} onChange={e=>setNewSup(p=>({...p,[k]:e.target.value}))} style={S.inp} placeholder={l} />
                  </div>
                ))}
              </div>
              <div style={{ display:"flex",gap:8,marginTop:10 }}>
                <button onClick={()=>{setSuppliers(prev=>[...prev,{...newSup,id:Date.now()}]);setNewSup({name:"",type:"",phone:"",active:true});setShowAddSup(false);}} style={{...S.btn(),flex:1}}>บันทึก</button>
                <button onClick={()=>setShowAddSup(false)} style={S.ghost}>ยกเลิก</button>
              </div>
            </Card>
          )}
          {suppliers.map(s=>(
            <Card key={s.id} style={{ borderColor:s.active?T.border:T.red+"33" }}>
              {editSup===s.id?(
                <div>
                  <div style={{ color:T.orange,fontWeight:700,fontSize:15,marginBottom:10 }}>✏️ แก้ไข</div>
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10 }}>
                    {[["ชื่อ","name"],["ประเภท","type"],["เบอร์โทร","phone"]].map(([l,k])=>(
                      <div key={k}><div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>{l}</div>
                        <input value={editSupData[k]??s[k]} onChange={e=>setEditSupData(p=>({...p,[k]:e.target.value}))} style={S.inp} /></div>
                    ))}
                  </div>
                  <div style={{ display:"flex",gap:8 }}>
                    <button onClick={()=>{setSuppliers(prev=>prev.map(x=>x.id===s.id?{...x,...editSupData}:x));setEditSup(null);}} style={{...S.btn(),flex:1}}>บันทึก</button>
                    <button onClick={()=>setEditSup(null)} style={S.ghost}>ยกเลิก</button>
                  </div>
                </div>
              ):(
                <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                  <div>
                    <div style={{ fontWeight:700,fontSize:16 }}>{s.name}</div>
                    <div style={{ color:T.textSm,fontSize:13 }}>{s.type} • {s.phone}</div>
                  </div>
                  <div style={{ display:"flex",gap:8 }}>
                    <button onClick={()=>setSuppliers(prev=>prev.map(x=>x.id===s.id?{...x,active:!x.active}:x))}
                      style={{ background:s.active?T.greenLt:T.redLt,border:`1px solid ${s.active?T.green:T.red}44`,borderRadius:7,padding:"5px 10px",color:s.active?T.green:T.red,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit" }}>{s.active?"เปิด":"ปิด"}</button>
                    <button onClick={()=>{setEditSup(s.id);setEditSupData({});}} style={{...S.ghost,padding:"5px 10px",fontSize:12}}>✏️</button>
                  </div>
                </div>
              )}
            </Card>
          ))}
        </div>
      )}

      {tab==="fixed"&&(
        <div style={{ display:"flex",flexDirection:"column",gap:12 }}>
          <Card style={{ background:T.orangeLt,borderColor:T.borderOr }}>
            <div style={{ color:T.orange,fontWeight:800,fontSize:16,marginBottom:4 }}>🔒 ค่าใช้จ่ายคงที่รายเดือน</div>
            <div style={{ color:T.textMd,fontSize:13 }}>แก้ไขได้ตลอด ระบบคำนวณกำไรและ BEP อัตโนมัติ</div>
          </Card>
          {fixedCosts.map((fc,i)=>(
            <Card key={i} style={{ padding:"12px 16px" }}>
              <div style={{ display:"grid",gridTemplateColumns:"1fr auto auto",gap:10,alignItems:"flex-end" }}>
                <div><div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>ชื่อค่าใช้จ่าย</div>
                  <input value={fc.name} onChange={e=>setFixedCosts(prev=>prev.map((f,j)=>j===i?{...f,name:e.target.value}:f))} style={{...S.inp,fontSize:15}} /></div>
                <div style={{ minWidth:120 }}><div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>฿/เดือน</div>
                  <input type="number" value={fc.amount} onChange={e=>setFixedCosts(prev=>prev.map((f,j)=>j===i?{...f,amount:+e.target.value||0}:f))} style={{...S.inp,fontSize:15,textAlign:"right"}} /></div>
                <button onClick={()=>setFixedCosts(prev=>prev.filter((_,j)=>j!==i))}
                  style={{ background:T.redLt,border:`1px solid ${T.red}33`,borderRadius:8,padding:"10px 12px",color:T.red,cursor:"pointer",fontSize:14,fontFamily:"inherit" }}>🗑</button>
              </div>
            </Card>
          ))}
          <button onClick={()=>setFixedCosts(prev=>[...prev,{name:"รายการใหม่",amount:0}])}
            style={{...S.ghost,width:"100%",padding:12,fontSize:15}}>+ เพิ่มรายการ</button>
          <Card style={{ background:T.orangeLt,borderColor:T.borderOr }}>
            <div style={{ fontWeight:700,fontSize:15,marginBottom:10 }}>สรุป</div>
            {fixedCosts.map((fc,i)=>(
              <div key={i} style={{ display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${T.borderOr}`,fontSize:14 }}>
                <span style={{ color:T.textMd }}>{fc.name}</span>
                <span style={{ color:T.orange,fontWeight:600 }}>฿{fmt(fc.amount)}</span>
              </div>
            ))}
            <div style={{ display:"flex",justifyContent:"space-between",paddingTop:10,marginTop:4,fontWeight:800,fontSize:17 }}>
              <span>รวม</span><span style={{ color:T.orange }}>฿{fmt(fixedCosts.reduce((a,b)=>a+b.amount,0))}</span>
            </div>
            <div style={{ display:"flex",gap:20,marginTop:10,paddingTop:10,borderTop:`1px solid ${T.borderOr}`,fontSize:13 }}>
              <div><div style={{ color:T.textSm }}>BEP/วัน</div><div style={{ color:T.orange,fontWeight:700,fontSize:16 }}>฿{fmt(Math.ceil(fixedCosts.reduce((a,b)=>a+b.amount,0)/30))}</div></div>
              <div><div style={{ color:T.textSm }}>เป้า/วัน (×2.5)</div><div style={{ color:T.green,fontWeight:700,fontSize:16 }}>฿{fmt(Math.ceil(fixedCosts.reduce((a,b)=>a+b.amount,0)/30*2.5))}</div></div>
            </div>
          </Card>
        </div>
      )}

      {tab==="line"&&(
        <Card>
          <div style={{ color:T.orange,fontWeight:800,fontSize:17,marginBottom:14 }}>📲 LINE Notify Token</div>
          <div style={{ color:T.textSm,fontSize:13,marginBottom:8 }}>รับ Token จาก notify-bot.line.me</div>
          <input type="password" value={lineToken} onChange={e=>setLineToken(e.target.value)} style={S.inp} placeholder="ใส่ Token ที่นี่..." />
          {lineToken&&<div style={{ color:T.green,fontSize:13,marginTop:6,fontWeight:600 }}>✅ Token บันทึกแล้ว</div>}
        </Card>
      )}
    </div>
  );
}

// ─── FRANCHISE PAGE ──────────────────────────
function FranchisePage({ cf, user, franchises, setFranchises, staff, setStaff }) {
  const [showAdd, setShowAdd] = useState(false);
  const [editId, setEditId] = useState(null);
  const [editData, setEditData] = useState({});
  const [newFr, setNewFr] = useState({ name:"",owner:"",phone:"",royaltyPct:5,monthlyTarget:100000,active:true });
  const mk = today().slice(0,7);

  const frStats = franchises.map(fr=>{
    const frCF=cf.filter(e=>e.branch===fr.id&&e.date.startsWith(mk));
    const frIn=frCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
    const frOut=frCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
    return { ...fr,frIn,frOut,royalty:frIn*(fr.royaltyPct/100),pct:fr.monthlyTarget>0?(frIn/fr.monthlyTarget*100).toFixed(0):0 };
  });

  const myFr=franchises.find(f=>f.id===user.franchiseId);
  const myCF=cf.filter(e=>e.branch===user.franchiseId&&e.date.startsWith(mk));
  const myIn=myCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const myOut=myCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const myRoyalty=myIn*((myFr?.royaltyPct||5)/100);

  if (user.role==="franchise") return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <div>
        <div style={{ fontSize:22,fontWeight:900 }}>🏪 {myFr?.name||"สาขาของฉัน"}</div>
        <div style={{ color:T.textSm,fontSize:14 }}>เจ้าของ: {myFr?.owner}</div>
      </div>
      <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:12 }}>
        {[["💰","รายรับ",`฿${fmt(myIn)}`,T.green],["💸","รายจ่าย",`฿${fmt(myOut)}`,T.red],["📈","กำไร",`฿${fmt(myIn-myOut)}`,myIn-myOut>=0?T.green:T.red],["🤝","Royalty",`฿${fmt(Math.round(myRoyalty))}`,T.orange]].map(([ic,l,v,c])=>(
          <Card key={l} style={{ padding:"14px 16px" }}><div style={{ fontSize:22,marginBottom:3 }}>{ic}</div><div style={{ color:T.textSm,fontSize:13 }}>{l}</div><div style={{ color:c,fontWeight:800,fontSize:20 }}>{v}</div></Card>
        ))}
      </div>
      <Card style={{ background:T.orangeLt,borderColor:T.borderOr }}>
        <div style={{ display:"flex",justifyContent:"space-between",marginBottom:8 }}>
          <div style={{ color:T.orange,fontWeight:700,fontSize:15 }}>เป้ายอดขาย</div>
          <div style={{ textAlign:"right" }}>
            <div style={{ color:T.orange,fontWeight:800 }}>฿{fmt(myIn)} / ฿{fmt(myFr?.monthlyTarget||0)}</div>
            <div style={{ color:T.textSm,fontSize:12 }}>{(myFr?.monthlyTarget>0?(myIn/myFr.monthlyTarget*100):0).toFixed(0)}%</div>
          </div>
        </div>
        <div style={{ background:"rgba(255,255,255,0.5)",borderRadius:8,height:12 }}>
          <div style={{ background:T.orange,height:"100%",borderRadius:8,width:`${Math.min(myFr?.monthlyTarget>0?(myIn/myFr.monthlyTarget*100):0,100)}%` }}/>
        </div>
      </Card>
    </div>
  );

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <Hdr title="🏪 จัดการแฟรนไชส์" action={<button onClick={()=>setShowAdd(!showAdd)} style={S.btn()}>+ เพิ่มแฟรนไชส์</button>} />
      {showAdd&&(
        <Card style={{ borderColor:T.borderOr }}>
          <div style={{ color:T.orange,fontWeight:800,fontSize:16,marginBottom:12 }}>➕ เพิ่มแฟรนไชส์</div>
          <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
            {[["ชื่อสาขา","name"],["เจ้าของ","owner"],["เบอร์โทร","phone"]].map(([l,k])=>(
              <div key={k} style={{ gridColumn:k==="name"?"1/-1":"auto" }}>
                <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>{l}</div>
                <input value={newFr[k]} onChange={e=>setNewFr(p=>({...p,[k]:e.target.value}))} style={S.inp} />
              </div>
            ))}
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>Royalty (%)</div><input type="number" value={newFr.royaltyPct} onChange={e=>setNewFr(p=>({...p,royaltyPct:+e.target.value}))} style={S.inp} /></div>
            <div><div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>เป้า/เดือน</div><input type="number" value={newFr.monthlyTarget} onChange={e=>setNewFr(p=>({...p,monthlyTarget:+e.target.value}))} style={S.inp} /></div>
          </div>
          <div style={{ display:"flex",gap:8,marginTop:12 }}>
            <button onClick={()=>{const nId="fr"+Date.now();setFranchises(prev=>[...prev,{...newFr,id:nId}]);setStaff(prev=>[...prev,{id:nId,name:`แฟรนไชส์ ${newFr.name}`,pin:"0000",role:"franchise",franchiseId:nId,active:true,perms:{cashflow:true,stock:true,purchase:false,report:true,ai:false,admin:false,viewPrice:true}}]);setShowAdd(false);}} style={{...S.btn(),flex:1}}>สร้าง</button>
            <button onClick={()=>setShowAdd(false)} style={S.ghost}>ยกเลิก</button>
          </div>
        </Card>
      )}
      <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:12 }}>
        {[["🏪","สาขา",`${franchises.filter(f=>f.active).length} สาขา`,T.blue],["💰","ยอดรวม",`฿${fmt(frStats.reduce((a,b)=>a+b.frIn,0))}`,T.green],["🤝","Royalty รวม",`฿${fmt(Math.round(frStats.reduce((a,b)=>a+b.royalty,0)))}`,T.orange]].map(([ic,l,v,c])=>(
          <Card key={l} style={{ padding:"14px 16px" }}><div style={{ fontSize:22,marginBottom:3 }}>{ic}</div><div style={{ color:T.textSm,fontSize:13 }}>{l}</div><div style={{ color:c,fontWeight:800,fontSize:20 }}>{v}</div></Card>
        ))}
      </div>
      {frStats.map(fr=>(
        <Card key={fr.id} style={{ borderColor:fr.active?T.border:T.red+"33" }}>
          {editId===fr.id?(
            <div>
              <div style={{ color:T.orange,fontWeight:700,fontSize:15,marginBottom:10 }}>✏️ แก้ไข</div>
              <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10 }}>
                {[["ชื่อ","name"],["เจ้าของ","owner"],["โทร","phone"],["Royalty%","royaltyPct"],["เป้า/เดือน","monthlyTarget"]].map(([l,k])=>(
                  <div key={k}><div style={{ color:T.textSm,fontSize:12,marginBottom:3 }}>{l}</div>
                    <input value={editData[k]??fr[k]} onChange={e=>setEditData(p=>({...p,[k]:e.target.value}))} style={{...S.inp,fontSize:14}} /></div>
                ))}
              </div>
              <div style={{ display:"flex",gap:8 }}>
                <button onClick={()=>{setFranchises(prev=>prev.map(f=>f.id===editId?{...f,...editData,royaltyPct:+editData.royaltyPct||f.royaltyPct,monthlyTarget:+editData.monthlyTarget||f.monthlyTarget}:f));setEditId(null);}} style={{...S.btn(),flex:1}}>บันทึก</button>
                <button onClick={()=>setEditId(null)} style={S.ghost}>ยกเลิก</button>
              </div>
            </div>
          ):(
            <>
              <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10 }}>
                <div>
                  <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                    <span style={{ fontWeight:800,fontSize:17 }}>{fr.name}</span>
                    <span style={{ background:fr.active?T.greenLt:T.redLt,color:fr.active?T.green:T.red,borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700 }}>{fr.active?"เปิด":"ปิด"}</span>
                  </div>
                  <div style={{ color:T.textSm,fontSize:13,marginTop:2 }}>{fr.owner} • {fr.phone} • Royalty {fr.royaltyPct}%</div>
                </div>
                <button onClick={()=>{setEditId(fr.id);setEditData({});}} style={{...S.ghost,padding:"6px 10px",fontSize:13}}>✏️</button>
              </div>
              <div style={{ marginBottom:8 }}>
                <div style={{ display:"flex",justifyContent:"space-between",marginBottom:4,fontSize:14 }}>
                  <span style={{ color:T.textMd }}>ยอดขาย {mk}</span>
                  <span style={{ color:T.orange,fontWeight:700 }}>฿{fmt(fr.frIn)} / ฿{fmt(fr.monthlyTarget)} ({fr.pct}%)</span>
                </div>
                <div style={{ background:T.bg,borderRadius:6,height:10 }}>
                  <div style={{ background:`linear-gradient(90deg,${T.orange},${T.orangeDk})`,width:`${Math.min(+fr.pct,100)}%`,height:"100%",borderRadius:6 }}/>
                </div>
              </div>
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

// ─── DASHBOARD PAGE ──────────────────────────
function DashboardPage({ cf, stock, movements, user, fixedCosts, setPage }) {
  const today_str = today();
  const mk = today_str.slice(0,7);
  const yd = new Date(); yd.setDate(yd.getDate()-1);
  const yd_str = yd.toISOString().split("T")[0];

  const myCF = user.role==="owner" ? cf : cf.filter(e=>e.branch===user.franchiseId||e.staffId===user.id);

  // ยอดวันนี้
  const todayCF  = myCF.filter(e=>e.date===today_str);
  const todayIn  = todayCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const todayOut = todayCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);

  // ยอดเมื่อวาน
  const ydCF  = myCF.filter(e=>e.date===yd_str);
  const ydIn  = ydCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);

  // ยอดเดือนนี้
  const mCF  = myCF.filter(e=>e.date.startsWith(mk));
  const mIn  = mCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mOut = mCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const totalFixed = fixedCosts.reduce((a,b)=>a+b.amount,0);
  const netP = mIn - mOut - totalFixed;

  // สต็อคแจ้งเตือน
  const outStock      = stock.filter(s=>stockSt(s)==="out");
  const criticalStock = stock.filter(s=>stockSt(s)==="critical");
  const lowStock      = stock.filter(s=>stockSt(s)==="low");
  const alertStock    = [...outStock,...criticalStock,...lowStock];

  // 7 วันล่าสุด
  const last7 = Array.from({length:7},(_,i)=>{
    const d=new Date(); d.setDate(d.getDate()-i);
    const k=d.toISOString().split("T")[0];
    return { d:k, inc:myCF.filter(e=>e.date===k&&e.flow==="in").reduce((a,b)=>a+b.amount,0) };
  }).reverse();
  const maxInc = Math.max(...last7.map(d=>d.inc),1);

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <div>
        <div style={{ color:T.text,fontSize:22,fontWeight:900 }}>🏠 ภาพรวมร้าน</div>
        <div style={{ color:T.textSm,fontSize:14 }}>{today_str}</div>
      </div>

      {/* KPI วันนี้ */}
      <div style={{ display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12 }}>
        <Card style={{ padding:"16px 18px",borderLeft:`4px solid ${T.green}` }}>
          <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>💰 รายรับวันนี้</div>
          <div style={{ color:T.green,fontWeight:900,fontSize:26 }}>฿{fmt(todayIn)}</div>
          {ydIn>0&&<div style={{ fontSize:12,color:todayIn>=ydIn?T.green:T.red,marginTop:4 }}>
            {todayIn>=ydIn?"▲":"▼"} {ydIn>0?Math.abs(((todayIn-ydIn)/ydIn)*100).toFixed(0):0}% vs เมื่อวาน
          </div>}
        </Card>
        <Card style={{ padding:"16px 18px",borderLeft:`4px solid ${T.red}` }}>
          <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>💸 รายจ่ายวันนี้</div>
          <div style={{ color:T.red,fontWeight:900,fontSize:26 }}>฿{fmt(todayOut)}</div>
          <div style={{ fontSize:12,color:T.textSm,marginTop:4 }}>กำไรวันนี้ <b style={{ color:todayIn-todayOut>=0?T.green:T.red }}>฿{fmt(todayIn-todayOut)}</b></div>
        </Card>
        <Card style={{ padding:"16px 18px",borderLeft:`4px solid ${T.orange}` }}>
          <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>📅 รายรับเดือนนี้</div>
          <div style={{ color:T.orange,fontWeight:900,fontSize:22 }}>฿{fmt(mIn)}</div>
          <div style={{ fontSize:12,color:T.textSm,marginTop:4 }}>เป้า BEP ฿{fmt(totalFixed)}</div>
        </Card>
        <Card style={{ padding:"16px 18px",borderLeft:`4px solid ${netP>=0?T.green:T.red}` }}>
          <div style={{ color:T.textSm,fontSize:13,marginBottom:4 }}>📈 กำไรสุทธิเดือนนี้</div>
          <div style={{ color:netP>=0?T.green:T.red,fontWeight:900,fontSize:22 }}>฿{fmt(netP)}</div>
          <div style={{ fontSize:12,color:T.textSm,marginTop:4 }}>หลังหักต้นทุนคงที่ ฿{fmt(totalFixed)}</div>
        </Card>
      </div>

      {/* กราฟ 7 วัน */}
      <Card>
        <div style={{ fontWeight:700,fontSize:15,marginBottom:12,color:T.text }}>📊 รายรับ 7 วันล่าสุด</div>
        <div style={{ display:"flex",alignItems:"flex-end",gap:6,height:100 }}>
          {last7.map(d=>{
            const h = maxInc>0 ? Math.max((d.inc/maxInc)*88,d.inc>0?12:2) : 2;
            const isToday = d.d===today_str;
            return (
              <div key={d.d} style={{ flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4 }}>
                <div style={{ fontSize:10,color:T.textXs,fontWeight:600 }}>
                  {d.inc>0?`฿${Math.round(d.inc/1000)}k`:""}
                </div>
                <div style={{ width:"100%",height:h,
                  background:isToday?T.orange:`linear-gradient(180deg,${T.orange}88,${T.orange}44)`,
                  borderRadius:"4px 4px 0 0",transition:"height .3s" }}/>
                <div style={{ fontSize:10,color:isToday?T.orange:T.textXs,fontWeight:isToday?700:400 }}>
                  {d.d.slice(8)}
                </div>
              </div>
            );
          })}
        </div>
      </Card>

      {/* แจ้งเตือนสต็อค */}
      {alertStock.length>0&&(
        <Card style={{ borderColor:T.red+"44",background:T.redLt }}>
          <div style={{ color:T.red,fontWeight:800,fontSize:16,marginBottom:10 }}>
            🚨 สต็อคต้องดูแล ({alertStock.length} รายการ)
          </div>
          {alertStock.map(s=>{
            const st=stockSt(s);
            return (
              <div key={s.id} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",
                padding:"8px 0",borderBottom:`1px solid ${T.red}22` }}>
                <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                  <Badge status={st} />
                  <span style={{ fontWeight:600,fontSize:15 }}>{s.name}</span>
                </div>
                <div style={{ textAlign:"right" }}>
                  <div style={{ color:ST_COLOR[st],fontWeight:700,fontSize:15 }}>{s.qty} {s.unit}</div>
                  <div style={{ color:T.textXs,fontSize:11 }}>ขั้นต่ำ {s.minQty} {s.unit}</div>
                </div>
              </div>
            );
          })}
          <button onClick={()=>setPage("stock")}
            style={{...S.btn(T.red),width:"100%",marginTop:10,padding:10,fontSize:14}}>
            ไปจัดการสต็อค →
          </button>
        </Card>
      )}

      {/* สต็อคคงเหลือ - ทุกคนเห็นได้ */}
      <Card>
        <div style={{ fontWeight:700,fontSize:15,marginBottom:10,color:T.text }}>📦 สถานะสต็อควัตถุดิบ</div>
        {stock.map(s=>{
          const st=stockSt(s);
          const pct=s.minQty>0?Math.min((s.qty/s.minQty)*100,200):100;
          return (
            <div key={s.id} style={{ marginBottom:12 }}>
              <div style={{ display:"flex",justifyContent:"space-between",marginBottom:4 }}>
                <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                  <span style={{ fontSize:14,fontWeight:600 }}>{s.name}</span>
                  <Badge status={st} />
                </div>
                <span style={{ color:ST_COLOR[st],fontWeight:700,fontSize:14 }}>{s.qty} {s.unit}</span>
              </div>
              <div style={{ background:T.bg,borderRadius:4,height:7,overflow:"hidden" }}>
                <div style={{ background:pct<50?T.red:pct<100?T.yellow:T.green,
                  width:`${Math.min(pct,100)}%`,height:"100%",borderRadius:4,transition:"width .3s" }}/>
              </div>
              <div style={{ fontSize:11,color:T.textXs,marginTop:2 }}>ขั้นต่ำ {s.minQty} {s.unit} • เหลือ {s.dailyUse>0?(s.qty/s.dailyUse).toFixed(1):"∞"} วัน</div>
            </div>
          );
        })}
        {stock.length===0&&<div style={{ color:T.textSm,textAlign:"center",padding:16 }}>ยังไม่มีสต็อค</div>}
      </Card>

      {/* ลิงก์ด่วน */}
      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
        {[
          ...(user.perms?.cashflow!==false?[["💵","กรอก Cash Flow","cashflow"]]:[]),
          ...(user.perms?.stock!==false?[["📦","บันทึกสต็อค",user.role==="owner"?"stock":"staffstock"]]:[]),
        ].map(([ic,l,pg])=>(
          <Card key={pg} style={{ padding:"16px 18px",cursor:"pointer",textAlign:"center" }} onClick={()=>setPage(pg)}>
            <div style={{ fontSize:28,marginBottom:6 }}>{ic}</div>
            <div style={{ fontWeight:700,fontSize:14,color:T.text }}>{l}</div>
          </Card>
        ))}
      </div>
    </div>
  );
}


// ─── REPORT PAGE ─────────────────────────────
function ReportPage({ cf, stock, movements, user, fixedCosts }) {
  const mk = today().slice(0,7);
  const myCF = user.role==="owner"?cf:cf.filter(e=>e.branch===user.franchiseId);
  const mCF = myCF.filter(e=>e.date.startsWith(mk));
  const mIn  = mCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mOut = mCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const totalFixed = fixedCosts.reduce((a,b)=>a+b.amount,0);
  const cogs = mCF.filter(e=>e.flow==="out"&&["วัตถุดิบ/ผัก","หมู/เนื้อ/ทะเล","บรรจุภัณฑ์"].includes(e.cat)).reduce((a,b)=>a+b.amount,0);
  const grossP = mIn - cogs;
  const netP   = mIn - mOut - totalFixed;
  const grossM = mIn>0?(grossP/mIn*100).toFixed(1):0;

  return (
    <div style={{ display:"flex",flexDirection:"column",gap:16 }}>
      <Hdr title="📊 รายงานกำไรขาดทุน" />
      <Card style={{ background:T.orangeLt,borderColor:T.borderOr }}>
        <div style={{ color:T.orange,fontWeight:700,fontSize:14,marginBottom:10 }}>เดือน {mk}</div>
        {[["💰 รายรับรวม",mIn,T.green,false],["− ต้นทุนวัตถุดิบ",cogs,T.red,true],
          ["= กำไรขั้นต้น",grossP,T.green,false],["− ค่าใช้จ่ายดำเนินการ",mOut-cogs+totalFixed,T.yellow,true],
          ["= กำไรสุทธิ",netP,netP>=0?T.green:T.red,false]].map(([l,v,c,indent])=>(
          <div key={l} style={{ display:"flex",justifyContent:"space-between",padding:"10px 0",
            borderBottom:`1px solid ${T.borderOr}`,paddingLeft:indent?20:0 }}>
            <span style={{ color:T.textMd,fontSize:15 }}>{l}</span>
            <span style={{ color:c,fontWeight:800,fontSize:16 }}>฿{fmt(v)}</span>
          </div>
        ))}
        <div style={{ display:"flex",gap:20,marginTop:12,paddingTop:10,borderTop:`1px solid ${T.borderOr}`,fontSize:13 }}>
          <div><div style={{ color:T.textSm }}>Gross Margin</div><div style={{ color:T.orange,fontWeight:700,fontSize:16 }}>{grossM}%</div></div>
          <div><div style={{ color:T.textSm }}>BEP/วัน</div><div style={{ color:T.orange,fontWeight:700,fontSize:16 }}>฿{fmt(Math.ceil(totalFixed/30))}</div></div>
        </div>
      </Card>
      <Card>
        <div style={{ fontWeight:800,fontSize:16,marginBottom:10 }}>สต็อคคงเหลือ</div>
        {stock.filter(s=>stockSt(s)!=="ok").map(s=>(
          <div key={s.id} style={{ display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.bg}`,fontSize:14 }}>
            <div style={{ display:"flex",alignItems:"center",gap:8 }}>
              <span style={{ fontWeight:600 }}>{s.name}</span>
              <Badge status={stockSt(s)} />
            </div>
            <span style={{ color:ST_COLOR[stockSt(s)],fontWeight:700 }}>{s.qty} {s.unit}</span>
          </div>
        ))}
        {stock.every(s=>stockSt(s)==="ok")&&<div style={{ color:T.green,textAlign:"center",padding:16 }}>✅ สต็อคปกติทั้งหมด</div>}
      </Card>
    </div>
  );
}

// ─── MAIN APP ────────────────────────────────
export default function App() {
  const [user, setUser]               = useState(null);
  const [page, setPage]               = useState("cashflow");
  const [stock, setStock]             = useState(INIT_STOCK);
  const [cf, setCF]                   = useState([]);
  const [movements, setMovements]     = useState([]);
  const [staff, setStaff]             = useState(INIT_STAFF);
  const [suppliers, setSuppliers]     = useState(INIT_SUPPLIERS);
  const [franchises, setFranchises]   = useState(INIT_FRANCHISES);
  const [fixedCosts, setFixedCosts]   = useState(INIT_FIXED);
  const [lineToken, setLineToken]     = useState("");
  const [notifOpen, setNotifOpen]     = useState(false);
  const [dbReady, setDbReady]         = useState(false);
  const [dbLoading, setDbLoading]     = useState(true);

  // Load from Supabase
  useEffect(()=>{
    async function load() {
      setDbLoading(true);
      try {
        const [cfData,stockData,mvData] = await Promise.all([db.getCF(),db.getStock(),db.getMvs()]);
        if (cfData?.length>0) setCF(cfData.map(r=>({id:r.id,date:r.date,flow:r.flow,cat:r.cat,itemName:r.item_name||"",amount:r.amount,method:r.method,note:r.note||"",branch:r.branch||"main",staffId:r.staff_id||"owner"})));
        if (stockData?.length>0) setStock(stockData.map(r=>({id:r.id,name:r.name,unit:r.unit,qty:r.qty,minQty:r.min_qty,dailyUse:r.daily_use,supplierId:r.supplier_id||1,costHistory:r.cost_history||[]})));
        if (mvData?.length>0) setMovements(mvData.map(r=>({id:r.id,itemId:r.item_id,type:r.type,qty:r.qty,unitCost:r.unit_cost||0,date:r.date,staffId:r.staff_id||"",note:r.note||"",branch:r.branch||"main"})));
        setDbReady(true);
      } catch { setDbReady(false); }
      setDbLoading(false);
    }
    load();
  },[]);

  // Save stock to Supabase when changed
  const saveStock = useCallback(async (newStock) => {
    setStock(newStock);
    if (dbReady) {
      const rows=newStock.map(s=>({id:s.id,name:s.name,unit:s.unit,qty:s.qty,min_qty:s.minQty,daily_use:s.dailyUse,supplier_id:s.supplierId||1,cost_history:s.costHistory||[]}));
      db.upsertStock(rows).catch(()=>{});
    }
  },[dbReady]);

  const notifications = useMemo(()=>{
    const notes=[];
    stock.filter(s=>["out","critical"].includes(stockSt(s))).forEach(s=>
      notes.push({id:s.id,icon:"📦",title:`${s.name} ${stockSt(s)==="out"?"หมด":"น้อยมาก"}`,body:`เหลือ ${s.qty} ${s.unit}`}));
    return notes;
  },[stock]);

  if (!user) return <LoginPage staff={staff} onLogin={u=>{setUser(u);setPage("dashboard");}} />;

  // Emergency page
  if (user.id==="emergency") return (
    <div style={{ minHeight:"100vh",background:T.bg,fontFamily:"'Noto Sans Thai','Noto Sans',sans-serif",
      display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,gap:16 }}>
      <div style={{ width:80,height:80,borderRadius:20,background:T.red,display:"flex",alignItems:"center",justifyContent:"center",fontSize:44 }}>🚨</div>
      <div style={{ color:T.red,fontWeight:900,fontSize:22 }}>Emergency Mode</div>
      <div style={{ color:T.textMd,fontSize:15,textAlign:"center",maxWidth:320 }}>
        โหมดฉุกเฉิน — ใช้สำหรับล้างข้อมูลเท่านั้น
      </div>
      <div style={{ display:"flex",flexDirection:"column",gap:10,width:"100%",maxWidth:320 }}>
        <button onClick={()=>{if(window.confirm("ล้าง Cash Flow ทั้งหมด?\nข้อมูลจะหายถาวร")){setCF([]);if(dbReady)db.clearCF().catch(()=>{});alert("ล้าง CF เรียบร้อย");}}} style={{...S.btn(T.red),padding:14,fontSize:16,width:"100%"}}>🗑 ล้าง Cash Flow</button>
        <button onClick={()=>{if(window.confirm("ล้างสต็อค + ประวัติทั้งหมด?\nข้อมูลจะหายถาวร")){setStock(INIT_STOCK);setMovements([]);if(dbReady)db.clearMvs().catch(()=>{});alert("ล้างสต็อคเรียบร้อย");}}} style={{...S.btn(T.orange),padding:14,fontSize:16,width:"100%"}}>🗑 ล้างสต็อค + ประวัติ</button>
        <button onClick={()=>{if(window.confirm("⚠️ ล้างข้อมูลทั้งหมด?\nCash Flow + สต็อค + ประวัติ จะหายหมด")){setCF([]);setStock(INIT_STOCK);setMovements([]);if(dbReady){db.clearCF().catch(()=>{});db.clearMvs().catch(()=>{});}alert("ล้างทั้งหมดเรียบร้อย");}}} style={{...S.btn(T.red),padding:14,fontSize:16,width:"100%",border:"2px solid #7f1d1d"}}>⚠️ ล้างข้อมูลทั้งหมด</button>
        <button onClick={()=>setUser(null)} style={{...S.ghost,padding:14,fontSize:16,width:"100%"}}>← ออกจากระบบ</button>
      </div>
    </div>
  );

  const p = user.perms;
  const isOwner = user.role==="owner";
  const isFr = user.role==="franchise";

  const ownerNav = [
    {id:"dashboard", icon:"🏠",label:"หลัก"},
    {id:"cashflow",  icon:"💵",label:"Cash Flow"},
    {id:"stock",     icon:"📦",label:"สต็อค"},
    {id:"report",    icon:"📊",label:"รายงาน"},
    {id:"franchise", icon:"🏪",label:"แฟรนไชส์"},
    {id:"settings",  icon:"⚙️",label:"ตั้งค่า"},
  ];
  const frNav = [
    {id:"dashboard", icon:"🏠",label:"หลัก"},
    {id:"cashflow",  icon:"💵",label:"Cash Flow"},
    {id:"staffstock",icon:"📦",label:"สต็อค"},
    {id:"franchise", icon:"🏪",label:"สาขาฉัน"},
    {id:"report",    icon:"📊",label:"รายงาน"},
  ];
  const staffNav = [
    {id:"dashboard", icon:"🏠",label:"หลัก"},
    ...(p.cashflow ?[{id:"cashflow",  icon:"💵",label:"Cash Flow"}]:[]),
    ...(p.stock    ?[{id:"staffstock",icon:"📦",label:"สต็อค"}]:[]),
  ];
  const nav = isOwner ? ownerNav : isFr ? frNav : staffNav;

  const pages = {
    dashboard:  <DashboardPage cf={cf} stock={stock} movements={movements} user={user} fixedCosts={fixedCosts} setPage={setPage} />,
    cashflow:   <CashflowPage cf={cf} setCF={setCF} user={user} dbReady={dbReady} />,
    stock:      <StockPage stock={stock} setStock={saveStock} movements={movements} setMovements={setMovements} user={user} suppliers={suppliers} />,
    staffstock: <StaffStockPage stock={stock} setStock={saveStock} movements={movements} setMovements={setMovements} user={user} />,
    report:     <ReportPage cf={cf} stock={stock} movements={movements} user={user} fixedCosts={fixedCosts} />,
    franchise:  <FranchisePage cf={cf} user={user} franchises={franchises} setFranchises={setFranchises} staff={staff} setStaff={setStaff} />,
    settings:   <SettingsPage staff={staff} setStaff={setStaff} lineToken={lineToken} setLineToken={setLineToken} suppliers={suppliers} setSuppliers={setSuppliers} fixedCosts={fixedCosts} setFixedCosts={setFixedCosts} />,
  };

  if (dbLoading) return (
    <div style={{ minHeight:"100vh",background:T.bg,fontFamily:"'Noto Sans Thai','Noto Sans',sans-serif",
      display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16 }}>
      <div style={{ width:80,height:80,borderRadius:20,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:44 }}>🫕</div>
      <div style={{ color:T.orange,fontWeight:800,fontSize:20 }}>ไท่กั๋วหม่าล่า</div>
      <div style={{ color:T.textSm,fontSize:15 }}>กำลังโหลดข้อมูล...</div>
      <style>{`@keyframes pulse{0%,100%{opacity:0.3}50%{opacity:1}}`}</style>
      <div style={{ display:"flex",gap:6 }}>
        {[0,1,2].map(i=><div key={i} style={{ width:10,height:10,borderRadius:"50%",background:T.orange,animation:`pulse 1.2s ${i*0.2}s infinite` }}/>)}
      </div>
    </div>
  );

  return (
    <div style={{ minHeight:"100vh",background:T.bg,fontFamily:"'Noto Sans Thai','Noto Sans',sans-serif",fontSize:16 }}>
      {/* Header */}
      <div style={{ background:T.card,borderBottom:`1px solid ${T.border}`,padding:"11px 18px",
        display:"flex",alignItems:"center",gap:12,position:"sticky",top:0,zIndex:50,boxShadow:T.shadow }}>
        <div style={{ width:36,height:36,borderRadius:10,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,
          display:"flex",alignItems:"center",justifyContent:"center",fontSize:20 }}>🫕</div>
        <div>
          <div style={{ color:T.orange,fontWeight:900,fontSize:15 }}>ไท่กั๋วหม่าล่า</div>
          <div style={{ color:T.textXs,fontSize:11 }}>{user.name} • {isOwner?"👑":isFr?"🏪":"👷"}</div>
        </div>
        {dbReady&&<div style={{ marginLeft:"auto",background:T.greenLt,border:`1px solid ${T.green}33`,borderRadius:6,padding:"3px 8px",fontSize:11,color:T.green,fontWeight:600 }}>● DB</div>}
        {/* Notification bell */}
        <button onClick={()=>setNotifOpen(!notifOpen)}
          style={{ position:"relative",background:"transparent",border:"none",cursor:"pointer",padding:8,borderRadius:10,marginLeft:dbReady?4:"auto" }}>
          <span style={{ fontSize:22 }}>🔔</span>
          {notifications.length>0&&<span style={{ position:"absolute",top:4,right:4,background:T.red,color:"#fff",borderRadius:"50%",width:17,height:17,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800 }}>{notifications.length}</span>}
        </button>
        <button onClick={()=>setUser(null)}
          style={{ background:T.bg,border:`1px solid ${T.border}`,borderRadius:8,padding:"5px 11px",color:T.textMd,cursor:"pointer",fontSize:13,fontFamily:"inherit" }}>ออก</button>
      </div>

      {/* Notification dropdown */}
      {notifOpen&&<div onClick={()=>setNotifOpen(false)} style={{ position:"fixed",inset:0,zIndex:190 }}/>}
      {notifOpen&&(
        <div style={{ position:"fixed",top:62,right:12,zIndex:200,width:300,background:T.card,
          border:`1px solid ${T.border}`,borderRadius:14,boxShadow:T.shadowMd,padding:"14px 16px",maxHeight:"60vh",overflowY:"auto" }}>
          <div style={{ fontWeight:800,fontSize:15,marginBottom:10 }}>🔔 แจ้งเตือน</div>
          {notifications.length===0
            ? <div style={{ color:T.textSm,textAlign:"center",padding:16 }}>ไม่มีการแจ้งเตือน ✅</div>
            : notifications.map(n=>(
              <div key={n.id} style={{ display:"flex",gap:10,padding:"8px 0",borderBottom:`1px solid ${T.bg}` }}>
                <span style={{ fontSize:20 }}>{n.icon}</span>
                <div><div style={{ fontWeight:600,fontSize:14 }}>{n.title}</div><div style={{ color:T.textSm,fontSize:12 }}>{n.body}</div></div>
              </div>
            ))
          }
        </div>
      )}

      {/* Content */}
      <div style={{ padding:"18px 14px 95px",maxWidth:900,margin:"0 auto" }}>
        {pages[page]||pages["cashflow"]}
      </div>

      {/* Bottom nav */}
      <div style={{ position:"fixed",bottom:0,left:0,right:0,background:T.card,
        borderTop:`1px solid ${T.border}`,display:"flex",justifyContent:"space-around",
        padding:"9px 0 15px",zIndex:100,boxShadow:"0 -2px 8px rgba(0,0,0,0.06)" }}>
        {nav.map(n=>(
          <button key={n.id} onClick={()=>setPage(n.id)}
            style={{ background:"none",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2,padding:"2px 8px",minWidth:50 }}>
            <span style={{ fontSize:22 }}>{n.icon}</span>
            <span style={{ fontSize:10,color:page===n.id?T.orange:T.textXs,fontWeight:page===n.id?800:400 }}>{n.label}</span>
            {page===n.id&&<div style={{ width:18,height:3,borderRadius:2,background:T.orange }}/>}
          </button>
        ))}
      </div>
    </div>
  );
}
