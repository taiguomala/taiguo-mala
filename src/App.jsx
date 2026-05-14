import { useState, useMemo, useEffect, useCallback, useRef } from "react";

const SUPA_URL = "https://klmowpluuvjmbvvmqzep.supabase.co";
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtsbW93cGx1dXZqbWJ2dm1xemVwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzgzMTIzMTMsImV4cCI6MjA5Mzg4ODMxM30.aXQz6WBqE8US5_-ij6GvvY0XaCykMag8x6W2a6uAwMU";

async function supa(path, opts={}) {
  try {
    const r = await fetch(`${SUPA_URL}/rest/v1/${path}`, {
      ...opts,
      headers: { apikey:SUPA_KEY, Authorization:`Bearer ${SUPA_KEY}`, "Content-Type":"application/json", Prefer:opts.prefer||"return=representation", ...opts.headers }
    });
    if (!r.ok) return null;
    const t = await r.text(); return t ? JSON.parse(t) : [];
  } catch { return null; }
}
const db = {
  getCF: () => supa("cashflow?order=date.desc,id.desc"),
  addCF: r => supa("cashflow", { method:"POST", body:JSON.stringify(r) }),
  delCF: id => supa(`cashflow?id=eq.${id}`, { method:"DELETE", prefer:"" }),
  clearCF: () => supa("cashflow?id=gt.0", { method:"DELETE", prefer:"" }),
  getStock: () => supa("stock?order=name.asc"),
  upsertStock: rows => supa("stock", { method:"POST", body:JSON.stringify(rows), headers:{ Prefer:"resolution=merge-duplicates,return=representation" }}),
  getMvs: () => supa("movements?order=date.desc,id.desc"),
  getAtt: () => supa("attendance?order=date.desc,id.desc"),
  addAtt: r => supa("attendance", { method:"POST", body:JSON.stringify(r) }),
  updateAtt: (id,data) => supa(`attendance?id=eq.${id}`, { method:"PATCH", body:JSON.stringify(data) }),
  clearAtt: () => supa("attendance?id=gt.0", { method:"DELETE", prefer:"" }),
  addMv: r => supa("movements", { method:"POST", body:JSON.stringify(r) }),
  clearMvs: () => supa("movements?id=gt.0", { method:"DELETE", prefer:"" }),
};

const T = {
  bg:"#f4f4f5", card:"#fff", border:"#e4e4e7", text:"#18181b",
  textMd:"#52525b", textSm:"#71717a", textXs:"#a1a1aa",
  orange:"#f97316", orangeDk:"#ea580c", orangeLt:"#fff7ed", borderOr:"#fed7aa",
  green:"#16a34a", greenLt:"#f0fdf4", red:"#dc2626", redLt:"#fef2f2",
  yellow:"#d97706", yellowLt:"#fffbeb", blue:"#2563eb", blueLt:"#eff6ff",
};
const F = "'Noto Sans Thai',sans-serif";
const S = {
  inp: { width:"100%", padding:"10px 12px", border:`1px solid #e4e4e7`, borderRadius:10, fontSize:15, fontFamily:F, background:"#fff", outline:"none", boxSizing:"border-box", color:"#18181b" },
  btn: (bg="#f97316") => ({ background:bg, color:"#fff", border:"none", borderRadius:10, padding:"9px 18px", cursor:"pointer", fontWeight:700, fontSize:15, fontFamily:F }),
  ghost: { background:"transparent", border:"1px solid #e4e4e7", borderRadius:10, padding:"9px 14px", cursor:"pointer", fontSize:14, color:"#52525b", fontFamily:F },
  card: { background:"#fff", border:"1px solid #e4e4e7", borderRadius:14, padding:"16px 18px", boxShadow:"0 1px 3px rgba(0,0,0,.08)" },
};

const today = () => new Date().toISOString().split("T")[0];

// Simple QR Code SVG generator (encodes URL as visual pattern)
function makeQRSvg(text, size=200) {
  // Use a deterministic hash to generate a QR-like pattern
  // This is a visual representation - for actual scanning use the URL directly
  const hash = text.split("").reduce((a,c)=>((a<<5)-a+c.charCodeAt(0))|0, 0);
  const cells = 21;
  const cell = size/cells;
  const modules = [];
  // Generate deterministic pattern based on text
  for(let r=0;r<cells;r++){
    for(let col=0;col<cells;col++){
      // finder patterns (corners)
      const inFinder=(rr,cc)=>(rr<7&&cc<7)||(rr<7&&cc>cells-8)||(rr>cells-8&&cc<7);
      if(inFinder(r,col)){
        const edge=(rr,cc)=>rr===0||rr===6||cc===0||cc===6;
        const inner=(rr,cc)=>rr>=2&&rr<=4&&cc>=2&&cc<=4;
        const topRight=(rr,cc)=>rr<7&&cc>cells-8;
        const botLeft=(rr,cc)=>rr>cells-8&&cc<7;
        let base=r<7&&col<7;
        let c2=r<7&&col>cells-8;
        let c3=r>cells-8&&col<7;
        const fr=c2?r:c3?r-cells+7:r;
        const fc=c2?col-cells+7:col;
        modules.push(edge(fr,fc)||inner(fr,fc)?1:0);
      } else {
        // Data pattern based on hash
        const v=Math.abs((hash^(r*31+col*17)^(r*col))%7);
        modules.push(v<3?1:0);
      }
    }
  }
  const rects=[];
  for(let r=0;r<cells;r++){
    for(let cc=0;cc<cells;cc++){
      if(modules[r*cells+cc]){
        rects.push(`<rect x="${cc*cell}" y="${r*cell}" width="${cell}" height="${cell}" fill="#000"/>`);
      }
    }
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
    <rect width="${size}" height="${size}" fill="#fff"/>
    ${rects.join("")}
  </svg>`;
}

function QRCodeSvg({value, size=200}){
  const svg = makeQRSvg(value, size);
  return <div dangerouslySetInnerHTML={{__html:svg}} style={{lineHeight:0}} />;
}
const fmt = n => Math.round(n).toLocaleString("th-TH");
const IN_CATS  = ["ยอดขาย dine-in","ยอดขาย delivery","เงินโอนเข้า","อื่นๆ"];
const OUT_CATS = ["วัตถุดิบ/ผัก","หมู/เนื้อ/ทะเล","ค่าเช่า","ค่าพนักงาน","ค่าไฟ/น้ำ","บรรจุภัณฑ์","อื่นๆ"];
const PAY = ["เงินสด","โอนธนาคาร","QR Code","GrabFood","LINE MAN"];
const stockSt = s => s.qty<=0?"out":s.qty<s.minQty*.5?"critical":s.qty<s.minQty?"low":"ok";
const ST_C = { ok:"#16a34a", low:"#d97706", critical:"#dc2626", out:"#dc2626" };
const ST_L = { ok:"ปกติ", low:"ใกล้หมด", critical:"น้อยมาก", out:"หมด" };
const wac = s => { const h=s.costHistory||[]; const q=h.reduce((a,b)=>a+b.qty,0); return q>0?h.reduce((a,b)=>a+b.total,0)/q:0; };

const INIT_STOCK = [
  {id:1,name:"หมูสามชั้น",unit:"kg",qty:15,minQty:5,dailyUse:3,supplierId:2,costHistory:[]},
  {id:2,name:"กุ้งแวนนาไม",unit:"kg",qty:3,minQty:4,dailyUse:2,supplierId:3,costHistory:[]},
  {id:3,name:"เต้าหู้ขาว",unit:"kg",qty:10,minQty:5,dailyUse:2,supplierId:1,costHistory:[]},
  {id:4,name:"ผักกาดขาว",unit:"kg",qty:2,minQty:3,dailyUse:2,supplierId:1,costHistory:[]},
  {id:5,name:"น้ำซุปมาล่า",unit:"ถุง",qty:20,minQty:8,dailyUse:4,supplierId:4,costHistory:[]},
  {id:6,name:"วุ้นเส้น",unit:"kg",qty:5,minQty:3,dailyUse:1,supplierId:1,costHistory:[]},
];
const INIT_SUPS = [
  {id:1,name:"ตลาดสดนครชัย",type:"ผัก",phone:"081-234-5678",active:true},
  {id:2,name:"ฟาร์มหมูสยาม",type:"หมู",phone:"082-345-6789",active:true},
  {id:3,name:"อาหารทะเลสด",type:"ทะเล",phone:"083-456-7890",active:true},
  {id:4,name:"ซอสมาล่าพรีเมียม",type:"ซอส",phone:"084-567-8901",active:true},
];
// HR Config per staff member - stored in staff object
// {wage:0, wageType:'day'|'month', otRate:1.5, shiftStart:'09:00', shiftEnd:'18:00', bonusPct:0}

const INIT_STAFF = [
  {id:"owner",name:"DR.Fresh (เจ้าของ)",pin:"1234",role:"owner",active:true,perms:{cashflow:true,stock:true,purchase:true,report:true,admin:true,viewPrice:true},hr:{wage:500,wageType:"day",otRate:1.5,shiftStart:"09:00",shiftEnd:"18:00",bonusPct:0}},
  {id:"s1",name:"มิ้ว",pin:"1111",role:"staff",active:true,perms:{cashflow:true,stock:true,purchase:false,report:false,admin:false,viewPrice:false,viewFinance:false}},
  {id:"s2",name:"ปาล์ม",pin:"2222",role:"staff",active:true,perms:{cashflow:true,stock:true,purchase:false,report:false,admin:false,viewPrice:false,viewFinance:false}},
  {id:"s3",name:"เจ",pin:"3333",role:"staff",active:true,perms:{cashflow:false,stock:true,purchase:false,report:false,admin:false,viewPrice:false},hr:{wage:500,wageType:"day",otRate:1.5,shiftStart:"09:00",shiftEnd:"18:00",bonusPct:0}},
  {id:"emergency",name:"Emergency",pin:"0000",role:"owner",active:true,perms:{cashflow:false,stock:false,purchase:false,report:false,admin:true,viewPrice:false}},
];
const INIT_WASTE = [];   // {id,itemId,qty,reason,date,staffId,cost}
const INIT_PROMO = [];   // {id,name,date,amount,note}

const INIT_FIXED = [{name:"ค่าเช่า",amount:4500},{name:"ค่าพนักงาน",amount:35000},{name:"ค่าไฟ",amount:8000},{name:"ค่าน้ำ",amount:1000},{name:"อื่นๆ",amount:1000}];

// ─── UI Components ───
function Card({children,style={},onClick}){return <div style={{...S.card,...style}} onClick={onClick}>{children}</div>;}
function Badge({status}){const c=ST_C[status]||ST_C.ok;return <span style={{background:c+"18",color:c,borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{ST_L[status]||"ปกติ"}</span>;}
function Tabs({tabs,active,onChange}){return <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{tabs.map(([v,l])=><button key={v} onClick={()=>onChange(v)} style={{padding:"8px 14px",borderRadius:8,border:`1px solid ${active===v?T.orange:T.border}`,background:active===v?T.orange:"transparent",color:active===v?"#fff":T.textMd,cursor:"pointer",fontSize:13,fontFamily:F,fontWeight:active===v?700:400}}>{l}</button>)}</div>;}
function Hdr({title,action}){return <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}><div style={{fontSize:20,fontWeight:900}}>{title}</div>{action}</div>;}

// ─── Login ───
function CheckinPage({staff,onCheckin}){
  // Read ?sid= from URL
  const sid = new URLSearchParams(window.location.search).get("sid");
  const s = staff.find(x=>x.id===sid);
  const [done,setDone]=useState(null); // {type,time,name}
  const [loading,setLoading]=useState(false);
  const nowStr=()=>new Date().toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"});
  const todayStr=()=>new Date().toISOString().split("T")[0];

  const stamp=async(type)=>{
    if(!s)return;
    setLoading(true);
    const t=nowStr();
    const d=todayStr();
    if(type==="in"){
      await supa("attendance",{method:"POST",body:JSON.stringify({
        id:Date.now(),staff_id:s.id,date:d,check_in:t,check_out:"",note:""
      })});
    } else {
      // find today's open record
      const recs=await supa(`attendance?staff_id=eq.${s.id}&date=eq.${d}&check_out=eq.`);
      if(recs&&recs.length>0){
        await supa(`attendance?id=eq.${recs[0].id}`,{method:"PATCH",body:JSON.stringify({check_out:t})});
      }
    }
    setDone({type,time:t,name:s.name});
    setLoading(false);
    // callback to update local state too
    onCheckin&&onCheckin(s.id,type,t,d);
  };

  if(!sid||!s) return (
    <div style={{minHeight:"100vh",background:"#f4f4f5",fontFamily:"'Noto Sans Thai',sans-serif",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:32,textAlign:"center",maxWidth:320,width:"100%",boxShadow:"0 4px 12px rgba(0,0,0,.1)"}}>
        <div style={{fontSize:40,marginBottom:12}}>❌</div>
        <div style={{fontWeight:700,fontSize:18,color:"#dc2626"}}>QR Code ไม่ถูกต้อง</div>
        <div style={{color:"#71717a",fontSize:14,marginTop:8}}>กรุณาสแกน QR Code ที่หน้าร้านใหม่</div>
      </div>
    </div>
  );

  if(done) return (
    <div style={{minHeight:"100vh",background:`linear-gradient(135deg,${done.type==="in"?"#f0fdf4":"#fef2f2"},#fff)`,fontFamily:"'Noto Sans Thai',sans-serif",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,textAlign:"center",maxWidth:360,width:"100%",boxShadow:"0 8px 24px rgba(0,0,0,.12)"}}>
        <div style={{fontSize:60,marginBottom:16}}>{done.type==="in"?"✅":"👋"}</div>
        <div style={{fontWeight:900,fontSize:24,color:done.type==="in"?"#16a34a":"#dc2626"}}>
          {done.type==="in"?"เช็คอินสำเร็จ!":"เช็คออกสำเร็จ!"}
        </div>
        <div style={{fontWeight:700,fontSize:20,marginTop:12}}>{done.name}</div>
        <div style={{fontSize:36,fontWeight:900,color:"#f97316",marginTop:8}}>{done.time}</div>
        <div style={{color:"#71717a",fontSize:14,marginTop:16}}>บันทึกเวลาเรียบร้อยแล้ว<br/>สามารถปิดหน้านี้ได้เลย</div>
      </div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#fff7ed,#fff)",fontFamily:"'Noto Sans Thai',sans-serif",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,textAlign:"center",maxWidth:360,width:"100%",boxShadow:"0 8px 24px rgba(0,0,0,.12)"}}>
        <div style={{width:80,height:80,borderRadius:20,background:`linear-gradient(135deg,#f97316,#ea580c)`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:44,margin:"0 auto 16px",boxShadow:"0 6px 20px rgba(249,115,22,.3)"}}>🫕</div>
        <div style={{color:"#f97316",fontWeight:900,fontSize:20,marginBottom:4}}>ไท่กั๋วหม่าล่า</div>
        <div style={{fontWeight:800,fontSize:26,marginBottom:4}}>{s.name}</div>
        <div style={{color:"#71717a",fontSize:15,marginBottom:24}}>
          {new Date().toLocaleDateString("th-TH",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}
        </div>
        <div style={{fontSize:42,fontWeight:900,color:"#18181b",marginBottom:28,letterSpacing:2}}>
          {new Date().toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"})}
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <button onClick={()=>stamp("in")} disabled={loading} style={{background:loading?"#a1a1aa":"#16a34a",color:"#fff",border:"none",borderRadius:14,padding:"18px",fontSize:20,fontWeight:800,cursor:loading?"default":"pointer",fontFamily:"'Noto Sans Thai',sans-serif",boxShadow:"0 4px 12px rgba(22,163,74,.3)"}}>
            {loading?"⏳ กำลังบันทึก...":"📥 เข้างาน (Check-in)"}
          </button>
          <button onClick={()=>stamp("out")} disabled={loading} style={{background:loading?"#a1a1aa":"#dc2626",color:"#fff",border:"none",borderRadius:14,padding:"18px",fontSize:20,fontWeight:800,cursor:loading?"default":"pointer",fontFamily:"'Noto Sans Thai',sans-serif",boxShadow:"0 4px 12px rgba(220,38,38,.3)"}}>
            {loading?"⏳ กำลังบันทึก...":"📤 ออกงาน (Check-out)"}
          </button>
        </div>
        <div style={{color:"#a1a1aa",fontSize:12,marginTop:20}}>เวลาจะถูกบันทึกและส่งไปยังเจ้าของร้านอัตโนมัติ</div>
      </div>
    </div>
  );
}

function LoginPage({staff,onLogin}){
  const [pin,setPin]=useState(""); const [err,setErr]=useState(false);
  const go=()=>{const u=staff.find(s=>s.pin===pin&&s.active);if(u)onLogin(u);else{setErr(true);setTimeout(()=>setErr(false),1500);}};
  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#fff7ed,#fff)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:20,gap:20,fontFamily:F}}>
      <div style={{textAlign:"center"}}>
        <div style={{width:100,height:100,borderRadius:24,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:54,margin:"0 auto 12px",boxShadow:"0 8px 28px rgba(249,115,22,.35)"}}>🫕</div>
        <div style={{color:T.orange,fontWeight:900,fontSize:26}}>ไท่กั๋วหม่าล่า</div>
        <div style={{color:T.textSm,fontSize:14,marginTop:4}}>TAI GUO MALA • ระบบจัดการร้าน</div>
      </div>
      <Card style={{width:"100%",maxWidth:360,boxShadow:"0 4px 12px rgba(0,0,0,.12)"}}>
        <div style={{color:T.textMd,fontSize:15,marginBottom:12,textAlign:"center"}}>กรอก PIN เพื่อเข้าใช้งาน</div>
        <input type="password" maxLength={4} value={pin} onChange={e=>{setPin(e.target.value);setErr(false);}} onKeyDown={e=>e.key==="Enter"&&go()}
          style={{...S.inp,fontSize:28,letterSpacing:10,textAlign:"center",marginBottom:10,border:`2px solid ${err?T.red:T.border}`}} placeholder="••••" autoFocus />
        <button onClick={go} style={{...S.btn(),width:"100%",padding:13,fontSize:17}}>เข้าสู่ระบบ</button>
        <div style={{marginTop:14,textAlign:"center",color:T.textXs,fontSize:13}}>กรอก PIN ที่ได้รับจากเจ้าของร้าน</div>
      </Card>
    </div>
  );
}

// ─── Dashboard ───
function DashboardPage({cf,stock,user,fixedCosts,waste,promos,setPage}){
  const ts=today(), mk=ts.slice(0,7);
  const yd=new Date(); yd.setDate(yd.getDate()-1); const ys=yd.toISOString().split("T")[0];
  const myCF=user.role==="owner"?cf:cf.filter(e=>e.staffId===user.id||e.branch===user.franchiseId);
  const todayIn=myCF.filter(e=>e.date===ts&&e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const todayOut=myCF.filter(e=>e.date===ts&&e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const ydIn=myCF.filter(e=>e.date===ys&&e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mIn=myCF.filter(e=>e.date.startsWith(mk)&&e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mOut=myCF.filter(e=>e.date.startsWith(mk)&&e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const totalFixed=fixedCosts.reduce((a,b)=>a+b.amount,0);
  const netP=mIn-mOut-totalFixed;
  const alerts=stock.filter(s=>["out","critical","low"].includes(stockSt(s)));
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div><div style={{fontSize:22,fontWeight:900}}>🏠 ภาพรวมร้าน</div><div style={{color:T.textSm,fontSize:14}}>{ts}</div></div>
      {(user.role==="owner"||user.perms?.viewFinance)&&(
        <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12}}>
          {[
            ["💰","รายรับวันนี้",todayIn,T.green,ydIn>0?`${todayIn>=ydIn?"▲":"▼"}${Math.abs(((todayIn-ydIn)/ydIn)*100).toFixed(0)}% vs เมื่อวาน`:""],
            ["💸","รายจ่ายวันนี้",todayOut,T.red,`กำไรวันนี้ ฿${fmt(todayIn-todayOut)}`],
            ["📅","รายรับเดือน",mIn,T.orange,`BEP ฿${fmt(totalFixed)}`],
            ["📈","กำไรสุทธิ",netP,netP>=0?T.green:T.red,"หลังต้นทุนคงที่"],
          ].map(([ic,l,v,col,sub])=>(
            <Card key={l} style={{padding:"16px 18px",borderLeft:`4px solid ${col}`}}>
              <div style={{color:T.textSm,fontSize:13,marginBottom:4}}>{ic} {l}</div>
              <div style={{color:col,fontWeight:900,fontSize:24}}>฿{fmt(v)}</div>
              {sub&&<div style={{fontSize:11,color:T.textSm,marginTop:4}}>{sub}</div>}
            </Card>
          ))}
        </div>
      )}
      {user.role!=="owner"&&!user.perms?.viewFinance&&(
        <Card style={{padding:"16px 18px",background:T.bg,borderColor:T.border,textAlign:"center"}}>
          <div style={{fontSize:28,marginBottom:6}}>🔒</div>
          <div style={{color:T.textMd,fontSize:14,fontWeight:600}}>ข้อมูลการเงินถูกซ่อน</div>
          <div style={{color:T.textXs,fontSize:12,marginTop:4}}>เจ้าของร้านจะเปิดสิทธิ์ใน Settings</div>
        </Card>
      )}
      {alerts.length>0&&(
        <Card style={{borderColor:T.red+"44",background:T.redLt}}>
          <div style={{color:T.red,fontWeight:800,fontSize:16,marginBottom:10}}>🚨 สต็อคต้องดูแล ({alerts.length} รายการ)</div>
          {alerts.map(s=>(
            <div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:"1px solid rgba(220,38,38,.1)"}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}><Badge status={stockSt(s)} /><span style={{fontWeight:600,fontSize:15}}>{s.name}</span></div>
              <div style={{textAlign:"right"}}><div style={{color:ST_C[stockSt(s)],fontWeight:700}}>{s.qty} {s.unit}</div><div style={{color:T.textXs,fontSize:11}}>ขั้นต่ำ {s.minQty}</div></div>
            </div>
          ))}
          <button onClick={()=>setPage("stock")} style={{...S.btn(T.red),width:"100%",marginTop:10,padding:10,fontSize:14}}>ไปจัดการสต็อค →</button>
        </Card>
      )}
      {(()=>{
        const mk2=today().slice(0,7);
        const wCost=waste&&waste.filter(w=>w.date.startsWith(mk2)).reduce((a,b)=>a+b.cost,0)||0;
        const highCostItems=stock.filter(s=>{const h=s.costHistory||[];if(h.length<2)return false;const avg=wac(s);const lastP=h[h.length-1]?.unitCost||0;return lastP>avg*1.15;});
        if(wCost===0&&highCostItems.length===0)return null;
        return(
          <Card style={{borderColor:T.yellow+"44",background:T.yellowLt}}>
            <div style={{color:T.yellow,fontWeight:800,fontSize:15,marginBottom:8}}>⚠️ แจ้งเตือนความเสี่ยง</div>
            {wCost>0&&<div style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${T.yellow}22`,fontSize:14}}>
              <span>🗑 Waste เดือนนี้</span>
              <span style={{color:T.red,fontWeight:700}}>฿{fmt(wCost)}</span>
            </div>}
            {highCostItems.map(s=>{const lastP=s.costHistory[s.costHistory.length-1]?.unitCost||0;const avg=wac(s);return(
              <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${T.yellow}22`,fontSize:14}}>
                <span>📈 {s.name} ต้นทุนสูงขึ้น</span>
                <span style={{color:T.red,fontWeight:700}}>+{(((lastP-avg)/avg)*100).toFixed(0)}%</span>
              </div>
            );})}
            <button onClick={()=>setPage("report")} style={{...S.btn(T.yellow),width:"100%",marginTop:8,padding:8,fontSize:13}}>ดูรายงานเพิ่มเติม →</button>
          </Card>
        );
      })()}
      <Card>
        <div style={{fontWeight:700,fontSize:15,marginBottom:10}}>📦 สถานะสต็อค</div>
        {stock.map(s=>{const st=stockSt(s);const pct=s.minQty>0?Math.min((s.qty/s.minQty)*100,200):100;return(
          <div key={s.id} style={{marginBottom:12}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:14,fontWeight:600}}>{s.name}</span><Badge status={st} /></div>
              <span style={{color:ST_C[st],fontWeight:700,fontSize:14}}>{s.qty} {s.unit}</span>
            </div>
            <div style={{background:T.bg,borderRadius:4,height:6}}><div style={{background:pct<50?T.red:pct<100?T.yellow:T.green,width:`${Math.min(pct,100)}%`,height:"100%",borderRadius:4}}/></div>
          </div>
        );})}
      </Card>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {[["💵","กรอก Cash Flow","cashflow"],["📦","บันทึกสต็อค",user.role==="owner"?"stock":"staffstock"]].map(([ic,l,pg])=>(
          <Card key={pg} style={{padding:"16px 18px",cursor:"pointer",textAlign:"center"}} onClick={()=>setPage(pg)}>
            <div style={{fontSize:28,marginBottom:6}}>{ic}</div><div style={{fontWeight:700,fontSize:14}}>{l}</div>
          </Card>
        ))}
      </div>
    </div>
  );
}

// ─── Cash Flow ───
function CashflowPage({cf,setCF,user,dbReady}){
  const [showForm,setShowForm]=useState(false);
  const [viewTab,setViewTab]=useState("daily");
  const [filterMode,setFilterMode]=useState("month");
  const [filterMonth,setFilterMonth]=useState(today().slice(0,7));
  const [filterDay,setFilterDay]=useState(today());
  const [openCash,setOpenCash]=useState(0); const [openBank,setOpenBank]=useState(0);
  const [showSetup,setShowSetup]=useState(false);
  const [form,setForm]=useState({date:today(),flow:"in",cat:IN_CATS[0],itemName:"",amount:"",method:"เงินสด",note:""});
  const isCash=e=>e.method==="เงินสด";
  const myCF=user.role==="owner"?cf:cf.filter(e=>e.staffId===user.id||e.branch===user.franchiseId);
  const months=[...new Set(myCF.map(e=>e.date.slice(0,7)))].sort((a,b)=>b.localeCompare(a));
  const filtered=filterMode==="day"?myCF.filter(e=>e.date===filterDay):myCF.filter(e=>e.date.startsWith(filterMonth));
  const sum=(arr,flow,fn)=>arr.filter(e=>e.flow===flow&&(fn?fn(e):true)).reduce((a,b)=>a+b.amount,0);
  const totalIn=sum(filtered,"in"); const totalOut=sum(filtered,"out");
  const cashIn=sum(filtered,"in",isCash); const cashOut=sum(filtered,"out",isCash);
  const bankIn=sum(filtered,"in",e=>!isCash(e)); const bankOut=sum(filtered,"out",e=>!isCash(e));
  const cashBal=openCash+cashIn-cashOut; const bankBal=openBank+bankIn-bankOut;
  const allDates=[...new Set(myCF.map(e=>e.date))].sort();
  let rC=openCash, rB=openBank;
  const dailyRows=allDates.map(d=>{
    const ent=myCF.filter(e=>e.date===d);
    const dCI=sum(ent,"in",isCash),dCO=sum(ent,"out",isCash),dBI=sum(ent,"in",e=>!isCash(e)),dBO=sum(ent,"out",e=>!isCash(e));
    rC+=dCI-dCO; rB+=dBI-dBO;
    return {d,ent,dCI,dCO,dBI,dBO,balCash:rC,balBank:rB};
  });
  const dispDays=[...(filterMode==="day"?dailyRows.filter(r=>r.d===filterDay):dailyRows.filter(r=>r.d.startsWith(filterMonth)))].reverse();
  const addEntry=async()=>{
    if(!form.amount||!+form.amount)return;
    const entry={...form,id:Date.now(),amount:+form.amount,branch:user.franchiseId||"main",staffId:user.id};
    setCF(p=>[entry,...p]); setShowForm(false);
    setForm({date:today(),flow:"in",cat:IN_CATS[0],itemName:"",amount:"",method:"เงินสด",note:""});
    if(dbReady)db.addCF({id:entry.id,date:entry.date,flow:entry.flow,cat:entry.cat,item_name:entry.itemName,amount:entry.amount,method:entry.method,note:entry.note,branch:entry.branch,staff_id:entry.staffId}).catch(()=>{});
  };
  const delEntry=id=>{setCF(p=>p.filter(e=>e.id!==id));if(dbReady)db.delCF(id).catch(()=>{});};
  const clearAll=()=>{if(!window.confirm("ล้าง Cash Flow ทั้งหมด?"))return;setCF([]);if(dbReady)db.clearCF().catch(()=>{});};
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="💵 Cash Flow" action={<div style={{display:"flex",gap:8}}>
        <button onClick={()=>setShowSetup(!showSetup)} style={{...S.ghost,fontSize:13,padding:"7px 12px"}}>⚙️ ยอดเปิด</button>
        <button onClick={clearAll} style={{...S.btn(T.red),fontSize:13,padding:"7px 12px"}}>🗑 ล้าง</button>
        <button onClick={()=>setShowForm(!showForm)} style={S.btn()}>+ กรอก</button>
      </div>} />
      {showSetup&&<Card style={{borderColor:T.borderOr,background:T.orangeLt}}>
        <div style={{color:T.orange,fontWeight:800,fontSize:16,marginBottom:8}}>⚙️ ตั้งยอดเปิดบัญชี</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>💵 เงินสด (฿)</div><input type="number" value={openCash} onChange={e=>setOpenCash(+e.target.value||0)} style={S.inp} /></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>🏦 ธนาคาร (฿)</div><input type="number" value={openBank} onChange={e=>setOpenBank(+e.target.value||0)} style={S.inp} /></div>
        </div>
        <button onClick={()=>setShowSetup(false)} style={{...S.btn(),width:"100%",marginTop:10,padding:10}}>✅ บันทึก</button>
      </Card>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Card style={{padding:"16px 18px",borderLeft:"4px solid #78716c",background:"#fafaf9"}}>
          <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}><span style={{fontSize:20}}>💵</span><span style={{color:"#78716c",fontWeight:700,fontSize:14}}>เงินสด</span></div>
          <div style={{color:cashBal>=0?T.green:T.red,fontWeight:900,fontSize:24}}>฿{fmt(cashBal)}</div>
          <div style={{fontSize:12,color:T.textSm,marginTop:4}}><div>เปิด ฿{fmt(openCash)}</div><div style={{color:T.green}}>+฿{fmt(cashIn)}</div><div style={{color:T.red}}>-฿{fmt(cashOut)}</div></div>
        </Card>
        <Card style={{padding:"16px 18px",borderLeft:`4px solid ${T.blue}`,background:T.blueLt}}>
          <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}><span style={{fontSize:20}}>🏦</span><span style={{color:T.blue,fontWeight:700,fontSize:14}}>ธนาคาร</span></div>
          <div style={{color:bankBal>=0?T.green:T.red,fontWeight:900,fontSize:24}}>฿{fmt(bankBal)}</div>
          <div style={{fontSize:12,color:T.textSm,marginTop:4}}><div>เปิด ฿{fmt(openBank)}</div><div style={{color:T.green}}>+฿{fmt(bankIn)}</div><div style={{color:T.red}}>-฿{fmt(bankOut)}</div></div>
        </Card>
      </div>
      <Card style={{padding:"12px 18px",background:T.orangeLt,borderColor:T.borderOr}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{color:T.textSm,fontSize:13}}>รวมทั้งหมด</div><div style={{color:T.orange,fontWeight:900,fontSize:22}}>฿{fmt(cashBal+bankBal)}</div></div>
          <div style={{textAlign:"right",fontSize:13}}><div style={{color:T.green}}>เข้า ฿{fmt(totalIn)}</div><div style={{color:T.red}}>ออก ฿{fmt(totalOut)}</div></div>
        </div>
      </Card>
      <Card style={{padding:"14px 16px"}}>
        <div style={{display:"flex",gap:6,marginBottom:10}}>
          {[["month","📅 เดือน"],["day","📆 วัน"]].map(([v,l])=>(
            <button key={v} onClick={()=>setFilterMode(v)} style={{flex:1,padding:"8px 4px",borderRadius:8,cursor:"pointer",fontSize:13,fontFamily:F,background:filterMode===v?T.orange:"transparent",border:`1px solid ${filterMode===v?T.orange:T.border}`,color:filterMode===v?"#fff":T.textMd,fontWeight:filterMode===v?700:400}}>{l}</button>
          ))}
        </div>
        {filterMode==="month"&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{(months.length?months:[today().slice(0,7)]).map(m=><button key={m} onClick={()=>setFilterMonth(m)} style={{background:filterMonth===m?T.orange:"transparent",border:`1px solid ${filterMonth===m?T.orange:T.border}`,borderRadius:8,padding:"6px 12px",color:filterMonth===m?"#fff":T.textMd,cursor:"pointer",fontSize:13,fontFamily:F}}>{m}</button>)}</div>}
        {filterMode==="day"&&<input type="date" value={filterDay} max={today()} onChange={e=>setFilterDay(e.target.value)} style={{...S.inp,fontSize:16}} />}
      </Card>
      <Tabs tabs={[["daily","📆 รายวัน+ยอดยก"],["list","📋 รายการ"]]} active={viewTab} onChange={setViewTab} />
      {showForm&&<Card style={{borderColor:T.borderOr}}>
        <div style={{color:T.orange,fontWeight:800,fontSize:17,marginBottom:12}}>📝 บันทึกรายการ</div>
        <div style={{display:"flex",gap:8,marginBottom:12}}>
          {[["in","💰 รายรับ",T.green],["out","💸 รายจ่าย",T.red]].map(([v,l,c])=>(
            <button key={v} onClick={()=>setForm(p=>({...p,flow:v,cat:v==="in"?IN_CATS[0]:OUT_CATS[0],itemName:""}))}
              style={{flex:1,padding:12,background:form.flow===v?c+"18":"transparent",border:`1.5px solid ${form.flow===v?c:T.border}`,borderRadius:10,color:form.flow===v?c:T.textMd,fontWeight:700,cursor:"pointer",fontSize:16,fontFamily:F}}>{l}</button>
          ))}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>วันที่</div><input type="date" value={form.date} max={today()} onChange={e=>setForm(p=>({...p,date:e.target.value}))} style={S.inp} /></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>จำนวน (฿)</div><input type="number" value={form.amount} onChange={e=>setForm(p=>({...p,amount:e.target.value}))} style={S.inp} placeholder="0" /></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>หมวดหมู่</div><select value={form.cat} onChange={e=>setForm(p=>({...p,cat:e.target.value}))} style={S.inp}>{(form.flow==="in"?IN_CATS:OUT_CATS).map(c=><option key={c}>{c}</option>)}</select></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ช่องทาง</div><select value={form.method} onChange={e=>setForm(p=>({...p,method:e.target.value}))} style={S.inp}>{PAY.map(m=><option key={m}>{m}</option>)}</select></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ชื่อรายการ</div><input value={form.itemName} onChange={e=>setForm(p=>({...p,itemName:e.target.value}))} style={S.inp} placeholder="เช่น หมูสามชั้น" /></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>หมายเหตุ</div><input value={form.note} onChange={e=>setForm(p=>({...p,note:e.target.value}))} style={S.inp} /></div>
        </div>
        <div style={{display:"flex",gap:8,marginTop:12}}>
          <button onClick={addEntry} style={{...S.btn(form.flow==="in"?T.green:T.red),flex:1,padding:12,fontSize:16}}>✅ บันทึก</button>
          <button onClick={()=>setShowForm(false)} style={{...S.ghost,padding:"12px 14px"}}>ยกเลิก</button>
        </div>
      </Card>}
      {viewTab==="daily"&&<div style={{display:"flex",flexDirection:"column",gap:10}}>
        {dispDays.length===0&&<div style={{color:T.textSm,textAlign:"center",padding:32}}>ไม่มีรายการ</div>}
        {dispDays.map((row,i)=>{const prev=dispDays[i+1];return(
          <Card key={row.d} style={{padding:"14px 18px"}}>
            <div style={{display:"flex",justifyContent:"space-between",paddingBottom:10,marginBottom:10,borderBottom:`1px solid ${T.bg}`}}>
              <span style={{fontWeight:800,fontSize:15}}>{row.d}</span>
              <div style={{fontSize:13}}><span style={{color:T.green}}>+฿{fmt(row.dCI+row.dBI)}</span> <span style={{color:T.red}}>-฿{fmt(row.dCO+row.dBO)}</span></div>
            </div>
            {prev&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8,padding:"6px 10px",background:T.bg,borderRadius:8,fontSize:12}}>
              <div style={{color:T.textSm}}>💵 ยกมา: <b>฿{fmt(prev.balCash)}</b></div>
              <div style={{color:T.textSm}}>🏦 ยกมา: <b>฿{fmt(prev.balBank)}</b></div>
            </div>}
            {row.ent.map(e=>(
              <div key={e.id} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:`1px solid ${T.bg}`}}>
                <div style={{width:26,height:26,borderRadius:7,flexShrink:0,background:e.flow==="in"?T.greenLt:T.redLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:e.flow==="in"?T.green:T.red}}>{e.flow==="in"?"↓":"↑"}</div>
                <div style={{flex:1}}><div style={{fontSize:14,fontWeight:600}}>{e.itemName||e.cat}</div><div style={{color:T.textSm,fontSize:11}}>{isCash(e)?"💵":"🏦"} {e.method}{e.note?` • ${e.note}`:""}</div></div>
                <span style={{color:e.flow==="in"?T.green:T.red,fontWeight:700,fontSize:14}}>{e.flow==="in"?"+":"-"}฿{fmt(e.amount)}</span>
                <button onClick={()=>delEntry(e.id)} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:18,padding:"0 4px"}}>×</button>
              </div>
            ))}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:10,padding:"8px 10px",background:T.orangeLt,borderRadius:8,fontSize:13}}>
              <div><div style={{color:T.textXs,fontSize:11}}>💵 สิ้นวัน</div><div style={{color:row.balCash>=0?T.green:T.red,fontWeight:800,fontSize:16}}>฿{fmt(row.balCash)}</div></div>
              <div><div style={{color:T.textXs,fontSize:11}}>🏦 สิ้นวัน</div><div style={{color:row.balBank>=0?T.green:T.red,fontWeight:800,fontSize:16}}>฿{fmt(row.balBank)}</div></div>
            </div>
          </Card>
        );})}
      </div>}
      {viewTab==="list"&&<div style={{display:"flex",flexDirection:"column",gap:8}}>
        {filtered.length===0&&<div style={{color:T.textSm,textAlign:"center",padding:32}}>ไม่มีรายการ</div>}
        {[...filtered].sort((a,b)=>b.date.localeCompare(a.date)||b.id-a.id).map(e=>(
          <Card key={e.id} style={{padding:"12px 16px"}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:32,height:32,borderRadius:9,flexShrink:0,background:e.flow==="in"?T.greenLt:T.redLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:e.flow==="in"?T.green:T.red}}>{e.flow==="in"?"↓":"↑"}</div>
              <div style={{flex:1}}><div style={{fontSize:15,fontWeight:600}}>{e.itemName||e.cat}</div><div style={{color:T.textSm,fontSize:12}}>{e.date} • {isCash(e)?"💵":"🏦"} {e.method}</div></div>
              <span style={{color:e.flow==="in"?T.green:T.red,fontWeight:700,fontSize:16}}>{e.flow==="in"?"+":"-"}฿{fmt(e.amount)}</span>
              <button onClick={()=>delEntry(e.id)} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:18,padding:"0 4px"}}>×</button>
            </div>
          </Card>
        ))}
      </div>}
    </div>
  );
}

// ─── Stock (Owner) ───
function StockPage({stock,setStock,movements,setMovements,user,suppliers}){
  const [tab,setTab]=useState("list");
  const [selId,setSelId]=useState(""); const [mvType,setMvType]=useState("in");
  const [qty,setQty]=useState(""); const [cost,setCost]=useState(""); const [note,setNote]=useState("");
  const [msg,setMsg]=useState({text:"",ok:false});
  const [editId,setEditId]=useState(null); const [editData,setEditData]=useState({});
  const [showAdd,setShowAdd]=useState(false);
  const [newItem,setNewItem]=useState({name:"",unit:"kg",qty:0,minQty:3,dailyUse:1,supplierId:1});
  const canPrice=user.perms?.viewPrice===true;
  const selItem=selId?stock.find(s=>String(s.id)===String(selId)):null;
  const showMsg=(t,ok)=>{setMsg({text:t,ok});setTimeout(()=>setMsg({text:"",ok:false}),3000);};
  const save=()=>{
    const q=parseFloat(qty);
    if(!selId||!q||q<=0){showMsg("กรุณาเลือกรายการและกรอกจำนวน",false);return;}
    const item=stock.find(s=>String(s.id)===String(selId));
    if(!item)return;
    if(mvType==="out"&&q>item.qty){showMsg(`มีแค่ ${item.qty} ${item.unit}`,false);return;}
    const uc=parseFloat(cost)||0;
    const nq=mvType==="in"?item.qty+q:item.qty-q;
    const nh=mvType==="in"&&uc>0?[...(item.costHistory||[]),{date:today(),unitCost:uc,qty:q,total:uc*q}]:item.costHistory;
    setStock(stock.map(s=>String(s.id)===String(selId)?{...s,qty:nq,costHistory:nh}:s));
    setMovements(p=>[...p,{id:Date.now(),itemId:item.id,type:mvType,qty:q,unitCost:uc,date:today(),staffId:user.id,note:note||(mvType==="in"?"รับเข้า":"จ่ายออก"),branch:"main"}]);
    showMsg(`✅ ${mvType==="in"?"รับเข้า":"จ่ายออก"} ${item.name} ${q} ${item.unit}`,true);
    setQty(""); setCost(""); setNote("");
  };
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="📦 สต็อควัตถุดิบ" action={<button onClick={()=>setShowAdd(!showAdd)} style={S.btn()}>+ เพิ่ม</button>} />
      {showAdd&&<Card style={{borderColor:T.borderOr}}>
        <div style={{color:T.orange,fontWeight:800,fontSize:16,marginBottom:12}}>➕ เพิ่มวัตถุดิบ</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          {[["ชื่อ","name","text"],["หน่วย","unit","text"],["จำนวน","qty","number"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"]].map(([l,k,t])=>(
            <div key={k}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>{l}</div><input type={t} value={newItem[k]} onChange={e=>setNewItem(p=>({...p,[k]:e.target.value}))} style={S.inp} /></div>
          ))}
        </div>
        <div style={{display:"flex",gap:8,marginTop:12}}>
          <button onClick={()=>{setStock([...stock,{...newItem,id:Date.now(),qty:+newItem.qty,minQty:+newItem.minQty,dailyUse:+newItem.dailyUse,costHistory:[]}]);setShowAdd(false);}} style={{...S.btn(),flex:1}}>บันทึก</button>
          <button onClick={()=>setShowAdd(false)} style={S.ghost}>ยกเลิก</button>
        </div>
      </Card>}
      <Tabs tabs={[["list","📋 รายการ"],["move","📥📤 รับ/จ่าย"],["history","📊 ประวัติ"]]} active={tab} onChange={t=>{setTab(t);setSelId("");setQty("");setMsg({text:"",ok:false});}} />
      {tab==="list"&&stock.map(item=>{
        const st=stockSt(item); const sup=suppliers?.find(x=>x.id===item.supplierId);
        return(<Card key={item.id} style={{borderColor:st!=="ok"?ST_C[st]+"44":T.border}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div style={{flex:1}}>
              <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}><span style={{fontWeight:700,fontSize:17}}>{item.name}</span><Badge status={st} /></div>
              <div style={{color:T.textSm,fontSize:13,marginTop:4}}>ซัพฯ: {sup?.name||"-"} • ใช้/วัน {item.dailyUse} {item.unit}{canPrice&&wac(item)>0?` • ฿${wac(item).toFixed(2)}/${item.unit}`:""}</div>
              <div style={{color:T.textXs,fontSize:12,marginTop:2}}>ขั้นต่ำ {item.minQty} {item.unit} • เหลือ {item.dailyUse>0?(item.qty/item.dailyUse).toFixed(1):"∞"} วัน</div>
            </div>
            <div style={{textAlign:"right",marginLeft:12}}><div style={{color:ST_C[st],fontWeight:900,fontSize:28}}>{item.qty}</div><div style={{color:T.textSm,fontSize:14}}>{item.unit}</div></div>
          </div>
          {editId===item.id?(
            <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${T.bg}`}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                {[["ชื่อ","name","text"],["หน่วย","unit","text"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"]].map(([l,k,t])=>(
                  <div key={k}><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>{l}</div><input type={t} value={editData[k]??item[k]} onChange={e=>setEditData(p=>({...p,[k]:e.target.value}))} style={S.inp} /></div>
                ))}
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>{setStock(p=>p.map(s=>s.id===item.id?{...s,...editData,minQty:+editData.minQty||s.minQty,dailyUse:+editData.dailyUse||s.dailyUse}:s));setEditId(null);setEditData({});}} style={{...S.btn(),flex:1}}>บันทึก</button>
                <button onClick={()=>{setEditId(null);setEditData({});}} style={S.ghost}>ยกเลิก</button>
              </div>
            </div>
          ):(
            <div style={{borderTop:`1px solid ${T.bg}`,marginTop:10,paddingTop:8,display:"flex",justifyContent:"space-between"}}>
              <button onClick={()=>{setEditId(item.id);setEditData({name:item.name,unit:item.unit,minQty:item.minQty,dailyUse:item.dailyUse});}} style={{background:"none",border:"none",color:T.orange,cursor:"pointer",fontSize:12,fontWeight:600}}>✏️ แก้ไข</button>
              <button onClick={()=>{if(window.confirm(`ลบ "${item.name}"?`))setStock(p=>p.filter(s=>s.id!==item.id));}} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:12}}>🗑 ลบ</button>
            </div>
          )}
        </Card>);
      })}
      {tab==="move"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        <div style={{display:"flex",gap:8}}>
          {[["in","📥 รับเข้า",T.green],["out","📤 จ่ายออก",T.red]].map(([v,l,c])=>(
            <button key={v} onClick={()=>setMvType(v)} style={{flex:1,padding:12,fontWeight:800,fontSize:16,cursor:"pointer",fontFamily:F,borderRadius:10,background:mvType===v?c+"18":"#fff",border:`2px solid ${mvType===v?c:T.border}`,color:mvType===v?c:T.textMd}}>{l}</button>
          ))}
        </div>
        <div><div style={{color:T.textSm,fontSize:13,marginBottom:6,fontWeight:600}}>เลือกรายการ</div>
          <select value={selId} onChange={e=>{setSelId(e.target.value);setQty("");setCost("");}} style={{...S.inp,fontSize:16,height:48}}>
            <option value="">— กรุณาเลือก —</option>
            {[...stock].sort((a,b)=>({out:0,critical:1,low:2,ok:3}[stockSt(a)]-{out:0,critical:1,low:2,ok:3}[stockSt(b)])).map(s=>{const st=stockSt(s);return <option key={s.id} value={String(s.id)}>{st==="out"?"🔴":st==="critical"?"🟠":st==="low"?"🟡":"🟢"} {s.name} (เหลือ {s.qty} {s.unit})</option>;})}
          </select>
        </div>
        {selItem&&<div style={{background:T.bg,borderRadius:10,padding:"10px 14px",display:"flex",justifyContent:"space-between",fontSize:14}}><span>คงเหลือ</span><span style={{color:ST_C[stockSt(selItem)],fontWeight:800,fontSize:18}}>{selItem.qty} {selItem.unit}</span></div>}
        <div style={{display:"grid",gridTemplateColumns:mvType==="in"&&canPrice?"1fr 1fr":"1fr",gap:10}}>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>จำนวน ({selItem?.unit||"หน่วย"})</div><input type="number" value={qty} onChange={e=>{setQty(e.target.value);setMsg({text:"",ok:false});}} style={{...S.inp,fontSize:22,fontWeight:800,textAlign:"center"}} placeholder="0" /></div>
          {mvType==="in"&&canPrice&&<div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ราคา/หน่วย (฿)</div><input type="number" value={cost} onChange={e=>setCost(e.target.value)} style={{...S.inp,fontSize:18,borderColor:T.orange}} placeholder="0" /></div>}
        </div>
        <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>หมายเหตุ</div><input value={note} onChange={e=>setNote(e.target.value)} style={S.inp} /></div>
        {selItem&&qty&&+qty>0&&<div style={{background:mvType==="in"?T.greenLt:T.redLt,borderRadius:8,padding:"10px 14px",fontSize:14}}>หลังบันทึก: <b style={{color:mvType==="in"?T.green:T.red,fontSize:18}}>{mvType==="in"?selItem.qty+(+qty):Math.max(0,selItem.qty-(+qty))} {selItem.unit}</b></div>}
        {msg.text&&<div style={{background:msg.ok?T.greenLt:T.yellowLt,borderRadius:8,padding:"10px 14px",fontSize:14,fontWeight:600,color:msg.ok?T.green:T.yellow}}>{msg.text}</div>}
        <button onClick={save} style={{...S.btn(mvType==="in"?T.green:T.red),width:"100%",padding:14,fontSize:17}}>✅ บันทึก {mvType==="in"?"รับเข้า":"จ่ายออก"}</button>
      </div>}
      {tab==="history"&&<Card>
        <div style={{fontWeight:800,fontSize:16,marginBottom:12}}>ประวัติ</div>
        {[...movements].sort((a,b)=>b.date.localeCompare(a.date)).slice(0,30).map(m=>{const item=stock.find(s=>s.id==m.itemId);const col=m.type==="in"?T.green:T.red;return(
          <div key={m.id} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:`1px solid ${T.bg}`}}>
            <div style={{width:32,height:32,borderRadius:8,background:col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,flexShrink:0}}>{m.type==="in"?"📥":"📤"}</div>
            <div style={{flex:1}}><div style={{fontWeight:600,fontSize:14}}>{item?.name||"?"}</div><div style={{color:T.textSm,fontSize:12}}>{m.date} • {m.note}</div></div>
            <span style={{color:col,fontWeight:800,fontSize:15}}>{m.type==="in"?"+":"-"}{m.qty} {item?.unit}</span>
            <button onClick={()=>setMovements(p=>p.filter(x=>x.id!==m.id))} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:18,padding:"0 4px"}}>×</button>
          </div>
        );})}
        {movements.length===0&&<div style={{color:T.textSm,textAlign:"center",padding:24}}>ยังไม่มีประวัติ</div>}
      </Card>}
    </div>
  );
}

// ─── Staff Stock ───
function StaffStockPage({stock,setStock,movements,setMovements,user}){
  const [tab,setTab]=useState("quick");
  const [selId,setSelId]=useState(""); const [mvType,setMvType]=useState("in");
  const [qty,setQty]=useState(""); const [note,setNote]=useState("");
  const [msg,setMsg]=useState({text:"",ok:false});
  const [checked,setChecked]=useState({}); const [submitted,setSubmitted]=useState(false);
  const selItem=selId?stock.find(s=>String(s.id)===String(selId)):null;
  const showMsg=(t,ok)=>{setMsg({text:t,ok});setTimeout(()=>setMsg({text:"",ok:false}),3000);};
  const save=()=>{
    const q=parseFloat(qty);
    if(!selId||!q||q<=0){showMsg("กรุณาเลือกรายการและกรอกจำนวน",false);return;}
    const item=stock.find(s=>String(s.id)===String(selId));
    if(!item){showMsg("ไม่พบรายการ",false);return;}
    if(mvType==="out"&&q>item.qty){showMsg(`มีแค่ ${item.qty} ${item.unit}`,false);return;}
    setStock(stock.map(s=>String(s.id)===String(selId)?{...s,qty:mvType==="in"?s.qty+q:s.qty-q}:s));
    setMovements(p=>[...p,{id:Date.now(),itemId:item.id,type:mvType,qty:q,unitCost:0,date:today(),staffId:user.id,note:note||(mvType==="in"?"รับเข้า":"จ่ายออก"),branch:"main"}]);
    showMsg(`✅ ${mvType==="in"?"รับเข้า":"จ่ายออก"} ${item.name} ${q} ${item.unit}`,true);
    setQty(""); setNote("");
  };
  const saveChecklist=()=>{
    const entries=Object.entries(checked).filter(([,v])=>v!==""&&!isNaN(+v));
    if(!entries.length)return;
    setStock(stock.map(s=>{const v=checked[String(s.id)];return(!v||isNaN(+v))?s:{...s,qty:+v};}));
    entries.forEach(([id,v])=>{const item=stock.find(s=>String(s.id)===id);if(!item)return;setMovements(p=>[...p,{id:Date.now()+Math.random(),itemId:item.id,type:"check",qty:+v,unitCost:0,date:today(),staffId:user.id,note:"นับสต็อครายวัน",branch:"main"}]);});
    setChecked({}); setSubmitted(true); setTimeout(()=>setSubmitted(false),3000);
  };
  return (
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div><div style={{fontSize:22,fontWeight:900}}>📦 บันทึกสต็อค</div><div style={{color:T.textSm,fontSize:14}}>สวัสดี {user.name} • {today()}</div></div>
      {stock.filter(s=>["critical","out"].includes(stockSt(s))).length>0&&(
        <Card style={{background:T.redLt,borderColor:T.red+"44"}}>
          <div style={{color:T.red,fontWeight:800,fontSize:15,marginBottom:6}}>🚨 แจ้งเจ้าของด่วน!</div>
          {stock.filter(s=>["critical","out"].includes(stockSt(s))).map(s=>(
            <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",fontSize:14}}><span style={{fontWeight:600}}>{s.name}</span><span style={{color:T.red,fontWeight:700}}>เหลือ {s.qty} {s.unit}</span></div>
          ))}
        </Card>
      )}
      <Tabs tabs={[["quick","⚡ รับเข้า/จ่ายออก"],["checklist","✅ นับรายวัน"],["history","📋 ประวัติ"]]} active={tab} onChange={t=>{setTab(t);setSelId("");setQty("");setMsg({text:"",ok:false});}} />
      {tab==="quick"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        <div style={{display:"flex",gap:8}}>
          {[["in","📥 รับเข้า",T.green],["out","📤 จ่ายออก",T.red]].map(([v,l,c])=>(
            <button key={v} onClick={()=>setMvType(v)} style={{flex:1,padding:12,fontWeight:800,fontSize:16,cursor:"pointer",fontFamily:F,borderRadius:10,background:mvType===v?c+"18":"#fff",border:`2px solid ${mvType===v?c:T.border}`,color:mvType===v?c:T.textMd}}>{l}</button>
          ))}
        </div>
        <div><div style={{color:T.textSm,fontSize:13,marginBottom:6,fontWeight:600}}>เลือกรายการ</div>
          <select value={selId} onChange={e=>{setSelId(e.target.value);setQty("");}} style={{...S.inp,fontSize:16,height:48}}>
            <option value="">— กรุณาเลือกรายการ —</option>
            {[...stock].sort((a,b)=>({out:0,critical:1,low:2,ok:3}[stockSt(a)]-{out:0,critical:1,low:2,ok:3}[stockSt(b)])).map(s=>{const st=stockSt(s);return <option key={s.id} value={String(s.id)}>{st==="out"?"🔴":st==="critical"?"🟠":st==="low"?"🟡":"🟢"} {s.name} (เหลือ {s.qty} {s.unit})</option>;})}
          </select>
        </div>
        {selItem&&<div style={{background:T.bg,borderRadius:10,padding:"10px 14px",display:"flex",justifyContent:"space-between",fontSize:14}}><span>คงเหลือ</span><span style={{color:ST_C[stockSt(selItem)],fontWeight:800,fontSize:18}}>{selItem.qty} {selItem.unit}</span></div>}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>จำนวน ({selItem?.unit||"หน่วย"})</div><input type="number" inputMode="numeric" value={qty} onChange={e=>{setQty(e.target.value);setMsg({text:"",ok:false});}} style={{...S.inp,fontSize:22,fontWeight:800,textAlign:"center"}} placeholder="0" /></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>หมายเหตุ</div><input value={note} onChange={e=>setNote(e.target.value)} style={S.inp} /></div>
        </div>
        {selItem&&qty&&+qty>0&&<div style={{background:mvType==="in"?T.greenLt:T.redLt,borderRadius:8,padding:"10px 14px",fontSize:14}}>หลังบันทึก: <b style={{color:mvType==="in"?T.green:T.red,fontSize:18}}>{mvType==="in"?selItem.qty+(+qty):Math.max(0,selItem.qty-(+qty))} {selItem.unit}</b></div>}
        {msg.text&&<div style={{background:msg.ok?T.greenLt:T.yellowLt,borderRadius:8,padding:"10px 14px",fontSize:14,fontWeight:600,color:msg.ok?T.green:T.yellow}}>{msg.text}</div>}
        <button onClick={save} style={{...S.btn(mvType==="in"?T.green:T.red),width:"100%",padding:14,fontSize:17}}>✅ บันทึก {mvType==="in"?"รับเข้า":"จ่ายออก"}</button>
      </div>}
      {tab==="checklist"&&<div style={{display:"flex",flexDirection:"column",gap:10}}>
        {submitted&&<div style={{background:T.greenLt,borderRadius:10,padding:"12px 16px",color:T.green,fontWeight:800}}>✅ บันทึกสำเร็จ!</div>}
        {stock.map(item=>{const st=stockSt(item);const val=checked[String(item.id)]??"";const diff=val!==""&&!isNaN(+val)?+val-item.qty:null;return(
          <Card key={item.id} style={{padding:"12px 16px",borderColor:st!=="ok"?ST_C[st]+"44":T.border}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{display:"flex",alignItems:"center",gap:7}}><span style={{fontWeight:700,fontSize:15}}>{item.name}</span><Badge status={st} /></div>
                <div style={{color:T.textSm,fontSize:12}}>ระบบ: {item.qty} {item.unit}</div>
                {diff!==null&&<div style={{fontSize:12,fontWeight:600,marginTop:2,color:diff<0?T.yellow:diff>0?T.green:T.textSm}}>{diff<0?`⚠️ ขาด ${Math.abs(diff)}`:diff>0?`+เกิน ${diff}`:"✓ ตรง"}</div>}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                <input type="number" inputMode="numeric" value={val} onChange={e=>setChecked(p=>({...p,[String(item.id)]:e.target.value}))} style={{...S.inp,width:80,fontSize:18,fontWeight:700,textAlign:"center"}} placeholder="นับได้" />
                <span style={{color:T.textSm,fontSize:12}}>{item.unit}</span>
              </div>
            </div>
          </Card>
        );})}
        <button onClick={saveChecklist} style={{...S.btn(T.green),width:"100%",padding:13,fontSize:16}}>✅ ส่งรายงาน ({Object.entries(checked).filter(([,v])=>v!=="").length} รายการ)</button>
      </div>}
      {tab==="history"&&<Card>
        <div style={{fontWeight:800,fontSize:16,marginBottom:12}}>ประวัติของฉัน</div>
        {[...movements].filter(m=>m.staffId===user.id).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,20).map(m=>{const item=stock.find(s=>s.id==m.itemId);const col=m.type==="in"?T.green:m.type==="check"?T.blue:T.red;return(
          <div key={m.id} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:`1px solid ${T.bg}`}}>
            <div style={{width:32,height:32,borderRadius:8,background:col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15}}>{m.type==="in"?"📥":m.type==="check"?"✅":"📤"}</div>
            <div style={{flex:1}}><div style={{fontWeight:600,fontSize:14}}>{item?.name||"?"}</div><div style={{color:T.textSm,fontSize:12}}>{m.date} • {m.note}</div></div>
            <span style={{color:col,fontWeight:800,fontSize:15}}>{m.type==="check"?"":m.type==="in"?"+":"-"}{m.qty} {item?.unit}</span>
          </div>
        );})}
        {movements.filter(m=>m.staffId===user.id).length===0&&<div style={{color:T.textSm,textAlign:"center",padding:24}}>ยังไม่มีประวัติ</div>}
      </Card>}
    </div>
  );
}

// ─── Purchase ───
function PurchasePage({stock,suppliers,lineToken}){
  const [selected,setSelected]=useState({}); const [orderQty,setOrderQty]=useState({});
  const [note,setNote]=useState(""); const [sent,setSent]=useState(false);
  const [sending,setSending]=useState(false); const [preview,setPreview]=useState(false);
  const needOrder=stock.filter(s=>s.qty<s.minQty);
  const selectAll=()=>{const sel={};const qty={};needOrder.forEach(s=>{sel[s.id]=true;qty[s.id]=String(s.minQty*2-s.qty);});setSelected(sel);setOrderQty(qty);};
  const toggle=id=>{setSelected(p=>({...p,[id]:!p[id]}));if(!orderQty[id]){const it=stock.find(s=>s.id===id);if(it)setOrderQty(p=>({...p,[id]:String(Math.max(it.minQty*2-it.qty,1))}))}};
  const selItems=stock.filter(s=>selected[s.id]);
  const buildMsg=()=>{
    const lines=["🛒 ใบสั่งซื้อวัตถุดิบ",`📅 ${today()}`,"─────────"];
    const bySup={};
    selItems.forEach(item=>{const sup=suppliers.find(x=>x.id===item.supplierId);const sn=sup?.name||"ไม่ระบุ";if(!bySup[sn])bySup[sn]=[];bySup[sn].push(item);});
    Object.entries(bySup).forEach(([sn,items])=>{lines.push(`\n🏪 ${sn}`);items.forEach(item=>lines.push(`  • ${item.name}: ${orderQty[item.id]||item.minQty} ${item.unit} (มี ${item.qty})`));});
    if(note)lines.push(`\n📝 ${note}`);
    lines.push("\n─────────\nไท่กั๋วหม่าล่า");
    return lines.join("\n");
  };
  const sendLine=async()=>{
    if(!lineToken){alert("ตั้งค่า LINE Token ใน Settings ก่อน");return;}
    if(!selItems.length){alert("เลือกรายการก่อน");return;}
    setSending(true);
    try{
      const res=await fetch("https://notify-api.line.me/api/notify",{method:"POST",headers:{"Authorization":`Bearer ${lineToken}`,"Content-Type":"application/x-www-form-urlencoded"},body:`message=${encodeURIComponent(buildMsg())}`});
      if(res.ok){setSent(true);setTimeout(()=>setSent(false),5000);}else alert("ส่งไม่สำเร็จ ตรวจสอบ Token");
    }catch{alert("เกิดข้อผิดพลาด");}
    setSending(false);
  };
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="🛒 สั่งซื้อวัตถุดิบ" />
      {sent&&<div style={{background:T.greenLt,border:`1px solid ${T.green}44`,borderRadius:10,padding:"12px 16px",color:T.green,fontWeight:800}}>✅ ส่ง LINE สำเร็จ!</div>}
      {needOrder.length>0?(
        <Card style={{borderColor:T.red+"44",background:T.redLt}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <div style={{color:T.red,fontWeight:800,fontSize:16}}>🚨 ต้องสั่งด่วน ({needOrder.length})</div>
            <button onClick={selectAll} style={{...S.btn(T.red),fontSize:13,padding:"6px 12px"}}>เลือกทั้งหมด</button>
          </div>
          {needOrder.map(s=>{const isSel=selected[s.id];const sup=suppliers.find(x=>x.id===s.supplierId);return(
            <div key={s.id} style={{background:"#fff",borderRadius:10,padding:"12px 14px",marginBottom:8,border:`2px solid ${isSel?T.orange:T.border}`,cursor:"pointer"}} onClick={()=>toggle(s.id)}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:22,height:22,borderRadius:6,border:`2px solid ${isSel?T.orange:T.border}`,background:isSel?T.orange:"transparent",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,color:"#fff"}}>{isSel?"✓":""}</div>
                  <div><div style={{fontWeight:700,fontSize:15}}>{s.name}</div><div style={{color:T.textSm,fontSize:12}}>ซัพฯ: {sup?.name||"-"}</div></div>
                </div>
                <div style={{textAlign:"right"}}><div style={{color:T.red,fontWeight:700}}>เหลือ {s.qty} {s.unit}</div><div style={{color:T.textXs,fontSize:11}}>ขั้นต่ำ {s.minQty}</div></div>
              </div>
              {isSel&&<div style={{marginTop:10,paddingTop:10,borderTop:`1px solid ${T.bg}`,display:"flex",alignItems:"center",gap:10}} onClick={e=>e.stopPropagation()}>
                <span style={{color:T.textSm,fontSize:13}}>สั่ง:</span>
                <input type="number" value={orderQty[s.id]||""} onChange={e=>setOrderQty(p=>({...p,[s.id]:e.target.value}))} style={{...S.inp,width:90,fontSize:16,fontWeight:700,textAlign:"center"}} placeholder="0" />
                <span style={{color:T.textSm,fontSize:13}}>{s.unit}</span>
              </div>}
            </div>
          );})}
        </Card>
      ):<Card style={{background:T.greenLt,borderColor:T.green+"44"}}><div style={{color:T.green,fontWeight:800,fontSize:16,textAlign:"center"}}>✅ สต็อคปกติทั้งหมด</div></Card>}
      <div><div style={{color:T.textSm,fontSize:13,marginBottom:6}}>หมายเหตุ</div><input value={note} onChange={e=>setNote(e.target.value)} style={S.inp} placeholder="เช่น ส่งด่วนก่อน 8 โมง" /></div>
      {selItems.length>0&&<button onClick={()=>setPreview(!preview)} style={{...S.ghost,width:"100%",padding:10,fontSize:14}}>{preview?"▲ ซ่อน":"▼ ดูตัวอย่าง"} ใบสั่งซื้อ ({selItems.length} รายการ)</button>}
      {preview&&<Card style={{background:"#1a1a2e"}}><pre style={{color:"#e2e8f0",fontSize:13,whiteSpace:"pre-wrap",margin:0,lineHeight:1.8}}>{buildMsg()}</pre></Card>}
      {!lineToken&&<div style={{background:T.yellowLt,borderRadius:8,padding:"10px 14px",fontSize:13,color:T.yellow,fontWeight:600}}>⚠️ ตั้งค่า LINE Token ใน Settings → LINE ก่อน</div>}
      <button onClick={sendLine} disabled={sending||!selItems.length} style={{...S.btn(!selItems.length?"#a1a1aa":T.green),width:"100%",padding:15,fontSize:17,opacity:!selItems.length?0.5:1}}>
        {sending?"⏳ กำลังส่ง...":`📲 ส่ง LINE (${selItems.length} รายการ)`}
      </button>
    </div>
  );
}

// ─── Settings ───
function SettingsPage({staff,setStaff,lineToken,setLineToken,fixedCosts,setFixedCosts,suppliers,setSuppliers}){
  const [tab,setTab]=useState("staff");
  const [showAddSup,setShowAddSup]=useState(false);
  const [newSup,setNewSup]=useState({name:"",type:"",phone:"",active:true});
  const [editSupId,setEditSupId]=useState(null);
  const [editSupData,setEditSupData]=useState({});
  const PERMS={cashflow:"💵 Cash Flow",stock:"📦 สต็อค",purchase:"🛒 สั่งซื้อ",report:"📊 รายงาน",admin:"⚙️ Admin",viewPrice:"💰 ดูราคา",viewFinance:"💎 ดูการเงิน"};
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="⚙️ ตั้งค่า" />
      <Tabs tabs={[["staff","👷 พนักงาน"],["supplier","🏪 ซัพพลายเออร์"],["fixed","🔒 ต้นทุน"],["line","📲 LINE"]]} active={tab} onChange={setTab} />
      {tab==="staff"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        {staff.filter(s=>s.id!=="owner"&&s.id!=="emergency").map(s=>(
          <Card key={s.id} style={{borderColor:s.active?T.border:T.red+"33"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:16}}>{s.name}</div>
                <div style={{display:"flex",alignItems:"center",gap:8,marginTop:4}}>
                  <span style={{color:T.textSm,fontSize:13}}>PIN:</span>
                  <input maxLength={4} value={s.pin}
                    onChange={e=>setStaff(p=>p.map(x=>x.id===s.id?{...x,pin:e.target.value.replace(/\D/,"")}:x))}
                    style={{...S.inp,width:80,fontSize:16,letterSpacing:4,textAlign:"center",padding:"4px 8px"}}
                    placeholder="••••" />
                </div>
              </div>
              <button onClick={()=>setStaff(p=>p.map(x=>x.id===s.id?{...x,active:!x.active}:x))} style={{background:s.active?T.greenLt:T.redLt,border:`1px solid ${s.active?T.green:T.red}44`,borderRadius:8,padding:"6px 12px",color:s.active?T.green:T.red,cursor:"pointer",fontSize:13,fontWeight:700,fontFamily:F,flexShrink:0}}>{s.active?"ใช้งาน":"ระงับ"}</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6}}>
              {Object.entries(PERMS).map(([perm,label])=>(
                <button key={perm} onClick={()=>setStaff(p=>p.map(x=>x.id===s.id?{...x,perms:{...x.perms,[perm]:!x.perms[perm]}}:x))} style={{background:s.perms[perm]?T.orange:"transparent",border:`1px solid ${s.perms[perm]?T.orange:T.border}`,borderRadius:8,padding:"7px 4px",color:s.perms[perm]?"#fff":T.textMd,cursor:"pointer",fontSize:11,fontFamily:F,fontWeight:s.perms[perm]?700:400,textAlign:"center"}}>{label}</button>
              ))}
            </div>
          </Card>
        ))}
      </div>}
      {tab==="supplier"&&(
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <button onClick={()=>setShowAddSup(!showAddSup)} style={{...S.btn(),width:"100%",padding:12,fontSize:15}}>+ เพิ่มซัพพลายเออร์</button>
          {showAddSup&&<Card style={{borderColor:T.borderOr}}>
            <div style={{color:T.orange,fontWeight:800,fontSize:16,marginBottom:10}}>➕ เพิ่มซัพพลายเออร์</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ชื่อ</div><input value={newSup.name} onChange={e=>setNewSup(p=>({...p,name:e.target.value}))} style={S.inp} placeholder="ชื่อซัพพลายเออร์" /></div>
              <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ประเภท</div><input value={newSup.type} onChange={e=>setNewSup(p=>({...p,type:e.target.value}))} style={S.inp} placeholder="เช่น ผัก, หมู" /></div>
              <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>เบอร์โทร</div><input value={newSup.phone} onChange={e=>setNewSup(p=>({...p,phone:e.target.value}))} style={S.inp} placeholder="0xx-xxx-xxxx" /></div>
            </div>
            <div style={{display:"flex",gap:8,marginTop:10}}>
              <button onClick={()=>{if(!newSup.name)return;setSuppliers(p=>[...p,{...newSup,id:Date.now()}]);setNewSup({name:"",type:"",phone:"",active:true});setShowAddSup(false);}} style={{...S.btn(),flex:1}}>บันทึก</button>
              <button onClick={()=>setShowAddSup(false)} style={S.ghost}>ยกเลิก</button>
            </div>
          </Card>}
          {suppliers.map(s=>(
            <Card key={s.id} style={{borderColor:s.active?T.border:T.red+"33"}}>
              {editSupId===s.id?(
                <div>
                  <div style={{color:T.orange,fontWeight:700,fontSize:15,marginBottom:10}}>✏️ แก้ไข</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
                    <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>ชื่อ</div><input value={editSupData.name??s.name} onChange={e=>setEditSupData(p=>({...p,name:e.target.value}))} style={S.inp} /></div>
                    <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>ประเภท</div><input value={editSupData.type??s.type} onChange={e=>setEditSupData(p=>({...p,type:e.target.value}))} style={S.inp} /></div>
                    <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>เบอร์โทร</div><input value={editSupData.phone??s.phone} onChange={e=>setEditSupData(p=>({...p,phone:e.target.value}))} style={S.inp} /></div>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>{setSuppliers(p=>p.map(x=>x.id===s.id?{...x,...editSupData}:x));setEditSupId(null);setEditSupData({});}} style={{...S.btn(),flex:1}}>บันทึก</button>
                    <button onClick={()=>{setEditSupId(null);setEditSupData({});}} style={S.ghost}>ยกเลิก</button>
                  </div>
                </div>
              ):(
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <div style={{fontWeight:700,fontSize:16}}>{s.name}</div>
                    <div style={{color:T.textSm,fontSize:13}}>{s.type} • {s.phone}</div>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>setSuppliers(p=>p.map(x=>x.id===s.id?{...x,active:!x.active}:x))} style={{background:s.active?T.greenLt:T.redLt,border:`1px solid ${s.active?T.green:T.red}44`,borderRadius:7,padding:"5px 10px",color:s.active?T.green:T.red,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:F}}>{s.active?"เปิด":"ปิด"}</button>
                    <button onClick={()=>{setEditSupId(s.id);setEditSupData({});}} style={{...S.ghost,padding:"5px 10px",fontSize:12}}>✏️</button>
                    <button onClick={()=>{if(window.confirm(`ลบ "${s.name}"?`))setSuppliers(p=>p.filter(x=>x.id!==s.id));}} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:14}}>🗑</button>
                  </div>
                </div>
              )}
            </Card>
          ))}
        </div>
      )}

            {tab==="fixed"&&<div style={{display:"flex",flexDirection:"column",gap:10}}>
        {fixedCosts.map((fc,i)=>(
          <Card key={i} style={{padding:"12px 16px"}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr auto auto",gap:10,alignItems:"flex-end"}}>
              <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>ชื่อ</div><input value={fc.name} onChange={e=>setFixedCosts(p=>p.map((f,j)=>j===i?{...f,name:e.target.value}:f))} style={{...S.inp,fontSize:15}} /></div>
              <div style={{minWidth:110}}><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>฿/เดือน</div><input type="number" value={fc.amount} onChange={e=>setFixedCosts(p=>p.map((f,j)=>j===i?{...f,amount:+e.target.value||0}:f))} style={{...S.inp,fontSize:15,textAlign:"right"}} /></div>
              <button onClick={()=>setFixedCosts(p=>p.filter((_,j)=>j!==i))} style={{background:T.redLt,border:`1px solid ${T.red}33`,borderRadius:8,padding:"10px 12px",color:T.red,cursor:"pointer",fontSize:14,fontFamily:F}}>🗑</button>
            </div>
          </Card>
        ))}
        <button onClick={()=>setFixedCosts(p=>[...p,{name:"รายการใหม่",amount:0}])} style={{...S.ghost,width:"100%",padding:12,fontSize:15}}>+ เพิ่มรายการ</button>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:6}}>รวม: <span style={{color:T.orange}}>฿{fmt(fixedCosts.reduce((a,b)=>a+b.amount,0))}</span>/เดือน</div>
          <div style={{display:"flex",gap:20,fontSize:13}}>
            <div><div style={{color:T.textSm}}>BEP/วัน</div><div style={{color:T.orange,fontWeight:700,fontSize:16}}>฿{fmt(Math.ceil(fixedCosts.reduce((a,b)=>a+b.amount,0)/30))}</div></div>
            <div><div style={{color:T.textSm}}>เป้า/วัน (×2.5)</div><div style={{color:T.green,fontWeight:700,fontSize:16}}>฿{fmt(Math.ceil(fixedCosts.reduce((a,b)=>a+b.amount,0)/30*2.5))}</div></div>
          </div>
        </Card>
      </div>}
      {tab==="line"&&<Card>
        <div style={{color:T.orange,fontWeight:800,fontSize:17,marginBottom:14}}>📲 LINE Notify Token</div>
        <div style={{color:T.textSm,fontSize:13,marginBottom:8}}>รับ Token จาก notify-bot.line.me</div>
        <input type="password" value={lineToken} onChange={e=>setLineToken(e.target.value)} style={S.inp} placeholder="ใส่ Token ที่นี่..." />
        {lineToken&&<div style={{color:T.green,fontSize:13,marginTop:6,fontWeight:600}}>✅ Token บันทึกแล้ว</div>}
      </Card>}
    </div>
  );
}

// ─── Report ───
function ReportPage({cf,stock,movements,user,fixedCosts,waste,setWaste,promos,setPromos}){
  const [tab,setTab]=useState("pl");
  const [wForm,setWForm]=useState({itemId:"",qty:"",reason:"",date:today()});
  const [pForm,setPForm]=useState({name:"",date:today(),amount:"",note:""});
  const [showWForm,setShowWForm]=useState(false);
  const [showPForm,setShowPForm]=useState(false);
  const mk=today().slice(0,7);
  const myCF=user.role==="owner"?cf:cf.filter(e=>e.branch===user.franchiseId);
  const mCF=myCF.filter(e=>e.date.startsWith(mk));
  const mIn=mCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mOut=mCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const totalFixed=fixedCosts.reduce((a,b)=>a+b.amount,0);
  const cogs=mCF.filter(e=>e.flow==="out"&&["วัตถุดิบ/ผัก","หมู/เนื้อ/ทะเล"].includes(e.cat)).reduce((a,b)=>a+b.amount,0);
  const netP=mIn-mOut-totalFixed;

  // Cost analysis per item
  const costAnalysis = stock.map(s=>{
    const h = s.costHistory||[];
    if(h.length===0) return {s, avg:0, trend:0, risk:"none", last3:[]};
    const avg = wac(s);
    // last 3 purchases
    const last3 = h.slice(-3);
    // trend: compare first half vs second half
    const mid = Math.floor(h.length/2);
    const firstHalfAvg = h.slice(0,mid).reduce((a,b)=>a+b.unitCost,0)/Math.max(mid,1);
    const secondHalfAvg = h.slice(mid).reduce((a,b)=>a+b.unitCost,0)/Math.max(h.length-mid,1);
    const trend = firstHalfAvg>0 ? ((secondHalfAvg-firstHalfAvg)/firstHalfAvg*100) : 0;
    // price spike detection
    const prices = h.map(x=>x.unitCost);
    const maxP = Math.max(...prices);
    const minP = Math.min(...prices);
    const volatility = minP>0?(maxP-minP)/minP*100:0;
    const risk = trend>20?"high":trend>10?"medium":volatility>30?"medium":"low";
    return {s, avg, trend, risk, last3, volatility, maxP, minP};
  }).filter(x=>x.s.costHistory?.length>0);

  // Stock count history from movements
  const checkMovements = movements.filter(m=>m.type==="check")
    .sort((a,b)=>b.date.localeCompare(a.date));
  const checkDates = [...new Set(checkMovements.map(m=>m.date))].slice(0,7);

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="📊 รายงาน" />
      <Tabs tabs={[["pl","💰 กำไรขาดทุน"],["cost","📈 ต้นทุน"],["forecast","🔮 พยากรณ์"],["waste","🗑 Waste"],["promo","🎁 โปร"],["stockcheck","📋 นับสต็อค"]]} active={tab} onChange={setTab} />

      {/* P&L */}
      {tab==="pl"&&<>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{color:T.orange,fontWeight:700,fontSize:14,marginBottom:10}}>เดือน {mk}</div>
          {[["💰 รายรับรวม",mIn,T.green,false],["− ต้นทุนวัตถุดิบ",cogs,T.red,true],["= กำไรขั้นต้น",mIn-cogs,T.green,false],["− ค่าใช้จ่ายดำเนินการ",mOut-cogs+totalFixed,T.yellow,true],["= กำไรสุทธิ",netP,netP>=0?T.green:T.red,false]].map(([l,v,c2,indent])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"10px 0",borderBottom:`1px solid ${T.borderOr}`,paddingLeft:indent?20:0}}>
              <span style={{color:T.textMd,fontSize:15}}>{l}</span><span style={{color:c2,fontWeight:800,fontSize:16}}>฿{fmt(v)}</span>
            </div>
          ))}
          <div style={{display:"flex",gap:20,marginTop:12,paddingTop:10,borderTop:`1px solid ${T.borderOr}`,fontSize:13}}>
            <div><div style={{color:T.textSm}}>Gross Margin</div><div style={{color:T.orange,fontWeight:700,fontSize:16}}>{mIn>0?((mIn-cogs)/mIn*100).toFixed(1):0}%</div></div>
            <div><div style={{color:T.textSm}}>BEP/วัน</div><div style={{color:T.orange,fontWeight:700,fontSize:16}}>฿{fmt(Math.ceil(totalFixed/30))}</div></div>
          </div>
        </Card>
        <Card>
          <div style={{fontWeight:800,fontSize:16,marginBottom:10}}>📦 สต็อคคงเหลือ</div>
          {stock.filter(s=>stockSt(s)!=="ok").map(s=>(
            <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.bg}`,fontSize:14}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontWeight:600}}>{s.name}</span><Badge status={stockSt(s)} /></div>
              <span style={{color:ST_C[stockSt(s)],fontWeight:700}}>{s.qty} {s.unit}</span>
            </div>
          ))}
          {stock.every(s=>stockSt(s)==="ok")&&<div style={{color:T.green,textAlign:"center",padding:16}}>✅ สต็อคปกติทั้งหมด</div>}
        </Card>
      </>}

      {/* Cost Analysis */}
      {tab==="cost"&&<>
        {costAnalysis.filter(x=>x.risk==="high").length>0&&(
          <Card style={{background:T.redLt,borderColor:T.red+"44"}}>
            <div style={{color:T.red,fontWeight:800,fontSize:16,marginBottom:8}}>⚠️ ความเสี่ยงต้นทุนสูง</div>
            {costAnalysis.filter(x=>x.risk==="high").map(({s,avg,trend})=>(
              <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.red}22`,fontSize:14}}>
                <div><div style={{fontWeight:700}}>{s.name}</div><div style={{color:T.textSm,fontSize:12}}>ต้นทุนเพิ่มขึ้น {trend.toFixed(1)}% จากช่วงแรก</div></div>
                <div style={{textAlign:"right"}}><div style={{color:T.red,fontWeight:800}}>฿{avg.toFixed(2)}</div><div style={{color:T.textXs,fontSize:11}}>เฉลี่ย/{s.unit}</div></div>
              </div>
            ))}
          </Card>
        )}
        {costAnalysis.filter(x=>x.risk==="medium").length>0&&(
          <Card style={{background:T.yellowLt,borderColor:T.yellow+"44"}}>
            <div style={{color:T.yellow,fontWeight:800,fontSize:15,marginBottom:8}}>🟡 ควรติดตาม</div>
            {costAnalysis.filter(x=>x.risk==="medium").map(({s,avg,trend,volatility})=>(
              <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.yellow}22`,fontSize:14}}>
                <div><div style={{fontWeight:700}}>{s.name}</div><div style={{color:T.textSm,fontSize:12}}>{trend>0?`ขึ้น ${trend.toFixed(1)}%`:`ผันผวน ${volatility.toFixed(0)}%`}</div></div>
                <div style={{textAlign:"right"}}><div style={{color:T.yellow,fontWeight:800}}>฿{avg.toFixed(2)}</div><div style={{color:T.textXs,fontSize:11}}>เฉลี่ย/{s.unit}</div></div>
              </div>
            ))}
          </Card>
        )}
        <Card>
          <div style={{fontWeight:800,fontSize:16,marginBottom:12}}>📈 ต้นทุนเฉลี่ยแต่ละสินค้า</div>
          {costAnalysis.length===0&&<div style={{color:T.textSm,textAlign:"center",padding:24}}>ยังไม่มีข้อมูลราคา<br/><span style={{fontSize:12}}>กรอกราคาตอนรับสินค้าเข้าก่อนครับ</span></div>}
          {costAnalysis.map(({s,avg,trend,risk,last3,maxP,minP})=>(
            <div key={s.id} style={{marginBottom:18,paddingBottom:16,borderBottom:`1px solid ${T.bg}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                <div>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontWeight:700,fontSize:16}}>{s.name}</span>
                    {risk==="high"&&<span style={{background:T.redLt,color:T.red,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>⚠️ ระวัง</span>}
                    {risk==="medium"&&<span style={{background:T.yellowLt,color:T.yellow,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>🟡 ติดตาม</span>}
                    {risk==="low"&&<span style={{background:T.greenLt,color:T.green,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>✅ ปกติ</span>}
                  </div>
                  <div style={{color:T.textSm,fontSize:12,marginTop:2}}>
                    ต่ำสุด ฿{minP.toFixed(2)} • สูงสุด ฿{maxP.toFixed(2)} • รับ {s.costHistory.length} ครั้ง
                  </div>
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{fontWeight:900,fontSize:22,color:risk==="high"?T.red:risk==="medium"?T.yellow:T.green}}>฿{avg.toFixed(2)}</div>
                  <div style={{color:T.textXs,fontSize:11}}>เฉลี่ย/{s.unit}</div>
                  {trend!==0&&<div style={{fontSize:12,fontWeight:600,color:trend>0?T.red:T.green,marginTop:2}}>
                    {trend>0?"▲":"▼"} {Math.abs(trend).toFixed(1)}%
                  </div>}
                </div>
              </div>
              {/* Price bar range */}
              <div style={{marginBottom:8}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:T.textXs,marginBottom:3}}>
                  <span>฿{minP.toFixed(0)}</span><span>ช่วงราคา</span><span>฿{maxP.toFixed(0)}</span>
                </div>
                <div style={{background:T.bg,borderRadius:4,height:8,position:"relative"}}>
                  <div style={{position:"absolute",left:`${maxP>minP?(avg-minP)/(maxP-minP)*100:50}%`,transform:"translateX(-50%)",width:12,height:12,borderRadius:"50%",background:T.orange,top:-2,border:"2px solid #fff",boxShadow:"0 1px 3px rgba(0,0,0,.2)"}} />
                  <div style={{background:`linear-gradient(90deg,${T.green},${T.yellow},${T.red})`,height:"100%",borderRadius:4,opacity:.4}} />
                </div>
              </div>
              {/* Last 3 receipts */}
              {last3.length>0&&(
                <div style={{display:"flex",gap:6}}>
                  {last3.map((h,i)=>(
                    <div key={i} style={{flex:1,background:T.bg,borderRadius:8,padding:"6px 8px",textAlign:"center"}}>
                      <div style={{fontSize:11,color:T.textXs}}>{h.date?.slice(5)||"-"}</div>
                      <div style={{fontWeight:700,fontSize:14,color:h.unitCost>avg*1.1?T.red:h.unitCost<avg*.9?T.green:T.text}}>฿{h.unitCost.toFixed(0)}</div>
                      <div style={{fontSize:10,color:T.textXs}}>{h.qty} {s.unit}</div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          ))}
        </Card>
      </>}


      {/* Forecast */}
      {tab==="forecast"&&<>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{color:T.orange,fontWeight:700,fontSize:15,marginBottom:4}}>🔮 พยากรณ์การจัดซื้อ</div>
          <div style={{color:T.textMd,fontSize:13}}>คำนวณจากค่าเฉลี่ยการใช้ 7 วันล่าสุด</div>
        </Card>
        {stock.map(s=>{
          const last7=Array.from({length:7},(_,i)=>{
            const d=new Date(); d.setDate(d.getDate()-i);
            return d.toISOString().split("T")[0];
          });
          const outMvs=movements.filter(m=>m.itemId===s.id&&m.type==="out"&&last7.includes(m.date));
          const dailyAvg=outMvs.reduce((a,b)=>a+b.qty,0)/7;
          const daysLeft=dailyAvg>0?s.qty/dailyAvg:999;
          const need7=dailyAvg*7; const need14=dailyAvg*14;
          const urgent=daysLeft<3;
          const warn=daysLeft<7;
          const avgCost=wac(s);
          return(
            <Card key={s.id} style={{borderColor:urgent?T.red+"44":warn?T.yellow+"44":T.border}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div style={{flex:1}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                    <span style={{fontWeight:700,fontSize:16}}>{s.name}</span>
                    {urgent&&<span style={{background:T.redLt,color:T.red,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>🚨 สั่งด่วน</span>}
                    {!urgent&&warn&&<span style={{background:T.yellowLt,color:T.yellow,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>⚠️ ใกล้หมด</span>}
                  </div>
                  <div style={{color:T.textSm,fontSize:13}}>ใช้เฉลี่ย <b>{dailyAvg.toFixed(2)} {s.unit}/วัน</b> • เหลือ {daysLeft<999?daysLeft.toFixed(1):"∞"} วัน</div>
                  {dailyAvg>0&&<>
                    <div style={{display:"flex",gap:12,marginTop:8,fontSize:13}}>
                      <div style={{background:T.bg,borderRadius:8,padding:"6px 10px"}}>
                        <div style={{color:T.textXs,fontSize:11}}>แนะนำสั่ง 7 วัน</div>
                        <div style={{fontWeight:700,color:T.orange}}>{need7.toFixed(1)} {s.unit}</div>
                        {avgCost>0&&<div style={{color:T.textXs,fontSize:11}}>≈ ฿{fmt(need7*avgCost)}</div>}
                      </div>
                      <div style={{background:T.bg,borderRadius:8,padding:"6px 10px"}}>
                        <div style={{color:T.textXs,fontSize:11}}>แนะนำสั่ง 14 วัน</div>
                        <div style={{fontWeight:700,color:T.blue}}>{need14.toFixed(1)} {s.unit}</div>
                        {avgCost>0&&<div style={{color:T.textXs,fontSize:11}}>≈ ฿{fmt(need14*avgCost)}</div>}
                      </div>
                    </div>
                  </>}
                  {dailyAvg===0&&<div style={{color:T.textXs,fontSize:12,marginTop:4}}>ไม่มีข้อมูลการใช้ 7 วันล่าสุด</div>}
                </div>
                <div style={{textAlign:"right",marginLeft:12}}>
                  <div style={{fontWeight:900,fontSize:24,color:urgent?T.red:warn?T.yellow:T.green}}>{s.qty}</div>
                  <div style={{color:T.textSm,fontSize:12}}>{s.unit}</div>
                </div>
              </div>
            </Card>
          );
        })}

        {/* สรุปใบสั่งซื้อล่วงหน้า */}
        {(()=>{
          const items7=[], items14=[];
          stock.forEach(s=>{
            const last7d=Array.from({length:7},(_,i)=>{const d=new Date();d.setDate(d.getDate()-i);return d.toISOString().split("T")[0];});
            const outMvs=movements.filter(m=>m.itemId===s.id&&m.type==="out"&&last7d.includes(m.date));
            const dailyAvg=outMvs.reduce((a,b)=>a+b.qty,0)/7;
            if(dailyAvg===0)return;
            const need7=Math.ceil(dailyAvg*7);
            const need14=Math.ceil(dailyAvg*14);
            const avgCost=wac(s);
            const daysLeft=s.qty/dailyAvg;
            if(daysLeft<7) items7.push({s,need:need7,avgCost,daysLeft,urgent:daysLeft<3});
            items14.push({s,need7,need14,avgCost,daysLeft});
          });
          if(items14.length===0)return <Card><div style={{color:T.textSm,textAlign:"center",padding:24}}>ยังไม่มีข้อมูลเพียงพอสำหรับสรุป<br/><span style={{fontSize:12}}>ต้องมีข้อมูลการใช้สต็อคอย่างน้อย 1 วัน</span></div></Card>;
          const total7=items7.reduce((a,b)=>a+(b.avgCost>0?b.need*b.avgCost:0),0);
          const total14=items14.reduce((a,b)=>a+(b.avgCost>0?b.need14*b.avgCost:0),0);
          return (
            <Card style={{borderColor:T.orange,background:T.orangeLt}}>
              <div style={{color:T.orange,fontWeight:900,fontSize:17,marginBottom:12}}>📋 สรุปใบสั่งซื้อล่วงหน้า</div>

              {/* ด่วน - ต้องสั่งภายใน 7 วัน */}
              {items7.length>0&&<>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                  <div style={{color:T.red,fontWeight:700,fontSize:15}}>🚨 ต้องสั่งเร็วๆ นี้ (เหลือไม่ถึง 7 วัน)</div>
                  {total7>0&&<div style={{color:T.red,fontWeight:800,fontSize:15}}>฿{fmt(total7)}</div>}
                </div>
                {items7.map(({s,need,avgCost,daysLeft,urgent})=>(
                  <div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 12px",marginBottom:6,background:"#fff",borderRadius:10,border:`2px solid ${urgent?T.red:T.yellow}`}}>
                    <div>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontWeight:700,fontSize:15}}>{s.name}</span>
                        <span style={{background:urgent?T.redLt:T.yellowLt,color:urgent?T.red:T.yellow,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>
                          {urgent?"🚨 สั่งด่วน":"⚠️ ใกล้หมด"}
                        </span>
                      </div>
                      <div style={{color:T.textSm,fontSize:12,marginTop:2}}>
                        เหลือ {daysLeft.toFixed(1)} วัน • สั่ง {need} {s.unit}
                      </div>
                    </div>
                    <div style={{textAlign:"right"}}>
                      <div style={{fontWeight:800,fontSize:16,color:T.orange}}>{need} {s.unit}</div>
                      {avgCost>0&&<div style={{color:T.textSm,fontSize:12}}>≈ ฿{fmt(need*avgCost)}</div>}
                    </div>
                  </div>
                ))}
                <div style={{borderBottom:`1px solid ${T.borderOr}`,marginBottom:12,paddingBottom:12}}/>
              </>}

              {/* แผน 14 วัน */}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <div style={{color:T.text,fontWeight:700,fontSize:15}}>📅 แผนสั่งซื้อ 14 วัน</div>
                {total14>0&&<div style={{color:T.orange,fontWeight:800,fontSize:15}}>฿{fmt(total14)}</div>}
              </div>
              {items14.map(({s,need14,avgCost,daysLeft})=>(
                <div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 12px",marginBottom:6,background:"#fff",borderRadius:10,border:`1px solid ${T.border}`}}>
                  <div>
                    <div style={{fontWeight:600,fontSize:15}}>{s.name}</div>
                    <div style={{color:T.textSm,fontSize:12,marginTop:1}}>
                      เหลือ {daysLeft.toFixed(1)} วัน • ใช้เฉลี่ย {(movements.filter(m=>m.itemId===s.id&&m.type==="out").reduce((a,b)=>a+b.qty,0)/7).toFixed(2)} {s.unit}/วัน
                    </div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontWeight:800,fontSize:16,color:T.blue}}>{need14} {s.unit}</div>
                    {avgCost>0&&<div style={{color:T.textSm,fontSize:12}}>≈ ฿{fmt(need14*avgCost)}</div>}
                  </div>
                </div>
              ))}

              {/* รวมงบประมาณ */}
              {total14>0&&(
                <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${T.borderOr}`}}>
                  <div style={{fontWeight:700,fontSize:15,marginBottom:8}}>💰 ประมาณงบประมาณ</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                    <div style={{background:"#fff",borderRadius:10,padding:"10px 14px",textAlign:"center",border:`1px solid ${T.red}44`}}>
                      <div style={{color:T.textSm,fontSize:12}}>สั่งด่วน (7 วัน)</div>
                      <div style={{color:T.red,fontWeight:900,fontSize:20}}>฿{fmt(total7)}</div>
                    </div>
                    <div style={{background:"#fff",borderRadius:10,padding:"10px 14px",textAlign:"center",border:`1px solid ${T.blue}44`}}>
                      <div style={{color:T.textSm,fontSize:12}}>แผน 14 วัน</div>
                      <div style={{color:T.blue,fontWeight:900,fontSize:20}}>฿{fmt(total14)}</div>
                    </div>
                  </div>
                </div>
              )}
            </Card>
          );
        })()}
      </>}

      {/* Waste */}
      {tab==="waste"&&<>
        <Card style={{background:T.redLt,borderColor:T.red+"44"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{color:T.red,fontWeight:700,fontSize:15}}>🗑 บันทึกของเสีย / Waste</div>
              <div style={{color:T.textMd,fontSize:13}}>เดือนนี้: <b>฿{fmt(waste.filter(w=>w.date.startsWith(mk)).reduce((a,b)=>a+b.cost,0))}</b></div>
            </div>
            <button onClick={()=>setShowWForm(!showWForm)} style={{...S.btn(T.red),fontSize:13,padding:"7px 12px"}}>+ บันทึก</button>
          </div>
        </Card>
        {showWForm&&<Card style={{borderColor:T.red+"44"}}>
          <div style={{color:T.red,fontWeight:800,fontSize:16,marginBottom:12}}>🗑 บันทึก Waste</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div style={{gridColumn:"1/-1"}}>
              <div style={{color:T.textSm,fontSize:13,marginBottom:4}}>สินค้า</div>
              <select value={wForm.itemId} onChange={e=>setWForm(p=>({...p,itemId:e.target.value}))} style={{...S.inp,height:44}}>
                <option value="">— เลือกสินค้า —</option>
                {stock.map(s=><option key={s.id} value={s.id}>{s.name} (คงเหลือ {s.qty} {s.unit})</option>)}
              </select>
            </div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>จำนวนที่เสียหาย</div>
              <input type="number" value={wForm.qty} onChange={e=>setWForm(p=>({...p,qty:e.target.value}))} style={S.inp} placeholder="0" /></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>วันที่</div>
              <input type="date" value={wForm.date} max={today()} onChange={e=>setWForm(p=>({...p,date:e.target.value}))} style={S.inp} /></div>
            <div style={{gridColumn:"1/-1"}}>
              <div style={{color:T.textSm,fontSize:13,marginBottom:4}}>สาเหตุ</div>
              <input value={wForm.reason} onChange={e=>setWForm(p=>({...p,reason:e.target.value}))} style={S.inp} placeholder="เช่น ผักเน่า, ของหล่น, หมดอายุ" /></div>
          </div>
          {wForm.itemId&&wForm.qty&&(()=>{
            const item=stock.find(s=>String(s.id)===String(wForm.itemId));
            const costPerUnit=wac(item||{costHistory:[]});
            const totalCost=costPerUnit*(+wForm.qty||0);
            return item&&<div style={{background:T.redLt,borderRadius:8,padding:"8px 12px",marginTop:8,fontSize:14}}>
              มูลค่าความเสียหาย: <b style={{color:T.red}}>฿{fmt(totalCost)}</b>
              {costPerUnit===0&&<span style={{color:T.textXs,fontSize:12}}> (ยังไม่มีราคาต้นทุน)</span>}
            </div>;
          })()}
          <div style={{display:"flex",gap:8,marginTop:12}}>
            <button onClick={()=>{
              if(!wForm.itemId||!wForm.qty)return;
              const item=stock.find(s=>String(s.id)===String(wForm.itemId));
              const cost=wac(item||{costHistory:[]})*(+wForm.qty||0);
              setWaste(p=>[{...wForm,id:Date.now(),cost,qty:+wForm.qty,itemName:item?.name||""},  ...p]);
              setWForm({itemId:"",qty:"",reason:"",date:today()});
              setShowWForm(false);
            }} style={{...S.btn(T.red),flex:1}}>✅ บันทึก</button>
            <button onClick={()=>setShowWForm(false)} style={S.ghost}>ยกเลิก</button>
          </div>
        </Card>}
        {waste.length===0&&<Card><div style={{color:T.textSm,textAlign:"center",padding:32}}>ยังไม่มีการบันทึก Waste</div></Card>}
        {[...waste].sort((a,b)=>b.date.localeCompare(a.date)).map(w=>(
          <Card key={w.id} style={{padding:"12px 16px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontWeight:700,fontSize:15}}>{w.itemName}</div>
                <div style={{color:T.textSm,fontSize:13}}>{w.date} • {w.reason||"ไม่ระบุสาเหตุ"}</div>
                <div style={{color:T.red,fontSize:13,fontWeight:600}}>เสียหาย {w.qty} {stock.find(s=>String(s.id)===String(w.itemId))?.unit||""}</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{color:T.red,fontWeight:800,fontSize:18}}>฿{fmt(w.cost)}</div>
                <button onClick={()=>setWaste(p=>p.filter(x=>x.id!==w.id))} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:13,marginTop:4}}>🗑 ลบ</button>
              </div>
            </div>
          </Card>
        ))}
        {waste.length>0&&<Card style={{background:T.redLt,borderColor:T.red+"44"}}>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:15,fontWeight:700}}>
            <span>รวม Waste ทั้งหมด</span>
            <span style={{color:T.red}}>฿{fmt(waste.reduce((a,b)=>a+b.cost,0))}</span>
          </div>
          <div style={{color:T.textSm,fontSize:13,marginTop:4}}>เดือนนี้: ฿{fmt(waste.filter(w=>w.date.startsWith(mk)).reduce((a,b)=>a+b.cost,0))}</div>
        </Card>}
      </>}

      {/* Promo */}
      {tab==="promo"&&<>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{color:T.orange,fontWeight:700,fontSize:15}}>🎁 โปรโมชั่น / แลกรางวัล</div>
              <div style={{color:T.textMd,fontSize:13}}>เดือนนี้: <b>฿{fmt(promos.filter(p=>p.date.startsWith(mk)).reduce((a,b)=>a+b.amount,0))}</b></div>
            </div>
            <button onClick={()=>setShowPForm(!showPForm)} style={{...S.btn(),fontSize:13,padding:"7px 12px"}}>+ บันทึก</button>
          </div>
        </Card>
        {showPForm&&<Card style={{borderColor:T.borderOr}}>
          <div style={{color:T.orange,fontWeight:800,fontSize:16,marginBottom:12}}>🎁 บันทึกโปรโมชั่น</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ชื่อโปรโมชั่น</div>
              <input value={pForm.name} onChange={e=>setPForm(p=>({...p,name:e.target.value}))} style={S.inp} placeholder="เช่น ส่วนลด 20%, แลกของรางวัล" /></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>มูลค่า (฿)</div>
              <input type="number" value={pForm.amount} onChange={e=>setPForm(p=>({...p,amount:e.target.value}))} style={S.inp} placeholder="0" /></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>วันที่</div>
              <input type="date" value={pForm.date} max={today()} onChange={e=>setPForm(p=>({...p,date:e.target.value}))} style={S.inp} /></div>
            <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>หมายเหตุ</div>
              <input value={pForm.note} onChange={e=>setPForm(p=>({...p,note:e.target.value}))} style={S.inp} placeholder="รายละเอียดเพิ่มเติม" /></div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:12}}>
            <button onClick={()=>{
              if(!pForm.name||!pForm.amount)return;
              setPromos(p=>[{...pForm,id:Date.now(),amount:+pForm.amount},...p]);
              setPForm({name:"",date:today(),amount:"",note:""});
              setShowPForm(false);
            }} style={{...S.btn(),flex:1}}>✅ บันทึก</button>
            <button onClick={()=>setShowPForm(false)} style={S.ghost}>ยกเลิก</button>
          </div>
        </Card>}
        {promos.length===0&&<Card><div style={{color:T.textSm,textAlign:"center",padding:32}}>ยังไม่มีการบันทึกโปรโมชั่น</div></Card>}
        {[...promos].sort((a,b)=>b.date.localeCompare(a.date)).map(p=>(
          <Card key={p.id} style={{padding:"12px 16px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontWeight:700,fontSize:15}}>{p.name}</div>
                <div style={{color:T.textSm,fontSize:13}}>{p.date}{p.note?` • ${p.note}`:""}</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{color:T.orange,fontWeight:800,fontSize:18}}>฿{fmt(p.amount)}</div>
                <button onClick={()=>setPromos(prev=>prev.filter(x=>x.id!==p.id))} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:13}}>🗑 ลบ</button>
              </div>
            </div>
          </Card>
        ))}
        {promos.length>0&&<Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:15,fontWeight:700}}>
            <span>รวมโปรโมชั่นทั้งหมด</span>
            <span style={{color:T.orange}}>฿{fmt(promos.reduce((a,b)=>a+b.amount,0))}</span>
          </div>
          <div style={{color:T.textSm,fontSize:13,marginTop:4}}>เดือนนี้: ฿{fmt(promos.filter(p=>p.date.startsWith(mk)).reduce((a,b)=>a+b.amount,0))}</div>
        </Card>}
      </>}

            {/* Stock Count History */}
      {tab==="stockcheck"&&<>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{color:T.orange,fontWeight:700,fontSize:15,marginBottom:4}}>📋 ประวัติการนับสต็อค</div>
          <div style={{color:T.textMd,fontSize:13}}>แสดงผลการนับสต็อครายวันของพนักงาน</div>
        </Card>
        {checkDates.length===0&&<Card><div style={{color:T.textSm,textAlign:"center",padding:32}}>ยังไม่มีการนับสต็อค<br/><span style={{fontSize:12}}>พนักงานนับสต็อคได้ที่หน้าสต็อค → นับรายวัน</span></div></Card>}
        {checkDates.map(date=>{
          const dayChecks=checkMovements.filter(m=>m.date===date);
          return (
            <Card key={date}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <div style={{fontWeight:800,fontSize:15}}>{date}</div>
                <div style={{background:T.greenLt,color:T.green,borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{dayChecks.length} รายการ</div>
              </div>
              {dayChecks.map(m=>{
                const item=stock.find(s=>s.id==m.itemId);
                const diff=item?m.qty-item.qty:0;
                return (
                  <div key={m.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${T.bg}`,fontSize:14}}>
                    <div>
                      <div style={{fontWeight:600}}>{item?.name||"?"}</div>
                      <div style={{color:T.textSm,fontSize:12}}>นับได้ {m.qty} {item?.unit}</div>
                    </div>
                    <div style={{textAlign:"right"}}>
                      <div style={{fontWeight:700,color:diff<0?T.red:diff>0?T.green:T.green}}>{m.qty} {item?.unit}</div>
                      {diff!==0&&<div style={{fontSize:11,color:diff<0?T.red:T.green}}>{diff>0?"+":""}{diff.toFixed(1)}</div>}
                    </div>
                  </div>
                );
              })}
            </Card>
          );
        })}
      </>}
    </div>
  );
}

// ─── MyQR Page (พนักงานถือ) ───
function MyQRPage({user}){
  const dt = today();
  // Daily token = base64(staffId + date) so it changes every day
  const checkinUrl=`${window.location.origin}${window.location.pathname}?sid=${user.id}`;

  const [timeStr, setTimeStr] = useState(new Date().toLocaleTimeString("th-TH"));
  useEffect(()=>{
    const t = setInterval(()=>setTimeStr(new Date().toLocaleTimeString("th-TH")),1000);
    return ()=>clearInterval(t);
  },[]);
  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:16,padding:"8px 0"}}>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:22,fontWeight:900,color:T.text}}>QR เช็คเข้างาน</div>
        <div style={{color:T.textSm,fontSize:14,marginTop:4}}>{user.name}</div>
      </div>
      <div style={{background:"#fff",padding:16,borderRadius:20,border:`3px solid ${T.orange}`,boxShadow:"0 8px 24px rgba(249,115,22,.2)",textAlign:"center"}}>
        <QRCodeSvg value={checkinUrl} size={200} />
        <div style={{fontWeight:800,fontSize:18,marginTop:10,color:T.text}}>{user.name}</div>
        <div style={{color:T.orange,fontSize:13,fontWeight:600,marginTop:2}}>📅 {dt}</div>
      </div>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:32,fontWeight:900,color:T.orange,letterSpacing:2}}>{timeStr}</div>
        <div style={{color:T.textSm,fontSize:13,marginTop:4}}>📲 สแกนด้วยกล้องมือถือเพื่อเช็คเวลา</div>
      </div>
      <Card style={{width:"100%",maxWidth:360,background:T.orangeLt,borderColor:T.borderOr}}>
        <div style={{color:T.orange,fontWeight:700,fontSize:14,marginBottom:6}}>ℹ️ วิธีใช้</div>
        <div style={{color:T.textMd,fontSize:13,lineHeight:1.8}}>
          1. แสดง QR นี้ให้เจ้าของสแกนตอนเข้างาน<br/>
          2. QR เปลี่ยนทุกวัน ใช้ได้แค่วันนี้เท่านั้น<br/>
          3. เจ้าของสแกนที่หน้า HR → สแกนออกงาน
        </div>
      </Card>
    </div>
  );
}


// ─── HR Page ───
function HRPage({staff,setStaff,attendance,setAttendance,cf}){
  const [tab,setTab]=useState("checkin");
  const [selStaff,setSelStaff]=useState("");
  const [qrMode,setQrMode]=useState(null); // null | 'in' | 'out'
  const [salaryMonth,setSalaryMonth]=useState(today().slice(0,7));
  const [editHrId,setEditHrId]=useState(null);
  const [hrEdit,setHrEdit]=useState({});
  const [editingOtId,setEditingOtId]=useState(null); // staffId being OT-edited
  const [otVal,setOtVal]=useState("");
  const now=()=>new Date().toTimeString().slice(0,5);
  const workers=staff.filter(s=>s.role==="staff"&&s.active);

  const calcHours=(inn,out)=>{
    if(!inn||!out)return 0;
    const [ih,im]=inn.split(":").map(Number);
    const [oh,om]=out.split(":").map(Number);
    return Math.max(0,(oh*60+om-ih*60-im)/60);
  };

  // OT rule: ฿50/hr หลัง 21:00, ปัด 30 นาที = ฿25, ต่ำกว่า 30 นาที = 0
  const OT_START = "21:00";
  const OT_RATE_PER_HOUR = 50;
  const calcOT=(checkOut)=>{
    if(!checkOut)return {otMins:0,otPay:0};
    const [sh,sm]=OT_START.split(":").map(Number);
    const [oh,om]=checkOut.split(":").map(Number);
    const otMins=(oh*60+om)-(sh*60+sm);
    if(otMins<=0)return {otMins:0,otPay:0};
    // round down to nearest 30 min
    const rounded=Math.floor(otMins/30)*30;
    const otPay=rounded>0?(rounded/60)*OT_RATE_PER_HOUR:0;
    return {otMins:rounded,otPay};
  };

  const calcPay=(s,att)=>{
    const hr=s.hr||{};
    const wage=hr.wage||0; const wt=hr.wageType||"day";
    let base=0,ot=0,bonus=0;
    if(wt==="month"){ base=wage; }
    else { base=att.filter(a=>a.checkIn).length*wage; }
    // OT = ฿50/hr หลัง 21:00 (ปัดลงทุก 30 นาที)
    ot=att.reduce((sum,a)=>{
      const override=a.otOverride; // manual override
      if(override!==undefined&&override!==null)return sum+override;
      return sum+calcOT(a.checkOut).otPay;
    },0);
    // bonus from sales
    if(hr.bonusPct>0){
      const mk=salaryMonth;
      const sales=cf.filter(e=>e.flow==="in"&&e.staffId===s.id&&e.date.startsWith(mk)).reduce((a,b)=>a+b.amount,0);
      bonus=sales*(hr.bonusPct/100);
    }
    return {base,ot,bonus,total:base+ot+bonus};
  };

  const handleCheckIn=(sid)=>{
    const existing=attendance.find(a=>a.staffId===sid&&a.date===today()&&!a.checkOut);
    if(existing)return alert("เช็คอินวันนี้แล้ว");
    setAttendance(p=>[...p,{id:Date.now(),staffId:sid,date:today(),checkIn:now(),checkOut:"",note:""}]);
  };
  const handleCheckOut=(sid)=>{
    setAttendance(p=>p.map(a=>a.staffId===sid&&a.date===today()&&!a.checkOut?{...a,checkOut:now()}:a));
  };

  // Camera QR Scanner
  const [scanning,setScanning]=useState(false);
  const [scanResult,setScanResult]=useState(null);
  const [scanMode,setScanMode]=useState("in"); // "in" | "out"
  const [camErr,setCamErr]=useState("");
  const videoRef=useRef(null);
  const canvasRef=useRef(null);
  const scanRef=useRef(null);

  const stopScan=()=>{
    setScanning(false);
    if(scanRef.current){clearInterval(scanRef.current);scanRef.current=null;}
    if(videoRef.current?.srcObject){
      videoRef.current.srcObject.getTracks().forEach(t=>t.stop());
      videoRef.current.srcObject=null;
    }
  };

  const startScan=async(mode)=>{
    setCamErr(""); setScanResult(null); setScanMode(mode);
    try{
      const stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:"environment"}});
      setScanning(true);
      setTimeout(()=>{
        if(videoRef.current){
          videoRef.current.srcObject=stream;
          videoRef.current.play();
        }
        // Use BarcodeDetector if available, else jsQR fallback
        if("BarcodeDetector" in window){
          const bd=new window.BarcodeDetector({formats:["qr_code"]});
          scanRef.current=setInterval(async()=>{
            if(!videoRef.current||videoRef.current.readyState<2)return;
            try{
              const barcodes=await bd.detect(videoRef.current);
              if(barcodes.length>0){
                handleScanResult(barcodes[0].rawValue,mode);
                stopScan();
              }
            }catch{}
          },300);
        } else {
          // Canvas fallback
          scanRef.current=setInterval(()=>{
            const v=videoRef.current; const cv=canvasRef.current;
            if(!v||!cv||v.readyState<2)return;
            const ctx=cv.getContext("2d");
            cv.width=v.videoWidth; cv.height=v.videoHeight;
            ctx.drawImage(v,0,0);
            // Can't decode without jsQR - show manual fallback
          },300);
        }
      },500);
    }catch(e){
      setCamErr(e.name==="NotAllowedError"?"❌ กรุณาอนุญาตการใช้กล้อง":"❌ ไม่พบกล้อง กรุณาใช้ปุ่มด้านล่างแทน");
      setScanning(false);
    }
  };

  const handleScanResult=(val,mode)=>{
    const parts=val.split(":");
    if(parts[0]!=="TAIGUO"){ setScanResult({ok:false,msg:"QR Code ไม่ถูกต้อง"}); return; }
    const sid=parts[1];
    const s=staff.find(x=>x.id===sid);
    if(!s){ setScanResult({ok:false,msg:"ไม่พบพนักงานในระบบ"}); return; }
    if(mode==="in") handleCheckIn(sid);
    else handleCheckOut(sid);
    setScanResult({ok:true,msg:`✅ ${s.name} ${mode==="in"?"เช็คอิน":"เช็คออก"} ${now()} น.`});
  };

  useEffect(()=>()=>stopScan(),[]);

  // QR Code display
  const QRDisplay=({staffId})=>{
    const s=staff.find(x=>x.id===staffId);
    if(!s)return null;
    const url=`${window.location.origin}${window.location.pathname}?sid=${s.id}`;
    return(
      <div style={{textAlign:"center",padding:16}}>
        <div style={{background:"#fff",padding:16,borderRadius:16,border:`2px solid ${T.orange}`,boxShadow:"0 4px 16px rgba(249,115,22,.15)",display:"inline-block"}}>
          <QRCodeSvg value={url} size={180} />
          <div style={{fontWeight:800,fontSize:17,marginTop:8,textAlign:"center"}}>{s.name}</div>
          <div style={{color:T.textSm,fontSize:11,textAlign:"center",marginTop:2}}>สแกนเพื่อเช็คเข้า-ออกงาน</div>
        </div>
        <div style={{marginTop:10,background:T.bg,borderRadius:10,padding:"8px 12px",fontSize:12,color:T.textSm,wordBreak:"break-all",textAlign:"center"}}>
          🔗 {url}
        </div>
        <div style={{display:"flex",gap:8,justifyContent:"center",marginTop:8}}>
          <button onClick={()=>window.open(url,"_blank")} style={{...S.ghost,fontSize:13,padding:"7px 12px"}}>🔗 เปิด URL</button>
          <button onClick={()=>window.print()} style={{...S.ghost,fontSize:13,padding:"7px 12px"}}>🖨 พิมพ์ QR</button>
        </div>
        <div style={{background:T.yellowLt,borderRadius:10,padding:"10px 14px",marginTop:8,fontSize:13,color:T.yellow,fontWeight:600}}>
          ⚠️ ปุ่มด้านล่างสำหรับกรณีฉุกเฉินเท่านั้น (โทรศัพท์ไม่มีกล้อง ฯลฯ)
        </div>
        <div style={{marginTop:8,display:"flex",gap:8,justifyContent:"center"}}>
          <button onClick={()=>handleCheckIn(staffId)} style={{...S.btn(T.green),padding:"9px 18px",fontSize:14}}>📥 บันทึกเข้างาน</button>
          <button onClick={()=>handleCheckOut(staffId)} style={{...S.btn(T.red),padding:"9px 18px",fontSize:14}}>📤 บันทึกออกงาน</button>
        </div>
      </div>
    );
  };

  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="👥 HR & เงินเดือน" />
      <Tabs tabs={[["checkin","📲 เช็คเข้า-ออก"],["salary","💰 คำนวณเงินเดือน"],["config","⚙️ ตั้งค่า HR"]]} active={tab} onChange={setTab} />

      {/* Check-in/out */}
      {tab==="checkin"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{color:T.orange,fontWeight:700,fontSize:15,marginBottom:8}}>📲 สแกน QR เข้า-ออกงาน</div>
          <div style={{color:T.textMd,fontSize:13}}>พนักงานสแกน QR ที่หน้าร้าน → ระบบบันทึกเวลาอัตโนมัติ</div>
        </Card>

        {/* Scanner UI */}
        <Card style={{borderColor:scanning?T.green:T.border}}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:10}}>🎥 สแกน QR Code</div>
          {!scanning&&!scanResult&&<>
            <div style={{display:"flex",gap:8,marginBottom:10}}>
              <button onClick={()=>startScan("in")} style={{...S.btn(T.green),flex:1,padding:12,fontSize:15}}>📥 สแกนเข้างาน</button>
              <button onClick={()=>startScan("out")} style={{...S.btn(T.red),flex:1,padding:12,fontSize:15}}>📤 สแกนออกงาน</button>
            </div>
            {camErr&&<div style={{background:T.yellowLt,borderRadius:8,padding:"10px 14px",fontSize:13,color:T.yellow,fontWeight:600}}>{camErr}</div>}
          </>}
          {scanning&&<div style={{position:"relative"}}>
            <video ref={videoRef} style={{width:"100%",borderRadius:10,background:"#000"}} playsInline muted />
            <canvas ref={canvasRef} style={{display:"none"}} />
            <div style={{position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)",width:160,height:160,border:"3px solid "+T.orange,borderRadius:12,pointerEvents:"none"}} />
            <button onClick={stopScan} style={{...S.btn(T.red),width:"100%",marginTop:8,padding:10}}>❌ หยุดสแกน</button>
            <div style={{textAlign:"center",color:T.textSm,fontSize:13,marginTop:6}}>
              จับ QR Code ไว้ในกรอบสีส้ม — {scanMode==="in"?"กำลังเช็คอิน":"กำลังเช็คออก"}
            </div>
          </div>}
          {scanResult&&<div style={{background:scanResult.ok?T.greenLt:T.redLt,borderRadius:10,padding:"14px 16px",textAlign:"center"}}>
            <div style={{fontSize:18,fontWeight:800,color:scanResult.ok?T.green:T.red}}>{scanResult.msg}</div>
            <button onClick={()=>setScanResult(null)} style={{...S.ghost,marginTop:10,fontSize:13}}>สแกนต่อ</button>
          </div>}
        </Card>

        {/* Manual - Staff selector */}
        <Card>
          <div style={{fontWeight:700,fontSize:15,marginBottom:8}}>🖨 พิมพ์ QR ติดหน้าร้าน</div>
          <div style={{color:T.textSm,fontSize:13,marginBottom:10}}>เลือกพนักงาน → พิมพ์ QR ติดป้ายชื่อ → วางหน้าร้าน<br/>พนักงานสแกนด้วยมือถือตัวเองได้เลย</div>
          <select value={selStaff} onChange={e=>setSelStaff(e.target.value)} style={{...S.inp,fontSize:16,height:48}}>
            <option value="">— เลือกพนักงานเพื่อดู QR —</option>
            {workers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Card>

        {selStaff&&<QRDisplay staffId={selStaff} />}

        {/* Today attendance */}
        <Card>
          <div style={{fontWeight:700,fontSize:15,marginBottom:10}}>📋 การเข้างานวันนี้ — {today()}</div>
          {workers.map(s=>{
            const att=attendance.find(a=>a.staffId===s.id&&a.date===today());
            const hr=s.hr||{};
            const hours=att?calcHours(att.checkIn,att.checkOut||now()):0;
            const {otMins,otPay}=att?.checkOut?calcOT(att.checkOut):{otMins:0,otPay:0};
            const otOverride=att?.otOverride;
            const finalOtPay=otOverride!==undefined&&otOverride!==null?otOverride:otPay;
            const editingOt = editingOtId===s.id;
            return(
              <div key={s.id} style={{padding:"12px 0",borderBottom:`1px solid ${T.bg}`}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <div style={{fontWeight:600,fontSize:15}}>{s.name}</div>
                    <div style={{color:T.textSm,fontSize:12}}>
                      {att?.checkIn?`เข้า ${att.checkIn}`:"ยังไม่เช็คอิน"}
                      {att?.checkOut?` • ออก ${att.checkOut}`:att?.checkIn?" • ยังอยู่":""}
                    </div>
                    {att?.checkOut&&otMins>0&&(
                      <div style={{fontSize:12,marginTop:3}}>
                        <span style={{color:T.orange,fontWeight:600}}>⏰ OT {otMins} นาที = </span>
                        {editingOt?(
                          <span style={{display:"inline-flex",alignItems:"center",gap:4}}>
                            <span style={{color:T.textSm}}>฿</span>
                            <input type="number" value={otVal} onChange={e=>setOtVal(e.target.value)}
                              style={{width:70,padding:"2px 6px",border:`1px solid ${T.orange}`,borderRadius:6,fontSize:12}}
                              autoFocus />
                            <button onClick={()=>{setAttendance(p=>p.map(a=>a.id===att.id?{...a,otOverride:+otVal||0}:a));setEditingOtId(null);}}
                              style={{...S.btn(T.green),padding:"2px 8px",fontSize:11}}>✓</button>
                            <button onClick={()=>setEditingOtId(null)} style={{...S.ghost,padding:"2px 6px",fontSize:11}}>✕</button>
                          </span>
                        ):(
                          <span>
                            <span style={{color:T.green,fontWeight:700}}>฿{fmt(finalOtPay)}</span>
                            {otOverride!==undefined&&otOverride!==null&&<span style={{color:T.textXs,fontSize:11}}> (แก้ไขแล้ว)</span>}
                            <button onClick={()=>{setOtVal(String(finalOtPay));setEditingOtId(s.id);}} style={{background:"none",border:"none",color:T.orange,cursor:"pointer",fontSize:11,marginLeft:4}}>✏️</button>
                          </span>
                        )}
                      </div>
                    )}
                    {att?.checkOut&&otMins===0&&att.checkOut<="21:00"&&(
                      <div style={{fontSize:11,color:T.textXs,marginTop:2}}>ออกก่อน 21:00 — ไม่มี OT</div>
                    )}
                  </div>
                  <div style={{display:"flex",gap:6,flexShrink:0}}>
                    {!att?.checkIn&&<button onClick={()=>handleCheckIn(s.id)} style={{...S.btn(T.green),fontSize:12,padding:"5px 10px"}}>เข้า</button>}
                    {att?.checkIn&&!att?.checkOut&&<button onClick={()=>handleCheckOut(s.id)} style={{...S.btn(T.red),fontSize:12,padding:"5px 10px"}}>ออก</button>}
                    {att?.checkOut&&<span style={{color:T.green,fontSize:12,fontWeight:600}}>✅ {hours.toFixed(1)} ชม.</span>}
                  </div>
                </div>
              </div>
            );
          })}
        </Card>
      </div>}

      {/* Salary */}
      {tab==="salary"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Card style={{padding:"12px 16px"}}>
          <div style={{color:T.textSm,fontSize:13,marginBottom:6}}>เลือกเดือน</div>
          <input type="month" value={salaryMonth} onChange={e=>setSalaryMonth(e.target.value)} style={{...S.inp,fontSize:16}} />
        </Card>
        {workers.map(s=>{
          const att=attendance.filter(a=>a.staffId===s.id&&a.date.startsWith(salaryMonth)&&a.checkIn);
          const {base,ot,bonus,total}=calcPay(s,att);
          const hr=s.hr||{};
          const totalDays=att.filter(a=>a.checkIn).length;
          const totalH=att.reduce((a,b)=>a+calcHours(b.checkIn,b.checkOut||""),0);
          const totalOtPay=ot;
          const otDays=att.filter(a=>a.checkOut&&(a.otOverride>0||(a.otOverride===undefined&&calcOT(a.checkOut).otMins>0))).length;
          return(
            <Card key={s.id} style={{borderColor:T.orange}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                <div style={{fontWeight:800,fontSize:17}}>{s.name}</div>
                <div style={{color:T.orange,fontWeight:900,fontSize:22}}>฿{fmt(total)}</div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:10}}>
                {[["วันทำงาน",`${totalDays} วัน`,T.blue],["ชม.รวม",`${totalH.toFixed(1)} ชม.`,T.textMd],["OT วัน",`${otDays} วัน`,T.orange]].map(([l,v,col])=>(
                  <div key={l} style={{background:T.bg,borderRadius:8,padding:"8px",textAlign:"center"}}>
                    <div style={{color:T.textXs,fontSize:11}}>{l}</div>
                    <div style={{color:col,fontWeight:700,fontSize:14}}>{v}</div>
                  </div>
                ))}
              </div>
              <div style={{borderTop:`1px solid ${T.bg}`,paddingTop:10}}>
                {[["💼 เงินเดือน/ค่าจ้างฐาน",base,T.text],["⏰ ค่าล่วงเวลา (OT)",ot,T.orange],["🎁 โบนัสยอดขาย",bonus,T.green]].map(([l,v,col])=>v>0&&(
                  <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",fontSize:14}}>
                    <span style={{color:T.textMd}}>{l}</span>
                    <span style={{color:col,fontWeight:600}}>฿{fmt(v)}</span>
                  </div>
                ))}
                <div style={{display:"flex",justifyContent:"space-between",paddingTop:8,marginTop:4,borderTop:`1px solid ${T.bg}`,fontWeight:800,fontSize:16}}>
                  <span>รวมจ่าย</span><span style={{color:T.orange}}>฿{fmt(total)}</span>
                </div>
              </div>
              {hr.wageType==="day"&&<div style={{color:T.textXs,fontSize:11,marginTop:6}}>{hr.wageType==="day"?`฿${hr.wage}/วัน`:""} • OT ×{hr.otRate||1.5} • โบนัส {hr.bonusPct||0}% ของยอดขาย</div>}
            </Card>
          );
        })}
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:16}}>
            <span>รวมจ่ายทั้งหมด</span>
            <span style={{color:T.orange}}>฿{fmt(workers.reduce((sum,s)=>{const att=attendance.filter(a=>a.staffId===s.id&&a.date.startsWith(salaryMonth)&&a.checkIn);return sum+calcPay(s,att).total;},0))}</span>
          </div>
        </Card>
      </div>}

      {/* HR Config */}
      {tab==="config"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{color:T.orange,fontWeight:700,fontSize:15}}>⚙️ ตั้งค่าอัตราค่าจ้าง</div>
          <div style={{color:T.textMd,fontSize:13}}>กำหนดเวลาเข้า-ออก ค่าจ้าง และโบนัสแต่ละคน</div>
        </Card>
        {workers.map(s=>{
          const hr=s.hr||{wage:500,wageType:"day",otRate:1.5,shiftStart:"09:00",shiftEnd:"18:00",bonusPct:0};
          const isEdit=editHrId===s.id;
          const ed=isEdit?hrEdit:hr;
          return(
            <Card key={s.id}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:isEdit?12:0}}>
                <div style={{fontWeight:700,fontSize:16}}>{s.name}</div>
                {!isEdit&&<button onClick={()=>{setEditHrId(s.id);setHrEdit({...hr});}} style={{...S.ghost,fontSize:13,padding:"5px 12px"}}>✏️ แก้ไข</button>}
              </div>
              {!isEdit&&<div style={{color:T.textSm,fontSize:13,marginTop:4}}>
                {hr.wageType==="day"?`฿${hr.wage}/วัน`:`฿${hr.wage}/เดือน`} • เข้า {hr.shiftStart} ออก {hr.shiftEnd} • OT ×{hr.otRate} • โบนัส {hr.bonusPct}%
              </div>}
              {isEdit&&<>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>ประเภทค่าจ้าง</div>
                    <select value={ed.wageType||"day"} onChange={e=>setHrEdit(p=>({...p,wageType:e.target.value}))} style={{...S.inp,height:40}}>
                      <option value="day">รายวัน</option>
                      <option value="month">รายเดือน</option>
                    </select>
                  </div>
                  <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>{ed.wageType==="month"?"฿/เดือน":"฿/วัน"}</div>
                    <input type="number" value={ed.wage||0} onChange={e=>setHrEdit(p=>({...p,wage:+e.target.value}))} style={S.inp} /></div>
                  <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>เวลาเข้างาน</div>
                    <input type="time" value={ed.shiftStart||"09:00"} onChange={e=>setHrEdit(p=>({...p,shiftStart:e.target.value}))} style={S.inp} /></div>
                  <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>เวลาออกงาน</div>
                    <input type="time" value={ed.shiftEnd||"18:00"} onChange={e=>setHrEdit(p=>({...p,shiftEnd:e.target.value}))} style={S.inp} /></div>
                  <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>อัตรา OT (×เท่า)</div>
                    <input type="number" step="0.5" value={ed.otRate||1.5} onChange={e=>setHrEdit(p=>({...p,otRate:+e.target.value}))} style={S.inp} /></div>
                  <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>โบนัส % ของยอดขาย</div>
                    <input type="number" step="0.5" value={ed.bonusPct||0} onChange={e=>setHrEdit(p=>({...p,bonusPct:+e.target.value}))} style={S.inp} /></div>
                </div>
                <div style={{display:"flex",gap:8,marginTop:12}}>
                  <button onClick={()=>{setStaff(p=>p.map(x=>x.id===s.id?{...x,hr:hrEdit}:x));setEditHrId(null);}} style={{...S.btn(),flex:1}}>บันทึก</button>
                  <button onClick={()=>setEditHrId(null)} style={S.ghost}>ยกเลิก</button>
                </div>
              </>}
            </Card>
          );
        })}
      </div>}
    </div>
  );
}


// ─── App ───
export default function App(){
  const [user,setUser]=useState(null);
  const [page,setPage]=useState("dashboard");
  const [stock,setStock]=useState(INIT_STOCK);
  const [cf,setCF]=useState([]);
  const [movements,setMovements]=useState([]);
  const [staff,setStaff]=useState(INIT_STAFF);
  const [suppliers,setSuppliers]=useState(INIT_SUPS);
  const [fixedCosts,setFixedCosts]=useState(INIT_FIXED);
  const [lineToken,setLineToken]=useState("");
  const [waste,setWaste]=useState(INIT_WASTE);
  const [attendance,setAttendance]=useState([]); // {id,staffId,date,checkIn,checkOut,note}
  const [promos,setPromos]=useState(INIT_PROMO);
  const [dbReady,setDbReady]=useState(false);
  const [dbLoading,setDbLoading]=useState(true);

  useEffect(()=>{
    async function load(){
      setDbLoading(true);
      try{
        const [cfData,stockData,mvData]=await Promise.all([db.getCF(),db.getStock(),db.getMvs()]);
        if(cfData?.length>0)setCF(cfData.map(r=>({id:r.id,date:r.date,flow:r.flow,cat:r.cat,itemName:r.item_name||"",amount:r.amount,method:r.method,note:r.note||"",branch:r.branch||"main",staffId:r.staff_id||"owner"})));
        if(stockData?.length>0)setStock(stockData.map(r=>({id:r.id,name:r.name,unit:r.unit,qty:r.qty,minQty:r.min_qty,dailyUse:r.daily_use,supplierId:r.supplier_id||1,costHistory:r.cost_history||[]})));
        if(mvData?.length>0)setMovements(mvData.map(r=>({id:r.id,itemId:r.item_id,type:r.type,qty:r.qty,unitCost:r.unit_cost||0,date:r.date,staffId:r.staff_id||"",note:r.note||"",branch:r.branch||"main"})));
        setDbReady(true);
      }catch{setDbReady(false);}
      setDbLoading(false);
    }
    load();
  },[]);

  const saveStock=useCallback(async ns=>{setStock(ns);if(dbReady)db.upsertStock(ns.map(s=>({id:s.id,name:s.name,unit:s.unit,qty:s.qty,min_qty:s.minQty,daily_use:s.dailyUse,supplier_id:s.supplierId||1,cost_history:s.costHistory||[]}))).catch(()=>{});},[dbReady]);

  // Detect QR check-in URL (?sid=staffId)
  const urlSid = new URLSearchParams(window.location.search).get("sid");
  if(urlSid) return <CheckinPage staff={staff} onCheckin={(sid,type,time,date)=>{
    if(type==="in"){
      setAttendance(p=>[...p,{id:Date.now(),staffId:sid,date,checkIn:time,checkOut:"",note:""}]);
    } else {
      setAttendance(p=>p.map(a=>a.staffId===sid&&a.date===date&&!a.checkOut?{...a,checkOut:time}:a));
    }
  }} />;

  if(!user)return <LoginPage staff={staff} onLogin={u=>{setUser(u);setPage("dashboard");}} />;

  if(user.id==="emergency")return(
    <div style={{minHeight:"100vh",background:T.bg,fontFamily:F,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,gap:16}}>
      <div style={{width:80,height:80,borderRadius:20,background:T.red,display:"flex",alignItems:"center",justifyContent:"center",fontSize:44}}>🚨</div>
      <div style={{color:T.red,fontWeight:900,fontSize:22}}>Emergency Mode</div>
      <div style={{display:"flex",flexDirection:"column",gap:10,width:"100%",maxWidth:320}}>
        <button onClick={()=>{if(window.confirm("ล้าง Cash Flow ทั้งหมด?")){setCF([]);if(dbReady)db.clearCF().catch(()=>{});alert("เรียบร้อย");}}} style={{...S.btn(T.red),padding:14,fontSize:16,width:"100%"}}>🗑 ล้าง Cash Flow</button>
        <button onClick={()=>{if(window.confirm("ล้างสต็อค + ประวัติ?")){setStock(INIT_STOCK);setMovements([]);if(dbReady)db.clearMvs().catch(()=>{});alert("เรียบร้อย");}}} style={{...S.btn(T.orange),padding:14,fontSize:16,width:"100%"}}>🗑 ล้างสต็อค + ประวัติ</button>
        <button onClick={()=>{if(window.confirm("⚠️ ล้างทั้งหมด?")){setCF([]);setStock(INIT_STOCK);setMovements([]);if(dbReady){db.clearCF().catch(()=>{});db.clearMvs().catch(()=>{});}alert("เรียบร้อย");}}} style={{...S.btn(T.red),padding:14,fontSize:16,width:"100%",border:"2px solid #7f1d1d"}}>⚠️ ล้างทั้งหมด</button>
        <button onClick={()=>setUser(null)} style={{...S.ghost,padding:14,fontSize:16,width:"100%"}}>← ออก</button>
      </div>
    </div>
  );

  if(dbLoading)return(
    <div style={{minHeight:"100vh",background:T.bg,fontFamily:F,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16}}>
      <div style={{width:80,height:80,borderRadius:20,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:44}}>🫕</div>
      <div style={{color:T.orange,fontWeight:800,fontSize:20}}>กำลังโหลดข้อมูล...</div>
    </div>
  );

  const p=user.perms; const isOwner=user.role==="owner";
  const nav=isOwner?[
    {id:"dashboard",icon:"🏠",label:"หลัก"},{id:"cashflow",icon:"💵",label:"Cash Flow"},
    {id:"stock",icon:"📦",label:"สต็อค"},{id:"purchase",icon:"🛒",label:"สั่งซื้อ"},
    {id:"report",icon:"📊",label:"รายงาน"},{id:"hr",icon:"👥",label:"HR"},
    {id:"settings",icon:"⚙️",label:"ตั้งค่า"},
  ]:[
    {id:"dashboard",icon:"🏠",label:"หลัก"},
    {id:"myqr",     icon:"📱",label:"QR ของฉัน"},
    ...(p.cashflow ?[{id:"cashflow",  icon:"💵",label:"Cash Flow"}]:[]),
    ...(p.stock    ?[{id:"staffstock",icon:"📦",label:"สต็อค"}]:[]),
    ...(p.purchase ?[{id:"purchase",  icon:"🛒",label:"สั่งซื้อ"}]:[]),
    ...(p.report   ?[{id:"report",   icon:"📊",label:"รายงาน"}]:[]),
    ...(p.admin    ?[{id:"settings", icon:"⚙️",label:"ตั้งค่า"}]:[]),
  ];

  const pages={
    myqr:      <MyQRPage user={user} />,
    dashboard: <DashboardPage cf={cf} stock={stock} user={user} fixedCosts={fixedCosts} waste={waste} promos={promos} setPage={setPage} />,
    cashflow:  <CashflowPage cf={cf} setCF={setCF} user={user} dbReady={dbReady} />,
    stock:     <StockPage stock={stock} setStock={saveStock} movements={movements} setMovements={setMovements} user={user} suppliers={suppliers} />,
    staffstock:<StaffStockPage stock={stock} setStock={saveStock} movements={movements} setMovements={setMovements} user={user} />,
    purchase:  <PurchasePage stock={stock} suppliers={suppliers} lineToken={lineToken} />,
    report:    <ReportPage cf={cf} stock={stock} movements={movements} user={user} fixedCosts={fixedCosts} waste={waste} setWaste={setWaste} promos={promos} setPromos={setPromos} />,
    hr:        <HRPage staff={staff} setStaff={setStaff} attendance={attendance} setAttendance={setAttendance} cf={cf} />,
    settings:  <SettingsPage staff={staff} setStaff={setStaff} lineToken={lineToken} setLineToken={setLineToken} fixedCosts={fixedCosts} setFixedCosts={setFixedCosts} suppliers={suppliers} setSuppliers={setSuppliers} attendance={attendance} setAttendance={setAttendance} />,
  };

  return(
    <div style={{minHeight:"100vh",background:T.bg,fontFamily:F,fontSize:16}}>
      <div style={{background:"#fff",borderBottom:`1px solid ${T.border}`,padding:"11px 18px",display:"flex",alignItems:"center",gap:12,position:"sticky",top:0,zIndex:50,boxShadow:"0 1px 3px rgba(0,0,0,.08)"}}>
        <div style={{width:36,height:36,borderRadius:10,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20}}>🫕</div>
        <div><div style={{color:T.orange,fontWeight:900,fontSize:15}}>ไท่กั๋วหม่าล่า</div><div style={{color:T.textXs,fontSize:11}}>{user.name} • {isOwner?"👑":"👷"}</div></div>
        {dbReady&&<div style={{marginLeft:"auto",background:T.greenLt,border:`1px solid ${T.green}33`,borderRadius:6,padding:"3px 8px",fontSize:11,color:T.green,fontWeight:600}}>● DB</div>}
        <button onClick={()=>setUser(null)} style={{...S.ghost,fontSize:13,padding:"5px 11px",marginLeft:dbReady?4:"auto"}}>ออก</button>
      </div>
      <div style={{padding:"18px 14px 95px",maxWidth:900,margin:"0 auto"}}>{pages[page]||pages.dashboard}</div>
      <div style={{position:"fixed",bottom:0,left:0,right:0,background:"#fff",borderTop:`1px solid ${T.border}`,display:"flex",justifyContent:"space-around",padding:"9px 0 15px",zIndex:100,boxShadow:"0 -2px 8px rgba(0,0,0,.06)"}}>
        {nav.map(n=>(
          <button key={n.id} onClick={()=>setPage(n.id)} style={{background:"none",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2,padding:"2px 8px",minWidth:46}}>
            <span style={{fontSize:22}}>{n.icon}</span>
            <span style={{fontSize:10,color:page===n.id?T.orange:T.textXs,fontWeight:page===n.id?800:400}}>{n.label}</span>
            {page===n.id&&<div style={{width:18,height:3,borderRadius:2,background:T.orange}}/>}
          </button>
        ))}
      </div>
    </div>
  );
}
