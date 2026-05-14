import{useState,useEffect,useCallback,useRef}from"react";
import*as XLSX from"xlsx";

const SB="https://klmowpluuvjmbvvmqzep.supabase.co";
const SK="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtsbW93cGx1dXZqbWJ2dm1xemVwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzgzMTIzMTMsImV4cCI6MjA5Mzg4ODMxM30.aXQz6WBqE8US5_-ij6GvvY0XaCykMag8x6W2a6uAwMU";
async function sb(path,o={}){try{const r=await fetch(`${SB}/rest/v1/${path}`,{...o,headers:{apikey:SK,Authorization:`Bearer ${SK}`,"Content-Type":"application/json",Prefer:o.prefer||"return=representation",...o.headers}});if(!r.ok)return null;const t=await r.text();return t?JSON.parse(t):[];}catch{return null;}}
const db={getCF:()=>sb("cashflow?order=date.desc,id.desc"),addCF:r=>sb("cashflow",{method:"POST",body:JSON.stringify(r)}),delCF:id=>sb(`cashflow?id=eq.${id}`,{method:"DELETE",prefer:""}),clearCF:()=>sb("cashflow?id=gt.0",{method:"DELETE",prefer:""}),getStock:()=>sb("stock?order=name.asc"),upsertStock:rows=>sb("stock",{method:"POST",body:JSON.stringify(rows),headers:{Prefer:"resolution=merge-duplicates,return=representation"}}),getMvs:()=>sb("movements?order=date.desc,id.desc"),addMv:r=>sb("movements",{method:"POST",body:JSON.stringify(r)}),clearMvs:()=>sb("movements?id=gt.0",{method:"DELETE",prefer:""}),getAtt:()=>sb("attendance?order=date.desc"),addAtt:r=>sb("attendance",{method:"POST",body:JSON.stringify(r)}),patchAtt:(id,d)=>sb(`attendance?id=eq.${id}`,{method:"PATCH",body:JSON.stringify(d)})};

const T={bg:"#f4f4f5",card:"#fff",border:"#e4e4e7",text:"#18181b",textMd:"#52525b",textSm:"#71717a",textXs:"#a1a1aa",orange:"#f97316",orangeDk:"#ea580c",orangeLt:"#fff7ed",borderOr:"#fed7aa",green:"#16a34a",greenLt:"#f0fdf4",red:"#dc2626",redLt:"#fef2f2",yellow:"#d97706",yellowLt:"#fffbeb",blue:"#2563eb",blueLt:"#eff6ff"};
const F="'Noto Sans Thai',sans-serif";
const S={inp:{width:"100%",padding:"10px 12px",border:`1px solid #e4e4e7`,borderRadius:10,fontSize:15,fontFamily:F,background:"#fff",outline:"none",boxSizing:"border-box",color:"#18181b"},btn:(bg="#f97316")=>({background:bg,color:"#fff",border:"none",borderRadius:10,padding:"9px 18px",cursor:"pointer",fontWeight:700,fontSize:15,fontFamily:F}),ghost:{background:"transparent",border:"1px solid #e4e4e7",borderRadius:10,padding:"9px 14px",cursor:"pointer",fontSize:14,color:"#52525b",fontFamily:F},card:{background:"#fff",border:"1px solid #e4e4e7",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,.08)"}};

const today=()=>new Date().toISOString().split("T")[0];
const fmt=n=>Math.round(n).toLocaleString("th-TH");
const IN_CATS=["ยอดขาย dine-in","ยอดขาย delivery","เงินโอนเข้า","อื่นๆ"];
const OUT_CATS=["วัตถุดิบ/ผัก","หมู/เนื้อ/ทะเล","ค่าเช่า","ค่าพนักงาน","ค่าไฟ/น้ำ","บรรจุภัณฑ์","อื่นๆ"];
const PAY=["เงินสด","โอนธนาคาร","QR Code","GrabFood","LINE MAN"];
const stockSt=s=>s.qty<=0?"out":s.qty<s.minQty*.5?"critical":s.qty<s.minQty?"low":"ok";
const ST_C={ok:T.green,low:T.yellow,critical:T.red,out:T.red};
const ST_L={ok:"ปกติ",low:"ใกล้หมด",critical:"น้อยมาก",out:"หมด"};
const wac=s=>{const h=s.costHistory||[];const q=h.reduce((a,b)=>a+b.qty,0);return q>0?h.reduce((a,b)=>a+b.total,0)/q:0;};
const exportXlsx=(rows,sheet,file)=>{const ws=XLSX.utils.json_to_sheet(rows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheet);XLSX.writeFile(wb,`${file}_${today()}.xlsx`);};
const readXlsx=file=>new Promise((res,rej)=>{const r=new FileReader();r.onload=ev=>{try{const wb=XLSX.read(ev.target.result,{type:"array"});res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));}catch(e){rej(e);}};r.onerror=rej;r.readAsArrayBuffer(file);});
const calcOT=(out,rate=50)=>{if(!out)return{m:0,p:0};const[sh,sm]="21:00".split(":").map(Number);const[oh,om]=out.split(":").map(Number);const raw=(oh*60+om)-(sh*60+sm);if(raw<=0)return{m:0,p:0};const rounded=Math.floor(raw/30)*30;return{m:rounded,p:rounded>0?(rounded/60)*rate:0};};

const INIT_STOCK=[{id:1,name:"หมูสามชั้น",unit:"kg",qty:15,minQty:5,dailyUse:3,supplierId:2,costHistory:[]},{id:2,name:"กุ้งแวนนาไม",unit:"kg",qty:3,minQty:4,dailyUse:2,supplierId:3,costHistory:[]},{id:3,name:"เต้าหู้ขาว",unit:"kg",qty:10,minQty:5,dailyUse:2,supplierId:1,costHistory:[]},{id:4,name:"ผักกาดขาว",unit:"kg",qty:2,minQty:3,dailyUse:2,supplierId:1,costHistory:[]},{id:5,name:"น้ำซุปมาล่า",unit:"ถุง",qty:20,minQty:8,dailyUse:4,supplierId:4,costHistory:[]},{id:6,name:"วุ้นเส้น",unit:"kg",qty:5,minQty:3,dailyUse:1,supplierId:1,costHistory:[]}];
const INIT_SUPS=[{id:1,name:"ตลาดสดนครชัย",type:"ผัก",phone:"081-234-5678",active:true},{id:2,name:"ฟาร์มหมูสยาม",type:"หมู",phone:"082-345-6789",active:true},{id:3,name:"อาหารทะเลสด",type:"ทะเล",phone:"083-456-7890",active:true},{id:4,name:"ซอสมาล่าพรีเมียม",type:"ซอส",phone:"084-567-8901",active:true}];
const INIT_STAFF=[
  {id:"owner",name:"DR.Fresh (เจ้าของ)",pin:"1234",role:"owner",active:true,perms:{cashflow:true,stock:true,purchase:true,report:true,admin:true,viewPrice:true,viewFinance:true},hr:{wage:0,wageType:"day",otPerHour:50,shiftStart:"09:00",shiftEnd:"18:00",bonusPct:0}},
  {id:"s1",name:"มิ้ว",pin:"1111",role:"staff",active:true,perms:{cashflow:true,stock:true,purchase:false,report:false,admin:false,viewPrice:false,viewFinance:false},hr:{wage:350,wageType:"day",otPerHour:50,shiftStart:"09:00",shiftEnd:"18:00",bonusPct:0}},
  {id:"s2",name:"ปาล์ม",pin:"2222",role:"staff",active:true,perms:{cashflow:true,stock:true,purchase:false,report:false,admin:false,viewPrice:false,viewFinance:false},hr:{wage:350,wageType:"day",otPerHour:50,shiftStart:"09:00",shiftEnd:"18:00",bonusPct:0}},
  {id:"s3",name:"เจ",pin:"3333",role:"staff",active:true,perms:{cashflow:false,stock:true,purchase:false,report:false,admin:false,viewPrice:false,viewFinance:false},hr:{wage:350,wageType:"day",otPerHour:50,shiftStart:"09:00",shiftEnd:"18:00",bonusPct:0}},
  {id:"emergency",name:"Emergency",pin:"0000",role:"owner",active:true,perms:{cashflow:false,stock:false,purchase:false,report:false,admin:true,viewPrice:false,viewFinance:false},hr:{}}
];
const INIT_FIXED=[{name:"ค่าเช่า",amount:4500},{name:"ค่าพนักงาน",amount:35000},{name:"ค่าไฟ",amount:8000},{name:"ค่าน้ำ",amount:1000},{name:"อื่นๆ",amount:1000}];

function Card({children,style={},onClick}){return <div style={{...S.card,...style}} onClick={onClick}>{children}</div>;}
function Badge({status}){const c=ST_C[status]||T.green;return <span style={{background:c+"18",color:c,borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{ST_L[status]||"ปกติ"}</span>;}
function Tabs({tabs,active,onChange}){return <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{tabs.map(([v,l])=><button key={v} onClick={()=>onChange(v)} style={{padding:"8px 14px",borderRadius:8,border:`1px solid ${active===v?T.orange:T.border}`,background:active===v?T.orange:"transparent",color:active===v?"#fff":T.textMd,cursor:"pointer",fontSize:13,fontFamily:F,fontWeight:active===v?700:400}}>{l}</button>)}</div>;}
function Hdr({title,action}){return <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}><div style={{fontSize:20,fontWeight:900}}>{title}</div>{action}</div>;}
function IEBtn({onImport,onExport}){const ref=useRef();return <div style={{display:"flex",gap:6}}><input ref={ref} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={async e=>{if(e.target.files[0]){try{onImport(await readXlsx(e.target.files[0]));}catch{alert("ไฟล์ผิดรูปแบบ");}e.target.value="";}}} /><button onClick={()=>ref.current?.click()} style={{...S.ghost,fontSize:13,padding:"7px 10px"}}>📥</button><button onClick={onExport} style={{...S.ghost,fontSize:13,padding:"7px 10px"}}>📤</button></div>;}

function CheckinPage({staff,onCheckin}){
  const sid=new URLSearchParams(window.location.search).get("sid");
  const s=staff.find(x=>x.id===sid);
  const[done,setDone]=useState(null);
  const[loading,setLoading]=useState(false);
  const[loc,setLoc]=useState(null);
  const[locErr,setLocErr]=useState(false);
  const nowStr=()=>new Date().toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"});
  const todayStr=()=>new Date().toISOString().split("T")[0];
  const getLoc=()=>new Promise(res=>{if(!navigator.geolocation){res(null);return;}navigator.geolocation.getCurrentPosition(async p=>{const{latitude:lat,longitude:lng,accuracy:acc}=p.coords;let addr="";try{const r=await fetch(`https://nominatim.openstreetmap.org/reverse?lat=${lat}&lon=${lng}&format=json&accept-language=th`);const j=await r.json();addr=j.display_name?.split(",").slice(0,3).join(", ")||"";}catch{}const l={lat:lat.toFixed(5),lng:lng.toFixed(5),acc:Math.round(acc),addr};setLoc(l);res(l);},()=>{setLocErr(true);res(null);},{timeout:6000,enableHighAccuracy:true});});
  const stamp=async type=>{if(!s)return;setLoading(true);const t=nowStr();const d=todayStr();const gps=await getLoc();const note=gps?`📍 ${gps.addr||`${gps.lat},${gps.lng}`} ±${gps.acc}m`:"";if(type==="in"){await db.addAtt({id:Date.now(),staff_id:s.id,date:d,check_in:t,check_out:"",note,lat_in:gps?.lat||"",lng_in:gps?.lng||""});}else{const recs=await sb(`attendance?staff_id=eq.${s.id}&date=eq.${d}&check_out=eq.`);if(recs?.length>0)await db.patchAtt(recs[0].id,{check_out:t,lat_out:gps?.lat||"",lng_out:gps?.lng||""});}setDone({type,time:t,name:s.name,gps});setLoading(false);onCheckin&&onCheckin(s.id,type,t,d,gps);};
  if(!sid||!s)return <div style={{minHeight:"100vh",background:"#f4f4f5",fontFamily:F,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{background:"#fff",borderRadius:16,padding:32,textAlign:"center"}}><div style={{fontSize:40}}>❌</div><div style={{fontWeight:700,fontSize:18,color:T.red}}>QR Code ไม่ถูกต้อง</div></div></div>;
  if(done)return(
    <div style={{minHeight:"100vh",background:done.type==="in"?"#f0fdf4":"#fef2f2",fontFamily:F,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,textAlign:"center",maxWidth:360,width:"100%",boxShadow:"0 8px 24px rgba(0,0,0,.12)"}}>
        <div style={{fontSize:60,marginBottom:16}}>{done.type==="in"?"✅":"👋"}</div>
        <div style={{fontWeight:900,fontSize:24,color:done.type==="in"?T.green:T.red}}>{done.type==="in"?"เช็คอินสำเร็จ!":"เช็คออกสำเร็จ!"}</div>
        <div style={{fontWeight:700,fontSize:20,marginTop:10}}>{done.name}</div>
        <div style={{fontSize:36,fontWeight:900,color:T.orange,marginTop:6}}>{done.time}</div>
        {done.gps&&<div style={{background:"#f4f4f5",borderRadius:10,padding:"10px 14px",marginTop:12,fontSize:12,color:T.textSm,textAlign:"left"}}>
          <div style={{fontWeight:600,marginBottom:2}}>📍 ตำแหน่ง</div>
          {done.gps.addr&&<div>{done.gps.addr}</div>}
          <div style={{color:T.textXs}}>±{done.gps.acc}m</div>
          <a href={`https://maps.google.com/?q=${done.gps.lat},${done.gps.lng}`} target="_blank" rel="noreferrer" style={{color:T.blue,display:"block",marginTop:4}}>🗺 ดูบน Google Maps</a>
        </div>}
        <div style={{color:T.textSm,fontSize:13,marginTop:16}}>บันทึกเรียบร้อย สามารถปิดได้เลย</div>
      </div>
    </div>
  );
  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#fff7ed,#fff)",fontFamily:F,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,textAlign:"center",maxWidth:360,width:"100%",boxShadow:"0 8px 24px rgba(0,0,0,.12)"}}>
        <div style={{width:70,height:70,borderRadius:18,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:38,margin:"0 auto 12px"}}>🫕</div>
        <div style={{color:T.orange,fontWeight:900,fontSize:18}}>ไท่กั๋วหม่าล่า</div>
        <div style={{fontWeight:800,fontSize:24,marginTop:6}}>{s.name}</div>
        <div style={{color:T.textSm,fontSize:14,marginBottom:20}}>{new Date().toLocaleDateString("th-TH",{weekday:"long",day:"numeric",month:"long"})}</div>
        <div style={{fontSize:38,fontWeight:900,color:T.text,marginBottom:24,letterSpacing:2}}>{new Date().toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"})}</div>
        {locErr&&<div style={{background:T.yellowLt,borderRadius:8,padding:"8px 12px",fontSize:12,color:T.yellow,marginBottom:12}}>⚠️ ไม่สามารถดึง GPS ได้ — บันทึกเวลาได้ปกติ</div>}
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <button onClick={()=>stamp("in")} disabled={loading} style={{background:loading?"#a1a1aa":T.green,color:"#fff",border:"none",borderRadius:14,padding:18,fontSize:20,fontWeight:800,cursor:"pointer",fontFamily:F}}>{loading?"⏳ กำลังบันทึก...":"📥 เข้างาน"}</button>
          <button onClick={()=>stamp("out")} disabled={loading} style={{background:loading?"#a1a1aa":T.red,color:"#fff",border:"none",borderRadius:14,padding:18,fontSize:20,fontWeight:800,cursor:"pointer",fontFamily:F}}>{loading?"⏳ กำลังบันทึก...":"📤 ออกงาน"}</button>
        </div>
        <div style={{color:T.textXs,fontSize:12,marginTop:16}}>ระบบบันทึกเวลา+ตำแหน่งส่งเจ้าของอัตโนมัติ</div>
      </div>
    </div>
  );
}

function LoginPage({staff,onLogin}){
  const[pin,setPin]=useState("");const[err,setErr]=useState(false);
  const go=()=>{const u=staff.find(s=>s.pin===pin&&s.active);if(u)onLogin(u);else{setErr(true);setTimeout(()=>setErr(false),1500);}};
  return(
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

function DashboardPage({cf,stock,user,fixedCosts,waste,promos,setPage}){
  const ts=today(),mk=ts.slice(0,7);
  const yd=new Date();yd.setDate(yd.getDate()-1);const ys=yd.toISOString().split("T")[0];
  const myCF=user.role==="owner"?cf:cf.filter(e=>e.staffId===user.id||e.branch===user.franchiseId);
  const todayIn=myCF.filter(e=>e.date===ts&&e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const todayOut=myCF.filter(e=>e.date===ts&&e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const ydIn=myCF.filter(e=>e.date===ys&&e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mIn=myCF.filter(e=>e.date.startsWith(mk)&&e.flow==="in").reduce((a,b)=>a+b.amount,0);
  const mOut=myCF.filter(e=>e.date.startsWith(mk)&&e.flow==="out").reduce((a,b)=>a+b.amount,0);
  const totalFixed=fixedCosts.reduce((a,b)=>a+b.amount,0);
  const netP=mIn-mOut-totalFixed;
  const alerts=stock.filter(s=>["out","critical","low"].includes(stockSt(s)));
  const showFin=user.role==="owner"||user.perms?.viewFinance;
  const wCost=(waste||[]).filter(w=>w.date.startsWith(mk)).reduce((a,b)=>a+b.cost,0);
  const highCost=stock.filter(s=>{const h=s.costHistory||[];if(h.length<2)return false;const last=h[h.length-1]?.unitCost||0;return last>wac(s)*1.15;});
  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div><div style={{fontSize:22,fontWeight:900}}>🏠 ภาพรวมร้าน</div><div style={{color:T.textSm,fontSize:14}}>{ts}</div></div>
      {showFin?(
        <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12}}>
          {[["💰","รายรับวันนี้",todayIn,T.green,ydIn>0?`${todayIn>=ydIn?"▲":"▼"}${Math.abs(((todayIn-ydIn)/ydIn)*100).toFixed(0)}% vs เมื่อวาน`:""],
            ["💸","รายจ่ายวันนี้",todayOut,T.red,`กำไรวันนี้ ฿${fmt(todayIn-todayOut)}`],
            ["📅","รายรับเดือน",mIn,T.orange,`BEP ฿${fmt(totalFixed)}`],
            ["📈","กำไรสุทธิ",netP,netP>=0?T.green:T.red,"หลังต้นทุนคงที่"]
          ].map(([ic,l,v,col,sub])=>(
            <Card key={l} style={{padding:"16px 18px",borderLeft:`4px solid ${col}`}}>
              <div style={{color:T.textSm,fontSize:13,marginBottom:4}}>{ic} {l}</div>
              <div style={{color:col,fontWeight:900,fontSize:24}}>฿{fmt(v)}</div>
              {sub&&<div style={{fontSize:11,color:T.textSm,marginTop:4}}>{sub}</div>}
            </Card>
          ))}
        </div>
      ):(
        <Card style={{padding:"16px 18px",background:T.bg,textAlign:"center"}}>
          <div style={{fontSize:28,marginBottom:6}}>🔒</div>
          <div style={{color:T.textMd,fontSize:14,fontWeight:600}}>ข้อมูลการเงินถูกซ่อน</div>
          <div style={{color:T.textXs,fontSize:12,marginTop:4}}>เจ้าของจะเปิดสิทธิ์ใน Settings</div>
        </Card>
      )}
      {alerts.length>0&&(
        <Card style={{borderColor:T.red+"44",background:T.redLt}}>
          <div style={{color:T.red,fontWeight:800,fontSize:16,marginBottom:10}}>🚨 สต็อคต้องดูแล ({alerts.length})</div>
          {alerts.map(s=>(
            <div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid rgba(220,38,38,.1)"}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}><Badge status={stockSt(s)} /><span style={{fontWeight:600}}>{s.name}</span></div>
              <div style={{textAlign:"right"}}><div style={{color:ST_C[stockSt(s)],fontWeight:700}}>{s.qty} {s.unit}</div><div style={{color:T.textXs,fontSize:11}}>ขั้นต่ำ {s.minQty}</div></div>
            </div>
          ))}
          <button onClick={()=>setPage("stock")} style={{...S.btn(T.red),width:"100%",marginTop:10,padding:10,fontSize:14}}>ไปจัดการสต็อค →</button>
        </Card>
      )}
      {(wCost>0||highCost.length>0)&&(
        <Card style={{borderColor:T.yellow+"44",background:T.yellowLt}}>
          <div style={{color:T.yellow,fontWeight:800,fontSize:15,marginBottom:8}}>⚠️ แจ้งเตือนความเสี่ยง</div>
          {wCost>0&&<div style={{display:"flex",justifyContent:"space-between",padding:"5px 0",fontSize:14}}><span>🗑 Waste เดือนนี้</span><span style={{color:T.red,fontWeight:700}}>฿{fmt(wCost)}</span></div>}
          {highCost.map(s=>{const last=s.costHistory[s.costHistory.length-1]?.unitCost||0;const avg=wac(s);return <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",fontSize:14}}><span>📈 {s.name} ต้นทุนสูงขึ้น</span><span style={{color:T.red,fontWeight:700}}>+{(((last-avg)/avg)*100).toFixed(0)}%</span></div>;})}
          <button onClick={()=>setPage("report")} style={{...S.btn(T.yellow),width:"100%",marginTop:8,padding:8,fontSize:13}}>ดูรายงาน →</button>
        </Card>
      )}
      <Card>
        <div style={{fontWeight:700,fontSize:15,marginBottom:10}}>📦 สถานะสต็อค</div>
        {stock.map(s=>{const st=stockSt(s);const pct=s.minQty>0?Math.min((s.qty/s.minQty)*100,200):100;return(
          <div key={s.id} style={{marginBottom:11}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}><span style={{fontSize:14,fontWeight:600}}>{s.name}</span><Badge status={st} /></div>
              <span style={{color:ST_C[st],fontWeight:700,fontSize:14}}>{s.qty} {s.unit}</span>
            </div>
            <div style={{background:T.bg,borderRadius:4,height:6}}><div style={{background:pct<50?T.red:pct<100?T.yellow:T.green,width:`${Math.min(pct,100)}%`,height:"100%",borderRadius:4}}/></div>
          </div>
        );})}
      </Card>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {[["💵","Cash Flow","cashflow"],["📦","สต็อค",user.role==="owner"?"stock":"staffstock"]].map(([ic,l,pg])=>(
          <Card key={pg} style={{padding:"14px",cursor:"pointer",textAlign:"center"}} onClick={()=>setPage(pg)}>
            <div style={{fontSize:26,marginBottom:5}}>{ic}</div><div style={{fontWeight:700,fontSize:14}}>{l}</div>
          </Card>
        ))}
      </div>
    </div>
  );
}
