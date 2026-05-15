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

function ClearBtn({onClear,label="ล้างข้อมูล"}){
  return <button onClick={()=>{if(window.confirm(`ล้าง${label}?
ข้อมูลจะหายถาวร`))onClear();}} style={{...S.btn(T.red),fontSize:12,padding:"6px 10px"}}>🗑 ล้าง</button>;
}

function CheckinPage({staff,shopLat,shopLng,shopRadius,onCheckin}){
  const sid=new URLSearchParams(window.location.search).get("sid");
  const s=staff.find(x=>x.id===sid);
  const[done,setDone]=useState(null);
  const[loading,setLoading]=useState(false);
  const[locErr,setLocErr]=useState("");
  const[time,setTime]=useState(new Date().toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"}));
  useEffect(()=>{const t=setInterval(()=>setTime(new Date().toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"})),10000);return()=>clearInterval(t);},[]);
  const nowStr=()=>new Date().toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"});
  const todayStr=()=>new Date().toISOString().split("T")[0];
  const dist=(la1,lo1,la2,lo2)=>{const R=6371000;const dLat=(la2-la1)*Math.PI/180;const dLon=(lo2-lo1)*Math.PI/180;const a=Math.sin(dLat/2)**2+Math.cos(la1*Math.PI/180)*Math.cos(la2*Math.PI/180)*Math.sin(dLon/2)**2;return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));};
  const stamp=async type=>{
    if(!s)return;
    setLoading(true);setLocErr("");
    try{
      const pos=await new Promise((res,rej)=>navigator.geolocation.getCurrentPosition(res,rej,{timeout:8000,enableHighAccuracy:true}));
      const{latitude:lat,longitude:lng,accuracy:acc}=pos.coords;
      if(shopLat&&shopLng){
        const d=dist(lat,lng,+shopLat,+shopLng);
        const r=+(shopRadius||200);
        if(d>r+acc){setLocErr(`❌ ตำแหน่งของคุณห่างจากร้าน ${Math.round(d)}m (รัศมี ${r}m)
กรุณาอยู่ในบริเวณร้านก่อนเช็คเวลา`);setLoading(false);return;}
      }
      const t=nowStr();const d2=todayStr();
      const note=`📍 ${lat.toFixed(5)},${lng.toFixed(5)} ±${Math.round(acc)}m`;
      if(type==="in"){
        await sb("attendance",{method:"POST",body:JSON.stringify({id:Date.now(),staff_id:s.id,date:d2,check_in:t,check_out:"",note,lat_in:lat.toFixed(5),lng_in:lng.toFixed(5)})});
      } else {
        const recs=await sb(`attendance?staff_id=eq.${s.id}&date=eq.${d2}&check_out=eq.`);
        if(recs?.length>0)await sb(`attendance?id=eq.${recs[0].id}`,{method:"PATCH",body:JSON.stringify({check_out:t,lat_out:lat.toFixed(5),lng_out:lng.toFixed(5)})});
      }
      setDone({type,time:t,name:s.name,dist:shopLat?Math.round(dist(lat,lng,+shopLat,+shopLng)):null});
      onCheckin&&onCheckin(s.id,type,t,d2,{lat:lat.toFixed(5),lng:lng.toFixed(5),acc:Math.round(acc)});
    }catch(e){
      if(e.code===1)setLocErr("❌ กรุณาอนุญาตการใช้ GPS ก่อนเช็คเวลา");
      else setLocErr("❌ ไม่สามารถดึงตำแหน่งได้ ลองใหม่อีกครั้ง");
    }
    setLoading(false);
  };
  if(!sid||!s)return(
    <div style={{minHeight:"100vh",background:"#f4f4f5",fontFamily:F,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:32,textAlign:"center",maxWidth:320,width:"100%"}}>
        <div style={{fontSize:40,marginBottom:12}}>❌</div>
        <div style={{fontWeight:700,fontSize:18,color:T.red}}>QR Code ไม่ถูกต้อง</div>
        <div style={{color:T.textSm,fontSize:14,marginTop:8}}>กรุณาสแกน QR ที่หน้าร้านใหม่</div>
      </div>
    </div>
  );
  if(done)return(
    <div style={{minHeight:"100vh",background:done.type==="in"?"#f0fdf4":"#fef2f2",fontFamily:F,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,textAlign:"center",maxWidth:360,width:"100%",boxShadow:"0 8px 24px rgba(0,0,0,.12)"}}>
        <div style={{fontSize:64,marginBottom:12}}>{done.type==="in"?"✅":"👋"}</div>
        <div style={{fontWeight:900,fontSize:24,color:done.type==="in"?T.green:T.red}}>{done.type==="in"?"เช็คอินสำเร็จ!":"เช็คออกสำเร็จ!"}</div>
        <div style={{fontWeight:700,fontSize:20,marginTop:10}}>{done.name}</div>
        <div style={{fontSize:42,fontWeight:900,color:T.orange,marginTop:6,letterSpacing:2}}>{done.time}</div>
        {done.dist!==null&&<div style={{color:T.green,fontSize:14,marginTop:10}}>📍 ห่างจากร้าน {done.dist}m ✓</div>}
        <div style={{color:T.textSm,fontSize:13,marginTop:16}}>บันทึกเรียบร้อย สามารถปิดได้เลย</div>
      </div>
    </div>
  );
  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#fff7ed,#fff)",fontFamily:F,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:28,textAlign:"center",maxWidth:360,width:"100%",boxShadow:"0 8px 24px rgba(0,0,0,.12)"}}>
        <div style={{width:70,height:70,borderRadius:18,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:38,margin:"0 auto 10px"}}>🫕</div>
        <div style={{color:T.orange,fontWeight:900,fontSize:17}}>ไท่กั๋วหม่าล่า</div>
        <div style={{fontWeight:800,fontSize:24,marginTop:6}}>{s.name}</div>
        <div style={{color:T.textSm,fontSize:13,marginBottom:6}}>{new Date().toLocaleDateString("th-TH",{weekday:"long",day:"numeric",month:"long"})}</div>
        <div style={{fontSize:44,fontWeight:900,color:T.text,marginBottom:4,letterSpacing:2}}>{time}</div>
        {shopLat&&shopLng&&<div style={{color:T.textSm,fontSize:12,marginBottom:16}}>📍 ระบบจะตรวจสอบว่าอยู่ในรัศมี {shopRadius||200}m ของร้าน</div>}
        {!shopLat&&<div style={{color:T.yellow,fontSize:12,marginBottom:16}}>⚠️ เจ้าของยังไม่ได้ตั้งค่าพิกัดร้าน</div>}
        {locErr&&<div style={{background:T.redLt,border:`1px solid ${T.red}33`,borderRadius:10,padding:"10px 14px",fontSize:13,color:T.red,fontWeight:600,marginBottom:14,textAlign:"left",whiteSpace:"pre-line"}}>{locErr}</div>}
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <button onClick={()=>stamp("in")} disabled={loading} style={{background:loading?"#a1a1aa":T.green,color:"#fff",border:"none",borderRadius:14,padding:"18px 20px",fontSize:20,fontWeight:800,cursor:loading?"default":"pointer",fontFamily:F,boxShadow:`0 4px 12px ${T.green}44`}}>
            {loading?"⏳ กำลังตรวจสอบ GPS...":"📥 เข้างาน (Check-in)"}
          </button>
          <button onClick={()=>stamp("out")} disabled={loading} style={{background:loading?"#a1a1aa":T.red,color:"#fff",border:"none",borderRadius:14,padding:"18px 20px",fontSize:20,fontWeight:800,cursor:loading?"default":"pointer",fontFamily:F,boxShadow:`0 4px 12px ${T.red}44`}}>
            {loading?"⏳ กำลังตรวจสอบ GPS...":"📤 ออกงาน (Check-out)"}
          </button>
        </div>
        <div style={{color:T.textXs,fontSize:11,marginTop:16}}>กด Check-in/out แล้วระบบจะขอ GPS อัตโนมัติ</div>
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

function CashflowPage({cf,setCF,user,dbReady}){
  const[showForm,setShowForm]=useState(false);
  const[viewTab,setViewTab]=useState("daily");
  const[fMode,setFMode]=useState("month");
  const[fMonth,setFMonth]=useState(today().slice(0,7));
  const[fDay,setFDay]=useState(today());
  const[openCash,setOpenCash]=useState(0);
  const[openBank,setOpenBank]=useState(0);
  const[showSetup,setShowSetup]=useState(false);
  const[form,setForm]=useState({date:today(),flow:"in",cat:IN_CATS[0],itemName:"",amount:"",method:"เงินสด",note:""});
  const isCash=e=>e.method==="เงินสด";
  const myCF=user.role==="owner"?cf:cf.filter(e=>e.staffId===user.id||e.branch===user.franchiseId);
  const months=[...new Set(myCF.map(e=>e.date.slice(0,7)))].sort((a,b)=>b.localeCompare(a));
  const filtered=fMode==="day"?myCF.filter(e=>e.date===fDay):myCF.filter(e=>e.date.startsWith(fMonth));
  const sum=(arr,flow,fn)=>arr.filter(e=>e.flow===flow&&(fn?fn(e):true)).reduce((a,b)=>a+b.amount,0);
  const totalIn=sum(filtered,"in"),totalOut=sum(filtered,"out");
  const cashIn=sum(filtered,"in",isCash),cashOut=sum(filtered,"out",isCash);
  const bankIn=sum(filtered,"in",e=>!isCash(e)),bankOut=sum(filtered,"out",e=>!isCash(e));
  const cashBal=openCash+cashIn-cashOut,bankBal=openBank+bankIn-bankOut;
  const allDates=[...new Set(myCF.map(e=>e.date))].sort();
  let rC=openCash,rB=openBank;
  const dailyRows=allDates.map(d=>{const ent=myCF.filter(e=>e.date===d);const dCI=sum(ent,"in",isCash),dCO=sum(ent,"out",isCash),dBI=sum(ent,"in",e=>!isCash(e)),dBO=sum(ent,"out",e=>!isCash(e));rC+=dCI-dCO;rB+=dBI-dBO;return{d,ent,dCI,dCO,dBI,dBO,balCash:rC,balBank:rB};});
  const dispDays=[...(fMode==="day"?dailyRows.filter(r=>r.d===fDay):dailyRows.filter(r=>r.d.startsWith(fMonth)))].reverse();
  const addEntry=async()=>{if(!form.amount||!+form.amount)return;const entry={...form,id:Date.now(),amount:+form.amount,branch:user.franchiseId||"main",staffId:user.id};setCF(p=>[entry,...p]);setShowForm(false);setForm({date:today(),flow:"in",cat:IN_CATS[0],itemName:"",amount:"",method:"เงินสด",note:""});if(dbReady)db.addCF({id:entry.id,date:entry.date,flow:entry.flow,cat:entry.cat,item_name:entry.itemName,amount:entry.amount,method:entry.method,note:entry.note,branch:entry.branch,staff_id:entry.staffId}).catch(()=>{});};
  const delEntry=id=>{setCF(p=>p.filter(e=>e.id!==id));if(dbReady)db.delCF(id).catch(()=>{});};
  const clearAll=()=>{if(!window.confirm("ล้าง Cash Flow ทั้งหมด?"))return;setCF([]);if(dbReady)db.clearCF().catch(()=>{});};
  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="💵 Cash Flow" action={<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        <button onClick={()=>setShowSetup(!showSetup)} style={{...S.ghost,fontSize:13,padding:"7px 10px"}}>⚙️ ยอดเปิด</button>
        <IEBtn onExport={()=>exportXlsx(cf.map(e=>({วันที่:e.date,ประเภท:e.flow==="in"?"รายรับ":"รายจ่าย",หมวด:e.cat,ชื่อ:e.itemName||"",จำนวน:e.amount,ช่องทาง:e.method})),"cashflow","cashflow")} onImport={rows=>{const m=rows.map(r=>({id:Date.now()+Math.random(),date:r["วันที่"]||today(),flow:r["ประเภท"]==="รายรับ"?"in":"out",cat:r["หมวด"]||IN_CATS[0],itemName:r["ชื่อ"]||"",amount:+r["จำนวน"]||0,method:r["ช่องทาง"]||"เงินสด",note:"",branch:"main",staffId:"owner"})).filter(r=>r.amount>0);setCF(p=>[...m,...p]);alert(`นำเข้า ${m.length} รายการ`);}} />
        {user.role==="owner"&&<button onClick={clearAll} style={{...S.btn(T.red),fontSize:13,padding:"7px 10px"}}>🗑 ล้าง</button>}
        <button onClick={()=>setShowForm(!showForm)} style={S.btn()}>+ กรอก</button>
      </div>} />
      {showSetup&&<Card style={{borderColor:T.borderOr,background:T.orangeLt}}>
        <div style={{color:T.orange,fontWeight:800,fontSize:15,marginBottom:8}}>⚙️ ยอดเปิดบัญชี</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>💵 เงินสด</div><input type="number" value={openCash} onChange={e=>setOpenCash(+e.target.value||0)} style={S.inp} /></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>🏦 ธนาคาร</div><input type="number" value={openBank} onChange={e=>setOpenBank(+e.target.value||0)} style={S.inp} /></div>
        </div>
        <button onClick={()=>setShowSetup(false)} style={{...S.btn(),width:"100%",marginTop:10,padding:10}}>✅ บันทึก</button>
      </Card>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Card style={{padding:"14px 16px",borderLeft:"4px solid #78716c",background:"#fafaf9"}}>
          <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:5}}><span style={{fontSize:18}}>💵</span><span style={{color:"#78716c",fontWeight:700,fontSize:13}}>เงินสด</span></div>
          <div style={{color:cashBal>=0?T.green:T.red,fontWeight:900,fontSize:22}}>฿{fmt(cashBal)}</div>
          <div style={{fontSize:11,color:T.textSm,marginTop:3}}>เปิด ฿{fmt(openCash)} <span style={{color:T.green}}>+{fmt(cashIn)}</span> <span style={{color:T.red}}>-{fmt(cashOut)}</span></div>
        </Card>
        <Card style={{padding:"14px 16px",borderLeft:`4px solid ${T.blue}`,background:T.blueLt}}>
          <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:5}}><span style={{fontSize:18}}>🏦</span><span style={{color:T.blue,fontWeight:700,fontSize:13}}>ธนาคาร</span></div>
          <div style={{color:bankBal>=0?T.green:T.red,fontWeight:900,fontSize:22}}>฿{fmt(bankBal)}</div>
          <div style={{fontSize:11,color:T.textSm,marginTop:3}}>เปิด ฿{fmt(openBank)} <span style={{color:T.green}}>+{fmt(bankIn)}</span> <span style={{color:T.red}}>-{fmt(bankOut)}</span></div>
        </Card>
      </div>
      <Card style={{padding:"12px 16px",background:T.orangeLt,borderColor:T.borderOr}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{color:T.textSm,fontSize:13}}>รวมทั้งหมด</div><div style={{color:T.orange,fontWeight:900,fontSize:22}}>฿{fmt(cashBal+bankBal)}</div></div>
          <div style={{textAlign:"right",fontSize:13}}><div style={{color:T.green}}>เข้า ฿{fmt(totalIn)}</div><div style={{color:T.red}}>ออก ฿{fmt(totalOut)}</div></div>
        </div>
      </Card>
      <Card style={{padding:"12px 14px"}}>
        <div style={{display:"flex",gap:6,marginBottom:10}}>
          {[["month","📅 เดือน"],["day","📆 วัน"]].map(([v,l])=>(
            <button key={v} onClick={()=>setFMode(v)} style={{flex:1,padding:"7px 4px",borderRadius:8,cursor:"pointer",fontSize:13,fontFamily:F,background:fMode===v?T.orange:"transparent",border:`1px solid ${fMode===v?T.orange:T.border}`,color:fMode===v?"#fff":T.textMd}}>{l}</button>
          ))}
        </div>
        {fMode==="month"&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{(months.length?months:[today().slice(0,7)]).map(m=><button key={m} onClick={()=>setFMonth(m)} style={{background:fMonth===m?T.orange:"transparent",border:`1px solid ${fMonth===m?T.orange:T.border}`,borderRadius:8,padding:"6px 12px",color:fMonth===m?"#fff":T.textMd,cursor:"pointer",fontSize:13,fontFamily:F}}>{m}</button>)}</div>}
        {fMode==="day"&&<input type="date" value={fDay} max={today()} onChange={e=>setFDay(e.target.value)} style={{...S.inp,fontSize:15}} />}
      </Card>
      <Tabs tabs={[["daily","📆 รายวัน+ยอดยก"],["list","📋 รายการ"]]} active={viewTab} onChange={setViewTab} />
      {showForm&&<Card style={{borderColor:T.borderOr}}>
        <div style={{color:T.orange,fontWeight:800,fontSize:16,marginBottom:10}}>📝 บันทึกรายการ</div>
        <div style={{display:"flex",gap:8,marginBottom:12}}>
          {[["in","💰 รายรับ",T.green],["out","💸 รายจ่าย",T.red]].map(([v,l,c])=>(
            <button key={v} onClick={()=>setForm(p=>({...p,flow:v,cat:v==="in"?IN_CATS[0]:OUT_CATS[0]}))} style={{flex:1,padding:11,background:form.flow===v?c+"18":"transparent",border:`1.5px solid ${form.flow===v?c:T.border}`,borderRadius:10,color:form.flow===v?c:T.textMd,fontWeight:700,cursor:"pointer",fontSize:15,fontFamily:F}}>{l}</button>
          ))}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>วันที่</div><input type="date" value={form.date} max={today()} onChange={e=>setForm(p=>({...p,date:e.target.value}))} style={S.inp} /></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>จำนวน (฿)</div><input type="number" value={form.amount} onChange={e=>setForm(p=>({...p,amount:e.target.value}))} style={S.inp} placeholder="0" /></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>หมวด</div><select value={form.cat} onChange={e=>setForm(p=>({...p,cat:e.target.value}))} style={S.inp}>{(form.flow==="in"?IN_CATS:OUT_CATS).map(c=><option key={c}>{c}</option>)}</select></div>
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
        {dispDays.length===0&&<div style={{color:T.textSm,textAlign:"center",padding:28}}>ไม่มีรายการ</div>}
        {dispDays.map((row,i)=>{const prev=dispDays[i+1];return(
          <Card key={row.d} style={{padding:"13px 16px"}}>
            <div style={{display:"flex",justifyContent:"space-between",paddingBottom:8,marginBottom:8,borderBottom:`1px solid ${T.bg}`}}>
              <span style={{fontWeight:800,fontSize:15}}>{row.d}</span>
              <div style={{fontSize:13}}><span style={{color:T.green}}>+฿{fmt(row.dCI+row.dBI)}</span> <span style={{color:T.red}}>-฿{fmt(row.dCO+row.dBO)}</span></div>
            </div>
            {prev&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8,padding:"5px 10px",background:T.bg,borderRadius:8,fontSize:12}}>
              <div style={{color:T.textSm}}>💵 ยกมา: <b>฿{fmt(prev.balCash)}</b></div>
              <div style={{color:T.textSm}}>🏦 ยกมา: <b>฿{fmt(prev.balBank)}</b></div>
            </div>}
            {row.ent.map(e=>(
              <div key={e.id} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",borderBottom:`1px solid ${T.bg}`}}>
                <div style={{width:24,height:24,borderRadius:6,flexShrink:0,background:e.flow==="in"?T.greenLt:T.redLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,color:e.flow==="in"?T.green:T.red}}>{e.flow==="in"?"↓":"↑"}</div>
                <div style={{flex:1}}><div style={{fontSize:14,fontWeight:600}}>{e.itemName||e.cat}</div><div style={{color:T.textSm,fontSize:11}}>{isCash(e)?"💵":"🏦"} {e.method}{e.note?` • ${e.note}`:""}</div></div>
                <span style={{color:e.flow==="in"?T.green:T.red,fontWeight:700,fontSize:14}}>{e.flow==="in"?"+":"-"}฿{fmt(e.amount)}</span>
                <button onClick={()=>delEntry(e.id)} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:18,padding:"0 4px"}}>×</button>
              </div>
            ))}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:8,padding:"7px 10px",background:T.orangeLt,borderRadius:8,fontSize:13}}>
              <div><div style={{color:T.textXs,fontSize:11}}>💵 สิ้นวัน</div><div style={{color:row.balCash>=0?T.green:T.red,fontWeight:800,fontSize:15}}>฿{fmt(row.balCash)}</div></div>
              <div><div style={{color:T.textXs,fontSize:11}}>🏦 สิ้นวัน</div><div style={{color:row.balBank>=0?T.green:T.red,fontWeight:800,fontSize:15}}>฿{fmt(row.balBank)}</div></div>
            </div>
          </Card>
        );})}
      </div>}
      {viewTab==="list"&&<div style={{display:"flex",flexDirection:"column",gap:8}}>
        {filtered.length===0&&<div style={{color:T.textSm,textAlign:"center",padding:28}}>ไม่มีรายการ</div>}
        {[...filtered].sort((a,b)=>b.date.localeCompare(a.date)||b.id-a.id).map(e=>(
          <Card key={e.id} style={{padding:"11px 14px"}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:30,height:30,borderRadius:8,flexShrink:0,background:e.flow==="in"?T.greenLt:T.redLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,color:e.flow==="in"?T.green:T.red}}>{e.flow==="in"?"↓":"↑"}</div>
              <div style={{flex:1}}><div style={{fontSize:15,fontWeight:600}}>{e.itemName||e.cat}</div><div style={{color:T.textSm,fontSize:12}}>{e.date} • {isCash(e)?"💵":"🏦"} {e.method}</div></div>
              <span style={{color:e.flow==="in"?T.green:T.red,fontWeight:700,fontSize:15}}>{e.flow==="in"?"+":"-"}฿{fmt(e.amount)}</span>
              <button onClick={()=>delEntry(e.id)} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:18,padding:"0 4px"}}>×</button>
            </div>
          </Card>
        ))}
      </div>}
    </div>
  );
}

function StockPage({stock,setStock,movements,setMovements,user,suppliers}){
  const[tab,setTab]=useState("list");const[selId,setSelId]=useState("");const[mvType,setMvType]=useState("in");const[qty,setQty]=useState("");const[cost,setCost]=useState("");const[note,setNote]=useState("");const[msg,setMsg]=useState({t:"",ok:false});const[editId,setEditId]=useState(null);const[editData,setEditData]=useState({});const[showAdd,setShowAdd]=useState(false);const[newItem,setNewItem]=useState({name:"",unit:"kg",qty:0,minQty:3,dailyUse:1,supplierId:1});
  const canPrice=user.perms?.viewPrice;const selItem=selId?stock.find(s=>String(s.id)===String(selId)):null;
  const showMsg=(t,ok)=>{setMsg({t,ok});setTimeout(()=>setMsg({t:"",ok:false}),3000);};
  const save=()=>{const q=parseFloat(qty);if(!selId||!q||q<=0){showMsg("กรุณาเลือกรายการและกรอกจำนวน",false);return;}const item=stock.find(s=>String(s.id)===String(selId));if(!item)return;if(mvType==="out"&&q>item.qty){showMsg(`มีแค่ ${item.qty} ${item.unit}`,false);return;}const uc=parseFloat(cost)||0;const nh=mvType==="in"&&uc>0?[...(item.costHistory||[]),{date:today(),unitCost:uc,qty:q,total:uc*q}]:item.costHistory;setStock(stock.map(s=>String(s.id)===String(selId)?{...s,qty:mvType==="in"?s.qty+q:s.qty-q,costHistory:nh}:s));setMovements(p=>[...p,{id:Date.now(),itemId:item.id,type:mvType,qty:q,unitCost:uc,date:today(),staffId:user.id,note:note||(mvType==="in"?"รับเข้า":"จ่ายออก"),branch:"main"}]);showMsg(`✅ ${item.name} ${q} ${item.unit}`,true);setQty("");setCost("");setNote("");};
  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="📦 สต็อค" action={<div style={{display:"flex",gap:6}}>
        <IEBtn onExport={()=>exportXlsx(stock.map(s=>({ชื่อ:s.name,หน่วย:s.unit,จำนวน:s.qty,ขั้นต่ำ:s.minQty,ใช้ต่อวัน:s.dailyUse})),"stock","stock")} onImport={rows=>{let a=0,u=0;const ns=[...stock];rows.forEach(r=>{const name=r["ชื่อ"]||r["name"]||"";if(!name)return;const idx=ns.findIndex(s=>s.name===name);if(idx>=0){if(r["จำนวน"]!==undefined)ns[idx]={...ns[idx],qty:+r["จำนวน"]};u++;}else{ns.push({id:Date.now()+Math.random(),name,unit:r["หน่วย"]||"kg",qty:+(r["จำนวน"]||0),minQty:+(r["ขั้นต่ำ"]||3),dailyUse:+(r["ใช้ต่อวัน"]||1),supplierId:1,costHistory:[]});a++;}});setStock(ns);alert(`เพิ่ม ${a} อัพเดท ${u}`);}} />
        {user.role==="owner"&&<ClearBtn label="สต็อคและประวัติ" onClear={()=>{setStock(INIT_STOCK);setMovements([]);}} />}
        <button onClick={()=>setShowAdd(!showAdd)} style={S.btn()}>+ เพิ่ม</button>
      </div>} />
      {showAdd&&<Card style={{borderColor:T.borderOr}}>
        <div style={{color:T.orange,fontWeight:800,fontSize:15,marginBottom:10}}>➕ เพิ่มวัตถุดิบ</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          {[["ชื่อ","name","text"],["หน่วย","unit","text"],["จำนวน","qty","number"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"]].map(([l,k,t])=>(
            <div key={k}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>{l}</div><input type={t} value={newItem[k]} onChange={e=>setNewItem(p=>({...p,[k]:e.target.value}))} style={S.inp} /></div>
          ))}
        </div>
        <div style={{display:"flex",gap:8,marginTop:10}}>
          <button onClick={()=>{setStock([...stock,{...newItem,id:Date.now(),qty:+newItem.qty,minQty:+newItem.minQty,dailyUse:+newItem.dailyUse,costHistory:[]}]);setShowAdd(false);}} style={{...S.btn(),flex:1}}>บันทึก</button>
          <button onClick={()=>setShowAdd(false)} style={S.ghost}>ยกเลิก</button>
        </div>
      </Card>}
      <Tabs tabs={[["list","📋 รายการ"],["move","📥📤 รับ/จ่าย"],["history","📊 ประวัติ"]]} active={tab} onChange={t=>{setTab(t);setSelId("");setQty("");setMsg({t:"",ok:false});}} />
      {tab==="list"&&stock.map(item=>{const st=stockSt(item);const sup=suppliers?.find(x=>x.id===item.supplierId);return(<Card key={item.id} style={{borderColor:st!=="ok"?ST_C[st]+"44":T.border}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{flex:1}}><div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}><span style={{fontWeight:700,fontSize:16}}>{item.name}</span><Badge status={st} /></div>
            <div style={{color:T.textSm,fontSize:13,marginTop:3}}>ซัพฯ: {sup?.name||"-"} • ใช้/วัน {item.dailyUse}{canPrice&&wac(item)>0?` • ฿${wac(item).toFixed(2)}/${item.unit}`:""}</div>
            <div style={{color:T.textXs,fontSize:12,marginTop:2}}>ขั้นต่ำ {item.minQty} • เหลือ {item.dailyUse>0?(item.qty/item.dailyUse).toFixed(1):"∞"} วัน</div>
          </div>
          <div style={{textAlign:"right",marginLeft:10}}><div style={{color:ST_C[st],fontWeight:900,fontSize:26}}>{item.qty}</div><div style={{color:T.textSm,fontSize:13}}>{item.unit}</div></div>
        </div>
        {editId===item.id?(
          <div style={{marginTop:10,paddingTop:10,borderTop:`1px solid ${T.bg}`}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
              {[["ชื่อ","name","text"],["หน่วย","unit","text"],["ขั้นต่ำ","minQty","number"],["ใช้/วัน","dailyUse","number"]].map(([l,k,t])=>(
                <div key={k}><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>{l}</div><input type={t} value={editData[k]??item[k]} onChange={e=>setEditData(p=>({...p,[k]:e.target.value}))} style={S.inp} /></div>
              ))}
            </div>
            <div style={{display:"flex",gap:8}}><button onClick={()=>{setStock(p=>p.map(s=>s.id===item.id?{...s,...editData,minQty:+editData.minQty||s.minQty,dailyUse:+editData.dailyUse||s.dailyUse}:s));setEditId(null);setEditData({});}} style={{...S.btn(),flex:1}}>บันทึก</button><button onClick={()=>{setEditId(null);setEditData({});}} style={S.ghost}>ยกเลิก</button></div>
          </div>
        ):(
          <div style={{borderTop:`1px solid ${T.bg}`,marginTop:8,paddingTop:7,display:"flex",justifyContent:"space-between"}}>
            <button onClick={()=>{setEditId(item.id);setEditData({name:item.name,unit:item.unit,minQty:item.minQty,dailyUse:item.dailyUse});}} style={{background:"none",border:"none",color:T.orange,cursor:"pointer",fontSize:12,fontWeight:600}}>✏️ แก้ไข</button>
            <button onClick={()=>{if(window.confirm(`ลบ "${item.name}"?`))setStock(p=>p.filter(s=>s.id!==item.id));}} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:12}}>🗑 ลบ</button>
          </div>
        )}
      </Card>);})}
      {tab==="move"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        <div style={{display:"flex",gap:8}}>{[["in","📥 รับเข้า",T.green],["out","📤 จ่ายออก",T.red]].map(([v,l,c])=><button key={v} onClick={()=>setMvType(v)} style={{flex:1,padding:11,fontWeight:800,fontSize:15,cursor:"pointer",fontFamily:F,borderRadius:10,background:mvType===v?c+"18":"#fff",border:`2px solid ${mvType===v?c:T.border}`,color:mvType===v?c:T.textMd}}>{l}</button>)}</div>
        <div><div style={{color:T.textSm,fontSize:13,marginBottom:5,fontWeight:600}}>เลือกรายการ</div>
          <select value={selId} onChange={e=>{setSelId(e.target.value);setQty("");setCost("");}} style={{...S.inp,fontSize:15,height:46}}>
            <option value="">— กรุณาเลือก —</option>
            {[...stock].sort((a,b)=>({out:0,critical:1,low:2,ok:3}[stockSt(a)]-{out:0,critical:1,low:2,ok:3}[stockSt(b)])).map(s=>{const st=stockSt(s);return <option key={s.id} value={String(s.id)}>{st==="out"?"🔴":st==="critical"?"🟠":st==="low"?"🟡":"🟢"} {s.name} (เหลือ {s.qty} {s.unit})</option>;})}
          </select>
        </div>
        {selItem&&<div style={{background:T.bg,borderRadius:10,padding:"10px 14px",display:"flex",justifyContent:"space-between",fontSize:14}}><span>คงเหลือ</span><span style={{color:ST_C[stockSt(selItem)],fontWeight:800,fontSize:17}}>{selItem.qty} {selItem.unit}</span></div>}
        <div style={{display:"grid",gridTemplateColumns:mvType==="in"&&canPrice?"1fr 1fr":"1fr",gap:10}}>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>จำนวน</div><input type="number" value={qty} onChange={e=>{setQty(e.target.value);setMsg({t:"",ok:false});}} style={{...S.inp,fontSize:22,fontWeight:800,textAlign:"center"}} placeholder="0" /></div>
          {mvType==="in"&&canPrice&&<div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ราคา/หน่วย</div><input type="number" value={cost} onChange={e=>setCost(e.target.value)} style={{...S.inp,fontSize:17,borderColor:T.orange}} /></div>}
        </div>
        <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>หมายเหตุ</div><input value={note} onChange={e=>setNote(e.target.value)} style={S.inp} /></div>
        {selItem&&qty&&+qty>0&&<div style={{background:mvType==="in"?T.greenLt:T.redLt,borderRadius:8,padding:"9px 14px",fontSize:14}}>หลังบันทึก: <b style={{color:mvType==="in"?T.green:T.red,fontSize:17}}>{mvType==="in"?selItem.qty+(+qty):Math.max(0,selItem.qty-(+qty))} {selItem.unit}</b></div>}
        {msg.t&&<div style={{background:msg.ok?T.greenLt:T.yellowLt,borderRadius:8,padding:"9px 14px",fontSize:14,fontWeight:600,color:msg.ok?T.green:T.yellow}}>{msg.t}</div>}
        <button onClick={save} style={{...S.btn(mvType==="in"?T.green:T.red),width:"100%",padding:13,fontSize:16}}>✅ บันทึก {mvType==="in"?"รับเข้า":"จ่ายออก"}</button>
      </div>}
      {tab==="history"&&<Card>
        <div style={{fontWeight:800,fontSize:16,marginBottom:10}}>ประวัติ</div>
        {[...movements].sort((a,b)=>b.date.localeCompare(a.date)).slice(0,30).map(m=>{const item=stock.find(s=>s.id==m.itemId);const col=m.type==="in"?T.green:T.red;return(
          <div key={m.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:`1px solid ${T.bg}`}}>
            <div style={{width:30,height:30,borderRadius:8,background:col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>{m.type==="in"?"📥":"📤"}</div>
            <div style={{flex:1}}><div style={{fontWeight:600,fontSize:14}}>{item?.name||"?"}</div><div style={{color:T.textSm,fontSize:12}}>{m.date} • {m.note}</div></div>
            <span style={{color:col,fontWeight:800,fontSize:14}}>{m.type==="in"?"+":"-"}{m.qty} {item?.unit}</span>
            <button onClick={()=>setMovements(p=>p.filter(x=>x.id!==m.id))} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:16,padding:"0 4px"}}>×</button>
          </div>
        );})}
        {movements.length===0&&<div style={{color:T.textSm,textAlign:"center",padding:20}}>ยังไม่มีประวัติ</div>}
      </Card>}
    </div>
  );
}

function StaffStockPage({stock,setStock,movements,setMovements,user}){
  const[tab,setTab]=useState("quick");const[selId,setSelId]=useState("");const[mvType,setMvType]=useState("in");const[qty,setQty]=useState("");const[note,setNote]=useState("");const[msg,setMsg]=useState({t:"",ok:false});const[checked,setChecked]=useState({});const[done,setDone]=useState(false);
  const selItem=selId?stock.find(s=>String(s.id)===String(selId)):null;
  const showMsg=(t,ok)=>{setMsg({t,ok});setTimeout(()=>setMsg({t:"",ok:false}),3000);};
  const save=()=>{const q=parseFloat(qty);if(!selId||!q||q<=0){showMsg("กรุณาเลือกรายการและกรอกจำนวน",false);return;}const item=stock.find(s=>String(s.id)===String(selId));if(!item)return;if(mvType==="out"&&q>item.qty){showMsg(`มีแค่ ${item.qty} ${item.unit}`,false);return;}setStock(stock.map(s=>String(s.id)===String(selId)?{...s,qty:mvType==="in"?s.qty+q:s.qty-q}:s));setMovements(p=>[...p,{id:Date.now(),itemId:item.id,type:mvType,qty:q,unitCost:0,date:today(),staffId:user.id,note:note||(mvType==="in"?"รับเข้า":"จ่ายออก"),branch:"main"}]);showMsg(`✅ ${item.name} ${q} ${item.unit}`,true);setQty("");setNote("");};
  return(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div><div style={{fontSize:21,fontWeight:900}}>📦 บันทึกสต็อค</div><div style={{color:T.textSm,fontSize:13}}>สวัสดี {user.name} • {today()}</div></div>
      {stock.filter(s=>["critical","out"].includes(stockSt(s))).length>0&&<Card style={{background:T.redLt,borderColor:T.red+"44"}}>
        <div style={{color:T.red,fontWeight:800,fontSize:14,marginBottom:5}}>🚨 แจ้งเจ้าของด่วน!</div>
        {stock.filter(s=>["critical","out"].includes(stockSt(s))).map(s=><div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:"3px 0",fontSize:14}}><span style={{fontWeight:600}}>{s.name}</span><span style={{color:T.red,fontWeight:700}}>{s.qty} {s.unit}</span></div>)}
      </Card>}
      <Tabs tabs={[["quick","⚡ รับ/จ่าย"],["checklist","✅ นับรายวัน"],["history","📋 ประวัติ"]]} active={tab} onChange={t=>{setTab(t);setSelId("");setQty("");setMsg({t:"",ok:false});}} />
      {tab==="quick"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        <div style={{display:"flex",gap:8}}>{[["in","📥 รับเข้า",T.green],["out","📤 จ่ายออก",T.red]].map(([v,l,c])=><button key={v} onClick={()=>setMvType(v)} style={{flex:1,padding:11,fontWeight:800,fontSize:15,cursor:"pointer",fontFamily:F,borderRadius:10,background:mvType===v?c+"18":"#fff",border:`2px solid ${mvType===v?c:T.border}`,color:mvType===v?c:T.textMd}}>{l}</button>)}</div>
        <div><div style={{color:T.textSm,fontSize:13,marginBottom:5,fontWeight:600}}>เลือกรายการ</div>
          <select value={selId} onChange={e=>{setSelId(e.target.value);setQty("");}} style={{...S.inp,fontSize:15,height:46}}>
            <option value="">— กรุณาเลือก —</option>
            {[...stock].sort((a,b)=>({out:0,critical:1,low:2,ok:3}[stockSt(a)]-{out:0,critical:1,low:2,ok:3}[stockSt(b)])).map(s=>{const st=stockSt(s);return <option key={s.id} value={String(s.id)}>{st==="out"?"🔴":st==="critical"?"🟠":st==="low"?"🟡":"🟢"} {s.name} ({s.qty} {s.unit})</option>;})}
          </select>
        </div>
        {selItem&&<div style={{background:T.bg,borderRadius:10,padding:"9px 14px",display:"flex",justifyContent:"space-between",fontSize:14}}><span>คงเหลือ</span><span style={{color:ST_C[stockSt(selItem)],fontWeight:800,fontSize:17}}>{selItem.qty} {selItem.unit}</span></div>}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>จำนวน</div><input type="number" inputMode="numeric" value={qty} onChange={e=>{setQty(e.target.value);setMsg({t:"",ok:false});}} style={{...S.inp,fontSize:22,fontWeight:800,textAlign:"center"}} placeholder="0" /></div>
          <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>หมายเหตุ</div><input value={note} onChange={e=>setNote(e.target.value)} style={S.inp} /></div>
        </div>
        {selItem&&qty&&+qty>0&&<div style={{background:mvType==="in"?T.greenLt:T.redLt,borderRadius:8,padding:"9px 14px",fontSize:14}}>หลังบันทึก: <b style={{color:mvType==="in"?T.green:T.red,fontSize:17}}>{mvType==="in"?selItem.qty+(+qty):Math.max(0,selItem.qty-(+qty))} {selItem.unit}</b></div>}
        {msg.t&&<div style={{background:msg.ok?T.greenLt:T.yellowLt,borderRadius:8,padding:"9px 14px",fontSize:14,fontWeight:600,color:msg.ok?T.green:T.yellow}}>{msg.t}</div>}
        <button onClick={save} style={{...S.btn(mvType==="in"?T.green:T.red),width:"100%",padding:13,fontSize:16}}>✅ บันทึก</button>
      </div>}
      {tab==="checklist"&&<div style={{display:"flex",flexDirection:"column",gap:10}}>
        {done&&<div style={{background:T.greenLt,borderRadius:10,padding:"11px 16px",color:T.green,fontWeight:800}}>✅ บันทึกสำเร็จ!</div>}
        {stock.map(item=>{const st=stockSt(item);const val=checked[String(item.id)]??"";const diff=val!==""&&!isNaN(+val)?+val-item.qty:null;return(
          <Card key={item.id} style={{padding:"11px 14px",borderColor:st!=="ok"?ST_C[st]+"44":T.border}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{display:"flex",alignItems:"center",gap:7}}><span style={{fontWeight:700,fontSize:14}}>{item.name}</span><Badge status={st} /></div>
                <div style={{color:T.textSm,fontSize:12}}>ระบบ: {item.qty} {item.unit}</div>
                {diff!==null&&<div style={{fontSize:12,fontWeight:600,marginTop:2,color:diff<0?T.yellow:T.green}}>{diff<0?`⚠️ ขาด ${Math.abs(diff)}`:diff>0?`+เกิน ${diff}`:"✓ ตรง"}</div>}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:5}}>
                <input type="number" value={val} onChange={e=>setChecked(p=>({...p,[String(item.id)]:e.target.value}))} style={{...S.inp,width:76,fontSize:17,fontWeight:700,textAlign:"center"}} placeholder="นับ" />
                <span style={{color:T.textSm,fontSize:12}}>{item.unit}</span>
              </div>
            </div>
          </Card>
        );})}
        <button onClick={()=>{const entries=Object.entries(checked).filter(([,v])=>v!==""&&!isNaN(+v));if(!entries.length)return;setStock(stock.map(s=>{const v=checked[String(s.id)];return(!v||isNaN(+v))?s:{...s,qty:+v};}));entries.forEach(([id,v])=>{const item=stock.find(s=>String(s.id)===id);if(!item)return;setMovements(p=>[...p,{id:Date.now()+Math.random(),itemId:item.id,type:"check",qty:+v,unitCost:0,date:today(),staffId:user.id,note:"นับสต็อค",branch:"main"}]);});setChecked({});setDone(true);setTimeout(()=>setDone(false),3000);}} style={{...S.btn(T.green),width:"100%",padding:12,fontSize:15}}>✅ ส่งรายงาน ({Object.entries(checked).filter(([,v])=>v!=="").length} รายการ)</button>
      </div>}
      {tab==="history"&&<Card>
        <div style={{fontWeight:800,fontSize:15,marginBottom:10}}>ประวัติของฉัน</div>
        {[...movements].filter(m=>m.staffId===user.id).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,20).map(m=>{const item=stock.find(s=>s.id==m.itemId);const col=m.type==="in"?T.green:m.type==="check"?T.blue:T.red;return(
          <div key={m.id} style={{display:"flex",alignItems:"center",gap:9,padding:"8px 0",borderBottom:`1px solid ${T.bg}`}}>
            <div style={{width:30,height:30,borderRadius:8,background:col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>{m.type==="in"?"📥":m.type==="check"?"✅":"📤"}</div>
            <div style={{flex:1}}><div style={{fontWeight:600,fontSize:14}}>{item?.name||"?"}</div><div style={{color:T.textSm,fontSize:12}}>{m.date} • {m.note}</div></div>
            <span style={{color:col,fontWeight:800,fontSize:14}}>{m.type==="check"?"":m.type==="in"?"+":"-"}{m.qty} {item?.unit}</span>
          </div>
        );})}
        {movements.filter(m=>m.staffId===user.id).length===0&&<div style={{color:T.textSm,textAlign:"center",padding:20}}>ยังไม่มีประวัติ</div>}
      </Card>}
    </div>
  );
}

function PurchasePage({stock,suppliers,lineToken}){
  const[sel,setSel]=useState({});const[oQty,setOQty]=useState({});const[note,setNote]=useState("");const[sent,setSent]=useState(false);const[sending,setSending]=useState(false);const[prev,setPrev]=useState(false);
  const need=stock.filter(s=>s.qty<s.minQty);
  const tog=id=>{setSel(p=>({...p,[id]:!p[id]}));if(!oQty[id]){const it=stock.find(s=>s.id===id);if(it)setOQty(p=>({...p,[id]:String(Math.max(it.minQty*2-it.qty,1))}));}};
  const selItems=stock.filter(s=>sel[s.id]);
  const buildMsg=()=>{const lines=["🛒 ใบสั่งซื้อ",`📅 ${today()}`,"─────────"];const bySup={};selItems.forEach(it=>{const sn=suppliers.find(x=>x.id===it.supplierId)?.name||"ไม่ระบุ";if(!bySup[sn])bySup[sn]=[];bySup[sn].push(it);});Object.entries(bySup).forEach(([sn,items])=>{lines.push(`\n🏪 ${sn}`);items.forEach(it=>lines.push(`  • ${it.name}: ${oQty[it.id]||it.minQty} ${it.unit} (มี ${it.qty})`));});if(note)lines.push(`\n📝 ${note}`);lines.push("\n─────────\nไท่กั๋วหม่าล่า");return lines.join("\n");};
  const send=async()=>{if(!lineToken){alert("ตั้งค่า LINE Token ก่อน");return;}if(!selItems.length){alert("เลือกรายการก่อน");return;}setSending(true);try{const r=await fetch("https://notify-api.line.me/api/notify",{method:"POST",headers:{"Authorization":`Bearer ${lineToken}`,"Content-Type":"application/x-www-form-urlencoded"},body:`message=${encodeURIComponent(buildMsg())}`});if(r.ok){setSent(true);setTimeout(()=>setSent(false),5000);}else alert("ส่งไม่สำเร็จ");}catch{alert("ผิดพลาด");}setSending(false);};
  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="🛒 สั่งซื้อวัตถุดิบ" />
      {sent&&<div style={{background:T.greenLt,borderRadius:10,padding:"12px 16px",color:T.green,fontWeight:800}}>✅ ส่ง LINE สำเร็จ!</div>}
      {need.length>0?(
        <Card style={{borderColor:T.red+"44",background:T.redLt}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <div style={{color:T.red,fontWeight:800,fontSize:15}}>🚨 ต้องสั่งด่วน ({need.length})</div>
            <button onClick={()=>{const s={};const q={};need.forEach(x=>{s[x.id]=true;q[x.id]=String(x.minQty*2-x.qty);});setSel(s);setOQty(q);}} style={{...S.btn(T.red),fontSize:13,padding:"6px 12px"}}>เลือกทั้งหมด</button>
          </div>
          {need.map(s=>{const isSel=sel[s.id];const sup=suppliers.find(x=>x.id===s.supplierId);return(
            <div key={s.id} style={{background:"#fff",borderRadius:10,padding:"11px 14px",marginBottom:7,border:`2px solid ${isSel?T.orange:T.border}`,cursor:"pointer"}} onClick={()=>tog(s.id)}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div style={{display:"flex",alignItems:"center",gap:9}}>
                  <div style={{width:20,height:20,borderRadius:5,border:`2px solid ${isSel?T.orange:T.border}`,background:isSel?T.orange:"transparent",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:"#fff"}}>{isSel?"✓":""}</div>
                  <div><div style={{fontWeight:700,fontSize:14}}>{s.name}</div><div style={{color:T.textSm,fontSize:12}}>ซัพฯ: {sup?.name||"-"}</div></div>
                </div>
                <div style={{textAlign:"right"}}><div style={{color:T.red,fontWeight:700}}>เหลือ {s.qty} {s.unit}</div><div style={{color:T.textXs,fontSize:11}}>ขั้นต่ำ {s.minQty}</div></div>
              </div>
              {isSel&&<div style={{marginTop:9,paddingTop:9,borderTop:`1px solid ${T.bg}`,display:"flex",alignItems:"center",gap:9}} onClick={e=>e.stopPropagation()}>
                <span style={{color:T.textSm,fontSize:13}}>สั่ง:</span>
                <input type="number" value={oQty[s.id]||""} onChange={e=>setOQty(p=>({...p,[s.id]:e.target.value}))} style={{...S.inp,width:85,fontSize:16,fontWeight:700,textAlign:"center"}} placeholder="0" />
                <span style={{color:T.textSm,fontSize:13}}>{s.unit}</span>
              </div>}
            </div>
          );})}
        </Card>
      ):<Card style={{background:T.greenLt,borderColor:T.green+"44"}}><div style={{color:T.green,fontWeight:800,fontSize:15,textAlign:"center"}}>✅ สต็อคปกติทั้งหมด</div></Card>}
      <div><div style={{color:T.textSm,fontSize:13,marginBottom:5}}>หมายเหตุ</div><input value={note} onChange={e=>setNote(e.target.value)} style={S.inp} placeholder="เช่น ส่งด่วนก่อน 8 โมง" /></div>
      {selItems.length>0&&<button onClick={()=>setPrev(!prev)} style={{...S.ghost,width:"100%",padding:10,fontSize:13}}>{prev?"▲ ซ่อน":"▼ ดูตัวอย่าง"} ({selItems.length})</button>}
      {prev&&<Card style={{background:"#1a1a2e"}}><pre style={{color:"#e2e8f0",fontSize:12,whiteSpace:"pre-wrap",margin:0,lineHeight:1.7}}>{buildMsg()}</pre></Card>}
      {!lineToken&&<div style={{background:T.yellowLt,borderRadius:8,padding:"9px 14px",fontSize:13,color:T.yellow,fontWeight:600}}>⚠️ ตั้งค่า LINE Token ใน Settings → LINE ก่อน</div>}
      <button onClick={send} disabled={sending||!selItems.length} style={{...S.btn(!selItems.length?"#a1a1aa":T.green),width:"100%",padding:14,fontSize:16,opacity:!selItems.length?0.5:1}}>{sending?"⏳ กำลังส่ง...":`📲 ส่ง LINE (${selItems.length} รายการ)`}</button>
    </div>
  );
}

function ReportPage({cf,stock,movements,user,fixedCosts,waste,setWaste,promos,setPromos}){
  const[tab,setTab]=useState(user.role==="owner"?"pl":"forecast");const[wShow,setWShow]=useState(false);const[pShow,setPShow]=useState(false);const[wF,setWF]=useState({itemId:"",qty:"",reason:"",date:today()});const[pF,setPF]=useState({name:"",date:today(),amount:""});
  const mk=today().slice(0,7);const myCF=user.role==="owner"?cf:cf.filter(e=>e.branch===user.franchiseId);const mCF=myCF.filter(e=>e.date.startsWith(mk));
  const mIn=mCF.filter(e=>e.flow==="in").reduce((a,b)=>a+b.amount,0);const mOut=mCF.filter(e=>e.flow==="out").reduce((a,b)=>a+b.amount,0);const totalFixed=fixedCosts.reduce((a,b)=>a+b.amount,0);const cogs=mCF.filter(e=>e.flow==="out"&&["วัตถุดิบ/ผัก","หมู/เนื้อ/ทะเล"].includes(e.cat)).reduce((a,b)=>a+b.amount,0);const netP=mIn-mOut-totalFixed;
  const costAna=stock.map(s=>{const h=s.costHistory||[];if(!h.length)return null;const avg=wac(s);const mid=Math.floor(h.length/2);const f=h.slice(0,mid).reduce((a,b)=>a+b.unitCost,0)/Math.max(mid,1);const sc=h.slice(mid).reduce((a,b)=>a+b.unitCost,0)/Math.max(h.length-mid,1);const trend=f>0?((sc-f)/f*100):0;const prices=h.map(x=>x.unitCost);const risk=trend>20?"high":trend>10?"medium":"low";return{s,avg,trend,risk,last3:h.slice(-3),maxP:Math.max(...prices),minP:Math.min(...prices)};}).filter(Boolean);
  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="📊 รายงาน" action={<div style={{display:"flex",gap:6}}><IEBtn onExport={()=>{if(tab==="pl"&&user.role==="owner")exportXlsx([{เดือน:mk,รายรับ:mIn,รายจ่าย:mOut,กำไรสุทธิ:netP}],"P&L","pl");else if(tab==="waste")exportXlsx((waste||[]).map(w=>({วันที่:w.date,สินค้า:w.itemName,จำนวน:w.qty,มูลค่า:w.cost})),"waste","waste");else if(tab==="promo")exportXlsx((promos||[]).map(p=>({วันที่:p.date,โปร:p.name,มูลค่า:p.amount})),"promo","promo");else exportXlsx([],"data","report");}} onImport={rows=>{if(tab==="waste")setWaste(p=>[...p,...rows.map(r=>({id:Date.now()+Math.random(),date:r["วันที่"]||today(),itemName:r["สินค้า"]||"",qty:+r["จำนวน"]||0,reason:"",cost:+r["มูลค่า"]||0,itemId:""}))]);else if(tab==="promo")setPromos(p=>[...p,...rows.map(r=>({id:Date.now()+Math.random(),date:r["วันที่"]||today(),name:r["โปร"]||"",amount:+r["มูลค่า"]||0}))]);}} />{user.role==="owner"&&(tab==="waste"?<ClearBtn label="Waste ทั้งหมด" onClear={()=>setWaste([])} />:tab==="promo"?<ClearBtn label="โปรโมชั่นทั้งหมด" onClear={()=>setPromos([])} />:null)}</div>} />
      <Tabs tabs={[...(user.role==="owner"?[["pl","💰 P&L"],["cost","📈 ต้นทุน"]]:[]),["forecast","🔮 พยากรณ์"],["waste","🗑 Waste"],["promo","🎁 โปร"]]} active={tab} onChange={setTab} />
      {tab==="pl"&&user.role==="owner"&&<>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{color:T.orange,fontWeight:700,fontSize:13,marginBottom:9}}>เดือน {mk}</div>
          {[["💰 รายรับ",mIn,T.green,false],["− วัตถุดิบ",cogs,T.red,true],["= กำไรขั้นต้น",mIn-cogs,T.green,false],["− ค่าใช้จ่าย",mOut-cogs+totalFixed,T.yellow,true],["= กำไรสุทธิ",netP,netP>=0?T.green:T.red,false]].map(([l,v,c2,i])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.borderOr}`,paddingLeft:i?16:0}}><span style={{color:T.textMd,fontSize:14}}>{l}</span><span style={{color:c2,fontWeight:800,fontSize:15}}>฿{fmt(v)}</span></div>
          ))}
          <div style={{display:"flex",gap:16,marginTop:9,paddingTop:9,borderTop:`1px solid ${T.borderOr}`,fontSize:13}}>
            <div><div style={{color:T.textSm}}>Margin</div><div style={{color:T.orange,fontWeight:700,fontSize:15}}>{mIn>0?((mIn-cogs)/mIn*100).toFixed(1):0}%</div></div>
            <div><div style={{color:T.textSm}}>BEP/วัน</div><div style={{color:T.orange,fontWeight:700,fontSize:15}}>฿{fmt(Math.ceil(totalFixed/30))}</div></div>
            <div><div style={{color:T.textSm}}>Waste</div><div style={{color:T.red,fontWeight:700,fontSize:15}}>฿{fmt((waste||[]).filter(w=>w.date.startsWith(mk)).reduce((a,b)=>a+b.cost,0))}</div></div>
          </div>
        </Card>
        <Card><div style={{fontWeight:800,fontSize:14,marginBottom:9}}>📦 สต็อคใกล้หมด</div>
          {stock.filter(s=>stockSt(s)!=="ok").map(s=>(
            <div key={s.id} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.bg}`,fontSize:14}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontWeight:600}}>{s.name}</span><Badge status={stockSt(s)} /></div>
              <span style={{color:ST_C[stockSt(s)],fontWeight:700}}>{s.qty} {s.unit}</span>
            </div>
          ))}
          {stock.every(s=>stockSt(s)==="ok")&&<div style={{color:T.green,textAlign:"center",padding:12}}>✅ สต็อคปกติ</div>}
        </Card>
      </>}
      {tab==="cost"&&user.role==="owner"&&<Card>
        <div style={{fontWeight:800,fontSize:14,marginBottom:10}}>📈 ต้นทุนแต่ละสินค้า</div>
        {costAna.length===0&&<div style={{color:T.textSm,textAlign:"center",padding:20}}>ยังไม่มีข้อมูลราคา (กรอกราคาตอนรับสินค้าเข้า)</div>}
        {costAna.map(({s,avg,trend,risk,last3})=>(
          <div key={s.id} style={{marginBottom:13,paddingBottom:11,borderBottom:`1px solid ${T.bg}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}><span style={{fontWeight:700,fontSize:14}}>{s.name}</span>
                <span style={{background:risk==="high"?T.redLt:risk==="medium"?T.yellowLt:T.greenLt,color:risk==="high"?T.red:risk==="medium"?T.yellow:T.green,borderRadius:6,padding:"2px 7px",fontSize:10,fontWeight:700}}>{risk==="high"?"⚠️ ระวัง":risk==="medium"?"🟡 ติดตาม":"✅ ปกติ"}</span>
              </div>
              <div style={{textAlign:"right"}}><div style={{fontWeight:900,fontSize:18,color:risk==="high"?T.red:risk==="medium"?T.yellow:T.green}}>฿{avg.toFixed(2)}</div>{trend!==0&&<div style={{fontSize:12,color:trend>0?T.red:T.green}}>{trend>0?"▲":"▼"}{Math.abs(trend).toFixed(1)}%</div>}</div>
            </div>
            {last3.length>0&&<div style={{display:"flex",gap:5}}>{last3.map((h,i)=><div key={i} style={{flex:1,background:T.bg,borderRadius:7,padding:"5px 7px",textAlign:"center"}}><div style={{fontSize:10,color:T.textXs}}>{h.date?.slice(5)||"-"}</div><div style={{fontWeight:700,fontSize:13,color:h.unitCost>avg*1.1?T.red:h.unitCost<avg*.9?T.green:T.text}}>฿{h.unitCost.toFixed(0)}</div></div>)}</div>}
          </div>
        ))}
      </Card>}
      {tab==="forecast"&&stock.map(s=>{
        const last7d=Array.from({length:7},(_,i)=>{const d=new Date();d.setDate(d.getDate()-i);return d.toISOString().split("T")[0];});
        const da=movements.filter(m=>m.itemId===s.id&&m.type==="out"&&last7d.includes(m.date)).reduce((a,b)=>a+b.qty,0)/7;
        const dl=da>0?s.qty/da:999;const urgent=dl<3;const warn=dl<7;const ac=wac(s);
        return(<Card key={s.id} style={{borderColor:urgent?T.red+"44":warn?T.yellow+"44":T.border}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div style={{flex:1}}>
              <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:3}}>
                <span style={{fontWeight:700,fontSize:14}}>{s.name}</span>
                {urgent&&<span style={{background:T.redLt,color:T.red,borderRadius:6,padding:"2px 7px",fontSize:10,fontWeight:700}}>🚨 สั่งด่วน</span>}
                {!urgent&&warn&&<span style={{background:T.yellowLt,color:T.yellow,borderRadius:6,padding:"2px 7px",fontSize:10,fontWeight:700}}>⚠️ ใกล้หมด</span>}
              </div>
              <div style={{color:T.textSm,fontSize:12}}>ใช้เฉลี่ย {da.toFixed(2)} {s.unit}/วัน • เหลือ {dl<999?dl.toFixed(1):"∞"} วัน</div>
              {da>0&&<div style={{display:"flex",gap:8,marginTop:6}}>
                <div style={{background:T.bg,borderRadius:7,padding:"5px 9px"}}><div style={{color:T.textXs,fontSize:10}}>7 วัน</div><div style={{fontWeight:700,color:T.orange,fontSize:13}}>{Math.ceil(da*7)} {s.unit}</div>{ac>0&&<div style={{color:T.textXs,fontSize:10}}>฿{fmt(Math.ceil(da*7)*ac)}</div>}</div>
                <div style={{background:T.bg,borderRadius:7,padding:"5px 9px"}}><div style={{color:T.textXs,fontSize:10}}>14 วัน</div><div style={{fontWeight:700,color:T.blue,fontSize:13}}>{Math.ceil(da*14)} {s.unit}</div>{ac>0&&<div style={{color:T.textXs,fontSize:10}}>฿{fmt(Math.ceil(da*14)*ac)}</div>}</div>
              </div>}
            </div>
            <div style={{textAlign:"right",marginLeft:10}}><div style={{fontWeight:900,fontSize:22,color:urgent?T.red:warn?T.yellow:T.green}}>{s.qty}</div><div style={{color:T.textSm,fontSize:11}}>{s.unit}</div></div>
          </div>
        </Card>);
      })}
      {tab==="waste"&&<>
        <Card style={{background:T.redLt,borderColor:T.red+"44"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{color:T.red,fontWeight:700,fontSize:14}}>🗑 Waste เดือนนี้</div><div style={{color:T.red,fontWeight:800,fontSize:17}}>฿{fmt((waste||[]).filter(w=>w.date.startsWith(mk)).reduce((a,b)=>a+b.cost,0))}</div></div>
            <button onClick={()=>setWShow(!wShow)} style={{...S.btn(T.red),fontSize:13,padding:"7px 12px"}}>+ บันทึก</button>
          </div>
        </Card>
        {wShow&&<Card style={{borderColor:T.red+"44"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
            <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>สินค้า</div><select value={wF.itemId} onChange={e=>setWF(p=>({...p,itemId:e.target.value}))} style={{...S.inp,height:42}}><option value="">— เลือก —</option>{stock.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</select></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>จำนวน</div><input type="number" value={wF.qty} onChange={e=>setWF(p=>({...p,qty:e.target.value}))} style={S.inp} /></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>วันที่</div><input type="date" value={wF.date} max={today()} onChange={e=>setWF(p=>({...p,date:e.target.value}))} style={S.inp} /></div>
            <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>สาเหตุ</div><input value={wF.reason||""} onChange={e=>setWF(p=>({...p,reason:e.target.value}))} style={S.inp} placeholder="เช่น ผักเน่า" /></div>
          </div>
          {wF.itemId&&wF.qty&&<div style={{background:T.redLt,borderRadius:8,padding:"7px 12px",marginTop:8,fontSize:14}}>มูลค่า: <b style={{color:T.red}}>฿{fmt(wac(stock.find(s=>String(s.id)===String(wF.itemId))||{costHistory:[]})*(+wF.qty||0))}</b></div>}
          <div style={{display:"flex",gap:8,marginTop:9}}><button onClick={()=>{if(!wF.itemId||!wF.qty)return;const it=stock.find(s=>String(s.id)===String(wF.itemId));setWaste(p=>[{...wF,id:Date.now(),cost:wac(it||{costHistory:[]})*(+wF.qty),qty:+wF.qty,itemName:it?.name||""},...p]);setWF({itemId:"",qty:"",reason:"",date:today()});setWShow(false);}} style={{...S.btn(T.red),flex:1}}>✅ บันทึก</button><button onClick={()=>setWShow(false)} style={S.ghost}>ยกเลิก</button></div>
        </Card>}
        {(waste||[]).length===0&&<Card><div style={{color:T.textSm,textAlign:"center",padding:22}}>ยังไม่มี Waste</div></Card>}
        {[...(waste||[])].sort((a,b)=>b.date.localeCompare(a.date)).map(w=>(
          <Card key={w.id} style={{padding:"10px 14px"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontWeight:700,fontSize:14}}>{w.itemName}</div><div style={{color:T.textSm,fontSize:12}}>{w.date} • {w.reason||"ไม่ระบุ"}</div></div><div style={{textAlign:"right"}}><div style={{color:T.red,fontWeight:800,fontSize:15}}>฿{fmt(w.cost)}</div><button onClick={()=>setWaste(p=>p.filter(x=>x.id!==w.id))} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:12}}>🗑</button></div></div></Card>
        ))}
      </>}
      {tab==="promo"&&<>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{color:T.orange,fontWeight:700,fontSize:14}}>🎁 โปรเดือนนี้</div><div style={{color:T.orange,fontWeight:800,fontSize:17}}>฿{fmt((promos||[]).filter(p=>p.date.startsWith(mk)).reduce((a,b)=>a+b.amount,0))}</div></div>
            <button onClick={()=>setPShow(!pShow)} style={{...S.btn(),fontSize:13,padding:"7px 12px"}}>+ บันทึก</button>
          </div>
        </Card>
        {pShow&&<Card style={{borderColor:T.borderOr}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
            <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ชื่อโปร</div><input value={pF.name||""} onChange={e=>setPF(p=>({...p,name:e.target.value}))} style={S.inp} /></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>มูลค่า (฿)</div><input type="number" value={pF.amount||""} onChange={e=>setPF(p=>({...p,amount:e.target.value}))} style={S.inp} /></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>วันที่</div><input type="date" value={pF.date} max={today()} onChange={e=>setPF(p=>({...p,date:e.target.value}))} style={S.inp} /></div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:9}}><button onClick={()=>{if(!pF.name||!pF.amount)return;setPromos(p=>[{...pF,id:Date.now(),amount:+pF.amount},...p]);setPF({name:"",date:today(),amount:""});setPShow(false);}} style={{...S.btn(),flex:1}}>✅ บันทึก</button><button onClick={()=>setPShow(false)} style={S.ghost}>ยกเลิก</button></div>
        </Card>}
        {(promos||[]).length===0&&<Card><div style={{color:T.textSm,textAlign:"center",padding:22}}>ยังไม่มีโปรโมชั่น</div></Card>}
        {[...(promos||[])].sort((a,b)=>b.date.localeCompare(a.date)).map(p=>(
          <Card key={p.id} style={{padding:"10px 14px"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontWeight:700,fontSize:14}}>{p.name}</div><div style={{color:T.textSm,fontSize:12}}>{p.date}</div></div><div style={{textAlign:"right"}}><div style={{color:T.orange,fontWeight:800,fontSize:15}}>฿{fmt(p.amount)}</div><button onClick={()=>setPromos(prev=>prev.filter(x=>x.id!==p.id))} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:12}}>🗑</button></div></div></Card>
        ))}
      </>}
    </div>
  );
}

function StaffCheckinPage({user,attendance,setAttendance,shopLat,shopLng,shopRadius}){
  const[loading,setLoading]=useState(false);
  const[status,setStatus]=useState(null); // {ok,msg,type}
  const[locErr,setLocErr]=useState("");
  const nowFn=()=>new Date().toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"});
  const todayStr=today();
  const myAtt=attendance.find(a=>a.staffId===user.id&&a.date===todayStr);
  const dist=(la1,lo1,la2,lo2)=>{const R=6371000;const dLat=(la2-la1)*Math.PI/180;const dLon=(lo2-lo1)*Math.PI/180;const a=Math.sin(dLat/2)**2+Math.cos(la1*Math.PI/180)*Math.cos(la2*Math.PI/180)*Math.sin(dLon/2)**2;return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));};

  const stamp=async type=>{
    setLoading(true);setLocErr("");setStatus(null);
    try{
      const pos=await new Promise((res,rej)=>navigator.geolocation.getCurrentPosition(res,rej,{timeout:8000,enableHighAccuracy:true}));
      const{latitude:lat,longitude:lng,accuracy:acc}=pos.coords;
      if(shopLat&&shopLng){
        const d=dist(lat,lng,+shopLat,+shopLng);
        const r=+(shopRadius||200);
        if(d>r+acc){setLocErr(`❌ คุณอยู่ห่างจากร้าน ${Math.round(d)}m (รัศมี ${r}m)
ต้องอยู่ในบริเวณร้านถึงจะเช็คเวลาได้`);setLoading(false);return;}
      }
      const t=nowFn();
      const note=`📍 ${lat.toFixed(5)},${lng.toFixed(5)} ±${Math.round(acc)}m`;
      if(type==="in"){
        if(myAtt&&!myAtt.checkOut){setLocErr("⚠️ เช็คอินวันนี้แล้ว");setLoading(false);return;}
        const newAtt={id:Date.now(),staffId:user.id,date:todayStr,checkIn:t,checkOut:"",note};
        setAttendance(p=>[...p,newAtt]);
        // sync to supabase
        sb("attendance",{method:"POST",body:JSON.stringify({id:newAtt.id,staff_id:user.id,date:todayStr,check_in:t,check_out:"",note,lat_in:lat.toFixed(5),lng_in:lng.toFixed(5)})}).catch(()=>{});
        setStatus({ok:true,type:"in",msg:`เช็คอินสำเร็จ ${t} น.`});
      } else {
        if(!myAtt||myAtt.checkOut){setLocErr("⚠️ ยังไม่ได้เช็คอินวันนี้");setLoading(false);return;}
        setAttendance(p=>p.map(a=>a.id===myAtt.id?{...a,checkOut:t}:a));
        sb(`attendance?id=eq.${myAtt.id}`,{method:"PATCH",body:JSON.stringify({check_out:t,lat_out:lat.toFixed(5),lng_out:lng.toFixed(5)})}).catch(()=>{});
        setStatus({ok:true,type:"out",msg:`เช็คออกสำเร็จ ${t} น.`});
      }
    }catch(e){
      setLocErr(e.code===1?"❌ กรุณาอนุญาต GPS ก่อนเช็คเวลา":"❌ ดึง GPS ไม่ได้ ลองใหม่อีกครั้ง");
    }
    setLoading(false);
  };

  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div><div style={{fontSize:22,fontWeight:900}}>📲 เช็คเข้า-ออกงาน</div><div style={{color:T.textSm,fontSize:14}}>สวัสดี {user.name} • {todayStr}</div></div>

      {/* สถานะวันนี้ */}
      <Card style={{borderColor:myAtt?.checkIn?myAtt.checkOut?T.green+"44":T.orange+"44":T.border,background:myAtt?.checkIn?myAtt.checkOut?T.greenLt:T.orangeLt:"#fff"}}>
        <div style={{fontWeight:700,fontSize:15,marginBottom:8}}>📋 สถานะวันนี้</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div style={{textAlign:"center"}}>
            <div style={{color:T.textSm,fontSize:12,marginBottom:4}}>⏰ เข้างาน</div>
            <div style={{fontWeight:900,fontSize:26,color:myAtt?.checkIn?T.green:T.textXs}}>{myAtt?.checkIn||"--:--"}</div>
          </div>
          <div style={{textAlign:"center"}}>
            <div style={{color:T.textSm,fontSize:12,marginBottom:4}}>🏁 ออกงาน</div>
            <div style={{fontWeight:900,fontSize:26,color:myAtt?.checkOut?T.red:T.textXs}}>{myAtt?.checkOut||"--:--"}</div>
          </div>
        </div>
        {myAtt?.checkIn&&!myAtt.checkOut&&<div style={{textAlign:"center",marginTop:8,color:T.orange,fontWeight:600,fontSize:13}}>🟡 กำลังทำงานอยู่</div>}
        {myAtt?.checkOut&&<div style={{textAlign:"center",marginTop:8,color:T.green,fontWeight:600,fontSize:13}}>✅ ออกงานแล้ววันนี้</div>}
        {!myAtt?.checkIn&&<div style={{textAlign:"center",marginTop:8,color:T.textSm,fontSize:13}}>ยังไม่ได้เช็คอินวันนี้</div>}
        {myAtt?.note&&<div style={{marginTop:6,color:T.textXs,fontSize:11,textAlign:"center"}}>{myAtt.note}</div>}
      </Card>

      {/* แจ้งเตือนพิกัด */}
      {!shopLat&&<div style={{background:T.yellowLt,borderRadius:10,padding:"10px 14px",fontSize:13,color:T.yellow,fontWeight:600}}>⚠️ เจ้าของยังไม่ได้ตั้งพิกัดร้าน — เช็คเวลาได้แต่ไม่ตรวจสอบระยะทาง</div>}
      {shopLat&&shopLng&&<div style={{background:T.greenLt,borderRadius:10,padding:"8px 14px",fontSize:12,color:T.green,fontWeight:600}}>📍 ระบบจะตรวจสอบว่าอยู่ในรัศมี {shopRadius||200}m ของร้าน</div>}

      {/* Error */}
      {locErr&&<div style={{background:T.redLt,border:`1px solid ${T.red}33`,borderRadius:10,padding:"12px 14px",fontSize:14,color:T.red,fontWeight:600,whiteSpace:"pre-line"}}>{locErr}</div>}

      {/* Success */}
      {status?.ok&&<div style={{background:status.type==="in"?T.greenLt:T.orangeLt,border:`1px solid ${status.type==="in"?T.green:T.orange}44`,borderRadius:10,padding:"14px 16px",textAlign:"center"}}>
        <div style={{fontSize:40,marginBottom:8}}>{status.type==="in"?"✅":"👋"}</div>
        <div style={{fontWeight:800,fontSize:18,color:status.type==="in"?T.green:T.orange}}>{status.msg}</div>
      </div>}

      {/* ปุ่มเช็คอิน/ออก */}
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        <button onClick={()=>{setStatus(null);stamp("in");}} disabled={loading||!!(myAtt?.checkIn&&!myAtt?.checkOut)}
          style={{background:loading?"#a1a1aa":myAtt?.checkIn&&!myAtt?.checkOut?"#86efac":T.green,color:"#fff",border:"none",borderRadius:14,padding:"18px 20px",fontSize:20,fontWeight:800,cursor:loading||!!(myAtt?.checkIn&&!myAtt?.checkOut)?"default":"pointer",fontFamily:F,boxShadow:`0 4px 12px ${T.green}33`,opacity:myAtt?.checkIn&&!myAtt?.checkOut?0.6:1}}>
          {loading?"⏳ กำลังตรวจสอบ GPS...":myAtt?.checkIn&&!myAtt?.checkOut?"✅ เช็คอินแล้ว":"📥 เข้างาน (Check-in)"}
        </button>
        <button onClick={()=>{setStatus(null);stamp("out");}} disabled={loading||!!myAtt?.checkOut||!myAtt?.checkIn}
          style={{background:loading?"#a1a1aa":myAtt?.checkOut?"#fca5a5":T.red,color:"#fff",border:"none",borderRadius:14,padding:"18px 20px",fontSize:20,fontWeight:800,cursor:loading||!!myAtt?.checkOut||!myAtt?.checkIn?"default":"pointer",fontFamily:F,boxShadow:`0 4px 12px ${T.red}33`,opacity:!myAtt?.checkIn||myAtt?.checkOut?0.5:1}}>
          {loading?"⏳ กำลังตรวจสอบ GPS...":myAtt?.checkOut?"✅ ออกงานแล้ว":"📤 ออกงาน (Check-out)"}
        </button>
      </div>

      <div style={{color:T.textXs,fontSize:11,textAlign:"center"}}>กดปุ่มแล้วระบบจะขอ GPS อัตโนมัติ • เวลาถูกบันทึกและส่งเจ้าของทันที</div>
      <div style={{marginTop:8,paddingTop:14,borderTop:`1px dashed ${T.border}`}}>
        <div style={{color:T.textXs,fontSize:11,textAlign:"center",marginBottom:8}}>🧪 โหมดทดสอบระบบ</div>
        <button onClick={()=>{if(!window.confirm("ล้างประวัติการเข้างานวันนี้?\n(ใช้สำหรับทดสอบระบบเท่านั้น)"))return;setAttendance(p=>p.filter(a=>!(a.staffId===user.id&&a.date===todayStr)));setStatus(null);setLocErr("");}} style={{width:"100%",background:"transparent",border:`1px dashed ${T.textXs}`,borderRadius:10,padding:"9px 14px",color:T.textSm,cursor:"pointer",fontSize:13,fontFamily:F}}>
          🗑 ล้างประวัติวันนี้ (เทสระบบ)
        </button>
      </div>
    </div>
  );
}


function HRPage({staff,setStaff,attendance,setAttendance,cf,shopLat,shopLng}){
  const[tab,setTab]=useState("checkin");const[selStaff,setSelStaff]=useState("");const[salaryMonth,setSalaryMonth]=useState(today().slice(0,7));const[editHrId,setEditHrId]=useState(null);const[hrEdit,setHrEdit]=useState({});const[editOtId,setEditOtId]=useState(null);const[otVal,setOtVal]=useState("");
  const workers=staff.filter(s=>s.role==="staff"&&s.active);
  const nowFn=()=>new Date().toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit"});
  const checkIn=sid=>{if(attendance.find(a=>a.staffId===sid&&a.date===today()&&!a.checkOut)){alert("เช็คอินแล้ว");return;}setAttendance(p=>[...p,{id:Date.now(),staffId:sid,date:today(),checkIn:nowFn(),checkOut:"",note:""}]);};
  const checkOut=sid=>setAttendance(p=>p.map(a=>a.staffId===sid&&a.date===today()&&!a.checkOut?{...a,checkOut:nowFn()}:a));
  const calcPay=(s,att)=>{const hr=s.hr||{};const wage=hr.wage||0;const wt=hr.wageType||"day";let base=wt==="month"?wage:att.filter(a=>a.checkIn).length*wage;const otp=hr.otPerHour||50;const ot=att.reduce((sum,a)=>{const ov=a.otOverride;return sum+(ov!==undefined&&ov!==null?ov:calcOT(a.checkOut,otp).p);},0);let bonus=0;if(hr.bonusPct>0){const sales=cf.filter(e=>e.flow==="in"&&e.staffId===s.id&&e.date.startsWith(salaryMonth)).reduce((a,b)=>a+b.amount,0);bonus=sales*(hr.bonusPct/100);}return{base,ot,bonus,total:base+ot+bonus};};
  const QRCodeDisplay=({staffId})=>{const s=staff.find(x=>x.id===staffId);if(!s)return null;const url=`${window.location.origin}${window.location.pathname}?sid=${s.id}`;const qr=`https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(url)}`;return(
    <div style={{textAlign:"center",padding:12}}>
      <div style={{background:"#fff",padding:12,borderRadius:14,border:`2px solid ${T.orange}`,display:"inline-block"}}>
        <img src={qr} alt="QR" style={{width:160,height:160,display:"block"}} />
        <div style={{fontWeight:800,fontSize:15,marginTop:6}}>{s.name}</div>
        <div style={{color:T.textSm,fontSize:11}}>สแกนเปิดหน้าเช็คเวลา</div>
      </div>
      <div style={{marginTop:8,fontSize:11,color:T.textSm,wordBreak:"break-all"}}>{url}</div>
      <div style={{display:"flex",gap:6,justifyContent:"center",marginTop:7,flexWrap:"wrap"}}>
        <button onClick={()=>window.open(url,"_blank")} style={{...S.ghost,fontSize:12,padding:"5px 10px"}}>🔗 เปิด URL</button>
        <button onClick={()=>window.print()} style={{...S.ghost,fontSize:12,padding:"5px 10px"}}>🖨 พิมพ์ QR</button>
        <button onClick={()=>checkIn(staffId)} style={{...S.btn(T.green),fontSize:12,padding:"5px 10px"}}>📥 บันทึกเข้า</button>
        <button onClick={()=>checkOut(staffId)} style={{...S.btn(T.red),fontSize:12,padding:"5px 10px"}}>📤 บันทึกออก</button>
      </div>
    </div>
  );};
  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="👥 HR & เงินเดือน" action={<div style={{display:"flex",gap:6}}>
        <IEBtn onExport={()=>{if(tab==="salary")exportXlsx(workers.map(s=>{const att=attendance.filter(a=>a.staffId===s.id&&a.date.startsWith(salaryMonth)&&a.checkIn);const{base,ot,bonus,total}=calcPay(s,att);return{ชื่อ:s.name,วันทำงาน:att.length,ค่าจ้าง:Math.round(base),OT:Math.round(ot),โบนัส:Math.round(bonus),รวม:Math.round(total)};}), "เงินเดือน","salary");else exportXlsx(attendance.map(a=>{const s=workers.find(x=>x.id===a.staffId);return{ชื่อ:s?.name||a.staffId,วันที่:a.date,เข้า:a.checkIn,ออก:a.checkOut||"-",หมายเหตุ:a.note||""};}), "attendance","attendance");}} onImport={rows=>{const m=rows.map(r=>{const s=workers.find(x=>x.name===r["ชื่อ"]);return{id:Date.now()+Math.random(),staffId:s?.id||"",date:r["วันที่"]||today(),checkIn:r["เข้า"]||"",checkOut:r["ออก"]||"",note:""};}).filter(r=>r.staffId);setAttendance(p=>[...p,...m]);alert(`นำเข้า ${m.length} รายการ`);}} />
        {user.role==="owner"&&tab==="checkin"&&<ClearBtn label="ประวัติการเข้างาน" onClear={()=>setAttendance([])} />}
      </div>} />
      <Tabs tabs={[["checkin","📲 เช็คเข้า-ออก"],["salary","💰 เงินเดือน"],["config","⚙️ ตั้งค่า"]]} active={tab} onChange={setTab} />
      {tab==="checkin"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        {shopLat&&shopLng?<Card style={{background:T.greenLt,borderColor:T.green+"44"}}><div style={{color:T.green,fontWeight:700,fontSize:14}}>📍 ตั้งค่าพิกัดร้านแล้ว</div><div style={{color:T.textSm,fontSize:13}}>พนักงานต้องอยู่ในรัศมีร้านถึงจะเช็คเวลาได้</div></Card>
        :<Card style={{background:T.yellowLt,borderColor:T.yellow+"44"}}><div style={{color:T.yellow,fontWeight:700,fontSize:14}}>⚠️ ยังไม่ได้ตั้งพิกัดร้าน</div><div style={{color:T.textSm,fontSize:13}}>ไปที่ Settings → ร้าน เพื่อตั้งค่าพิกัด</div></Card>}
        <Card><div style={{fontWeight:700,fontSize:14,marginBottom:7}}>🖨 QR Code พนักงาน</div>
          <select value={selStaff} onChange={e=>setSelStaff(e.target.value)} style={{...S.inp,fontSize:15,height:44}}><option value="">— เลือกพนักงาน —</option>{workers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</select>
        </Card>
        {selStaff&&<QRCodeDisplay staffId={selStaff} />}
        <Card><div style={{fontWeight:700,fontSize:14,marginBottom:9}}>📋 เข้างานวันนี้</div>
          {workers.map(s=>{const att=attendance.find(a=>a.staffId===s.id&&a.date===today());const hr2=s.hr||{};const{m:otMins,p:otPay}=att?.checkOut?calcOT(att.checkOut,hr2.otPerHour||50):{m:0,p:0};const finalOt=att?.otOverride!==undefined&&att?.otOverride!==null?att.otOverride:otPay;const editOt=editOtId===s.id;
            return(<div key={s.id} style={{padding:"9px 0",borderBottom:`1px solid ${T.bg}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{fontWeight:600,fontSize:14}}>{s.name}</div>
                  <div style={{color:T.textSm,fontSize:12}}>{att?.checkIn?`เข้า ${att.checkIn}`:"ยังไม่เช็คอิน"}{att?.checkOut?` • ออก ${att.checkOut}`:att?.checkIn?" • ยังอยู่":""}</div>
                  {att?.checkOut&&otMins>0&&<div style={{fontSize:12,marginTop:2}}>
                    <span style={{color:T.orange,fontWeight:600}}>⏰ OT {otMins} นาที = </span>
                    {editOt?(<span style={{display:"inline-flex",alignItems:"center",gap:4}}>
                      <input type="number" value={otVal} onChange={e=>setOtVal(e.target.value)} style={{width:68,padding:"2px 5px",border:`1px solid ${T.orange}`,borderRadius:6,fontSize:12}} autoFocus />
                      <button onClick={()=>{setAttendance(p=>p.map(a=>a.id===att.id?{...a,otOverride:+otVal||0}:a));setEditOtId(null);}} style={{...S.btn(T.green),padding:"2px 7px",fontSize:11}}>✓</button>
                      <button onClick={()=>setEditOtId(null)} style={{...S.ghost,padding:"2px 5px",fontSize:11}}>✕</button>
                    </span>):(
                      <span><span style={{color:T.green,fontWeight:700}}>฿{fmt(finalOt)}</span>{att?.otOverride!==undefined&&att?.otOverride!==null&&<span style={{color:T.textXs,fontSize:11}}> (แก้ไข)</span>}<button onClick={()=>{setOtVal(String(finalOt));setEditOtId(s.id);}} style={{background:"none",border:"none",color:T.orange,cursor:"pointer",fontSize:11,marginLeft:4}}>✏️</button></span>
                    )}
                  </div>}
                  {att?.note&&<div style={{color:T.textXs,fontSize:11,marginTop:1}}>{att.note}</div>}
                </div>
                <div style={{display:"flex",gap:5}}>
                  {!att?.checkIn&&<button onClick={()=>checkIn(s.id)} style={{...S.btn(T.green),fontSize:12,padding:"5px 10px"}}>เข้า</button>}
                  {att?.checkIn&&!att?.checkOut&&<button onClick={()=>checkOut(s.id)} style={{...S.btn(T.red),fontSize:12,padding:"5px 10px"}}>ออก</button>}
                  {att?.checkOut&&<span style={{color:T.green,fontSize:12,fontWeight:600}}>✅</span>}
                </div>
              </div>
            </div>);
          })}
        </Card>
      </div>}
      {tab==="salary"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Card style={{padding:"11px 14px"}}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>เลือกเดือน</div><input type="month" value={salaryMonth} onChange={e=>setSalaryMonth(e.target.value)} style={{...S.inp,fontSize:15}} /></Card>
        {workers.map(s=>{const att=attendance.filter(a=>a.staffId===s.id&&a.date.startsWith(salaryMonth)&&a.checkIn);const{base,ot,bonus,total}=calcPay(s,att);const hr2=s.hr||{};return(
          <Card key={s.id} style={{borderColor:T.orange}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><div style={{fontWeight:800,fontSize:16}}>{s.name}</div><div style={{color:T.orange,fontWeight:900,fontSize:20}}>฿{fmt(total)}</div></div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:7,marginBottom:8}}>
              {[["วันทำงาน",`${att.length} วัน`,T.blue],["OT",`${att.filter(a=>a.checkOut&&calcOT(a.checkOut,hr2.otPerHour||50).m>0).length} วัน`,T.orange],["฿/วัน",`฿${hr2.wage||0}`,T.textMd]].map(([l,v,col])=>(
                <div key={l} style={{background:T.bg,borderRadius:8,padding:"6px",textAlign:"center"}}><div style={{color:T.textXs,fontSize:10}}>{l}</div><div style={{color:col,fontWeight:700,fontSize:13}}>{v}</div></div>
              ))}
            </div>
            <div style={{borderTop:`1px solid ${T.bg}`,paddingTop:7}}>
              {[["💼 ค่าจ้าง",base,T.text],["⏰ OT",ot,T.orange],["🎁 โบนัส",bonus,T.green]].map(([l,v,col])=>v>0&&<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"3px 0",fontSize:13}}><span style={{color:T.textMd}}>{l}</span><span style={{color:col,fontWeight:600}}>฿{fmt(v)}</span></div>)}
              <div style={{display:"flex",justifyContent:"space-between",paddingTop:5,marginTop:3,borderTop:`1px solid ${T.bg}`,fontWeight:800,fontSize:15}}><span>รวม</span><span style={{color:T.orange}}>฿{fmt(total)}</span></div>
            </div>
          </Card>
        );})}
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:14,marginBottom:9}}><span>รวมจ่ายทั้งหมด</span><span style={{color:T.orange}}>฿{fmt(workers.reduce((sum,s)=>{const att=attendance.filter(a=>a.staffId===s.id&&a.date.startsWith(salaryMonth)&&a.checkIn);return sum+calcPay(s,att).total;},0))}</span></div>
          <button onClick={()=>{const lines=["📋 สรุปเงินเดือน "+salaryMonth,"─────────────"];workers.forEach(s=>{const att=attendance.filter(a=>a.staffId===s.id&&a.date.startsWith(salaryMonth)&&a.checkIn);const{base,ot,bonus,total}=calcPay(s,att);lines.push("","👤 "+s.name,`  วัน: ${att.length}  ฐาน: ฿${fmt(base)}${ot>0?`  OT: ฿${fmt(ot)}`:""}${bonus>0?`  โบนัส: ฿${fmt(bonus)}`:""}`,`  รวม: ฿${fmt(total)}`);});lines.push("","─────────────","รวม: ฿"+fmt(workers.reduce((s2,s)=>{const att=attendance.filter(a=>a.staffId===s.id&&a.date.startsWith(salaryMonth)&&a.checkIn);return s2+calcPay(s,att).total;},0)));const msg=lines.join("\n");if(navigator.clipboard)navigator.clipboard.writeText(msg).then(()=>alert("คัดลอกแล้ว!"));else{const ta=document.createElement("textarea");ta.value=msg;document.body.appendChild(ta);ta.select();document.execCommand("copy");document.body.removeChild(ta);alert("คัดลอกแล้ว!");}}} style={{...S.btn(),width:"100%",padding:11,fontSize:14}}>📋 สรุปเงินเดือน (Copy)</button>
        </Card>
      </div>}
      {tab==="config"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        {workers.map(s=>{const hr=s.hr||{wage:350,wageType:"day",otPerHour:50,shiftStart:"09:00",shiftEnd:"18:00",bonusPct:0};const isEdit=editHrId===s.id;const ed=isEdit?hrEdit:hr;return(
          <Card key={s.id}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontWeight:700,fontSize:15}}>{s.name}</div>{!isEdit&&<button onClick={()=>{setEditHrId(s.id);setHrEdit({...hr});}} style={{...S.ghost,fontSize:12,padding:"5px 10px"}}>✏️ แก้ไข</button>}</div>
            {!isEdit&&<div style={{color:T.textSm,fontSize:13,marginTop:3}}>{hr.wageType==="day"?`฿${hr.wage}/วัน`:`฿${hr.wage}/เดือน`} • OT ฿{hr.otPerHour||50}/ชม. • โบนัส {hr.bonusPct||0}%</div>}
            {isEdit&&<><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:9}}>
              <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>ประเภท</div><select value={ed.wageType||"day"} onChange={e=>setHrEdit(p=>({...p,wageType:e.target.value}))} style={{...S.inp,height:40}}><option value="day">รายวัน</option><option value="month">รายเดือน</option></select></div>
              <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>{ed.wageType==="month"?"฿/เดือน":"฿/วัน"}</div><input type="number" value={ed.wage||0} onChange={e=>setHrEdit(p=>({...p,wage:+e.target.value}))} style={S.inp} /></div>
              <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>OT ฿/ชม.</div><input type="number" value={ed.otPerHour??50} onChange={e=>setHrEdit(p=>({...p,otPerHour:+e.target.value}))} style={S.inp} /></div>
              <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>โบนัส %</div><input type="number" step="0.5" value={ed.bonusPct||0} onChange={e=>setHrEdit(p=>({...p,bonusPct:+e.target.value}))} style={S.inp} /></div>
            </div>
            <div style={{display:"flex",gap:8,marginTop:9}}><button onClick={()=>{setStaff(p=>p.map(x=>x.id===s.id?{...x,hr:hrEdit}:x));setEditHrId(null);}} style={{...S.btn(),flex:1}}>บันทึก</button><button onClick={()=>setEditHrId(null)} style={S.ghost}>ยกเลิก</button></div></>}
          </Card>
        );})}
      </div>}
    </div>
  );
}

function SettingsPage({staff,setStaff,lineToken,setLineToken,fixedCosts,setFixedCosts,suppliers,setSuppliers,shopLat,setShopLat,shopLng,setShopLng,shopRadius,setShopRadius}){
  const[tab,setTab]=useState("staff");const[showAddSup,setShowAddSup]=useState(false);const[newSup,setNewSup]=useState({name:"",type:"",phone:"",active:true});const[editSupId,setEditSupId]=useState(null);const[editSupData,setEditSupData]=useState({});const[gettingLoc,setGettingLoc]=useState(false);
  const PERMS={cashflow:"💵 Cash Flow",stock:"📦 สต็อค",purchase:"🛒 สั่งซื้อ",report:"📊 รายงาน",admin:"⚙️ Admin",viewPrice:"💰 ดูราคา",viewFinance:"💎 ดูการเงิน"};
  const getShopLoc=()=>{setGettingLoc(true);navigator.geolocation.getCurrentPosition(p=>{setShopLat(p.coords.latitude.toFixed(6));setShopLng(p.coords.longitude.toFixed(6));setGettingLoc(false);alert(`📍 บันทึกพิกัดร้านแล้ว!\n${p.coords.latitude.toFixed(5)}, ${p.coords.longitude.toFixed(5)}`);},()=>{setGettingLoc(false);alert("ไม่สามารถดึง GPS ได้");},{timeout:8000,enableHighAccuracy:true});};
  return(
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <Hdr title="⚙️ ตั้งค่า" />
      <Tabs tabs={[["staff","👷 พนักงาน"],["shop","🏪 ร้าน"],["supplier","🚚 ซัพฯ"],["fixed","🔒 ต้นทุน"],["line","📲 LINE"]]} active={tab} onChange={setTab} />
      {tab==="staff"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        {staff.filter(s=>s.id!=="owner"&&s.id!=="emergency").map(s=>(
          <Card key={s.id} style={{borderColor:s.active?T.border:T.red+"33"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:9}}>
              <div style={{flex:1}}>
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}><span style={{color:T.textSm,fontSize:12}}>ชื่อ:</span><input value={s.name} onChange={e=>setStaff(p=>p.map(x=>x.id===s.id?{...x,name:e.target.value}:x))} style={{...S.inp,flex:1,fontSize:14,padding:"3px 8px"}} placeholder="ชื่อพนักงาน" /></div>
                <div style={{display:"flex",alignItems:"center",gap:6}}><span style={{color:T.textSm,fontSize:12}}>PIN:</span><input maxLength={4} value={s.pin} onChange={e=>setStaff(p=>p.map(x=>x.id===s.id?{...x,pin:e.target.value.replace(/\D/,"")}:x))} style={{...S.inp,width:72,fontSize:15,letterSpacing:4,textAlign:"center",padding:"3px 7px"}} /></div>
              </div>
              <button onClick={()=>setStaff(p=>p.map(x=>x.id===s.id?{...x,active:!x.active}:x))} style={{background:s.active?T.greenLt:T.redLt,border:`1px solid ${s.active?T.green:T.red}44`,borderRadius:8,padding:"5px 10px",color:s.active?T.green:T.red,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:F}}>{s.active?"ใช้งาน":"ระงับ"}</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:4}}>
              {Object.entries(PERMS).map(([perm,label])=>(
                <button key={perm} onClick={()=>setStaff(p=>p.map(x=>x.id===s.id?{...x,perms:{...x.perms,[perm]:!x.perms[perm]}}:x))} style={{background:s.perms[perm]?T.orange:"transparent",border:`1px solid ${s.perms[perm]?T.orange:T.border}`,borderRadius:7,padding:"5px 3px",color:s.perms[perm]?"#fff":T.textMd,cursor:"pointer",fontSize:9,fontFamily:F,fontWeight:s.perms[perm]?700:400,textAlign:"center"}}>{label}</button>
              ))}
            </div>
          </Card>
        ))}
      </div>}
      {tab==="shop"&&<div style={{display:"flex",flexDirection:"column",gap:12}}>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{color:T.orange,fontWeight:800,fontSize:15,marginBottom:6}}>📍 ตั้งค่าพิกัดร้าน</div>
          <div style={{color:T.textMd,fontSize:13,marginBottom:12}}>ระบบจะตรวจสอบว่าพนักงานอยู่ในรัศมีร้านก่อนเช็คเวลา</div>
          <button onClick={getShopLoc} disabled={gettingLoc} style={{...S.btn(),width:"100%",padding:12,fontSize:15,marginBottom:12}}>{gettingLoc?"⏳ กำลังดึง GPS...":"📍 ใช้ตำแหน่งปัจจุบันเป็นพิกัดร้าน"}</button>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>Latitude</div><input type="number" step="0.000001" value={shopLat||""} onChange={e=>setShopLat(e.target.value)} style={S.inp} placeholder="เช่น 16.428475" /></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>Longitude</div><input type="number" step="0.000001" value={shopLng||""} onChange={e=>setShopLng(e.target.value)} style={S.inp} placeholder="เช่น 101.324125" /></div>
            <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>รัศมีที่อนุญาต (เมตร)</div><input type="number" value={shopRadius||200} onChange={e=>setShopRadius(e.target.value)} style={S.inp} placeholder="200" /></div>
          </div>
          {shopLat&&shopLng&&<div style={{marginTop:10,padding:"8px 12px",background:"#fff",borderRadius:8,fontSize:13,color:T.green,fontWeight:600}}>✅ พิกัดร้าน: {(+shopLat).toFixed(4)}, {(+shopLng).toFixed(4)} • รัศมี {shopRadius||200}m
            <br/><a href={`https://maps.google.com/?q=${shopLat},${shopLng}`} target="_blank" rel="noreferrer" style={{color:T.blue,fontSize:12}}>🗺 ดูบน Google Maps</a>
          </div>}
        </Card>
      </div>}
      {tab==="supplier"&&<div style={{display:"flex",flexDirection:"column",gap:10}}>
        <button onClick={()=>setShowAddSup(!showAddSup)} style={{...S.btn(),width:"100%",padding:11,fontSize:14}}>+ เพิ่มซัพพลายเออร์</button>
        {showAddSup&&<Card style={{borderColor:T.borderOr}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
            <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ชื่อ</div><input value={newSup.name} onChange={e=>setNewSup(p=>({...p,name:e.target.value}))} style={S.inp} /></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>ประเภท</div><input value={newSup.type} onChange={e=>setNewSup(p=>({...p,type:e.target.value}))} style={S.inp} /></div>
            <div><div style={{color:T.textSm,fontSize:13,marginBottom:4}}>โทร</div><input value={newSup.phone} onChange={e=>setNewSup(p=>({...p,phone:e.target.value}))} style={S.inp} /></div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:9}}><button onClick={()=>{if(!newSup.name)return;setSuppliers(p=>[...p,{...newSup,id:Date.now()}]);setNewSup({name:"",type:"",phone:"",active:true});setShowAddSup(false);}} style={{...S.btn(),flex:1}}>บันทึก</button><button onClick={()=>setShowAddSup(false)} style={S.ghost}>ยกเลิก</button></div>
        </Card>}
        {suppliers.map(s=>(
          <Card key={s.id} style={{borderColor:s.active?T.border:T.red+"33"}}>
            {editSupId===s.id?(
              <div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:9}}>
                  <div style={{gridColumn:"1/-1"}}><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>ชื่อ</div><input value={editSupData.name??s.name} onChange={e=>setEditSupData(p=>({...p,name:e.target.value}))} style={S.inp} /></div>
                  <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>ประเภท</div><input value={editSupData.type??s.type} onChange={e=>setEditSupData(p=>({...p,type:e.target.value}))} style={S.inp} /></div>
                  <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>โทร</div><input value={editSupData.phone??s.phone} onChange={e=>setEditSupData(p=>({...p,phone:e.target.value}))} style={S.inp} /></div>
                </div>
                <div style={{display:"flex",gap:8}}><button onClick={()=>{setSuppliers(p=>p.map(x=>x.id===s.id?{...x,...editSupData}:x));setEditSupId(null);setEditSupData({});}} style={{...S.btn(),flex:1}}>บันทึก</button><button onClick={()=>setEditSupId(null)} style={S.ghost}>ยกเลิก</button></div>
              </div>
            ):(
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div><div style={{fontWeight:700,fontSize:14}}>{s.name}</div><div style={{color:T.textSm,fontSize:12}}>{s.type} • {s.phone}</div></div>
                <div style={{display:"flex",gap:5}}>
                  <button onClick={()=>setSuppliers(p=>p.map(x=>x.id===s.id?{...x,active:!x.active}:x))} style={{background:s.active?T.greenLt:T.redLt,border:`1px solid ${s.active?T.green:T.red}44`,borderRadius:7,padding:"4px 8px",color:s.active?T.green:T.red,cursor:"pointer",fontSize:11,fontWeight:700,fontFamily:F}}>{s.active?"เปิด":"ปิด"}</button>
                  <button onClick={()=>{setEditSupId(s.id);setEditSupData({});}} style={{...S.ghost,padding:"4px 8px",fontSize:11}}>✏️</button>
                  <button onClick={()=>{if(window.confirm(`ลบ "${s.name}"?`))setSuppliers(p=>p.filter(x=>x.id!==s.id));}} style={{background:"none",border:"none",color:T.textXs,cursor:"pointer",fontSize:13}}>🗑</button>
                </div>
              </div>
            )}
          </Card>
        ))}
      </div>}
      {tab==="fixed"&&<div style={{display:"flex",flexDirection:"column",gap:10}}>
        {fixedCosts.map((fc,i)=>(
          <Card key={i} style={{padding:"11px 14px"}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr auto auto",gap:9,alignItems:"flex-end"}}>
              <div><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>ชื่อ</div><input value={fc.name} onChange={e=>setFixedCosts(p=>p.map((f,j)=>j===i?{...f,name:e.target.value}:f))} style={{...S.inp,fontSize:14}} /></div>
              <div style={{minWidth:105}}><div style={{color:T.textSm,fontSize:12,marginBottom:3}}>฿/เดือน</div><input type="number" value={fc.amount} onChange={e=>setFixedCosts(p=>p.map((f,j)=>j===i?{...f,amount:+e.target.value||0}:f))} style={{...S.inp,fontSize:14,textAlign:"right"}} /></div>
              <button onClick={()=>setFixedCosts(p=>p.filter((_,j)=>j!==i))} style={{background:T.redLt,border:`1px solid ${T.red}33`,borderRadius:8,padding:"9px 11px",color:T.red,cursor:"pointer",fontSize:13,fontFamily:F}}>🗑</button>
            </div>
          </Card>
        ))}
        <button onClick={()=>setFixedCosts(p=>[...p,{name:"รายการใหม่",amount:0}])} style={{...S.ghost,width:"100%",padding:11,fontSize:14}}>+ เพิ่มรายการ</button>
        <Card style={{background:T.orangeLt,borderColor:T.borderOr}}>
          <div style={{fontWeight:700,fontSize:14,marginBottom:5}}>รวม: <span style={{color:T.orange}}>฿{fmt(fixedCosts.reduce((a,b)=>a+b.amount,0))}</span>/เดือน</div>
          <div style={{display:"flex",gap:16,fontSize:13}}>
            <div><div style={{color:T.textSm}}>BEP/วัน</div><div style={{color:T.orange,fontWeight:700,fontSize:15}}>฿{fmt(Math.ceil(fixedCosts.reduce((a,b)=>a+b.amount,0)/30))}</div></div>
            <div><div style={{color:T.textSm}}>เป้า/วัน (×2.5)</div><div style={{color:T.green,fontWeight:700,fontSize:15}}>฿{fmt(Math.ceil(fixedCosts.reduce((a,b)=>a+b.amount,0)/30*2.5))}</div></div>
          </div>
        </Card>
      </div>}
      {tab==="line"&&<Card>
        <div style={{color:T.orange,fontWeight:800,fontSize:16,marginBottom:12}}>📲 LINE Notify Token</div>
        <div style={{color:T.textSm,fontSize:13,marginBottom:8}}>รับ Token จาก notify-bot.line.me</div>
        <input type="password" value={lineToken} onChange={e=>setLineToken(e.target.value)} style={S.inp} placeholder="ใส่ Token ที่นี่..." />
        {lineToken&&<div style={{color:T.green,fontSize:13,marginTop:6,fontWeight:600}}>✅ Token บันทึกแล้ว</div>}
      </Card>}
    </div>
  );
}

export default function App(){
  const[user,setUser]=useState(null);const[page,setPage]=useState("dashboard");
  const[stock,setStock]=useState(INIT_STOCK);const[cf,setCF]=useState([]);const[movements,setMovements]=useState([]);
  const[staff,setStaff]=useState(INIT_STAFF);const[suppliers,setSuppliers]=useState(INIT_SUPS);
  const[fixedCosts,setFixedCosts]=useState(INIT_FIXED);const[lineToken,setLineToken]=useState("");
  const[waste,setWaste]=useState([]);const[promos,setPromos]=useState([]);const[attendance,setAttendance]=useState([]);
  const[shopLat,setShopLat]=useState("");const[shopLng,setShopLng]=useState("");const[shopRadius,setShopRadius]=useState("200");
  const[dbReady,setDbReady]=useState(false);const[loading,setLoading]=useState(true);

  useEffect(()=>{async function load(){setLoading(true);try{const[cfD,stD,mvD]=await Promise.all([db.getCF(),db.getStock(),db.getMvs()]);if(cfD?.length)setCF(cfD.map(r=>({id:r.id,date:r.date,flow:r.flow,cat:r.cat,itemName:r.item_name||"",amount:r.amount,method:r.method,note:r.note||"",branch:r.branch||"main",staffId:r.staff_id||"owner"})));if(stD?.length)setStock(stD.map(r=>({id:r.id,name:r.name,unit:r.unit,qty:r.qty,minQty:r.min_qty,dailyUse:r.daily_use,supplierId:r.supplier_id||1,costHistory:r.cost_history||[]})));if(mvD?.length)setMovements(mvD.map(r=>({id:r.id,itemId:r.item_id,type:r.type,qty:r.qty,unitCost:r.unit_cost||0,date:r.date,staffId:r.staff_id||"",note:r.note||"",branch:r.branch||"main"})));setDbReady(true);}catch{}setLoading(false);}load();},[]);

  const saveStock=useCallback(async ns=>{setStock(ns);if(dbReady)db.upsertStock(ns.map(s=>({id:s.id,name:s.name,unit:s.unit,qty:s.qty,min_qty:s.minQty,daily_use:s.dailyUse,supplier_id:s.supplierId||1,cost_history:s.costHistory||[]}))).catch(()=>{});},[dbReady]);

  const urlSid=new URLSearchParams(window.location.search).get("sid");
  if(urlSid)return <CheckinPage staff={staff} shopLat={shopLat} shopLng={shopLng} shopRadius={shopRadius} onCheckin={(sid,type,time,date,gps)=>{if(type==="in")setAttendance(p=>[...p,{id:Date.now(),staffId:sid,date,checkIn:time,checkOut:"",note:gps?`📍 ${gps.lat},${gps.lng} ±${gps.acc}m`:""}]);else setAttendance(p=>p.map(a=>a.staffId===sid&&a.date===date&&!a.checkOut?{...a,checkOut:time}:a));}} />;

  if(!user)return <LoginPage staff={staff} onLogin={u=>{setUser(u);setPage(u.role==="owner"?"dashboard":"staffcheckin");}}/>;

  if(user.id==="emergency")return(
    <div style={{minHeight:"100vh",background:T.bg,fontFamily:F,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,gap:14}}>
      <div style={{width:76,height:76,borderRadius:20,background:T.red,display:"flex",alignItems:"center",justifyContent:"center",fontSize:42}}>🚨</div>
      <div style={{color:T.red,fontWeight:900,fontSize:20}}>Emergency Mode</div>
      <div style={{display:"flex",flexDirection:"column",gap:9,width:"100%",maxWidth:310}}>
        <button onClick={()=>{if(window.confirm("ล้าง Cash Flow?")){{setCF([]);if(dbReady)db.clearCF().catch(()=>{});alert("เรียบร้อย");}}}} style={{...S.btn(T.red),padding:13,fontSize:15,width:"100%"}}>🗑 ล้าง Cash Flow</button>
        <button onClick={()=>{if(window.confirm("ล้างสต็อค?")){{setStock(INIT_STOCK);setMovements([]);if(dbReady)db.clearMvs().catch(()=>{});alert("เรียบร้อย");}}}} style={{...S.btn(T.orange),padding:13,fontSize:15,width:"100%"}}>🗑 ล้างสต็อค</button>
        <button onClick={()=>{if(window.confirm("⚠️ ล้างทั้งหมด?")){{setCF([]);setStock(INIT_STOCK);setMovements([]);if(dbReady){db.clearCF().catch(()=>{});db.clearMvs().catch(()=>{});}alert("เรียบร้อย");}}}} style={{...S.btn(T.red),padding:13,fontSize:15,width:"100%",border:"2px solid #7f1d1d"}}>⚠️ ล้างทั้งหมด</button>
        <button onClick={()=>setUser(null)} style={{...S.ghost,padding:13,fontSize:15,width:"100%"}}>← ออก</button>
      </div>
    </div>
  );

  if(loading)return(
    <div style={{minHeight:"100vh",background:T.bg,fontFamily:F,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:14}}>
      <div style={{width:76,height:76,borderRadius:20,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:42}}>🫕</div>
      <div style={{color:T.orange,fontWeight:800,fontSize:18}}>กำลังโหลดข้อมูล...</div>
    </div>
  );

  const p=user.perms;const isOwner=user.role==="owner";
  const nav=isOwner?[
    {id:"dashboard",icon:"🏠",label:"หลัก"},{id:"cashflow",icon:"💵",label:"Cash Flow"},
    {id:"stock",icon:"📦",label:"สต็อค"},{id:"purchase",icon:"🛒",label:"สั่งซื้อ"},
    {id:"report",icon:"📊",label:"รายงาน"},{id:"hr",icon:"👥",label:"HR"},
    {id:"settings",icon:"⚙️",label:"ตั้งค่า"},
  ]:[
    {id:"dashboard",icon:"🏠",label:"หลัก"},
    {id:"staffcheckin",icon:"📲",label:"เช็คเวลา"},
    ...(p.cashflow?[{id:"cashflow",icon:"💵",label:"Cash Flow"}]:[]),
    ...(p.stock?[{id:"staffstock",icon:"📦",label:"สต็อค"}]:[]),
    ...(p.purchase?[{id:"purchase",icon:"🛒",label:"สั่งซื้อ"}]:[]),
    ...(p.report?[{id:"report",icon:"📊",label:"รายงาน"}]:[]),
    ...(p.admin?[{id:"settings",icon:"⚙️",label:"ตั้งค่า"}]:[]),
  ];

  const pages={
    staffcheckin:<StaffCheckinPage user={user} attendance={attendance} setAttendance={setAttendance} shopLat={shopLat} shopLng={shopLng} shopRadius={shopRadius} />,
    dashboard:<DashboardPage cf={cf} stock={stock} user={user} fixedCosts={fixedCosts} waste={waste} promos={promos} setPage={setPage} />,
    cashflow:<CashflowPage cf={cf} setCF={setCF} user={user} dbReady={dbReady} />,
    stock:<StockPage stock={stock} setStock={saveStock} movements={movements} setMovements={setMovements} user={user} suppliers={suppliers} />,
    staffstock:<StaffStockPage stock={stock} setStock={saveStock} movements={movements} setMovements={setMovements} user={user} />,
    purchase:<PurchasePage stock={stock} suppliers={suppliers} lineToken={lineToken} />,
    report:<ReportPage cf={cf} stock={stock} movements={movements} user={user} fixedCosts={fixedCosts} waste={waste} setWaste={setWaste} promos={promos} setPromos={setPromos} />,
    hr:<HRPage staff={staff} setStaff={setStaff} attendance={attendance} setAttendance={setAttendance} cf={cf} shopLat={shopLat} shopLng={shopLng} />,
    settings:<SettingsPage staff={staff} setStaff={setStaff} lineToken={lineToken} setLineToken={setLineToken} fixedCosts={fixedCosts} setFixedCosts={setFixedCosts} suppliers={suppliers} setSuppliers={setSuppliers} shopLat={shopLat} setShopLat={setShopLat} shopLng={shopLng} setShopLng={setShopLng} shopRadius={shopRadius} setShopRadius={setShopRadius} />,
  };

  return(
    <div style={{minHeight:"100vh",background:T.bg,fontFamily:F,fontSize:16}}>
      <div style={{background:"#fff",borderBottom:`1px solid ${T.border}`,padding:"10px 16px",display:"flex",alignItems:"center",gap:11,position:"sticky",top:0,zIndex:50,boxShadow:"0 1px 3px rgba(0,0,0,.08)"}}>
        <div style={{width:34,height:34,borderRadius:9,background:`linear-gradient(135deg,${T.orange},${T.orangeDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:19}}>🫕</div>
        <div><div style={{color:T.orange,fontWeight:900,fontSize:14}}>ไท่กั๋วหม่าล่า</div><div style={{color:T.textXs,fontSize:11}}>{user.name} • {isOwner?"👑":"👷"}{dbReady?" ● DB":""}</div></div>
        <button onClick={()=>setUser(null)} style={{...S.ghost,fontSize:13,padding:"4px 10px",marginLeft:"auto"}}>ออก</button>
      </div>
      <div style={{padding:"16px 14px 90px",maxWidth:900,margin:"0 auto"}}>{pages[page]||pages.dashboard}</div>
      <div style={{position:"fixed",bottom:0,left:0,right:0,background:"#fff",borderTop:`1px solid ${T.border}`,display:"flex",justifyContent:"space-around",padding:"8px 0 14px",zIndex:100,boxShadow:"0 -2px 8px rgba(0,0,0,.06)"}}>
        {nav.map(n=>(
          <button key={n.id} onClick={()=>setPage(n.id)} style={{background:"none",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2,padding:"2px 6px",minWidth:42}}>
            <span style={{fontSize:20}}>{n.icon}</span>
            <span style={{fontSize:9,color:page===n.id?T.orange:T.textXs,fontWeight:page===n.id?800:400}}>{n.label}</span>
            {page===n.id&&<div style={{width:16,height:3,borderRadius:2,background:T.orange}}/>}
          </button>
        ))}
      </div>
    </div>
  );
}
