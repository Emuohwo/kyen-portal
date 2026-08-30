import { useState, useEffect, useCallback, useRef } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, PieChart, Pie, Cell, LineChart, Line,
} from "recharts";

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

const DEFAULT_SKUS = [
  { sku:"KYN-ONP-150g CUP",       name:"Onion Powder 150g Cup",       cat:"Onion Powder", sort_order:1  },
  { sku:"KYN-ONP-250g CUP",       name:"Onion Powder 250g Cup",       cat:"Onion Powder", sort_order:2  },
  { sku:"KYN-ONP 80g CUP",        name:"Onion Powder 80g Cup",        cat:"Onion Powder", sort_order:3  },
  { sku:"KYN-ONF-80g POUCH",      name:"Onion Flakes 80g Pouch",      cat:"Onion Flakes", sort_order:4  },
  { sku:"KYEN-DFS 350ML GLASS",   name:"Date Syrup 350ML Glass",      cat:"Date Syrup",   sort_order:5  },
  { sku:"KYEN-DFS 350ML PLASTIC", name:"Date Syrup 350ML Plastic",    cat:"Date Syrup",   sort_order:6  },
  { sku:"KYEN-DFS 50ML PLASTIC",  name:"Date Syrup 500ML Plastic",    cat:"Date Syrup",   sort_order:7  },
  { sku:"KYEN-DFS 1LTR. PLASTIC", name:"Date Syrup 1Ltr Plastic",     cat:"Date Syrup",   sort_order:8  },
  { sku:"KYN-C/DF 150g POUCH",    name:"Coconut & Date Flakes 150g",  cat:"C&D Flakes",   sort_order:9  },
  { sku:"KYN-C/DF 70g POUCH",     name:"Coconut & Date Flakes 70g",   cat:"C&D Flakes",   sort_order:10 },
  { sku:"KYN-C/DF 500g POUCH",    name:"C&D Flakes 500g Sachet",      cat:"C&D Flakes",   sort_order:11 },
  { sku:"KYN-C/DF 500g CUP",      name:"C&D Flakes 500g Cup",         cat:"C&D Flakes",   sort_order:12 },
];

const CHANNELS     = ["Supermarket","Retail Store","Wholesaler","End User"];
const PAYMENT_TYPES= ["Cash","Credit","Bank Transfer"];
const PAY_COL      = {Cash:G1,Credit:"#e74c3c","Bank Transfer":"#2980B9"};
const VISIT_TYPES = ["First Visit","Follow-up","Merchandising"];
const ZONES       = ["South South","South East","South West","North Central","North East","North West","FCT"];
const MFULL = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const DAYS  = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

const G1="#1A5C38"; const G2="#2E7D52"; const G3="#4CAF7D";
const AMB="#F4A900"; const GL="#D6EAD8"; const LG="#F7F9F7";
const PIE=[G1,G2,G3,AMB,"#E67E22","#8E44AD","#2980B9","#E74C3C"];
const STATUS_COL={ planned:AMB, completed:G1, cancelled:"#e74c3c" };
const VTYPE_COL={ "First Visit":G1, "Follow-up":G2, "Merchandising":AMB };

const fmt   = n  => "₦"+Number(n||0).toLocaleString("en-NG");
const pct   = (a,b)=> b>0?Math.min(Math.round(a/b*100),999):0;
const uid   = ()  => Math.random().toString(36).slice(2,11);
const isoD  = d   => d.toISOString().split("T")[0];
const mRange= (y,m)=>({ start:`${y}-${String(m+1).padStart(2,"0")}-01`, end:isoD(new Date(y,m+1,0)) });
const getP  = (prices,sku,ch)=>{ const p=prices?.[sku]; if(!p) return 0; return typeof p==="object"?(ch==="Wholesaler"?(p.wholesale||0):ch==="End User"?(p.end_user||0):(p.retail||0)):Number(p); };
const normPx= (raw,skus)=>{ const o={}; skus.forEach(s=>{ const p=raw?.[s.sku]; o[s.sku]=p&&typeof p==="object"?{retail:p.retail||0,wholesale:p.wholesale||0,end_user:p.end_user||0}:{retail:Number(p)||0,wholesale:Number(p)||0,end_user:Number(p)||0}; }); return o; };

// DB converters
const cfgFD =(r,sk)=>({ adminPassword:r?.admin_password||"kyen2024", visitBenchmark:r?.visit_benchmark||5, prices:normPx(r?.prices,sk) });
const cfgTD = c=>({ id:1, admin_password:c.adminPassword, visit_benchmark:c.visitBenchmark, prices:c.prices });
const saleFD= r=>({ id:r.id,repId:r.rep_id,repName:r.rep_name,date:r.date,channel:r.channel,customerId:r.customer_id,customerName:r.customer_name,sku:r.sku,qty:r.qty,price:r.price,discountPct:r.discount_pct||0,paymentType:r.payment_type||"Cash",amount:r.amount });
const saleTD= s=>({ id:s.id,rep_id:s.repId,rep_name:s.repName,date:s.date,channel:s.channel,customer_id:s.customerId||null,customer_name:s.customerName||null,sku:s.sku,qty:s.qty,price:s.price,discount_pct:s.discountPct||0,payment_type:s.paymentType||"Cash",amount:s.amount });
const visitFD=r=>({ id:r.id,repId:r.rep_id,repName:r.rep_name,date:r.date,customerId:r.customer_id,visitType:r.visit_type,notes:r.notes });
const visitTD=v=>({ id:v.id,rep_id:v.repId,rep_name:v.repName,date:v.date,customer_id:v.customerId||null,visit_type:v.visitType,notes:v.notes||"" });
const planFD =r=>({ id:r.id,repId:r.rep_id,repName:r.rep_name,plannedDate:r.planned_date,customerId:r.customer_id,customerName:r.customer_name,visitType:r.visit_type,notes:r.notes,status:r.status });
const planTD =p=>({ id:p.id,rep_id:p.repId,rep_name:p.repName,planned_date:p.plannedDate,customer_id:p.customerId||null,customer_name:p.customerName||null,visit_type:p.visitType,notes:p.notes||"",status:p.status });
const tgtsFD = rows=>{ const m={}; (rows||[]).forEach(r=>{ m[`${r.rep_id}_${r.year}_${r.month}`]={ valueNGN:r.value_ngn,visitsTarget:r.visits_target,skuTargets:r.sku_targets||{} }; }); return m; };

// Export
function dlXLSX(sheets,fn){ const wb=XLSX.utils.book_new(); sheets.forEach(({name,data})=>{ const ws=XLSX.utils.json_to_sheet(data); ws["!cols"]=data.length?Object.keys(data[0]).map(k=>({wch:Math.max(k.length,...data.map(r=>String(r[k]??"").length))+2})):[];  XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31)); }); XLSX.writeFile(wb,fn+".xlsx"); }
function dlCSV(rows,fn){ if(!rows.length) return; const h=Object.keys(rows[0]); const csv=[h.join(","),...rows.map(r=>h.map(k=>{const v=r[k]??"";return String(v).includes(",")?`"${v}"`:v;}).join(","))].join("\n"); const a=document.createElement("a"); a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"})); a.download=fn+".csv"; a.click(); }

// ── UI Primitives ─────────────────────────────────────────────────────────
const Inp=({label,value,onChange,type="text",placeholder="",className=""})=>(
  <div className={className}>
    {label&&<label className="block text-xs font-semibold text-gray-500 mb-1">{label}</label>}
    <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} className="w-full border rounded-xl p-2.5 text-sm outline-none focus:border-green-600"/>
  </div>
);
const Sel=({label,value,onChange,children,className=""})=>(
  <div className={className}>
    {label&&<label className="block text-xs font-semibold text-gray-500 mb-1">{label}</label>}
    <select value={value} onChange={e=>onChange(e.target.value)} className="w-full border rounded-xl p-2.5 text-sm outline-none bg-white">{children}</select>
  </div>
);
function Btn({children,onClick,disabled,color=G1,outline=false,sm=false,full=false}){
  return <button onClick={onClick} disabled={disabled} style={{background:outline?"transparent":(disabled?"#ccc":color),border:outline?`1.5px solid ${color}`:"none",color:outline?color:"white",opacity:disabled?.6:1}} className={`rounded-xl font-bold transition-all ${sm?"px-3 py-1.5 text-xs":"px-4 py-2.5 text-sm"} ${full?"w-full":""}`}>{children}</button>;
}
function Modal({title,children,onClose}){
  return <div className="fixed inset-0 z-50 flex items-center justify-center" style={{background:"rgba(0,0,0,0.45)"}}>
    <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg mx-4 max-h-[90vh] overflow-y-auto">
      <div style={{background:G1}} className="px-5 py-3 rounded-t-2xl flex justify-between items-center sticky top-0">
        <div className="text-white font-bold text-sm">{title}</div>
        <button onClick={onClose} className="text-white text-xl font-bold opacity-70 hover:opacity-100">×</button>
      </div>
      <div className="p-5">{children}</div>
    </div>
  </div>;
}
function Badge({children,color=G1}){ return <span style={{background:color+"22",color}} className="text-xs font-bold px-2 py-0.5 rounded-full whitespace-nowrap">{children}</span>; }
function StatCard({label,value,sub,pctVal,color=G1}){
  const p=pctVal!==undefined?Math.min(pctVal,100):null; const bc=p===null?G1:p>=100?G1:p>=60?G2:AMB;
  return <div className="bg-white rounded-xl shadow p-4 flex flex-col gap-1">
    <div className="text-xs font-semibold uppercase tracking-widest text-gray-400">{label}</div>
    <div style={{color}} className="text-2xl font-black leading-none">{value}</div>
    {sub&&<div className="text-xs text-gray-400">{sub}</div>}
    {p!==null&&<><div className="flex justify-between text-xs text-gray-400 mt-1"><span>vs target</span><span style={{color:bc,fontWeight:700}}>{pctVal}%</span></div><div className="h-1.5 rounded-full bg-gray-100"><div style={{width:`${p}%`,background:bc,transition:"width .4s"}} className="h-full rounded-full"/></div></>}
  </div>;
}
function Empty({text}){ return <div className="h-40 flex flex-col items-center justify-center text-gray-300"><div className="text-3xl mb-2">📊</div><div className="text-sm">{text}</div></div>; }

// ── Customer Selector ────────────────────────────────────────────────────
function CustomerSelector({customers,value,onChange,onNewCustomer,label="Customer / Outlet"}){
  const [open,setOpen]=useState(false); const [search,setSearch]=useState(""); const [adding,setAdding]=useState(false); const ref=useRef(null);
  useEffect(()=>{ const h=e=>{ if(ref.current&&!ref.current.contains(e.target))setOpen(false); }; document.addEventListener("mousedown",h); return ()=>document.removeEventListener("mousedown",h); },[]);
  const filtered=customers.filter(c=>c.name.toLowerCase().includes(search.toLowerCase())||(c.city||"").toLowerCase().includes(search.toLowerCase())||(c.zone||"").toLowerCase().includes(search.toLowerCase()));
  const selected=customers.find(c=>c.id===value);
  if(adding) return <div>{label&&<label className="block text-xs font-semibold text-gray-500 mb-1">{label}</label>}<InlineCustomerForm onSave={async c=>{await onNewCustomer(c);onChange(c.id);setAdding(false);}} onCancel={()=>setAdding(false)}/></div>;
  return <div ref={ref}>{label&&<label className="block text-xs font-semibold text-gray-500 mb-1">{label}</label>}
    <div className="relative">
      <div onClick={()=>setOpen(o=>!o)} style={{borderColor:open?G1:""}} className="w-full border rounded-xl p-2.5 text-sm cursor-pointer flex justify-between items-center">
        {selected?<span className="font-semibold text-gray-800">{selected.name}<span className="font-normal text-gray-400 text-xs ml-1">· {selected.city}</span></span>:<span className="text-gray-400">Select or search customer…</span>}
        <span className="text-gray-400 text-xs ml-2">{open?"▲":"▼"}</span>
      </div>
      {open&&<div className="absolute z-30 w-full mt-1 bg-white rounded-xl shadow-xl border max-h-60 overflow-y-auto">
        <div className="p-2 border-b sticky top-0 bg-white"><input autoFocus value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search name, city or zone…" className="w-full border rounded-lg p-2 text-xs outline-none"/></div>
        {filtered.map(c=><div key={c.id} onClick={()=>{onChange(c.id);setOpen(false);setSearch("");}} className="px-3 py-2 hover:bg-gray-50 cursor-pointer">
          <div className="text-sm font-semibold">{c.name}</div>
          <div className="text-xs text-gray-400">{[c.type,c.city,c.zone].filter(Boolean).join(" · ")}</div>
        </div>)}
        {filtered.length===0&&<div className="px-3 py-2 text-xs text-gray-400">No match found.</div>}
        <div onClick={()=>{setOpen(false);setAdding(true);}} style={{color:G1}} className="px-3 py-2.5 border-t text-xs font-bold cursor-pointer hover:bg-gray-50">+ Create new customer</div>
      </div>}
    </div>
  </div>;
}

function InlineCustomerForm({onSave,onCancel,initial={}}){
  const [f,setF]=useState({name:"",type:"Supermarket",address:"",contactName:"",contactPhone:"",city:"",zone:"South South",...initial});
  const upd=k=>v=>setF(p=>({...p,[k]:v}));
  const valid=f.name.trim()&&f.city.trim();
  return <div style={{background:GL}} className="rounded-xl p-4 space-y-3">
    <div className="text-xs font-black text-gray-700">New Customer Details</div>
    <div className="grid grid-cols-2 gap-3">
      <Inp label="Customer Name *" value={f.name} onChange={upd("name")} className="col-span-2"/>
      <Sel label="Type" value={f.type} onChange={upd("type")}>{CHANNELS.map(c=><option key={c}>{c}</option>)}</Sel>
      <Sel label="Zone" value={f.zone} onChange={upd("zone")}>{ZONES.map(z=><option key={z}>{z}</option>)}</Sel>
      <Inp label="City *" value={f.city} onChange={upd("city")} placeholder="e.g. Port Harcourt"/>
      <Inp label="Address" value={f.address} onChange={upd("address")} placeholder="Street / area"/>
      <Inp label="Contact Name" value={f.contactName} onChange={upd("contactName")}/>
      <Inp label="Contact Phone" type="tel" value={f.contactPhone} onChange={upd("contactPhone")}/>
    </div>
    <div className="flex gap-2 pt-1">
      <Btn onClick={()=>onSave({id:uid(),...f})} disabled={!valid} sm>Save Customer</Btn>
      <Btn onClick={onCancel} outline sm color="#999">Cancel</Btn>
    </div>
  </div>;
}

// ── Main App ──────────────────────────────────────────────────────────────
export default function App(){
  const [loaded,setLoaded]=useState(false); const [user,setUser]=useState(null); const [view,setView]=useState("dashboard");
  const [selMonth,setSelMonth]=useState(new Date().getMonth()); const [selYear,setSelYear]=useState(new Date().getFullYear());
  const [fetching,setFetching]=useState(false);
  const [isMobile,setIsMobile]=useState(()=>window.innerWidth<768);
  const [cfg,setCfg]=useState({adminPassword:"kyen2024",visitBenchmark:5,prices:{}});
  const [team,setTeam]=useState([]); const [skus,setSkus]=useState(DEFAULT_SKUS);
  const [customers,setCustomers]=useState([]); const [targets,setTargets]=useState({});
  const [monthSales,setMonthSales]=useState([]); const [monthVisits,setMonthVisits]=useState([]);
  const [visitPlans,setVisitPlans]=useState([]);

  useEffect(()=>{
    const handler=()=>setIsMobile(window.innerWidth<768);
    window.addEventListener("resize",handler);
    return ()=>window.removeEventListener("resize",handler);
  },[]);

  useEffect(()=>{
    (async()=>{
      // Always upsert default SKUs first — ensures they survive any re-deployment
      await supabase.from("skus").upsert(
        DEFAULT_SKUS.map(s=>({...s,is_active:true})),
        {onConflict:"sku",ignoreDuplicates:true}
      );
      const [cfgR,teamR,skuR,custR,tgtR]=await Promise.all([
        supabase.from("config").select("*").eq("id",1).single(),
        supabase.from("team").select("*").order("name"),
        supabase.from("skus").select("*").eq("is_active",true).order("sort_order"),
        supabase.from("customers").select("*").order("name"),
        supabase.from("targets").select("*"),
      ]);
      const activeSkus=skuR.data?.length?skuR.data:DEFAULT_SKUS;
      setSkus(activeSkus);
      if(cfgR.data) setCfg(cfgFD(cfgR.data,activeSkus));
      if(teamR.data) setTeam(teamR.data);
      if(custR.data) setCustomers(custR.data);
      if(tgtR.data) setTargets(tgtsFD(tgtR.data));
      setLoaded(true);
    })();
  },[]);

  const fetchMonth=useCallback(async(month,year)=>{
    setFetching(true);
    const {start,end}=mRange(year,month);
    const [sR,vR,pR]=await Promise.all([
      supabase.from("sales").select("*").gte("date",start).lte("date",end).order("date"),
      supabase.from("visits").select("*").gte("date",start).lte("date",end).order("date"),
      supabase.from("visit_plans").select("*").gte("planned_date",start).lte("planned_date",end).order("planned_date"),
    ]);
    setMonthSales((sR.data||[]).map(saleFD));
    setMonthVisits((vR.data||[]).map(visitFD));
    setVisitPlans((pR.data||[]).map(planFD));
    setFetching(false);
  },[]);

  useEffect(()=>{ if(loaded) fetchMonth(selMonth,selYear); },[loaded,selMonth,selYear]);

  useEffect(()=>{
    if(!loaded) return;
    const {start,end}=mRange(selMonth,selYear);
    const ch=supabase.channel("kyen-rt")
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"sales"},({new:r})=>{ if(r.date>=start&&r.date<=end) setMonthSales(p=>p.find(x=>x.id===r.id)?p:[...p,saleFD(r)]); })
      .on("postgres_changes",{event:"DELETE",schema:"public",table:"sales"},({old:r})=>setMonthSales(p=>p.filter(s=>s.id!==r.id)))
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"visits"},({new:r})=>{ if(r.date>=start&&r.date<=end) setMonthVisits(p=>p.find(x=>x.id===r.id)?p:[...p,visitFD(r)]); })
      .on("postgres_changes",{event:"*",schema:"public",table:"visit_plans"},()=>fetchMonth(selMonth,selYear))
      .subscribe();
    return ()=>supabase.removeChannel(ch);
  },[loaded,selMonth,selYear]);

  // DB ops
  const saveCfg=async c=>{setCfg(c);await supabase.from("config").upsert(cfgTD(c));};
  const addMember=async m=>{const{data}=await supabase.from("team").insert(m).select().single();if(data)setTeam(t=>[...t,data]);return data;};
  const removeMember=async id=>{await supabase.from("team").delete().eq("id",id);setTeam(t=>t.filter(m=>m.id!==id));};
  const addSku=async s=>{await supabase.from("skus").insert({...s,is_active:true});setSkus(p=>[...p,s]);};
  const deactivateSku=async sku=>{await supabase.from("skus").update({is_active:false}).eq("sku",sku);setSkus(p=>p.filter(s=>s.sku!==sku));};
  const addCustomer=async c=>{const{data}=await supabase.from("customers").insert(c).select().single();if(data)setCustomers(p=>[...p,data]);return data;};
  const updateCustomer=async c=>{await supabase.from("customers").update(c).eq("id",c.id);setCustomers(p=>p.map(x=>x.id===c.id?{...x,...c}:x));};
  const fetchCustHistory=async id=>{const{data}=await supabase.from("sales").select("*").eq("customer_id",id).order("date",{ascending:false});return(data||[]).map(saleFD);};
  const addSales=async entries=>{await supabase.from("sales").insert(entries.map(saleTD));setMonthSales(p=>[...p,...entries]);};
  const deleteSale=async id=>{await supabase.from("sales").delete().eq("id",id);setMonthSales(p=>p.filter(s=>s.id!==id));};
  const addVisit=async v=>{await supabase.from("visits").insert(visitTD(v));setMonthVisits(p=>[...p,v]);};
  const addPlan=async p=>{await supabase.from("visit_plans").insert(planTD(p));setVisitPlans(prev=>[...prev,p]);};
  const updatePlanStatus=async(id,status)=>{await supabase.from("visit_plans").update({status}).eq("id",id);setVisitPlans(p=>p.map(x=>x.id===id?{...x,status}:x));};
  const deletePlan=async id=>{await supabase.from("visit_plans").delete().eq("id",id);setVisitPlans(p=>p.filter(x=>x.id!==id));};
  const saveTarget=async(repId,year,month,vals)=>{
    const id=`${repId}_${year}_${month}`;
    await supabase.from("targets").upsert({id,rep_id:repId,year,month,value_ngn:vals.valueNGN,visits_target:vals.visitsTarget,sku_targets:vals.skuTargets});
    setTargets(t=>({...t,[id]:vals}));
  };

  if(!loaded) return <div style={{background:G1}} className="flex items-center justify-center h-screen"><div className="text-center"><div className="text-white text-2xl font-black tracking-widest mb-2">KYEN</div><div style={{color:AMB}} className="text-sm animate-pulse">Loading…</div></div></div>;
  if(!user) return <Login cfg={cfg} team={team} onLogin={async u=>{
    setUser(u);
    setView("dashboard");
    if(u.role==="rep"){
      const now=new Date().toISOString();
      await supabase.from("team").update({last_login:now}).eq("id",u.id);
      setTeam(t=>t.map(m=>m.id===u.id?{...m,last_login:now}:m));
    }
  }}/>;

  const isAdmin=user.role==="admin";
  const mySales =isAdmin?monthSales:monthSales.filter(s=>s.repId===user.id);
  const myVisits=isAdmin?monthVisits:monthVisits.filter(v=>v.repId===user.id);
  const myPlans =isAdmin?visitPlans:visitPlans.filter(p=>p.repId===user.id);

  return <div className="flex h-screen bg-gray-100 overflow-hidden" style={{fontFamily:"Inter,system-ui,sans-serif"}}>
    {/* Sidebar — desktop only */}
    {!isMobile&&<Sidebar isAdmin={isAdmin} view={view} setView={setView} user={user} onLogout={()=>{setUser(null);setView("dashboard");}}/>}

    <div className="flex-1 flex flex-col overflow-hidden">
      <TopBar selMonth={selMonth} setSelMonth={setSelMonth} selYear={selYear} setSelYear={setSelYear} fetching={fetching} isMobile={isMobile}/>
      <main className="flex-1 overflow-auto p-3 pb-24">
        {view==="dashboard"&&isAdmin&&<AdminDash monthSales={monthSales} monthVisits={monthVisits} team={team} targets={targets} selMonth={selMonth} selYear={selYear} skus={skus}/>}
        {view==="dashboard"&&!isAdmin&&<RepDash user={user} monthSales={mySales} monthVisits={myVisits} targets={targets} selMonth={selMonth} selYear={selYear} skus={skus}/>}
        {view==="log_sale"&&<LogSale user={user} cfg={cfg} skus={skus} customers={customers} onSave={async e=>{await addSales(e);setView("dashboard");}} onNewCustomer={addCustomer}/>}
        {view==="log_visits"&&<LogVisits user={user} customers={customers} onSave={async v=>{await addVisit(v);setView("dashboard");}} onNewCustomer={addCustomer}/>}
        {view==="all_sales"&&<AllSalesView sales={isAdmin?monthSales:mySales} customers={customers} team={team} onDelete={isAdmin?deleteSale:null} month={selMonth} year={selYear} skus={skus}/>}
        {view==="visits_log"&&<VisitsLog visits={isAdmin?monthVisits:myVisits} customers={customers} team={team} isAdmin={isAdmin} month={selMonth} year={selYear}/>}
        {view==="customers"&&<CustomerList customers={customers} onAdd={addCustomer} onEdit={isAdmin?updateCustomer:null} monthSales={monthSales} fetchCustHistory={fetchCustHistory}/>}
        {view==="my_account"&&!isAdmin&&<RepAccount user={user} onChangePassword={async(oldPw,newPw)=>{const match=team.find(t=>t.id===user.id&&t.password===oldPw);if(!match)return"Current password is incorrect.";await supabase.from("team").update({password:newPw}).eq("id",user.id);setTeam(t=>t.map(m=>m.id===user.id?{...m,password:newPw}:m));return null;}}/>}
        {view==="visit_plan"&&<VisitPlanView user={user} isAdmin={isAdmin} visitPlans={myPlans} customers={customers} team={team} onAdd={addPlan} onStatus={async(id,status,notes)=>{
          await updatePlanStatus(id,status);
          if(status==="completed"){
            const plan=visitPlans.find(p=>p.id===id);
            if(plan) await addVisit({id:uid(),repId:plan.repId,repName:plan.repName,date:plan.plannedDate,customerId:plan.customerId,visitType:plan.visitType,notes:notes||plan.notes||""});
          }
        }} onDelete={deletePlan} onNewCustomer={addCustomer} selMonth={selMonth} selYear={selYear}/>}
        {view==="team"&&isAdmin&&<TeamMgmt team={team} onAdd={addMember} onRemove={removeMember}/>}
        {view==="skus"&&isAdmin&&<SKUManager skus={skus} onAdd={addSku} onDeactivate={deactivateSku}/>}
        {view==="prices"&&isAdmin&&<PricesMgmt cfg={cfg} skus={skus} onSave={saveCfg}/>}
        {view==="targets_set"&&isAdmin&&<TargetsMgmt team={team} targets={targets} skus={skus} onSave={saveTarget} selMonth={selMonth} selYear={selYear}/>}
      </main>
    </div>

    {/* Bottom nav — mobile only */}
    {isMobile&&user&&<BottomNav isAdmin={isAdmin} view={view} setView={setView} onLogout={()=>{setUser(null);setView("dashboard");}}/>}
  </div>;
}

// ── Login ─────────────────────────────────────────────────────────────────
function Login({cfg,team,onLogin}){
  const [who,setWho]=useState("rep"); const [u,setU]=useState(""); const [pw,setPw]=useState(""); const [err,setErr]=useState("");
  const go=()=>{setErr(""); if(who==="admin"){pw===cfg.adminPassword?onLogin({role:"admin",name:"Admin"}):setErr("Incorrect password.");} else{const r=team.find(t=>t.username.toLowerCase()===u.toLowerCase()&&t.password===pw);r?onLogin({...r,role:"rep"}):setErr("Username or password incorrect.");}};
  return <div style={{background:G1}} className="flex items-center justify-center h-screen">
    <div className="bg-white rounded-2xl shadow-2xl p-8 w-full max-w-sm mx-4">
      <div className="text-center mb-6"><div style={{color:G1}} className="text-2xl font-black tracking-widest">KYEN PRODUCTS</div><div style={{color:AMB}} className="text-xs font-bold tracking-widest uppercase mt-0.5">Sales Hub</div></div>
      <div className="flex rounded-xl overflow-hidden border mb-5">{["rep","admin"].map(r=><button key={r} onClick={()=>{setWho(r);setErr("");}} style={who===r?{background:G1,color:"white"}:{color:"#666"}} className="flex-1 py-2 text-sm font-semibold">{r==="admin"?"Admin":"Sales Rep"}</button>)}</div>
      {who==="rep"&&<Inp value={u} onChange={setU} placeholder="Username" className="mb-3"/>}
      <Inp type="password" value={pw} onChange={setPw} placeholder="Password" className="mb-3"/>
      {/* Hack: real onKeyDown */}
      {err&&<p className="text-red-500 text-xs mb-3 text-center">{err}</p>}
      <button onClick={go} style={{background:G1}} className="w-full py-3 text-white rounded-xl font-bold text-sm">Sign In</button>
      <p className="text-center text-xs text-gray-300 mt-4">Default admin password: <span className="font-mono">kyen2024</span></p>
    </div>
  </div>;
}

// ── Bottom Navigation — mobile ────────────────────────────────────────────
function BottomNav({isAdmin,view,setView,onLogout}){
  const repItems=[
    {id:"dashboard",  i:"📊", l:"Home"},
    {id:"log_sale",   i:"🛒", l:"Sale"},
    {id:"log_visits", i:"🗺️", l:"Visit"},
    {id:"visit_plan", i:"📅", l:"Plan"},
    {id:"more",       i:"⋯",  l:"More"},
  ];
  const admItems=[
    {id:"dashboard",  i:"📊", l:"Home"},
    {id:"all_sales",  i:"📋", l:"Sales"},
    {id:"visits_log", i:"📍", l:"Visits"},
    {id:"customers",  i:"🏪", l:"Outlets"},
    {id:"more",       i:"⋯",  l:"More"},
  ];
  const items=isAdmin?admItems:repItems;
  const [showMore,setShowMore]=useState(false);

  const repMore=[
    {id:"visits_log",  i:"📍", l:"Visits History"},
    {id:"all_sales",   i:"📋", l:"My Sales"},
    {id:"customers",   i:"🏪", l:"Customers"},
    {id:"my_account",  i:"👤", l:"My Account"},
  ];
  const admMore=[
    {id:"team",        i:"👥", l:"Team"},
    {id:"skus",        i:"📦", l:"Manage SKUs"},
    {id:"prices",      i:"💰", l:"Prices"},
    {id:"targets_set", i:"🎯", l:"Targets"},
    {id:"visit_plan",  i:"📅", l:"Visit Plans"},
  ];
  const moreItems=isAdmin?admMore:repMore;

  const go=id=>{ setView(id); setShowMore(false); };

  return <>
    {/* More drawer */}
    {showMore&&<div className="fixed inset-0 z-40" onClick={()=>setShowMore(false)}>
      <div className="absolute bottom-16 left-0 right-0 bg-white border-t shadow-xl rounded-t-2xl p-4 z-50" onClick={e=>e.stopPropagation()}>
        <div className="w-10 h-1 bg-gray-200 rounded-full mx-auto mb-4"/>
        <div className="grid grid-cols-4 gap-3 mb-4">
          {moreItems.map(m=>(
            <button key={m.id} onClick={()=>go(m.id)}
              style={view===m.id?{background:G1+"15",color:G1}:{color:"#555"}}
              className="flex flex-col items-center gap-1 p-2 rounded-xl text-xs font-medium">
              <span className="text-2xl">{m.i}</span>
              <span>{m.l}</span>
            </button>
          ))}
        </div>
        <button onClick={onLogout} style={{color:"#e74c3c",borderColor:"#e74c3c22"}}
          className="w-full py-2.5 border rounded-xl text-sm font-semibold">
          Sign Out
        </button>
      </div>
    </div>}

    {/* Bottom bar */}
    <div className="fixed bottom-0 left-0 right-0 z-30 bg-white border-t shadow-lg" style={{paddingBottom:"env(safe-area-inset-bottom)"}}>
      <div className="flex items-center justify-around px-2 py-1">
        {items.map(item=>{
          const active=item.id==="more"?showMore:view===item.id;
          return (
            <button key={item.id}
              onClick={()=>item.id==="more"?setShowMore(s=>!s):go(item.id)}
              style={{color:active?G1:"#888"}}
              className="flex flex-col items-center gap-0.5 px-3 py-1.5 rounded-xl transition-all flex-1">
              <span className="text-xl leading-none">{item.i}</span>
              <span className="text-xs font-semibold" style={{fontSize:"10px"}}>{item.l}</span>
              {active&&<div style={{background:G1}} className="w-1 h-1 rounded-full mt-0.5"/>}
            </button>
          );
        })}
      </div>
    </div>
  </>;
}

// ── Sidebar ───────────────────────────────────────────────────────────────
function Sidebar({isAdmin,view,setView,user,onLogout}){
  const repNav=[{id:"dashboard",l:"Dashboard",i:"📊"},{id:"log_sale",l:"Log a Sale",i:"🛒"},{id:"log_visits",l:"Log Visits",i:"🗺️"},{id:"visits_log",l:"Visits History",i:"📍"},{id:"visit_plan",l:"Visit Plan",i:"📅"},{id:"all_sales",l:"My Sales",i:"📋"},{id:"customers",l:"Customers",i:"🏪"},{id:"my_account",l:"My Account",i:"👤"}];
  const admNav=[{id:"dashboard",l:"Dashboard",i:"📊"},{id:"all_sales",l:"All Sales",i:"📋"},{id:"visits_log",l:"Visits Log",i:"📍"},{id:"customers",l:"Customers",i:"🏪"},{id:"visit_plan",l:"Visit Plans",i:"📅"},{id:"team",l:"Team",i:"👥"},{id:"skus",l:"Manage SKUs",i:"📦"},{id:"prices",l:"Prices",i:"💰"},{id:"targets_set",l:"Targets",i:"🎯"}];
  const nav=isAdmin?admNav:repNav;
  return <div style={{background:G1}} className="w-52 flex-shrink-0 flex flex-col">
    <div className="px-5 pt-5 pb-3 border-b border-green-800"><div className="text-white font-black tracking-widest">KYEN</div><div style={{color:AMB}} className="text-xs font-bold tracking-widest">SALES HUB</div></div>
    <div className="px-4 py-2 text-xs" style={{color:"#A8D5B5"}}><span className="font-semibold">{user.name}</span><span className="ml-1 opacity-60">{isAdmin?"(admin)":"(rep)"}</span></div>
    <nav className="flex-1 px-2 py-2 space-y-0.5 overflow-y-auto">{nav.map(l=><button key={l.id} onClick={()=>setView(l.id)} style={view===l.id?{background:"rgba(255,255,255,0.15)",color:"white"}:{color:"#A8D5B5"}} className="w-full text-left px-3 py-2 rounded-lg text-sm flex items-center gap-2.5"><span>{l.i}</span><span className="font-medium">{l.l}</span></button>)}</nav>
    <button onClick={onLogout} className="m-3 py-2 text-xs rounded-lg border" style={{color:"#A8D5B5",borderColor:"rgba(255,255,255,0.2)"}}>Sign Out</button>
  </div>;
}

// ── TopBar ────────────────────────────────────────────────────────────────
function TopBar({selMonth,setSelMonth,selYear,setSelYear,fetching,isMobile}){
  return <div className="bg-white border-b px-3 py-2 flex items-center justify-between shadow-sm flex-shrink-0">
    <div className="flex items-center gap-2">
      {!isMobile&&<span className="text-sm font-bold text-gray-700">Kyen Products — Sales Reporting</span>}
      {isMobile&&<span style={{color:G1}} className="text-sm font-black tracking-widest">KYEN</span>}
      {fetching&&<div style={{borderTopColor:G1}} className="w-4 h-4 border-2 border-gray-200 rounded-full animate-spin"/>}
    </div>
    <div className="flex items-center gap-1.5">
      <select value={selMonth} onChange={e=>setSelMonth(Number(e.target.value))} className="border rounded-lg px-1.5 py-1 text-xs">
        {MFULL.map((m,i)=><option key={i} value={i}>{isMobile?m.slice(0,3):m}</option>)}
      </select>
      <select value={selYear} onChange={e=>setSelYear(Number(e.target.value))} className="border rounded-lg px-1.5 py-1 text-xs">
        {[2024,2025,2026,2027,2028].map(y=><option key={y}>{y}</option>)}
      </select>
    </div>
  </div>;
}

// ── Rep Dashboard ─────────────────────────────────────────────────────────
function RepDash({user,monthSales,monthVisits,targets,selMonth,selYear,skus}){
  const tk=`${user.id}_${selYear}_${selMonth}`; const t=targets[tk]||{valueNGN:0,visitsTarget:0,skuTargets:{}};
  const val=monthSales.reduce((a,s)=>a+Number(s.amount),0);
  const outs=monthVisits.length;
  const byDay={}; monthSales.forEach(s=>{byDay[s.date]=(byDay[s.date]||0)+Number(s.amount);});
  const lineData=Object.entries(byDay).sort().map(([d,v])=>({day:d.slice(5),value:v}));
  const skuPerf=skus.map(s=>{ const sold=monthSales.filter(x=>x.sku===s.sku).reduce((a,x)=>a+Number(x.qty),0); const tgt=Number(t.skuTargets?.[s.sku]||0); return {name:s.name,sku:s.sku,sold,tgt,p:pct(sold,tgt)}; }).filter(s=>s.sold>0||s.tgt>0);
  return <div className="space-y-4">
    <div><div className="text-lg font-black text-gray-800">Welcome, {user.name.split(" ")[0]} 👋</div><div className="text-xs text-gray-400">{MFULL[selMonth]} {selYear}</div></div>
    <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
      <StatCard label="Sales Value" value={fmt(val)} sub={t.valueNGN>0?`Target: ${fmt(t.valueNGN)}`:""} pctVal={t.valueNGN>0?pct(val,t.valueNGN):undefined} color={G1}/>
      <StatCard label="Visits Made" value={outs} sub={t.visitsTarget>0?`Target: ${t.visitsTarget}`:""} pctVal={t.visitsTarget>0?pct(outs,t.visitsTarget):undefined} color={AMB}/>
      <StatCard label="Sales Entries" value={monthSales.length} color="#555"/>
    </div>
    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
      <div className="bg-white rounded-xl shadow p-4"><div className="font-bold text-sm text-gray-700 mb-3">Daily Sales Value</div>
        {lineData.length>0?<ResponsiveContainer width="100%" height={200}><LineChart data={lineData}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="day" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}} tickFormatter={v=>"₦"+Math.round(v/1000)+"k"}/><Tooltip formatter={v=>fmt(v)}/><Line type="monotone" dataKey="value" stroke={G1} strokeWidth={2.5} dot={{fill:G1,r:3}}/></LineChart></ResponsiveContainer>:<Empty text="No sales yet"/>}
      </div>
      <div className="bg-white rounded-xl shadow p-4"><div className="font-bold text-sm text-gray-700 mb-3">SKU Performance vs Target</div>
        {skuPerf.length>0?<div className="space-y-2.5 max-h-52 overflow-y-auto">{skuPerf.map(s=><div key={s.sku}><div className="flex justify-between text-xs mb-0.5"><span className="text-gray-700 truncate">{s.name}</span><span className="font-bold ml-2 whitespace-nowrap">{s.sold}{s.tgt>0?` / ${s.tgt}`:""}</span></div><div className="h-1.5 bg-gray-100 rounded-full"><div style={{width:`${Math.min(s.p,100)}%`,background:s.p>=100?G1:s.p>=60?G2:AMB}} className="h-full rounded-full"/></div></div>)}</div>:<Empty text="Log sales to see SKU progress"/>}
      </div>
    </div>
  </div>;
}

// ── Admin Dashboard ───────────────────────────────────────────────────────
function AdminDash({monthSales,monthVisits,team,targets,selMonth,selYear,skus}){
  const tv=monthSales.reduce((a,s)=>a+Number(s.amount),0); const to=monthVisits.length;
  const repPerf=team.map(rep=>{
    const rs=monthSales.filter(s=>s.repId===rep.id); const rv=monthVisits.filter(v=>v.repId===rep.id);
    const t=targets[`${rep.id}_${selYear}_${selMonth}`]||{valueNGN:0,visitsTarget:0};
    const val=rs.reduce((a,s)=>a+Number(s.amount),0); const units=rs.reduce((a,s)=>a+Number(s.qty),0);
    return {name:rep.name.split(" ")[0],fullName:rep.name,val,valTgt:t.valueNGN,units,outs:rv.length,outTgt:t.visitsTarget};
  });
  const byChan={}; CHANNELS.forEach(c=>byChan[c]=0); monthSales.forEach(s=>{byChan[s.channel]=(byChan[s.channel]||0)+Number(s.amount);});
  const chanData=Object.entries(byChan).filter(([,v])=>v>0).map(([n,v])=>({name:n,value:v}));
  const byCat={}; monthSales.forEach(s=>{const cat=skus.find(k=>k.sku===s.sku)?.cat||"Other"; byCat[cat]=(byCat[cat]||0)+Number(s.qty);});
  const catData=Object.entries(byCat).map(([n,v])=>({name:n,value:v}));
  const byDay={}; monthSales.forEach(s=>{const d=s.date.slice(5);byDay[d]=(byDay[d]||0)+Number(s.amount);});
  const trend=Object.entries(byDay).sort().map(([d,v])=>({day:d,value:v}));
  const doExport=()=>dlXLSX([
    {name:"Team Performance",data:repPerf.map(r=>({Rep:r.fullName,"Sales Value":r.val,"Value Target":r.valTgt,"Value %":r.valTgt>0?pct(r.val,r.valTgt)+"%":"—",Units:r.units,Visits:r.outs,"Visit Target":r.outTgt,"Visits %":r.outTgt>0?pct(r.outs,r.outTgt)+"%":"—"}))},
    {name:"All Sales",data:monthSales.map(s=>({Date:s.date,Rep:s.repName,Customer:s.customerName||"",Channel:s.channel,SKU:s.sku,Qty:s.qty,"Unit Price":s.price,"Amount (NGN)":s.amount}))},
    {name:"Visits",data:monthVisits.map(v=>({Date:v.date,Rep:v.repName,"Visit Type":v.visitType,Notes:v.notes||""}))},
  ],`Kyen_Report_${MFULL[selMonth]}_${selYear}`);
  return <div className="space-y-4">
    <div className="flex justify-between items-start flex-wrap gap-2">
      <div><div className="text-lg font-black text-gray-800">Admin Dashboard</div><div className="text-xs text-gray-400">{MFULL[selMonth]} {selYear} · Live</div></div>
      <button onClick={doExport} style={{background:G1}} className="px-4 py-2 text-white text-xs font-bold rounded-xl">📦 Full Report Excel</button>
    </div>
    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
      <StatCard label="Total Value" value={fmt(tv)} color={G1}/><StatCard label="Total Units" value={monthSales.reduce((a,s)=>a+Number(s.qty),0)} color={G2}/>
      <StatCard label="Total Visits" value={to} color={AMB}/><StatCard label="Active Reps" value={team.length} color="#555"/>
    </div>
    <div className="bg-white rounded-xl shadow p-4"><div className="font-bold text-sm text-gray-800 mb-3">Team Performance</div>
      <div className="overflow-x-auto"><table className="w-full text-xs min-w-max">
        <thead><tr style={{background:G1}} className="text-white"><th className="p-2 text-left">Rep</th><th className="p-2 text-right">Value</th><th className="p-2 text-right">Val Tgt</th><th className="p-2 text-center">Val %</th><th className="p-2 text-right">Units</th><th className="p-2 text-right">Visits</th><th className="p-2 text-right">Visit Tgt</th><th className="p-2 text-center">Visit %</th></tr></thead>
        <tbody>
          {repPerf.length===0&&<tr><td colSpan={8} className="text-center p-6 text-gray-400">No team members yet.</td></tr>}
          {repPerf.map((r,i)=>{ const pc=(a,b)=>b>0?<span style={{color:pct(a,b)>=100?G1:pct(a,b)>=60?G2:"#e74c3c",fontWeight:700}}>{pct(a,b)}%</span>:<span className="text-gray-300">—</span>;
            return <tr key={r.name} style={{background:i%2===0?LG:"white"}}><td className="p-2 font-semibold">{r.fullName}</td><td className="p-2 text-right">{fmt(r.val)}</td><td className="p-2 text-right text-gray-400">{r.valTgt?fmt(r.valTgt):"—"}</td><td className="p-2 text-center">{pc(r.val,r.valTgt)}</td><td className="p-2 text-right">{r.units}</td><td className="p-2 text-right">{r.outs}</td><td className="p-2 text-right text-gray-400">{r.outTgt||"—"}</td><td className="p-2 text-center">{pc(r.outs,r.outTgt)}</td></tr>;
          })}
        </tbody>
      </table></div>
    </div>
    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
      <div className="bg-white rounded-xl shadow p-4"><div className="font-bold text-sm text-gray-700 mb-3">Sales Value vs Target by Rep</div>
        {repPerf.length>0?<ResponsiveContainer width="100%" height={220}><BarChart data={repPerf}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="name" tick={{fontSize:11}}/><YAxis tick={{fontSize:10}} tickFormatter={v=>"₦"+Math.round(v/1000)+"k"}/><Tooltip formatter={v=>fmt(v)}/><Legend/><Bar dataKey="val" name="Actual" fill={G1} radius={[4,4,0,0]}/><Bar dataKey="valTgt" name="Target" fill={AMB} radius={[4,4,0,0]}/></BarChart></ResponsiveContainer>:<Empty text="No data"/>}</div>
      <div className="bg-white rounded-xl shadow p-4"><div className="font-bold text-sm text-gray-700 mb-3">Sales by Channel</div>
        {chanData.length>0?<ResponsiveContainer width="100%" height={220}><PieChart><Pie data={chanData} cx="50%" cy="45%" outerRadius={80} dataKey="value" label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`}>{chanData.map((e,i)=><Cell key={i} fill={PIE[i]}/>)}</Pie><Tooltip formatter={v=>fmt(v)}/></PieChart></ResponsiveContainer>:<Empty text="No data"/>}</div>
      <div className="bg-white rounded-xl shadow p-4"><div className="font-bold text-sm text-gray-700 mb-3">Daily Sales Trend</div>
        {trend.length>0?<ResponsiveContainer width="100%" height={200}><LineChart data={trend}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="day" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}} tickFormatter={v=>"₦"+Math.round(v/1000)+"k"}/><Tooltip formatter={v=>fmt(v)}/><Line type="monotone" dataKey="value" stroke={G1} strokeWidth={2.5} dot={{fill:G1,r:3}}/></LineChart></ResponsiveContainer>:<Empty text="No data"/>}</div>
      <div className="bg-white rounded-xl shadow p-4"><div className="font-bold text-sm text-gray-700 mb-3">Units by Category</div>
        {catData.length>0?<ResponsiveContainer width="100%" height={200}><BarChart data={catData}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="name" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/><Tooltip/><Bar dataKey="value" name="Units" radius={[4,4,0,0]}>{catData.map((e,i)=><Cell key={i} fill={PIE[i]}/>)}</Bar></BarChart></ResponsiveContainer>:<Empty text="No data"/>}</div>
    </div>
  </div>;
}

// ── Log Sale ──────────────────────────────────────────────────────────────
function LogSale({user,cfg,skus,customers,onSave,onNewCustomer}){
  const [date,setDate]    = useState(isoD(new Date()));
  const [channel,setChannel] = useState("");
  const [custId,setCustId]   = useState("");
  const [lines,setLines]     = useState([{sku:"",qty:1,discount:0}]);
  const [saving,setSaving]   = useState(false);
  const [done,setDone]       = useState(false);

  const updLine = (i,field,val) => setLines(lines.map((x,j)=>j===i?{...x,[field]:val}:x));
  const unitPrice  = (l) => getP(cfg.prices,l.sku,channel);
  const lineGross  = (l) => unitPrice(l)*Number(l.qty||0);
  const lineSaving = (l) => lineGross(l)*(Number(l.discount||0)/100);
  const lineTotal  = (l) => lineGross(l)-lineSaving(l);
  const totalGross  = lines.reduce((a,l)=>a+lineGross(l),0);
  const totalSaving = lines.reduce((a,l)=>a+lineSaving(l),0);
  const totalNet    = lines.reduce((a,l)=>a+lineTotal(l),0);
  const hasDiscount = lines.some(l=>Number(l.discount||0)>0);

  const [paymentType,setPaymentType] = useState("Cash");
  const valid = date&&channel&&custId&&paymentType&&lines.every(l=>l.sku&&Number(l.qty)>0&&Number(l.discount||0)>=0&&Number(l.discount||0)<=100);
  const TIER_CONFIG = {
    "Supermarket":  {icon:"🏪",label:"Retail Price",    color:G2,          bg:"#EEF7F2",bc:G2},
    "Retail Store": {icon:"🏪",label:"Retail Price",    color:G2,          bg:"#EEF7F2",bc:G2},
    "Wholesaler":   {icon:"🏭",label:"Wholesale Price", color:"#B7800A",   bg:"#FFF8E7",bc:AMB},
    "End User":     {icon:"👤",label:"End User Price",  color:"#6B21A8",   bg:"#F5F3FF",bc:"#A855F7"},
  };
  const tier = channel ? TIER_CONFIG[channel] : null;
  const selCust = customers.find(c=>c.id===custId);
  const cats    = [...new Set(skus.map(s=>s.cat))];

  const handle = async () => {
    if(!valid) return; setSaving(true);
    const entries = lines.map(l=>({
      id:uid(), repId:user.id, repName:user.name, date, channel,
      customerId:custId, customerName:selCust?.name||"",
      sku:l.sku, qty:Number(l.qty),
      price:unitPrice(l),
      discountPct:Number(l.discount||0),
      paymentType,
      amount:lineTotal(l),
    }));
    await onSave(entries); setSaving(false); setDone(true);
  };

  if(done) return (
    <div className="max-w-md mx-auto mt-10 text-center bg-white rounded-2xl shadow p-10">
      <div className="text-5xl mb-3">✅</div>
      <div style={{color:G1}} className="text-xl font-black mb-1">Sale Logged!</div>
      <div className="text-gray-400 text-sm mb-2">Saved to database.</div>
      <div className="mb-6"><Badge color={PAY_COL[paymentType]}>{paymentType==="Cash"?"💵":paymentType==="Credit"?"📋":"🏦"} {paymentType}</Badge></div>
      <Btn onClick={()=>{setDone(false);setLines([{sku:"",qty:1,discount:0}]);setCustId("");setChannel("");setPaymentType("Cash");}}>Log Another Sale</Btn>
    </div>
  );

  return (
    <div className="max-w-xl mx-auto">
      <div className="bg-white rounded-xl shadow p-6">
        <div className="font-black text-gray-800 text-base mb-4">Log a Sale</div>

        <div className="grid grid-cols-2 gap-3 mb-4">
          <Inp label="Date of Sale *" type="date" value={date} onChange={setDate}/>
          <Sel label="Sales Channel *" value={channel} onChange={setChannel}>
            <option value="">Select…</option>{CHANNELS.map(c=><option key={c}>{c}</option>)}
          </Sel>
        </div>

        <div className="mb-4">
          <CustomerSelector customers={customers} value={custId} onChange={setCustId} onNewCustomer={onNewCustomer} label="Customer / Outlet *"/>
        </div>

        {selCust&&(
          <div style={{background:GL}} className="rounded-xl p-3 mb-3 text-xs space-y-0.5">
            <div className="font-bold text-gray-800">{selCust.name}</div>
            {selCust.address&&<div className="text-gray-500">📍 {selCust.address}, {selCust.city}</div>}
            {selCust.contact_phone&&<div className="text-gray-500">📞 {selCust.contact_phone}</div>}
          </div>
        )}

        {/* Payment Type */}
        <div className="mb-4">
          <label className="block text-xs font-semibold text-gray-500 mb-1">Payment Type *</label>
          <div className="flex gap-2">
            {PAYMENT_TYPES.map(pt=>(
              <button key={pt} onClick={()=>setPaymentType(pt)}
                style={paymentType===pt?{background:PAY_COL[pt],color:"white",border:`1.5px solid ${PAY_COL[pt]}`}:{border:`1.5px solid ${PAY_COL[pt]}44`,color:PAY_COL[pt]}}
                className="flex-1 py-2 rounded-xl text-xs font-bold transition-all">
                {pt==="Cash"?"💵":pt==="Credit"?"📋":"🏦"} {pt}
              </button>
            ))}
          </div>
          {paymentType==="Credit"&&(
            <div style={{background:"#FEF2F2",borderColor:"#e74c3c44"}} className="mt-2 rounded-lg px-3 py-2 border text-xs font-semibold text-red-600">
              ⚠️ Credit sale — payment not yet collected. Ensure this is approved.
            </div>
          )}
        </div>

        {tier&&(
          <div style={{background:tier.bg,borderColor:tier.bc+"66"}} className="rounded-xl px-3 py-2 mb-4 flex items-center gap-2 border">
            <span>{tier.icon}</span>
            <span style={{color:tier.color}} className="text-xs font-black">{tier.label} applied</span>
          </div>
        )}

        {/* Column headers */}
        <div className="flex items-center justify-between mb-2">
          <div className="text-sm font-semibold text-gray-700">Products Sold</div>
          <button onClick={()=>setLines([...lines,{sku:"",qty:1,discount:0}])} style={{color:G1}} className="text-xs font-bold">+ Add SKU</button>
        </div>
        {channel&&(
          <div className="grid gap-2 mb-1 px-1 text-xs text-gray-400" style={{gridTemplateColumns:"1fr 52px 72px 16px"}}>
            <span>SKU</span><span className="text-center">Qty</span><span className="text-right">Subtotal</span><span/>
          </div>
        )}

        <div className="space-y-2 mb-2">
          {lines.map((l,i)=>{
            const up  = unitPrice(l);
            const disc= Number(l.discount||0);
            const sub = lineTotal(l);
            return (
              <div key={i} className="rounded-xl border p-2 space-y-2" style={{borderColor:disc>0?"#F4A90066":"#e5e7eb",background:disc>0?"#FFFDF5":"white"}}>
                {/* SKU row */}
                <div className="grid gap-2 items-center" style={{gridTemplateColumns:"1fr 52px 72px 16px"}}>
                  <select value={l.sku} onChange={e=>updLine(i,"sku",e.target.value)} className="border rounded-lg p-1.5 text-xs bg-white">
                    <option value="">Select SKU…</option>
                    {cats.map(cat=>(
                      <optgroup key={cat} label={cat}>
                        {skus.filter(s=>s.cat===cat).map(s=><option key={s.sku} value={s.sku}>{s.sku}</option>)}
                      </optgroup>
                    ))}
                  </select>
                  <input type="number" min="1" value={l.qty} onChange={e=>updLine(i,"qty",e.target.value)}
                    className="border rounded-lg p-1.5 text-xs text-center"/>
                  <div className="text-xs text-right font-semibold" style={{color:G2}}>{fmt(sub)}</div>
                  {lines.length>1
                    ? <button onClick={()=>setLines(lines.filter((_,j)=>j!==i))} className="text-red-400 font-bold text-base leading-none">×</button>
                    : <span/>}
                </div>

                {/* Unit price + Discount row — always visible per line */}
                {l.sku&&channel&&(
                  <div className="flex items-center gap-2 px-0.5 flex-wrap">
                    <span className="text-xs text-gray-400">Unit: <span className="font-semibold text-gray-600">{fmt(up)}</span></span>
                    <span className="text-gray-200 text-xs">|</span>
                    <div className="flex items-center gap-1.5">
                      <label className="text-xs font-semibold" style={{color:AMB}}>🏷️ Discount %</label>
                      <div className="flex items-center border rounded-lg overflow-hidden" style={{borderColor:disc>0?"#F4A900":"#e5e7eb",background:disc>0?"#FFF8E7":"white"}}>
                        <input
                          type="number" min="0" max="100" step="0.5"
                          value={l.discount||""}
                          onChange={e=>updLine(i,"discount",e.target.value)}
                          placeholder="0"
                          className="w-14 p-1.5 text-xs text-center outline-none"
                          style={{background:"transparent"}}
                        />
                        <span className="pr-2 text-xs font-bold" style={{color:disc>0?"#B7800A":"#aaa"}}>%</span>
                      </div>
                      {disc>0&&(
                        <span className="text-xs font-bold" style={{color:"#B7800A"}}>
                          → saving {fmt(lineSaving(l))}
                        </span>
                      )}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* Grand total block */}
        <div style={{background:GL}} className="rounded-xl p-3 mb-5 space-y-1">
          {hasDiscount&&(
            <div className="flex justify-between text-xs text-gray-500">
              <span>Original total</span><span className="line-through">{fmt(totalGross)}</span>
            </div>
          )}
          {hasDiscount&&(
            <div className="flex justify-between text-xs font-semibold" style={{color:"#B7800A"}}>
              <span>🏷️ Total discount</span><span>— {fmt(totalSaving)}</span>
            </div>
          )}
          <div className="flex justify-between items-center pt-1">
            <div>
              <span className="text-sm font-semibold text-gray-600">Grand Total</span>
              {tier&&<span className="text-xs text-gray-400 ml-2">({tier.label})</span>}
            </div>
            <span style={{color:G1}} className="text-xl font-black">{fmt(totalNet)}</span>
          </div>
          <div className="flex justify-between items-center pt-1 border-t border-gray-100 mt-1">
            <span className="text-xs text-gray-500">Payment method</span>
            <Badge color={PAY_COL[paymentType]}>{paymentType==="Cash"?"💵":paymentType==="Credit"?"📋":"🏦"} {paymentType}</Badge>
          </div>
        </div>

        {!valid&&<p className="text-xs text-gray-400 mb-3">Fill in all required fields. Discount must be 0–100.</p>}
        <Btn onClick={handle} disabled={!valid||saving} full>{saving?"Saving…":"Submit Sale"}</Btn>
      </div>
    </div>
  );
}

// ── Log Visits ────────────────────────────────────────────────────────────
function LogVisits({user,customers,onSave,onNewCustomer}){
  const [date,setDate]=useState(isoD(new Date())); const [custId,setCustId]=useState(""); const [visitType,setVisitType]=useState("First Visit"); const [notes,setNotes]=useState(""); const [saving,setSaving]=useState(false); const [done,setDone]=useState(false);
  const valid=date&&custId&&visitType;
  const selCust=customers.find(c=>c.id===custId);
  const handle=async()=>{ if(!valid) return; setSaving(true); await onSave({id:uid(),repId:user.id,repName:user.name,date,customerId:custId,visitType,notes}); setSaving(false); setDone(true); };
  if(done) return <div className="max-w-md mx-auto mt-10 text-center bg-white rounded-2xl shadow p-10"><div className="text-5xl mb-3">🗺️</div><div style={{color:G1}} className="text-xl font-black mb-1">Visit Logged!</div><div className="text-gray-400 text-sm mb-6">Saved to database.</div><Btn onClick={()=>{setDone(false);setCustId("");setNotes("");setVisitType("First Visit");}}>Log Another Visit</Btn></div>;
  return <div className="max-w-lg mx-auto"><div className="bg-white rounded-xl shadow p-6">
    <div className="font-black text-gray-800 text-base mb-4">Log an Outlet Visit</div>
    <Inp label="Date of Visit *" type="date" value={date} onChange={setDate} className="mb-4"/>
    <div className="mb-4"><CustomerSelector customers={customers} value={custId} onChange={setCustId} onNewCustomer={onNewCustomer} label="Outlet / Customer *"/></div>
    {selCust&&<div style={{background:GL}} className="rounded-xl p-3 mb-4 text-xs space-y-0.5">
      <div className="font-bold text-gray-800">{selCust.name}</div>
      {selCust.address&&<div className="text-gray-500">📍 {selCust.address}, {selCust.city}</div>}
      {selCust.contact_phone&&<div className="text-gray-500">📞 {selCust.contact_phone} {selCust.contact_name&&`· ${selCust.contact_name}`}</div>}
      {selCust.zone&&<div className="text-gray-500">🗺️ Zone: {selCust.zone}</div>}
    </div>}
    <Sel label="Visit Type *" value={visitType} onChange={setVisitType} className="mb-4">{VISIT_TYPES.map(t=><option key={t}>{t}</option>)}</Sel>
    <div style={{background:VTYPE_COL[visitType]+"18",borderColor:VTYPE_COL[visitType]+"44"}} className="rounded-xl px-3 py-2 mb-4 flex items-center gap-2 border">
      <span>{visitType==="First Visit"?"🆕":visitType==="Follow-up"?"🔄":"🏷️"}</span>
      <span style={{color:VTYPE_COL[visitType]}} className="text-xs font-black">{visitType}</span>
    </div>
    <div className="mb-5"><label className="block text-xs font-semibold text-gray-500 mb-1">Notes / Outcome</label><textarea value={notes} onChange={e=>setNotes(e.target.value)} rows={3} placeholder="What was discussed? Any orders? Follow-up needed?" className="w-full border rounded-xl p-2.5 text-sm outline-none resize-none"/></div>
    <Btn onClick={handle} disabled={!valid||saving} full>{saving?"Saving…":"Submit Visit"}</Btn>
  </div></div>;
}

// ── Visit Plan ────────────────────────────────────────────────────────────
function VisitPlanView({user,isAdmin,visitPlans,customers,team,onAdd,onStatus,onDelete,onNewCustomer,selMonth,selYear}){
  const [mode,setMode]         = useState("week");
  const [weekOff,setWeekOff]   = useState(0);
  const [adding,setAdding]     = useState(false);
  const [completing,setCompleting] = useState(null); // plan being completed
  const [completeNotes,setCompleteNotes] = useState("");
  const [filterRep,setFilterRep] = useState("all");

  const fp  = filterRep==="all" ? visitPlans : visitPlans.filter(p=>p.repId===filterRep);
  const getWeek = off => {
    const t=new Date(); const m=new Date(t);
    m.setDate(t.getDate()-((t.getDay()||7)-1)+off*7);
    return Array.from({length:7},(_,i)=>{ const d=new Date(m); d.setDate(m.getDate()+i); return d; });
  };
  const week = getWeek(weekOff);
  const wd   = `${week[0].toLocaleDateString("en-GB",{day:"numeric",month:"short"})} – ${week[6].toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"})}`;
  const plansOn  = d => fp.filter(p=>p.plannedDate===isoD(d));
  const plansDay = day => {
    if(!day) return [];
    const ds=`${selYear}-${String(selMonth+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
    return fp.filter(p=>p.plannedDate===ds);
  };
  const stats={planned:fp.filter(p=>p.status==="planned").length,completed:fp.filter(p=>p.status==="completed").length,cancelled:fp.filter(p=>p.status==="cancelled").length};

  const daysInMonth = new Date(selYear,selMonth+1,0).getDate();
  const firstDay    = new Date(selYear,selMonth,1).getDay();
  const cells       = Array.from({length:firstDay+daysInMonth},(_,i)=>i<firstDay?null:i-firstDay+1);

  const handleComplete = async () => {
    if(!completing) return;
    await onStatus(completing.id,"completed",completeNotes);
    setCompleting(null); setCompleteNotes("");
  };

  const PlanCard = ({p, compact=false}) => {
    const c = customers.find(x=>x.id===p.customerId);
    return (
      <div style={{background:STATUS_COL[p.status]+"18",borderLeft:`3px solid ${STATUS_COL[p.status]}`}}
        className={`rounded-r-lg p-1.5 text-xs group ${compact?"":"mb-1"}`}>
        <div className="font-semibold text-gray-800 truncate">{c?.name||p.customerName||"—"}</div>
        <div style={{color:VTYPE_COL[p.visitType]}} className="text-xs">{p.visitType}</div>
        {isAdmin&&<div className="text-xs text-gray-400 truncate">{p.repName}</div>}
        {p.notes&&!compact&&<div className="text-xs text-gray-400 truncate mt-0.5 italic">{p.notes}</div>}
        <div className="hidden group-hover:flex gap-1 mt-1.5 flex-wrap items-center">
          {p.status==="planned"&&<>
            <button onClick={()=>{setCompleting(p);setCompleteNotes(p.notes||"");}}
              style={{background:G1,color:"white"}} className="text-xs font-bold rounded px-2 py-0.5">
              ✓ Complete
            </button>
            <button onClick={()=>onStatus(p.id,"cancelled","")}
              className="text-xs font-bold text-red-400 border border-red-200 rounded px-2 py-0.5">
              ✗ Cancel
            </button>
          </>}
          {p.status==="completed"&&<span style={{color:G1}} className="text-xs font-bold">✓ Logged</span>}
          {p.status==="cancelled"&&<span className="text-xs text-red-400 font-bold">✗ Cancelled</span>}
          <button onClick={()=>onDelete(p.id)} className="text-xs text-gray-300 hover:text-red-400 ml-auto">🗑</button>
        </div>
      </div>
    );
  };

  return <div className="space-y-4">
    <div className="flex items-start justify-between flex-wrap gap-2">
      <div><div className="text-lg font-black text-gray-800">Visit Plan</div><div className="text-xs text-gray-400">{MFULL[selMonth]} {selYear}</div></div>
      <div className="flex gap-2 flex-wrap items-center">
        {isAdmin&&<select value={filterRep} onChange={e=>setFilterRep(e.target.value)} className="border rounded-xl px-3 py-1.5 text-xs"><option value="all">All Reps</option>{team.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}</select>}
        <div className="flex rounded-xl overflow-hidden border">{["week","month"].map(m=><button key={m} onClick={()=>setMode(m)} style={mode===m?{background:G1,color:"white"}:{color:"#666"}} className="px-3 py-1.5 text-xs font-semibold capitalize">{m}</button>)}</div>
        <Btn sm onClick={()=>setAdding(true)}>+ Plan Visit</Btn>
      </div>
    </div>

    <div className="flex gap-2 flex-wrap">
      <Badge color={AMB}>⏳ {stats.planned} Planned</Badge>
      <Badge color={G1}>✓ {stats.completed} Completed</Badge>
      <Badge color="#e74c3c">✗ {stats.cancelled} Cancelled</Badge>
    </div>

    {/* Completion notice */}
    {stats.completed>0&&(
      <div style={{background:"#EEF7F2",borderColor:G2+"44"}} className="rounded-xl px-3 py-2 border flex items-center gap-2 text-xs">
        <span>✅</span>
        <span style={{color:G1}} className="font-semibold">Completed visits are automatically added to the Visits Log.</span>
      </div>
    )}

    {/* Week view */}
    {mode==="week"&&<div className="space-y-3">
      <div className="flex items-center gap-2">
        <button onClick={()=>setWeekOff(w=>w-1)} className="px-3 py-1 border rounded-lg text-xs hover:bg-gray-50">← Prev</button>
        <div className="text-sm font-semibold text-gray-700 flex-1 text-center">{wd}</div>
        <button onClick={()=>setWeekOff(w=>w+1)} className="px-3 py-1 border rounded-lg text-xs hover:bg-gray-50">Next →</button>
        <button onClick={()=>setWeekOff(0)} style={{background:GL,color:G1}} className="px-3 py-1 rounded-lg text-xs font-bold">Today</button>
      </div>
      <div className="grid grid-cols-7 gap-2">
        {week.map((d,i)=>{
          const today=isoD(d)===isoD(new Date()); const plans=plansOn(d);
          return <div key={i} className="bg-white rounded-xl shadow-sm min-h-32">
            <div style={{background:today?G1:"",color:today?"white":"#666"}} className="rounded-t-xl px-2 py-1.5 text-center">
              <div className="text-xs font-bold">{DAYS[d.getDay()]}</div>
              <div className="text-sm font-black">{d.getDate()}</div>
            </div>
            <div className="p-1.5 space-y-1">
              {plans.map(p=><PlanCard key={p.id} p={p}/>)}
              {plans.length===0&&<div className="text-xs text-gray-200 text-center pt-3">—</div>}
            </div>
          </div>;
        })}
      </div>
    </div>}

    {/* Month view */}
    {mode==="month"&&<div className="bg-white rounded-xl shadow p-4">
      <div className="grid grid-cols-7 gap-1 mb-1">{DAYS.map(d=><div key={d} className="text-center text-xs font-bold text-gray-400 py-1">{d}</div>)}</div>
      <div className="grid grid-cols-7 gap-1">
        {cells.map((day,i)=>{
          if(!day) return <div key={i}/>;
          const plans=plansDay(day);
          const ds=`${selYear}-${String(selMonth+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
          const today=ds===isoD(new Date());
          return <div key={i} style={{background:today?GL:""}} className="min-h-14 border rounded-lg p-1 hover:bg-gray-50">
            <div style={{color:today?G1:"#999",fontWeight:today?700:400}} className="text-xs text-right mb-1">{day}</div>
            {plans.slice(0,2).map(p=>{
              const c=customers.find(x=>x.id===p.customerId);
              return <div key={p.id} style={{background:STATUS_COL[p.status],color:"white"}} className="text-xs rounded px-1 mb-0.5 truncate cursor-pointer"
                onClick={()=>{if(p.status==="planned"){setCompleting(p);setCompleteNotes(p.notes||"");}}}>
                {c?.name||"—"}
              </div>;
            })}
            {plans.length>2&&<div className="text-xs text-gray-400">+{plans.length-2}</div>}
          </div>;
        })}
      </div>
    </div>}

    {/* Add plan modal */}
    {adding&&<Modal title="Plan a Visit" onClose={()=>setAdding(false)}>
      <AddPlanForm user={user} isAdmin={isAdmin} team={team} customers={customers} onSave={async p=>{await onAdd(p);setAdding(false);}} onClose={()=>setAdding(false)} onNewCustomer={onNewCustomer}/>
    </Modal>}

    {/* Complete visit modal — auto-creates visit log */}
    {completing&&<Modal title="Complete Visit" onClose={()=>{setCompleting(null);setCompleteNotes("");}}>
      <div className="space-y-4">
        {/* Plan summary */}
        <div style={{background:GL}} className="rounded-xl p-3">
          <div className="font-bold text-gray-800 text-sm">{completing.customerName||"—"}</div>
          <div className="text-xs text-gray-500 mt-0.5">{completing.plannedDate} · <span style={{color:VTYPE_COL[completing.visitType]}} className="font-semibold">{completing.visitType}</span></div>
          {isAdmin&&<div className="text-xs text-gray-400 mt-0.5">Rep: {completing.repName}</div>}
        </div>

        {/* Info banner */}
        <div style={{background:"#EEF7F2",borderColor:G2+"44"}} className="rounded-xl px-3 py-2.5 border">
          <div style={{color:G1}} className="text-xs font-black mb-0.5">✅ This will also create a Visit Log entry</div>
          <div className="text-xs text-gray-500">The visit will appear in Visits History automatically.</div>
        </div>

        {/* Notes */}
        <div>
          <label className="block text-xs font-semibold text-gray-500 mb-1">Visit Outcome / Notes <span className="font-normal text-gray-400">(optional)</span></label>
          <textarea value={completeNotes} onChange={e=>setCompleteNotes(e.target.value)}
            rows={3} placeholder="What was discussed? Any orders placed? Next steps?"
            className="w-full border rounded-xl p-2.5 text-sm outline-none resize-none"/>
        </div>

        <div className="flex gap-2">
          <Btn onClick={handleComplete} full>✓ Mark as Completed</Btn>
          <Btn onClick={()=>{setCompleting(null);setCompleteNotes("");}} outline color="#999" sm>Cancel</Btn>
        </div>
      </div>
    </Modal>}
  </div>;
}

function AddPlanForm({user,isAdmin,team,customers,onSave,onClose,onNewCustomer}){
  const [repId,setRepId]=useState(user.role==="rep"?user.id:""); const [date,setDate]=useState(isoD(new Date())); const [custId,setCustId]=useState(""); const [vt,setVt]=useState("First Visit"); const [notes,setNotes]=useState(""); const [saving,setSaving]=useState(false);
  const rep=isAdmin?team.find(t=>t.id===repId):user;
  const selCust=customers.find(c=>c.id===custId);
  const valid=date&&custId&&vt&&(isAdmin?repId:true);
  const handle=async()=>{ if(!valid) return; setSaving(true); await onSave({id:uid(),repId:isAdmin?repId:user.id,repName:rep?.name||user.name,plannedDate:date,customerId:custId,customerName:selCust?.name||"",visitType:vt,notes,status:"planned"}); setSaving(false); };
  return <div className="space-y-3">
    {isAdmin&&<Sel label="Assign to Rep *" value={repId} onChange={setRepId}><option value="">Select rep…</option>{team.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}</Sel>}
    <Inp label="Planned Date *" type="date" value={date} onChange={setDate}/>
    <CustomerSelector customers={customers} value={custId} onChange={setCustId} onNewCustomer={onNewCustomer} label="Customer / Outlet *"/>
    <Sel label="Visit Type *" value={vt} onChange={setVt}>{VISIT_TYPES.map(t=><option key={t}>{t}</option>)}</Sel>
    <div><label className="block text-xs font-semibold text-gray-500 mb-1">Notes (optional)</label><textarea value={notes} onChange={e=>setNotes(e.target.value)} rows={2} className="w-full border rounded-xl p-2.5 text-sm outline-none resize-none" placeholder="Objective of visit…"/></div>
    <div className="flex gap-2 pt-1"><Btn onClick={handle} disabled={!valid||saving}>{saving?"Saving…":"Save Plan"}</Btn><Btn onClick={onClose} outline color="#999">Cancel</Btn></div>
  </div>;
}

// ── All Sales View ────────────────────────────────────────────────────────
function AllSalesView({sales,customers,team,onDelete,month,year,skus}){
  const [filter,setFilter]=useState("all");
  const f=filter==="all"?sales:sales.filter(s=>s.repId===filter||s.channel===filter);
  const tv=f.reduce((a,s)=>a+Number(s.amount),0);
  const totalSaved=f.reduce((a,s)=>{
    const gross=(s.price*s.qty);
    return a+(gross-s.amount);
  },0);
  const hasAnyDiscount=f.some(s=>Number(s.discountPct||0)>0);

  const doXL=()=>dlXLSX([{name:"Sales",data:f.map(s=>({
    Date:s.date,Rep:s.repName,
    Customer:s.customerName||customers.find(c=>c.id===s.customerId)?.name||"",
    Channel:s.channel,SKU:s.sku,
    Product:skus.find(k=>k.sku===s.sku)?.name||"",
    Qty:s.qty,"Unit Price (₦)":s.price,
    "Discount %":s.discountPct||0,
    "Amount (₦)":s.amount,
  }))}],`Kyen_Sales_${MFULL[month]}_${year}`);

  return <div className="space-y-3">
    <div className="flex items-center justify-between flex-wrap gap-2">
      <div className="font-black text-gray-800">Sales Records <span className="font-normal text-sm text-gray-400">({f.length})</span></div>
      <div className="flex gap-2 flex-wrap items-center">
        <select value={filter} onChange={e=>setFilter(e.target.value)} className="border rounded-xl px-3 py-1.5 text-xs">
          <option value="all">All</option>
          {CHANNELS.map(c=><option key={c} value={c}>{c}</option>)}
          {team.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}
        </select>
        {f.length>0&&<>
          <button onClick={doXL} style={{background:G1}} className="px-3 py-1.5 text-white text-xs font-bold rounded-xl">⬇ Excel</button>
          <button onClick={()=>dlCSV(f.map(s=>({Date:s.date,Rep:s.repName,Customer:s.customerName||"",Channel:s.channel,SKU:s.sku,Qty:s.qty,"Discount %":s.discountPct||0,Amount:s.amount})),`Kyen_Sales_${MFULL[month]}_${year}`)} style={{background:G2}} className="px-3 py-1.5 text-white text-xs font-bold rounded-xl">⬇ CSV</button>
        </>}
      </div>
    </div>

    {/* Summary strip */}
    <div style={{background:GL}} className="rounded-xl px-4 py-2 flex gap-4 flex-wrap">
      <div><span className="text-xs text-gray-500">Net Value</span><div style={{color:G1}} className="font-black text-sm">{fmt(tv)}</div></div>
      <div><span className="text-xs text-gray-500">Units</span><div style={{color:G1}} className="font-black text-sm">{f.reduce((a,s)=>a+Number(s.qty),0)}</div></div>
      <div><span className="text-xs text-gray-500">💵 Cash</span><div style={{color:G1}} className="font-black text-sm">{fmt(f.filter(s=>s.paymentType==="Cash").reduce((a,s)=>a+s.amount,0))}</div></div>
      <div><span className="text-xs text-gray-500">📋 Credit</span><div style={{color:"#e74c3c"}} className="font-black text-sm">{fmt(f.filter(s=>s.paymentType==="Credit").reduce((a,s)=>a+s.amount,0))}</div></div>
      <div><span className="text-xs text-gray-500">🏦 Transfer</span><div style={{color:"#2980B9"}} className="font-black text-sm">{fmt(f.filter(s=>s.paymentType==="Bank Transfer").reduce((a,s)=>a+s.amount,0))}</div></div>
      {totalSaved>0&&<div><span className="text-xs text-gray-500">Discounts Given</span><div style={{color:"#B7800A"}} className="font-black text-sm">{fmt(totalSaved)}</div></div>}
    </div>

    <div className="bg-white rounded-xl shadow overflow-x-auto">
      <table className="w-full text-xs min-w-max">
        <thead>
          <tr style={{background:G1}} className="text-white text-left">
            <th className="p-2.5">Date</th>
            <th className="p-2.5">Rep</th>
            <th className="p-2.5">Customer</th>
            <th className="p-2.5">Channel</th>
            <th className="p-2.5">Payment</th>
            <th className="p-2.5">SKU</th>
            <th className="p-2.5 text-right">Qty</th>
            <th className="p-2.5 text-right">Unit ₦</th>
            {hasAnyDiscount&&<th className="p-2.5 text-center">Disc %</th>}
            <th className="p-2.5 text-right">Amount</th>
            {onDelete&&<th className="p-2.5"/>}
          </tr>
        </thead>
        <tbody>
          {f.length===0&&<tr><td colSpan={onDelete?9:8} className="text-center p-8 text-gray-400">No sales for this period.</td></tr>}
          {[...f].reverse().map((s,i)=>{
            const disc=Number(s.discountPct||0);
            return <tr key={s.id} style={{background:i%2===0?LG:"white"}}>
              <td className="p-2.5">{s.date}</td>
              <td className="p-2.5 font-medium">{s.repName}</td>
              <td className="p-2.5 font-medium">{s.customerName||customers.find(c=>c.id===s.customerId)?.name||"—"}</td>
              <td className="p-2.5"><Badge color={s.channel==="Wholesaler"?AMB:s.channel==="End User"?"#6B21A8":G1}>{s.channel}</Badge></td>
              <td className="p-2.5"><Badge color={PAY_COL[s.paymentType||"Cash"]}>{s.paymentType||"Cash"}</Badge></td>
              <td className="p-2.5 text-gray-600">{s.sku}</td>
              <td className="p-2.5 text-right font-semibold">{s.qty}</td>
              <td className="p-2.5 text-right text-gray-500">{fmt(s.price)}</td>
              {hasAnyDiscount&&<td className="p-2.5 text-center">
                {disc>0
                  ? <span style={{background:"#F4A90022",color:"#B7800A"}} className="font-bold px-2 py-0.5 rounded-full">🏷️ {disc}%</span>
                  : <span className="text-gray-300">—</span>}
              </td>}
              <td className="p-2.5 text-right font-bold" style={{color:G1}}>{fmt(s.amount)}</td>
              {onDelete&&<td className="p-2.5 text-center"><button onClick={()=>onDelete(s.id)} className="text-red-300 hover:text-red-500">🗑</button></td>}
            </tr>;
          })}
        </tbody>
      </table>
    </div>
  </div>;
}

// ── Customer List ─────────────────────────────────────────────────────────
// ── Visits Log ────────────────────────────────────────────────────────────
function VisitsLog({visits,customers,team,isAdmin,month,year}){
  const [filterRep,setFilterRep]=useState("all");
  const [filterType,setFilterType]=useState("all");
  const [search,setSearch]=useState("");

  const getCust=id=>customers.find(c=>c.id===id);

  let data=[...visits].sort((a,b)=>b.date.localeCompare(a.date));
  if(filterRep!=="all") data=data.filter(v=>v.repId===filterRep);
  if(filterType!=="all") data=data.filter(v=>v.visitType===filterType);
  if(search.trim()) data=data.filter(v=>{
    const c=getCust(v.customerId);
    return (c?.name||"").toLowerCase().includes(search.toLowerCase())||
           (c?.city||"").toLowerCase().includes(search.toLowerCase())||
           (v.notes||"").toLowerCase().includes(search.toLowerCase());
  });

  const typeColor={"First Visit":G1,"Follow-up":G2,"Merchandising":AMB};
  const typeIcon={"First Visit":"🆕","Follow-up":"🔄","Merchandising":"🏷️"};

  const doExport=()=>dlXLSX([{name:"Visits",data:data.map(v=>{const c=getCust(v.customerId);return{
    Date:v.date,Rep:v.repName,"Outlet Name":c?.name||"—","Outlet Type":c?.type||"—",
    Address:c?.address||"—",City:c?.city||"—",Zone:c?.zone||"—",
    "Contact Name":c?.contact_name||"—","Contact Phone":c?.contact_phone||"—",
    "Visit Type":v.visitType,"Notes / Outcome":v.notes||"—",
  };})}],`Kyen_Visits_${MFULL[month]}_${year}`);

  return <div className="space-y-3">
    <div className="flex items-center justify-between flex-wrap gap-2">
      <div>
        <div className="font-black text-gray-800">Visits Log <span className="font-normal text-sm text-gray-400">({data.length} visits)</span></div>
        <div className="text-xs text-gray-400 mt-0.5">{MFULL[month]} {year}{isAdmin?" · All reps":""}</div>
      </div>
      <div className="flex gap-2 flex-wrap items-center">
        {isAdmin&&<select value={filterRep} onChange={e=>setFilterRep(e.target.value)} className="border rounded-xl px-3 py-1.5 text-xs">
          <option value="all">All Reps</option>{team.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}
        </select>}
        <select value={filterType} onChange={e=>setFilterType(e.target.value)} className="border rounded-xl px-3 py-1.5 text-xs">
          <option value="all">All Visit Types</option>{["First Visit","Follow-up","Merchandising"].map(t=><option key={t}>{t}</option>)}
        </select>
        {data.length>0&&<button onClick={doExport} style={{background:G1}} className="px-3 py-1.5 text-white text-xs font-bold rounded-xl">⬇ Excel</button>}
      </div>
    </div>

    {/* Search */}
    <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search by outlet name, city or notes…" className="w-full border rounded-xl p-2.5 text-sm outline-none"/>

    {/* Summary strip */}
    <div style={{background:GL}} className="rounded-xl px-4 py-2 flex gap-6 text-sm flex-wrap">
      {["First Visit","Follow-up","Merchandising"].map(t=>{
        const n=data.filter(v=>v.visitType===t).length;
        return <div key={t}><span className="text-xs text-gray-500">{typeIcon[t]} {t}</span><div style={{color:typeColor[t]}} className="font-black text-sm">{n}</div></div>;
      })}
    </div>

    {/* Visit cards */}
    {data.length===0
      ? <div className="bg-white rounded-xl shadow p-10 text-center text-gray-400 text-sm">No visits recorded for this period.</div>
      : <div className="space-y-3">
          {data.map(v=>{
            const c=getCust(v.customerId);
            return <div key={v.id} className="bg-white rounded-xl shadow p-4">
              <div className="flex items-start justify-between gap-4 flex-wrap">
                {/* Left: outlet info */}
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 mb-1 flex-wrap">
                    <div className="font-black text-gray-800 text-sm">{c?.name||"Unknown Outlet"}</div>
                    {c&&<Badge color={c.type==="Wholesaler"?AMB:c.type==="Retail Store"?G2:G1}>{c.type||"—"}</Badge>}
                    <span style={{background:typeColor[v.visitType]+"18",color:typeColor[v.visitType]}} className="text-xs font-bold px-2 py-0.5 rounded-full">
                      {typeIcon[v.visitType]} {v.visitType}
                    </span>
                  </div>
                  {/* Outlet details row */}
                  <div className="flex gap-4 flex-wrap text-xs text-gray-500 mb-2">
                    {c?.address&&<span>📍 {c.address}</span>}
                    {c?.city&&<span>🏙️ {c.city}</span>}
                    {c?.zone&&<span>🗺️ {c.zone}</span>}
                    {c?.contact_phone&&<span>📞 {c.contact_phone}</span>}
                    {c?.contact_name&&<span>👤 {c.contact_name}</span>}
                  </div>
                  {/* Notes */}
                  {v.notes
                    ? <div style={{background:GL}} className="rounded-lg px-3 py-2 text-xs text-gray-700">
                        <span className="font-semibold text-gray-500 mr-1">Notes:</span>{v.notes}
                      </div>
                    : <div className="text-xs text-gray-300 italic">No notes recorded.</div>}
                </div>
                {/* Right: meta */}
                <div className="text-right flex-shrink-0">
                  <div className="text-xs font-semibold text-gray-600">{v.date}</div>
                  {isAdmin&&<div className="text-xs text-gray-400 mt-0.5">{v.repName}</div>}
                </div>
              </div>
            </div>;
          })}
        </div>}
  </div>;
}

// ── Customer Purchase History ─────────────────────────────────────────────
function CustomerHistory({customer,onClose,fetchHistory}){
  const [history,setHistory]=useState(null);
  const [loading,setLoading]=useState(true);
  useEffect(()=>{ fetchHistory(customer.id).then(d=>{setHistory(d);setLoading(false);}); },[customer.id]);
  const total  =(history||[]).reduce((a,s)=>a+Number(s.amount),0);
  const cash   =(history||[]).filter(s=>s.paymentType==="Cash").reduce((a,s)=>a+Number(s.amount),0);
  const credit =(history||[]).filter(s=>s.paymentType==="Credit").reduce((a,s)=>a+Number(s.amount),0);
  return <Modal title={`Purchase History — ${customer.name}`} onClose={onClose}>
    {loading
      ? <div className="text-center py-8 text-gray-400 text-sm">Loading history…</div>
      : <div className="space-y-3">
          <div style={{background:GL}} className="rounded-xl p-3 text-xs space-y-0.5">
            <div className="font-bold text-gray-800">{customer.name}</div>
            {customer.address&&<div className="text-gray-500">📍 {customer.address}, {customer.city}</div>}
            {customer.contact_phone&&<div className="text-gray-500">📞 {customer.contact_phone}{customer.contact_name&&` · ${customer.contact_name}`}</div>}
            {customer.zone&&<div className="text-gray-500">🗺️ {customer.zone}</div>}
          </div>
          <div className="grid grid-cols-3 gap-2">
            <div className="bg-white rounded-xl p-3 text-center shadow-sm border"><div className="text-xs text-gray-400">Total Spent</div><div style={{color:G1}} className="font-black text-sm">{fmt(total)}</div></div>
            <div className="bg-white rounded-xl p-3 text-center shadow-sm border"><div className="text-xs text-gray-400">💵 Cash</div><div style={{color:G1}} className="font-black text-sm">{fmt(cash)}</div></div>
            <div className="bg-white rounded-xl p-3 text-center shadow-sm border"><div className="text-xs text-gray-400">📋 Credit</div><div style={{color:credit>0?"#e74c3c":"#ccc"}} className="font-black text-sm">{fmt(credit)}</div></div>
          </div>
          {history.length===0
            ? <div className="text-center py-6 text-gray-400 text-sm">No purchases recorded yet.</div>
            : <div className="max-h-72 overflow-y-auto space-y-1.5">
                {history.map((s,i)=>(
                  <div key={s.id} style={{background:i%2===0?LG:"white"}} className="rounded-lg px-3 py-2 flex justify-between items-center gap-2">
                    <div className="flex-1 min-w-0">
                      <div className="text-xs font-semibold text-gray-800 truncate">{s.sku}</div>
                      <div className="text-xs text-gray-400">{s.date} · {s.repName} · qty {s.qty}{s.discountPct>0?` · ${s.discountPct}% disc`:""}</div>
                    </div>
                    <div className="flex items-center gap-2 flex-shrink-0">
                      <Badge color={PAY_COL[s.paymentType||"Cash"]}>{s.paymentType||"Cash"}</Badge>
                      <div style={{color:G1}} className="font-black text-xs">{fmt(s.amount)}</div>
                    </div>
                  </div>
                ))}
              </div>}
          <div className="text-xs text-gray-400 text-center">{history.length} transaction{history.length!==1?"s":""} · all time</div>
        </div>}
  </Modal>;
}

// ── Customer List ─────────────────────────────────────────────────────────
function CustomerList({customers,onAdd,onEdit,monthSales,fetchCustHistory}){
  const [adding,setAdding]=useState(false);
  const [editing,setEditing]=useState(null);
  const [viewing,setViewing]=useState(null);
  const [search,setSearch]=useState("");
  const [saved,setSaved]=useState(false);

  const fl=customers.filter(c=>
    c.name.toLowerCase().includes(search.toLowerCase())||
    (c.city||"").toLowerCase().includes(search.toLowerCase())||
    (c.zone||"").toLowerCase().includes(search.toLowerCase())
  );
  const sc=id=>monthSales.filter(s=>s.customerId===id).length;
  const handleEdit=async()=>{
    if(!editing||!editing.name.trim()) return;
    await onEdit(editing);
    setSaved(true); setTimeout(()=>{setSaved(false);setEditing(null);},1200);
  };

  return <div className="space-y-4">
    <div className="flex items-center justify-between gap-2">
      <div className="font-black text-gray-800">Customers / Outlets <span className="font-normal text-sm text-gray-400">({customers.length})</span></div>
      <Btn sm onClick={()=>setAdding(true)}>+ New Customer</Btn>
    </div>
    <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search by name, city or zone…" className="w-full border rounded-xl p-2.5 text-sm outline-none"/>
    <div className="bg-white rounded-xl shadow overflow-x-auto">
      <table className="w-full text-xs min-w-max">
        <thead><tr style={{background:G1}} className="text-white text-left">
          <th className="p-2.5">Name</th><th className="p-2.5">Type</th>
          <th className="p-2.5">City</th><th className="p-2.5">Zone</th>
          <th className="p-2.5">Address</th><th className="p-2.5">Contact</th>
          <th className="p-2.5">Phone</th>
          <th className="p-2.5 text-center">Sales</th>
          <th className="p-2.5 text-center">History</th>
          {onEdit&&<th className="p-2.5 text-center">Edit</th>}
        </tr></thead>
        <tbody>
          {fl.length===0&&<tr><td colSpan={onEdit?10:9} className="text-center p-8 text-gray-400">No customers yet.</td></tr>}
          {fl.map((c,i)=><tr key={c.id} style={{background:i%2===0?LG:"white"}}>
            <td className="p-2.5 font-semibold">{c.name}</td>
            <td className="p-2.5"><Badge color={c.type==="Wholesaler"?AMB:c.type==="End User"?"#6B21A8":G1}>{c.type||"—"}</Badge></td>
            <td className="p-2.5">{c.city||"—"}</td>
            <td className="p-2.5">{c.zone||"—"}</td>
            <td className="p-2.5">{c.address||"—"}</td>
            <td className="p-2.5">{c.contact_name||"—"}</td>
            <td className="p-2.5">{c.contact_phone||"—"}</td>
            <td className="p-2.5 text-center font-bold" style={{color:G1}}>{sc(c.id)}</td>
            <td className="p-2.5 text-center">
              <button onClick={()=>setViewing(c)} style={{color:"#2980B9"}} className="text-xs font-bold hover:underline">📋 History</button>
            </td>
            {onEdit&&<td className="p-2.5 text-center">
              <button onClick={()=>setEditing({...c})} style={{color:G2}} className="text-xs font-bold hover:underline">✏️ Edit</button>
            </td>}
          </tr>)}
        </tbody>
      </table>
    </div>
    {viewing&&<CustomerHistory customer={viewing} onClose={()=>setViewing(null)} fetchHistory={fetchCustHistory}/>}
    {adding&&<Modal title="New Customer" onClose={()=>setAdding(false)}>
      <InlineCustomerForm onSave={async c=>{await onAdd(c);setAdding(false);}} onCancel={()=>setAdding(false)}/>
    </Modal>}
    {editing&&<Modal title={`Edit — ${editing.name}`} onClose={()=>setEditing(null)}>
      <div className="space-y-3">
        <Inp label="Customer Name *" value={editing.name} onChange={v=>setEditing({...editing,name:v})}/>
        <div className="grid grid-cols-2 gap-3">
          <Sel label="Type" value={editing.type||"Supermarket"} onChange={v=>setEditing({...editing,type:v})}>
            {["Supermarket","Retail Store","Wholesaler","End User"].map(t=><option key={t}>{t}</option>)}
          </Sel>
          <Sel label="Zone" value={editing.zone||"South South"} onChange={v=>setEditing({...editing,zone:v})}>
            {["South South","South East","South West","North Central","North East","North West","FCT"].map(z=><option key={z}>{z}</option>)}
          </Sel>
          <Inp label="City" value={editing.city||""} onChange={v=>setEditing({...editing,city:v})} placeholder="e.g. Port Harcourt"/>
          <Inp label="Address" value={editing.address||""} onChange={v=>setEditing({...editing,address:v})} placeholder="Street / area"/>
          <Inp label="Contact Name" value={editing.contact_name||""} onChange={v=>setEditing({...editing,contact_name:v})}/>
          <Inp label="Contact Phone" value={editing.contact_phone||""} type="tel" onChange={v=>setEditing({...editing,contact_phone:v})}/>
        </div>
        {saved&&<p style={{color:G1}} className="text-xs font-bold text-center">✓ Saved!</p>}
        <div className="flex gap-2 pt-1">
          <Btn onClick={handleEdit} disabled={!editing.name.trim()||saved}>{saved?"✓ Saved!":"Save Changes"}</Btn>
          <Btn onClick={()=>setEditing(null)} outline color="#999">Cancel</Btn>
        </div>
      </div>
    </Modal>}
  </div>;
}

// ── Team Management ───────────────────────────────────────────────────────
function TeamMgmt({team,onAdd,onRemove}){
  const [form,setForm]=useState({name:"",username:"",password:""}); const [err,setErr]=useState(""); const [ok,setOk]=useState(""); const [saving,setSaving]=useState(false);
  const add=async()=>{ if(!form.name||!form.username||!form.password) return setErr("All fields required."); if(team.find(t=>t.username.toLowerCase()===form.username.toLowerCase())) return setErr("Username taken."); setSaving(true); const r=await onAdd({id:uid(),...form}); setSaving(false); if(r){setForm({name:"",username:"",password:""});setErr("");setOk("Added!");setTimeout(()=>setOk(""),2000);}else setErr("Failed."); };
  const fmtLogin=ts=>{if(!ts) return "Never logged in"; const d=new Date(ts); return d.toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"})+" at "+d.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"});};
  return <div className="max-w-xl mx-auto space-y-4">
    <div className="bg-white rounded-xl shadow p-6"><div className="font-black text-gray-800 mb-4">Add Team Member</div>
      <Inp label="Full Name" value={form.name} onChange={v=>setForm({...form,name:v})} className="mb-3"/>
      <Inp label="Username" value={form.username} onChange={v=>setForm({...form,username:v})} className="mb-3"/>
      <Inp label="Password" type="password" value={form.password} onChange={v=>setForm({...form,password:v})} className="mb-4"/>
      {err&&<p className="text-red-500 text-xs mb-3">{err}</p>}{ok&&<p style={{color:G1}} className="text-xs mb-3 font-bold">{ok}</p>}
      <Btn onClick={add} disabled={saving}>{saving?"Saving…":"Add Member"}</Btn>
    </div>
    <div className="bg-white rounded-xl shadow p-6"><div className="font-black text-gray-800 mb-4">Team ({team.length})</div>
      {team.length===0&&<p className="text-gray-400 text-sm">No members yet.</p>}
      {team.map(m=><div key={m.id} className="flex items-center justify-between py-3 border-b last:border-0">
        <div>
          <div className="font-semibold text-sm">{m.name}</div>
          <div className="text-xs text-gray-400">@{m.username}</div>
          <div className="flex items-center gap-1 mt-0.5">
            <span className="text-xs" style={{color:m.last_login?G2:"#ccc"}}>🕐</span>
            <span className="text-xs" style={{color:m.last_login?"#555":"#ccc"}}>{fmtLogin(m.last_login)}</span>
          </div>
        </div>
        <button onClick={()=>onRemove(m.id)} className="text-xs text-red-400 hover:text-red-600 font-semibold">Remove</button>
      </div>)}
    </div>
  </div>;
}

// ── SKU Manager ───────────────────────────────────────────────────────────
function SKUManager({skus,onAdd,onDeactivate}){
  const [form,setForm]=useState({sku:"",name:"",cat:""}); const [newCat,setNewCat]=useState(""); const [err,setErr]=useState(""); const [ok,setOk]=useState("");
  const cats=[...new Set(skus.map(s=>s.cat))];
  const add=async()=>{ if(!form.sku.trim()||!form.name.trim()) return setErr("SKU code and name required."); if(skus.find(s=>s.sku===form.sku.trim())) return setErr("SKU already exists."); const cat=form.cat||newCat.trim(); if(!cat) return setErr("Select or enter a category."); await onAdd({sku:form.sku.trim(),name:form.name.trim(),cat,sort_order:99}); setForm({sku:"",name:"",cat:""}); setNewCat(""); setErr(""); setOk("SKU added!"); setTimeout(()=>setOk(""),2000); };
  return <div className="max-w-2xl mx-auto space-y-4">
    <div className="bg-white rounded-xl shadow p-6"><div className="font-black text-gray-800 mb-4">Add New SKU</div>
      <div className="grid grid-cols-2 gap-3 mb-3">
        <Inp label="SKU Code *" value={form.sku} onChange={v=>setForm({...form,sku:v})} placeholder="e.g. KYN-NEW-100g"/>
        <Inp label="Product Name *" value={form.name} onChange={v=>setForm({...form,name:v})} placeholder="e.g. Onion Powder 100g Cup"/>
      </div>
      <div className="mb-4"><label className="block text-xs font-semibold text-gray-500 mb-1">Category</label>
        <div className="flex gap-2 flex-wrap mb-2">{cats.map(c=><button key={c} onClick={()=>{setForm(p=>({...p,cat:c}));setNewCat("");}} style={form.cat===c?{background:G1,color:"white"}:{border:`1.5px solid ${G1}`,color:G1}} className="text-xs px-3 py-1 rounded-lg font-semibold">{c}</button>)}</div>
        <Inp value={newCat} onChange={v=>{setNewCat(v);setForm(p=>({...p,cat:""}));}} placeholder="Or type a new category…"/>
      </div>
      {err&&<p className="text-red-500 text-xs mb-3">{err}</p>}{ok&&<p style={{color:G1}} className="text-xs mb-3 font-bold">{ok}</p>}
      <Btn onClick={add}>Add SKU</Btn>
    </div>
    <div className="bg-white rounded-xl shadow p-4"><div className="font-black text-gray-800 mb-3">Active SKUs ({skus.length})</div>
      {cats.map(cat=><div key={cat} className="mb-4"><div className="text-xs font-black mb-2" style={{color:G1}}>{cat}</div>
        {skus.filter(s=>s.cat===cat).map(s=><div key={s.sku} className="flex items-center justify-between py-2 border-b last:border-0">
          <div><div className="text-xs font-semibold text-gray-800">{s.sku}</div><div className="text-xs text-gray-400">{s.name}</div></div>
          {!DEFAULT_SKUS.find(d=>d.sku===s.sku)&&<button onClick={()=>window.confirm(`Remove ${s.sku}?`)&&onDeactivate(s.sku)} className="text-xs text-red-400 hover:text-red-600">Remove</button>}
        </div>)}
      </div>)}
    </div>
  </div>;
}

// ── Prices & Settings ─────────────────────────────────────────────────────
function PricesMgmt({cfg,skus,onSave}){
  const normP=()=>{ const o={}; skus.forEach(s=>{ const p=cfg.prices?.[s.sku]; o[s.sku]=p&&typeof p==="object"?{retail:p.retail||0,wholesale:p.wholesale||0,end_user:p.end_user||0}:{retail:Number(p)||0,wholesale:Number(p)||0,end_user:Number(p)||0}; }); return o; };
  const [prices,setPrices]=useState(normP); const [bm,setBm]=useState(cfg.visitBenchmark||5); const [pw,setPw]=useState(cfg.adminPassword); const [saved,setSaved]=useState(false);
  const setTier=(sku,tier,val)=>setPrices(p=>({...p,[sku]:{...p[sku],[tier]:Number(val)}}));
  const save=async()=>{ await onSave({...cfg,prices,visitBenchmark:Number(bm),adminPassword:pw}); setSaved(true); setTimeout(()=>setSaved(false),2000); };
  const cats=[...new Set(skus.map(s=>s.cat))];
  return <div className="max-w-3xl mx-auto space-y-4">
    <div className="flex gap-2 flex-wrap">
      <div style={{background:"#EEF7F2",borderColor:G2}} className="flex-1 min-w-36 rounded-xl px-3 py-2.5 border flex items-center gap-2"><span>🏪</span><div><div style={{color:G2}} className="text-xs font-black">Retail Price</div><div className="text-xs text-gray-400">Supermarkets & Retail</div></div></div>
      <div style={{background:"#FFF8E7",borderColor:AMB}} className="flex-1 min-w-36 rounded-xl px-3 py-2.5 border flex items-center gap-2"><span>🏭</span><div><div style={{color:"#B7800A"}} className="text-xs font-black">Wholesale Price</div><div className="text-xs text-gray-400">Wholesalers & Distributors</div></div></div>
      <div style={{background:"#F5F3FF",borderColor:"#A855F7"}} className="flex-1 min-w-36 rounded-xl px-3 py-2.5 border flex items-center gap-2"><span>👤</span><div><div style={{color:"#6B21A8"}} className="text-xs font-black">End User Price</div><div className="text-xs text-gray-400">Direct / individual buyers</div></div></div>
    </div>
    {cats.map(cat=><div key={cat} className="bg-white rounded-xl shadow p-5">
      <div className="font-bold text-sm mb-2" style={{color:G1}}>{cat}</div>
      <div className="flex gap-2 mb-2">
        <div className="flex-1"/>
        <div className="w-28 text-center text-xs font-black" style={{color:G2}}>🏪 Retail (₦)</div>
        <div className="w-28 text-center text-xs font-black" style={{color:"#B7800A"}}>🏭 Wholesale (₦)</div>
        <div className="w-28 text-center text-xs font-black" style={{color:"#6B21A8"}}>👤 End User (₦)</div>
      </div>
      {skus.filter(s=>s.cat===cat).map(s=><div key={s.sku} className="flex items-center gap-2 mb-2.5">
        <div className="flex-1"><div className="text-xs font-semibold text-gray-700">{s.sku}</div><div className="text-xs text-gray-400">{s.name}</div></div>
        <div style={{borderColor:G2+"55"}} className="flex items-center border rounded-xl overflow-hidden w-28"><span className="px-1.5 text-xs text-gray-400 bg-gray-50 border-r py-2">₦</span><input type="number" min="0" value={prices[s.sku]?.retail??""} onChange={e=>setTier(s.sku,"retail",e.target.value)} className="flex-1 p-1.5 text-sm text-right outline-none w-0"/></div>
        <div style={{borderColor:AMB+"88"}} className="flex items-center border rounded-xl overflow-hidden w-28"><span className="px-1.5 text-xs text-gray-400 bg-gray-50 border-r py-2">₦</span><input type="number" min="0" value={prices[s.sku]?.wholesale??""} onChange={e=>setTier(s.sku,"wholesale",e.target.value)} className="flex-1 p-1.5 text-sm text-right outline-none w-0"/></div>
        <div style={{borderColor:"#A855F755"}} className="flex items-center border rounded-xl overflow-hidden w-28"><span className="px-1.5 text-xs text-gray-400 bg-gray-50 border-r py-2">₦</span><input type="number" min="0" value={prices[s.sku]?.end_user??""} onChange={e=>setTier(s.sku,"end_user",e.target.value)} className="flex-1 p-1.5 text-sm text-right outline-none w-0"/></div>
      </div>)}
    </div>)}
    <div className="bg-white rounded-xl shadow p-5"><div className="font-bold text-sm mb-4">Settings</div>
      <div className="flex items-center gap-3 mb-3"><label className="text-sm text-gray-600 flex-1">Daily Visit Benchmark</label><input type="number" min="1" value={bm} onChange={e=>setBm(e.target.value)} className="w-20 border rounded-xl p-2 text-sm text-center"/></div>
      <div className="flex items-center gap-3"><label className="text-sm text-gray-600 flex-1">Admin Password</label><input type="password" value={pw} onChange={e=>setPw(e.target.value)} className="flex-1 border rounded-xl p-2 text-sm"/></div>
    </div>
    <Btn onClick={save} color={saved?"#27ae60":G1} full>{saved?"✓ Saved!":"Save All Prices & Settings"}</Btn>
  </div>;
}

// ── Targets by SKU ────────────────────────────────────────────────────────
// ── Rep Account — password self-service ──────────────────────────────────
function RepAccount({user,onChangePassword}){
  const [curr,setCurr]=useState("");
  const [next,setNext]=useState("");
  const [confirm,setConfirm]=useState("");
  const [err,setErr]=useState("");
  const [ok,setOk]=useState(false);
  const [saving,setSaving]=useState(false);

  const valid=curr&&next&&confirm&&next===confirm&&next.length>=6;

  const handle=async()=>{
    setErr(""); if(!valid) return;
    if(next!==confirm){setErr("New passwords do not match.");return;}
    if(next.length<6){setErr("Password must be at least 6 characters.");return;}
    setSaving(true);
    const errMsg=await onChangePassword(curr,next);
    setSaving(false);
    if(errMsg){setErr(errMsg);}
    else{setOk(true);setCurr("");setNext("");setConfirm("");}
  };

  return <div className="max-w-md mx-auto space-y-4">
    <div className="bg-white rounded-xl shadow p-6">
      {/* Profile card */}
      <div style={{background:GL}} className="rounded-xl p-4 mb-5 flex items-center gap-3">
        <div style={{background:G1,color:"white"}} className="w-12 h-12 rounded-full flex items-center justify-center text-lg font-black flex-shrink-0">
          {user.name.split(" ").map(n=>n[0]).join("").slice(0,2)}
        </div>
        <div>
          <div className="font-black text-gray-800">{user.name}</div>
          <div className="text-xs text-gray-400">@{user.username}</div>
          <div className="text-xs text-gray-400 mt-0.5">Sales Representative</div>
        </div>
      </div>

      <div className="font-bold text-sm text-gray-800 mb-4">Change Password</div>

      {ok&&(
        <div style={{background:"#EEF7F2",borderColor:G2}} className="rounded-xl p-3 mb-4 border flex items-center gap-2">
          <span>✅</span>
          <span style={{color:G1}} className="text-sm font-semibold">Password changed successfully.</span>
        </div>
      )}

      <Inp label="Current Password" type="password" value={curr} onChange={v=>{setCurr(v);setErr("");setOk(false);}} className="mb-3"/>
      <Inp label="New Password" type="password" value={next} onChange={v=>{setNext(v);setErr("");setOk(false);}} className="mb-3"/>
      <Inp label="Confirm New Password" type="password" value={confirm} onChange={v=>{setConfirm(v);setErr("");setOk(false);}} className="mb-4"/>

      {next&&confirm&&next!==confirm&&(
        <p className="text-red-500 text-xs mb-2">Passwords do not match.</p>
      )}
      {next&&next.length>0&&next.length<6&&(
        <p className="text-amber-600 text-xs mb-2">Password must be at least 6 characters.</p>
      )}
      {err&&<p className="text-red-500 text-xs mb-3">{err}</p>}

      <Btn onClick={handle} disabled={!valid||saving} full>{saving?"Saving…":"Update Password"}</Btn>
    </div>

    <div className="bg-white rounded-xl shadow p-4">
      <div className="font-semibold text-sm text-gray-700 mb-2">Password tips</div>
      <ul className="text-xs text-gray-400 space-y-1">
        <li>▸ Use at least 6 characters</li>
        <li>▸ Mix letters and numbers for a stronger password</li>
        <li>▸ Do not share your password with anyone</li>
        <li>▸ Contact admin if you are locked out</li>
      </ul>
    </div>
  </div>;
}

function TargetsMgmt({team,targets,skus,onSave,selMonth,selYear}){
  const [local,setLocal]=useState({}); const [saved,setSaved]=useState(false); const [saving,setSaving]=useState(false);
  useEffect(()=>setLocal({...targets}),[targets,selMonth,selYear]);
  const get=(id,f)=>local[`${id}_${selYear}_${selMonth}`]?.[f]??"";
  const getSku=(id,sku)=>local[`${id}_${selYear}_${selMonth}`]?.skuTargets?.[sku]??"";
  const setF=(id,f,v)=>{ const k=`${id}_${selYear}_${selMonth}`; setLocal(l=>({...l,[k]:{...(l[k]||{valueNGN:0,visitsTarget:0,skuTargets:{}}),[f]:Number(v)}})); };
  const setSku=(id,sku,v)=>{ const k=`${id}_${selYear}_${selMonth}`; setLocal(l=>({...l,[k]:{...(l[k]||{valueNGN:0,visitsTarget:0,skuTargets:{}}),skuTargets:{...((l[k]||{}).skuTargets||{}),[sku]:Number(v)}}})); };
  const save=async()=>{ setSaving(true); for(const rep of team){ const k=`${rep.id}_${selYear}_${selMonth}`; await onSave(rep.id,selYear,selMonth,local[k]||{valueNGN:0,visitsTarget:0,skuTargets:{}}); } setSaving(false); setSaved(true); setTimeout(()=>setSaved(false),2000); };
  const cats=[...new Set(skus.map(s=>s.cat))];
  return <div className="max-w-2xl mx-auto">
    <div className="bg-white rounded-xl shadow p-6">
      <div className="font-black text-gray-800 mb-1">Set Monthly Targets</div>
      <div className="text-xs text-gray-400 mb-5">For: <span style={{color:G1}} className="font-bold">{MFULL[selMonth]} {selYear}</span></div>
      {team.length===0&&<p className="text-gray-400 text-sm">No team members yet.</p>}
      {team.map(m=><div key={m.id} className="mb-6 pb-6 border-b last:border-0">
        <div className="font-bold text-base text-gray-800 mb-3">{m.name}</div>
        <div className="grid grid-cols-2 gap-3 mb-4">
          <div><label className="block text-xs text-gray-400 mb-1">💰 Value Target (₦)</label><input type="number" min="0" value={get(m.id,"valueNGN")} onChange={e=>setF(m.id,"valueNGN",e.target.value)} className="w-full border rounded-xl p-2 text-sm" placeholder="e.g. 500000"/></div>
          <div><label className="block text-xs text-gray-400 mb-1">🗺️ Outlet Visits Target</label><input type="number" min="0" value={get(m.id,"visitsTarget")} onChange={e=>setF(m.id,"visitsTarget",e.target.value)} className="w-full border rounded-xl p-2 text-sm" placeholder="e.g. 110"/></div>
        </div>
        <div className="font-semibold text-xs text-gray-600 mb-3">📦 Volume Targets by SKU (units / month)</div>
        {cats.map(cat=><div key={cat} className="mb-4">
          <div className="text-xs font-black mb-2" style={{color:G1}}>{cat}</div>
          <div className="grid grid-cols-2 gap-2">
            {skus.filter(s=>s.cat===cat).map(s=><div key={s.sku} className="flex items-center gap-2">
              <label className="text-xs text-gray-500 flex-1 truncate" title={`${s.sku} — ${s.name}`}>{s.sku}</label>
              <input type="number" min="0" value={getSku(m.id,s.sku)} onChange={e=>setSku(m.id,s.sku,e.target.value)} className="w-20 border rounded-xl p-1.5 text-xs text-center" placeholder="0"/>
            </div>)}
          </div>
        </div>)}
      </div>)}
      {team.length>0&&<Btn onClick={save} disabled={saving} color={saved?"#27ae60":G1} full>{saving?"Saving…":saved?"✓ Saved!":"Save Targets"}</Btn>}
    </div>
  </div>;
}
