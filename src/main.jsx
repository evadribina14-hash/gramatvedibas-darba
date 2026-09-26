import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  CalendarDays, Check, CheckCircle2, ClipboardCheck, Clock3,
  Download, FileText, Info, Plus, Search, Settings2, Trash2, Users, X
} from "lucide-react";
import "./styles.css";

const STORAGE = "gramatvedibas-darba-v5";

const uid = (p="id") => `${p}-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
const todayISO = () => new Date().toISOString().slice(0,10);

const HOLIDAYS_2026 = {
  "2026-01-01":"Jaungada diena",
  "2026-04-03":"Lielā Piektdiena",
  "2026-04-05":"Pirmās Lieldienas",
  "2026-04-06":"Otrās Lieldienas",
  "2026-05-01":"Darba svētki",
  "2026-05-04":"Latvijas Republikas Neatkarības atjaunošanas diena",
  "2026-05-10":"Mātes diena",
  "2026-05-24":"Vasarsvētki",
  "2026-06-23":"Līgo diena",
  "2026-06-24":"Jāņu diena",
  "2026-11-18":"Latvijas Republikas Proklamēšanas diena",
  "2026-12-24":"Ziemassvētki",
  "2026-12-25":"Ziemassvētki",
  "2026-12-26":"Ziemassvētki",
  "2026-12-31":"Vecgada diena"
};

const MOVED_DAYS_2026 = {
  "2026-01-02":"Brīvdiena — darba diena pārcelta uz 17.01.2026.",
  "2026-01-17":"Pārceltā darba diena no 02.01.2026.",
  "2026-06-22":"Brīvdiena — darba diena pārcelta uz 27.06.2026.",
  "2026-06-27":"Pārceltā darba diena no 22.06.2026."
};

const defaultInfo = [
  ["Nodokļu likmes","IIN","2026. gadā 25,5% līdz 105 300 EUR gada ienākumam; virs 105 300 EUR — 33%; papildu 3% gada ienākumu pārsniegumam virs 200 000 EUR.","2026"],
  ["Nodokļu likmes","Neapliekamais minimums","550 EUR mēnesī.","2026"],
  ["Nodokļu likmes","Atvieglojums par apgādājamo","250 EUR mēnesī par vienu apgādājamo.","2026"],
  ["Nodokļu likmes","Pensionāra neapliekamais minimums","1000 EUR mēnesī jeb 12 000 EUR gadā.","2026"],
  ["Nodokļu likmes","VSAOI","Darba ņēmējs 10,50%; darba devējs 23,59%; kopā 34,09%.","2026"],
  ["Nodokļu likmes","VSAOI pensionāram","Darba ņēmējs 9,25%; darba devējs 20,77%; kopā 30,02%.","2026"],
  ["PVN","PVN standarta likme","21%.","2026"],
  ["PVN","PVN reģistrācijas slieksnis","50 000 EUR.","2026"],
  ["PVN","ES preču iegādes slieksnis","10 000 EUR.","2026"],
  ["Uzņēmumi","Mikrosabiedrība","Bilance līdz 450 000 EUR; apgrozījums līdz 900 000 EUR; vidēji līdz 10 darbiniekiem.","2026"],
  ["Uzņēmumi","Maza sabiedrība","Bilance līdz 5 000 000 EUR; apgrozījums līdz 10 000 000 EUR; vidēji līdz 50 darbiniekiem.","2026"],
  ["Uzņēmumi","Vidēja sabiedrība","Bilance līdz 25 000 000 EUR; apgrozījums līdz 50 000 000 EUR; vidēji līdz 250 darbiniekiem.","2026"],
  ["Komandējumi","Dienas nauda Latvijā","8 EUR.","2026"],
  ["Darba ņēmēji","Ziņas par darba ņēmējiem","Iesniegt atbilstoši VID prasībām un MK noteikumiem Nr. 827.","2026"],
  ["Grāmatvedība","Pamatlīdzekļi","Pārbaudīt iegādi, nodošanu ekspluatācijā, inventāra numurus, nolietojumu, atlikumus un inventarizāciju.","2026"],
  ["Grāmatvedība","Reprezentācijas izdevumi","Pārbaudīt dokumentus, izdevumu ekonomisko būtību un nodokļu piemērošanu.","2026"]
].map(([category,title,value,year])=>({id:uid("info"),category,title,value,year}));

const annualGroups = [
  ["Pamatl.",["2310 atlikums pārbaudīts","2310 pareizā pusē","2380 atlikums pārbaudīts","Pamatlīdzekļu inventarizācija","Nolietojums pārbaudīts"]],
  ["Kase",["Kases atlikums pārbaudīts","Kases dokumenti pārbaudīti"]],
  ["Debitori",["Nosūtīti salīdzināšanas akti","Saņemti salīdzināšanas akti","Debitoru atlikumi pārbaudīti"]],
  ["Kreditori",["Nosūtīti salīdzināšanas akti kreditoriem","Saņemti salīdzināšanas akti no kreditoriem","Kreditoru atlikumi pārbaudīti"]],
  ["Ieņēmumi",["Ieņēmumi pārbaudīti","Ieņēmumu konti pārbaudīti"]],
  ["Izdevumi",["Izdevumi pārbaudīti","Izdevumu konti pārbaudīti","Uzkrājumi pārbaudīti"]],
  ["Citi",["Inventarizācija","Nodokļi pārbaudīti"]]
];

const defaultReports = [
  {id:"vat",name:"PVN",period:"monthly",dueDay:20,builtIn:true},
  {id:"ddz",name:"VSAOI / Darba devēja ziņojums",period:"monthly",dueDay:17,builtIn:true},
  {id:"paz",name:"Paziņojums par fiziskām personām izmaksātajām summām",period:"yearly",dueDay:31,builtIn:true},
  {id:"uin",name:"UIN",period:"monthly",dueDay:20,builtIn:true}
];

const defaultData = {
  clients: [],
  jobs: [
    {id:"job-1",name:"Ikdienas grāmatvedība"},
    {id:"job-2",name:"PVN deklarācija"},
    {id:"job-3",name:"Darba devēja ziņojums"}
  ],
  entries:[],
  info:defaultInfo,
  annual:{},
  reports:defaultReports,
  reportStatuses:{}
};

function loadData(){
  try{
    const raw=localStorage.getItem(STORAGE);
    if(!raw) return defaultData;
    const d=JSON.parse(raw);
    return {...defaultData,...d,
      clients:Array.isArray(d.clients)?d.clients:[],
      jobs:Array.isArray(d.jobs)?d.jobs:defaultData.jobs,
      entries:Array.isArray(d.entries)?d.entries:[],
      info:Array.isArray(d.info)?d.info:defaultInfo,
      annual:d.annual||{},
      reports:Array.isArray(d.reports)?d.reports:defaultReports,
      reportStatuses:d.reportStatuses||{}
    };
  }catch{return defaultData;}
}

const monthDays = m => {
  const [y,mo]=m.split("-").map(Number);
  const n=new Date(y,mo,0).getDate();
  return Array.from({length:n},(_,i)=>`${y}-${String(mo).padStart(2,"0")}-${String(i+1).padStart(2,"0")}`);
};
const fmtDate = s => new Date(`${s}T00:00:00`).toLocaleDateString("lv-LV");
const monthName = m => {
  const [y,mo]=m.split("-").map(Number);
  return new Date(y,mo-1,1).toLocaleDateString("lv-LV",{month:"long",year:"numeric"});
};
const durationText = mins => {
  const h=Math.floor(mins/60),m=mins%60;
  return h?`${h} h${m?` ${m} min`:""}`:`${m} min`;
};

function App(){
  const [data,setData]=useState(loadData);
  const [page,setPage]=useState("work");
  const [month,setMonth]=useState(new Date().toISOString().slice(0,7));
  const [clientId,setClientId]=useState("");
  const [jobId,setJobId]=useState("");
  const [newClient,setNewClient]=useState("");
  const [newJob,setNewJob]=useState("");
  const [note,setNote]=useState("");
  const [timer,setTimer]=useState(null);
  const [now,setNow]=useState(Date.now());

  const [infoSearch,setInfoSearch]=useState("");
  const [infoCategory,setInfoCategory]=useState("");
  const [infoTitle,setInfoTitle]=useState("");
  const [infoValue,setInfoValue]=useState("");
  const [infoYear,setInfoYear]=useState("2026");
  const [editInfo,setEditInfo]=useState(null);

  const [annualClient,setAnnualClient]=useState("");
  const [annualNew,setAnnualNew]=useState("");
  const [annualGroup,setAnnualGroup]=useState("Papildu darbi");
  const [annualView,setAnnualView]=useState("work");

  const [reportName,setReportName]=useState("");
  const [reportPeriod,setReportPeriod]=useState("monthly");
  const [reportDue,setReportDue]=useState("20");

  useEffect(()=>localStorage.setItem(STORAGE,JSON.stringify(data)),[data]);
  useEffect(()=>{
    if(!timer)return;
    const i=setInterval(()=>setNow(Date.now()),1000);
    return()=>clearInterval(i);
  },[timer]);
  useEffect(()=>{
    if(!clientId && data.clients[0])setClientId(data.clients[0].id);
    if(!jobId && data.jobs[0])setJobId(data.jobs[0].id);
    if(!annualClient && data.clients[0])setAnnualClient(data.clients[0].id);
  },[data.clients, data.jobs]);

  const entries=data.entries.filter(e=>e.date?.startsWith(month));
  const total=entries.reduce((s,e)=>s+e.duration,0);

  const addClient=()=>{
    if(!newClient.trim())return alert("Ievadi klienta nosaukumu.");
    const c={id:uid("client"),name:newClient.trim(),vatPeriod:"monthly"};
    setData(d=>({...d,clients:[...d.clients,c]}));
    setClientId(c.id);setAnnualClient(c.id);setNewClient("");
  };
  const deleteClient = id => {
    const c = data.clients.find(x => x.id === id);
    if (!c) return;

    setData(d => {
      const annual = { ...d.annual };
      delete annual[id];

      const statuses = { ...d.reportStatuses };
      Object.keys(statuses).forEach(k => {
        if (k.includes(`|${id}|`)) delete statuses[k];
      });

      return {
        ...d,
        clients: d.clients.filter(x => x.id !== id),
        annual,
        reportStatuses: statuses
      };
    });

    if (clientId === id) setClientId("");
    if (annualClient === id) setAnnualClient("");
  };
  const addJob=()=>{
    if(!newJob.trim())return;
    const j={id:uid("job"),name:newJob.trim()};
    setData(d=>({...d,jobs:[...d.jobs,j]}));setJobId(j.id);setNewJob("");
  };
  const deleteJob=id=>{
    const j=data.jobs.find(x=>x.id===id);
    if(!j)return;
    if(!confirm(`Dzēst darbu "${j.name}"?\n\nVēsturiskie ieraksti paliks.`))return;
    setData(d=>({...d,jobs:d.jobs.filter(x=>x.id!==id)}));
    if(jobId===id)setJobId("");
  };
  const start=()=>{
    if(!clientId||!jobId)return alert("Izvēlies klientu un darbu.");
    if(timer)return;
    setTimer({clientId,jobId,start:Date.now(),date:todayISO(),note});
  };
  const stop=()=>{
    if(!timer)return;
    const end=Date.now();
    const duration=Math.max(1,Math.round((end-timer.start)/60000));
    setData(d=>({...d,entries:[{id:uid("entry"),clientId:timer.clientId,jobId:timer.jobId,date:timer.date,duration,note:timer.note||""},...d.entries]}));
    setTimer(null);setNote("");
  };
  const manual=()=>{
    if(!clientId||!jobId)return alert("Izvēlies klientu un darbu.");
    const mins=Number(prompt("Cik minūtes?","60"));if(!mins||mins<=0)return;
    const date=prompt("Datums YYYY-MM-DD",todayISO());if(!date)return;
    setData(d=>({...d,entries:[{id:uid("entry"),clientId,jobId,date,duration:Math.round(mins),note},...d.entries]}));setNote("");
  };

  const saveInfo=()=>{
    if(!infoCategory.trim()||!infoTitle.trim()||!infoValue.trim())return alert("Aizpildi kategoriju, nosaukumu un informāciju.");
    setData(d=>({...d,info:editInfo?d.info.map(x=>x.id===editInfo?{...x,category:infoCategory.trim(),title:infoTitle.trim(),value:infoValue.trim(),year:infoYear}:{...x}):[...d.info,{id:uid("info"),category:infoCategory.trim(),title:infoTitle.trim(),value:infoValue.trim(),year:infoYear}]}));
    setInfoCategory("");setInfoTitle("");setInfoValue("");setInfoYear("2026");setEditInfo(null);
  };
  const startEditInfo=x=>{setInfoCategory(x.category);setInfoTitle(x.title);setInfoValue(x.value);setInfoYear(x.year||"");setEditInfo(x.id);window.scrollTo({top:0,behavior:"smooth"});};
  const deleteInfo=id=>{if(confirm("Dzēst šo informāciju?"))setData(d=>({...d,info:d.info.filter(x=>x.id!==id)}));};

  const annualTasks=useMemo(()=>{
    const standard=annualGroups.flatMap(([group,tasks])=>tasks.map(t=>({id:`${group}-${t}`,group,text:t,custom:false})));
    return [...standard,...(data.annual[annualClient]?.custom||[])];
  },[data.annual,annualClient]);
  const toggleAnnual=id=>setData(d=>{
    const a=d.annual[annualClient]||{done:{},custom:[]};
    return {...d,annual:{...d.annual,[annualClient]:{...a,done:{...a.done,[id]:!a.done?.[id]}}}};
  });
  const addAnnual=()=>{
    if(!annualNew.trim()||!annualClient)return;
    const a=data.annual[annualClient]||{done:{},custom:[]};
    const task={id:uid("annual"),group:annualGroup||"Papildu darbi",text:annualNew.trim(),custom:true};
    setData(d=>({...d,annual:{...d.annual,[annualClient]:{...a,custom:[...a.custom,task]}}}));
    setAnnualNew("");
  };
  const deleteAnnual=id=>setData(d=>({...d,annual:{...d.annual,[annualClient]:{...d.annual[annualClient],custom:(d.annual[annualClient]?.custom||[]).filter(x=>x.id!==id)}}}));

  const addReport=()=>{
    if(!reportName.trim())return alert("Ievadi atskaites nosaukumu.");
    const day=Number(reportDue);if(day<1||day>31)return alert("Termiņa dienai jābūt 1–31.");
    setData(d=>({...d,reports:[...d.reports,{id:uid("report"),name:reportName.trim(),period:reportPeriod,dueDay:day,builtIn:false}]}));
    setReportName("");setReportDue("20");
  };
  const deleteReport=id=>{const r=data.reports.find(x=>x.id===id);if(!r||r.builtIn)return;if(confirm(`Dzēst atskaiti "${r.name}"?`))setData(d=>({...d,reports:d.reports.filter(x=>x.id!==id)}));};
  const statusKey=(rid,cid)=>`${month}|${cid}|${rid}`;
  const status=(rid,cid)=>data.reportStatuses[statusKey(rid,cid)]||{submitted:false,taxSent:false};
  const toggleStatus=(rid,cid,field)=>setData(d=>{const k=statusKey(rid,cid),s=d.reportStatuses[k]||{submitted:false,taxSent:false};return {...d,reportStatuses:{...d.reportStatuses,[k]:{...s,[field]:!s[field]}}};});

  const exportExcel=()=>{
    const dates=monthDays(month);
    const rows=data.clients.map(c=>{
      const r={Klients:c.name};
      dates.forEach(date=>{const mins=entries.filter(e=>e.clientId===c.id&&e.date===date).reduce((s,e)=>s+e.duration,0);r[date]=mins?Number((mins/60).toFixed(2)):"";});
      const mins=entries.filter(e=>e.clientId===c.id).reduce((s,e)=>s+e.duration,0);r["Mēnesī"]=Math.round(mins/60);return r;
    });
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Darba laiks");
    const notes=entries.map(e=>({Datums:fmtDate(e.date),Klients:data.clients.find(c=>c.id===e.clientId)?.name||"Dzēsts klients","Paveiktais darbs":e.note||data.jobs.find(j=>j.id===e.jobId)?.name||""}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(notes),"Piezīmes");
    XLSX.writeFile(wb,`gramatvedibas-darba-${month}.xlsx`);
  };

  const nav=[
    ["work","Darba laiks",Clock3],
    ["table","Atskaite",ClipboardCheck],
    ["calendar","Kalendārs",CalendarDays],
    ["info","Informācija",Info],
    ["annual","Gada pārskats",FileText],
    ["deadlines","Termiņi",CheckCircle2]
  ];

  return <div className="app">
    <header className="topbar">
      <h1>Grāmatvedības darba pārvaldība</h1>
      <p>Darba laiks, klienti, atskaites, termiņi un informācija vienuviet.</p>
    </header>
    <nav className="nav">{nav.map(([id,label,Icon])=><button key={id} className={`nav-btn ${page===id?"active":""}`} onClick={()=>setPage(id)}><Icon/><span>{label}</span></button>)}</nav>

    <main className="content">
      {page==="work"&&<WorkPage {...{data,clientId,setClientId,jobId,setJobId,newClient,setNewClient,newJob,setNewJob,note,setNote,timer,start,stop,manual,deleteClient,addClient,addJob,deleteJob}}/>}
      {page==="table"&&<TablePage {...{data,month,setMonth,entries,total,exportExcel}}/>}
      {page==="calendar"&&<CalendarPage month={month} setMonth={setMonth}/>}
      {page==="info"&&<InfoPage {...{data,infoSearch,setInfoSearch,infoCategory,setInfoCategory,infoTitle,setInfoTitle,infoValue,setInfoValue,infoYear,setInfoYear,editInfo,startEditInfo,saveInfo,deleteInfo}}/>}
      {page==="annual"&&<AnnualPage {...{data,annualClient,setAnnualClient,annualTasks,toggleAnnual,annualNew,setAnnualNew,annualGroup,setAnnualGroup,addAnnual,deleteAnnual,annualView,setAnnualView}}/>}
      {page==="deadlines"&&<ReportsPage {...{data,month,setMonth,reportName,setReportName,reportPeriod,setReportPeriod,reportDue,setReportDue,addReport,deleteReport,status,toggleStatus}}/>}
    </main>
  </div>;
}

function WorkPage(p){
  return <>
    <div className="card">
      <Section title="Darba laiks" text="Uzskaiti laiku pēc klienta un darba." icon={<Clock3/>}/>
      <div className="form-grid">
        <div><label>Klients</label><div className="inline">
          <select value={p.clientId} onChange={e=>p.setClientId(e.target.value)}><option value="">Izvēlies klientu</option>{p.data.clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>
          {p.clientId&&<button className="danger-btn" onClick={()=>p.deleteClient(p.clientId)}><Trash2 size={16}/>Dzēst</button>}
        </div>
        <div className="inline mt"><input placeholder="Jauns klients..." value={p.newClient} onChange={e=>p.setNewClient(e.target.value)}/><button className="icon-btn" onClick={p.addClient}><Plus/></button></div></div>
        <div><label>Darbs</label><select value={p.jobId} onChange={e=>p.setJobId(e.target.value)}><option value="">Izvēlies darbu</option>{p.data.jobs.map(j=><option key={j.id} value={j.id}>{j.name}</option>)}</select>
        <div className="inline mt"><input placeholder="Jauns darbs..." value={p.newJob} onChange={e=>p.setNewJob(e.target.value)}/><button className="icon-btn" onClick={p.addJob}><Plus/></button></div></div>
      </div>
      <label className="mt-label">Ko izdarīji?</label><textarea value={p.note} onChange={e=>p.setNote(e.target.value)} placeholder="Piemēram: pārbaudīti bankas izraksti..."/>
      {!p.timer?<button className="primary-btn" onClick={p.start}><Clock3/>Sākt darbu</button>:<div className="timer-box"><strong>Darbs tiek uzskaitīts</strong><div className="timer">00:00:00</div><button className="primary-btn" onClick={p.stop}><CheckCircle2/>Pabeigt darbu</button></div>}
      <button className="secondary-btn mt" onClick={p.manual}><Plus/>Pievienot laiku manuāli</button>
    </div>
    <div className="card"><Section title="Klientu pārvaldība" text="Jebkuru klientu vari izdzēst, kad vairs ar viņu nestrādā." icon={<Users/>}/>
      {p.data.clients.length?<div className="manage-list">{p.data.clients.map(c=><div className="manage-row" key={c.id}><div><strong>{c.name}</strong><span>PVN: {c.vatPeriod==="monthly"?"katru mēnesi":"nav PVN"}</span></div><button className="danger-btn" onClick={()=>p.deleteClient(c.id)}><Trash2/>Dzēst</button></div>)}</div>:<div className="empty-state"><Users/><strong>Nav klientu</strong><span>Pievieno pirmo klientu augstāk.</span></div>}
    </div>
    <div className="card"><Section title="Darbu pārvaldība" text="Arī darbus vari dzēst." icon={<Settings2/>}/><div className="manage-list">{p.data.jobs.map(j=><div className="manage-row" key={j.id}><strong>{j.name}</strong><button className="danger-btn" onClick={()=>p.deleteJob(j.id)}><Trash2/>Dzēst</button></div>)}</div></div>
  </>;
}

function TablePage({data,month,setMonth,entries,total,exportExcel}){
  const dates=monthDays(month);
  return <div className="card"><Section title="Darba laika atskaite" text={`${monthName(month)} — mēneša kopsumma noapaļota līdz pilnām stundām.`} icon={<ClipboardCheck/>}/>
    <div className="toolbar"><input type="month" value={month} onChange={e=>setMonth(e.target.value)}/><button className="secondary-btn" onClick={exportExcel}><Download/>Excel</button></div>
    <div className="table-wrap"><table><thead><tr><th>Klients</th>{dates.map(d=><th key={d}>{d.slice(-2)}</th>)}<th>Mēnesī</th></tr></thead><tbody>{data.clients.map(c=>{const cm=entries.filter(e=>e.clientId===c.id).reduce((s,e)=>s+e.duration,0);return <tr key={c.id}><td><strong>{c.name}</strong></td>{dates.map(d=>{const m=entries.filter(e=>e.clientId===c.id&&e.date===d).reduce((s,e)=>s+e.duration,0);return <td key={d}>{m?(m/60).toFixed(2):""}</td>})}<td><strong>{Math.round(cm/60)} h</strong></td></tr>})}<tr className="total-row"><td><strong>KOPĀ</strong></td>{dates.map(d=>{const m=entries.filter(e=>e.date===d).reduce((s,e)=>s+e.duration,0);return <td key={d}>{m?(m/60).toFixed(2):""}</td>})}<td><strong>{Math.round(total/60)} h</strong></td></tr></tbody></table></div>
  </div>;
}

function CalendarPage({month,setMonth}){
  const [y,m]=month.split("-").map(Number);const first=new Date(y,m-1,1);const n=new Date(y,m,0).getDate();const start=(first.getDay()+6)%7;const cells=[];
  for(let i=0;i<start;i++)cells.push(<div className="calendar-day empty" key={"e"+i}/>);
  for(let d=1;d<=n;d++){const s=`${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`;const wd=new Date(y,m-1,d).getDay();const weekend=wd===0||wd===6;const holiday=HOLIDAYS_2026[s];const moved=MOVED_DAYS_2026[s];cells.push(<div key={s} title={holiday||moved||""} className={`calendar-day ${weekend?"red":""} ${holiday?"holiday":""} ${moved?"moved":""}`}><b>{d}</b>{holiday&&<small>Sv</small>}{moved&&<small>P</small>}</div>);}
  return <div className="card"><Section title="Kalendārs" text="Sarkani — brīvdienas. Ar īpašu atzīmi — Latvijas svētku dienas un 2026. gada pārceltās darba dienas." icon={<CalendarDays/>}/><div className="toolbar"><input type="month" value={month} onChange={e=>setMonth(e.target.value)}/></div><div className="calendar">{["P","O","T","C","P","S","Sv"].map(x=><div className="calendar-head" key={x}>{x}</div>)}{cells}</div><div className="legend"><i className="legend-red"/> Brīvdiena <i className="legend-holiday"/> Svētku diena <i className="legend-moved"/> Pārceltā darba diena</div></div>;
}

function InfoPage(p){
  const filtered=p.data.info.filter(x=>`${x.category} ${x.title} ${x.value}`.toLowerCase().includes(p.infoSearch.toLowerCase()));
  const cats=[...new Set(filtered.map(x=>x.category))];
  return <><div className="card"><Section title="Informācija" text="Meklē jebkuru nodokļu, grāmatvedības vai darba informāciju." icon={<Info/>}/><div className="search-box"><Search/><input placeholder="Meklēt informāciju..." value={p.infoSearch} onChange={e=>p.setInfoSearch(e.target.value)}/>{p.infoSearch&&<button onClick={()=>p.setInfoSearch("")}><X/></button>}</div><div className="form-grid mt"><input placeholder="Kategorija" value={p.infoCategory} onChange={e=>p.setInfoCategory(e.target.value)}/><input placeholder="Nosaukums" value={p.infoTitle} onChange={e=>p.setInfoTitle(e.target.value)}/></div><textarea className="mt" placeholder="Informācija..." value={p.infoValue} onChange={e=>p.setInfoValue(e.target.value)}/><div className="form-grid"><input placeholder="Gads" value={p.infoYear} onChange={e=>p.setInfoYear(e.target.value)}/><button className="primary-btn" onClick={p.saveInfo}>{p.editInfo?<><Check/>Saglabāt izmaiņas</>:<><Plus/>Pievienot informāciju</>}</button></div></div>
    {cats.map(cat=><div className="card" key={cat}><Section title={cat} text="" icon={<FileText/>}/><div className="info-grid">{filtered.filter(x=>x.category===cat).map(x=><div className="info-card" key={x.id}><div className="info-head"><strong>{x.title}</strong><span>{x.year}</span></div><p>{x.value}</p><div className="row-actions"><button className="secondary-btn" onClick={()=>p.startEditInfo(x)}>Labot</button><button className="danger-btn" onClick={()=>p.deleteInfo(x.id)}><Trash2/></button></div></div>)}</div></div>)}
    {!filtered.length&&<div className="card empty-state"><Search/><strong>Nekas netika atrasts</strong><span>Pamēģini citu vārdu.</span></div>}
  </>;
}

function AnnualPage(p){
  if(!p.data.clients.length)return <div className="card empty-state"><Users/><strong>Vispirms pievieno klientu</strong><span>Tad šeit parādīsies gada pārskata klienti.</span></div>;

  const selected=p.data.clients.find(c=>c.id===p.annualClient)||p.data.clients[0];
  const custom=p.data.annual[selected.id]?.custom||[];
  const allTasks=[
    ...annualGroups.flatMap(([group,list])=>list.map(text=>({id:`${group}-${text}`,group,text,custom:false}))),
    ...custom
  ];
  const done=allTasks.filter(x=>p.data.annual[selected.id]?.done?.[x.id]).length;
  const progress=allTasks.length?Math.round(done/allTasks.length*100):0;

  const moveClient=(dir)=>{
    const i=p.data.clients.findIndex(c=>c.id===selected.id);
    const next=p.data.clients[(i+dir+p.data.clients.length)%p.data.clients.length];
    p.setAnnualClient(next.id);
  };

  if(p.annualView==="progress") return <div className="card">
    <div className="annual-head"><div><h2>Gada pārskata progress</h2><p>Atsevišķs pārskats par visu klientu izpildi.</p></div><button className="secondary-btn" onClick={()=>p.setAnnualView("work")}>Atgriezties</button></div>
    <div className="year-bar"><button className="icon-btn" onClick={()=>{}}><span>‹</span></button><div><strong>2026. gads</strong><small>Gada pārskats</small></div><button className="icon-btn" onClick={()=>{}}><span>›</span></button></div>
    <div className="progress-summary"><div><strong>Kopējais progress</strong><b>{Math.round(p.data.clients.reduce((sum,c)=>{const ts=[...annualGroups.flatMap(([g,list])=>list.map(t=>`${g}-${t}`)),...(p.data.annual[c.id]?.custom||[]).map(t=>t.id)];const d=ts.filter(id=>p.data.annual[c.id]?.done?.[id]).length;return sum+(ts.length?d/ts.length:0)},0)/Math.max(1,p.data.clients.length)*100)}%</b></div><div className="progress"><i style={{width:`${Math.round(p.data.clients.reduce((sum,c)=>{const ts=[...annualGroups.flatMap(([g,list])=>list.map(t=>`${g}-${t}`)),...(p.data.annual[c.id]?.custom||[]).map(t=>t.id)];const d=ts.filter(id=>p.data.annual[c.id]?.done?.[id]).length;return sum+(ts.length?d/ts.length:0)},0)/Math.max(1,p.data.clients.length)*100)}%`}}/></div></div>
    <div className="progress-client-list">{p.data.clients.map(c=>{
      const ts=[...annualGroups.flatMap(([g,list])=>list.map(t=>`${g}-${t}`)),...(p.data.annual[c.id]?.custom||[]).map(t=>t.id)];
      const d=ts.filter(id=>p.data.annual[c.id]?.done?.[id]).length;
      const pct=ts.length?Math.round(d/ts.length*100):0;
      return <button className="progress-client" key={c.id} onClick={()=>{p.setAnnualClient(c.id);p.setAnnualView("work")}}><div><strong>{c.name}</strong><span>{d}/{ts.length}</span></div><div className="progress"><i style={{width:`${pct}%`}}/></div><b>{pct}%</b><span>›</span></button>;
    })}</div>
  </div>;

  const groups=[...annualGroups, ...(custom.length?[['Mani papildu darbi',custom.map(x=>x.text)]]:[])];
  const taskByGroup={};
  allTasks.forEach(t=>{if(!taskByGroup[t.group])taskByGroup[t.group]=[];taskByGroup[t.group].push(t);});

  return <div className="card">
    <div className="annual-head"><div><h2>Gada pārskats</h2><p>Katram klientam sava pārbaudes tabula.</p></div><button className="secondary-btn" onClick={()=>p.setAnnualView("progress")}>Progress</button></div>

    <div className="year-bar client-switcher"><button className="icon-btn" onClick={()=>moveClient(-1)}><span>‹</span></button><div><strong>{selected.name}</strong><small>2026. gads</small></div><button className="icon-btn" onClick={()=>moveClient(1)}><span>›</span></button></div>
    <select className="client-picker" value={selected.id} onChange={e=>p.setAnnualClient(e.target.value)}>{p.data.clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>

    <div className="annual-progress-line"><div><span>{done} / {allTasks.length} izdarīti</span><b>{progress}%</b></div><div className="progress"><i style={{width:`${progress}%`}}/></div></div>

    <div className="annual-simple-table-wrap">
      <table className="annual-simple-table">
        <thead><tr><th>Darbība</th><th>Statuss</th></tr></thead>
        <tbody>
          {Object.entries(taskByGroup).map(([group,tasks])=> <React.Fragment key={group}>
            <tr className="annual-group-row"><th colSpan="2">{group}</th></tr>
            {tasks.map(t=>{
              const isDone=!!p.data.annual[selected.id]?.done?.[t.id];
              return <tr key={t.id} className={isDone?"annual-task-done":""}>
                <td>{t.text}</td>
                <td><div className="annual-status-actions"><button className={`annual-status-button ${isDone?"done":"pending"}`} onClick={()=>p.toggleAnnual(t.id)} aria-label={isDone?"Izdarīts":"Nav izdarīts"}>{isDone?<Check/>:<span>○</span>}</button>{t.custom&&<button className="annual-delete" onClick={()=>p.deleteAnnual(t.id)} aria-label="Dzēst darbību"><Trash2/></button>}</div></td>
              </tr>;
            })}
          </React.Fragment>)}
        </tbody>
      </table>
    </div>

    <div className="annual-legend"><span><b className="legend-dot green">✓</b> Izdarīts</span><span><b className="legend-dot yellow">○</b> Nav izdarīts</span></div>

    <div className="notes-box annual-add-box"><h3>Pievienot darbību šim klientam</h3><p>Piemēram, 5310 konts vai cita pārbaude, kas vajadzīga tikai {selected.name}.</p><div className="form-grid"><input placeholder="Grupa, piem. Pamatlīdzekļi" value={p.annualGroup} onChange={e=>p.setAnnualGroup(e.target.value)}/><input placeholder="Darbība vai konts" value={p.annualNew} onChange={e=>p.setAnnualNew(e.target.value)}/></div><button className="secondary-btn mt" onClick={p.addAnnual}><Plus/>Pievienot darbību</button></div>
  </div>;
}
function ReportsPage(p){
  const shortName=(r)=>{
    const n=r.name.toLowerCase();
    if(n.includes("pvn")) return "PVN";
    if(n.includes("vsaoi")||n.includes("darba devēja")) return "VSAOI";
    if(n.includes("paziņojums")) return "Paziņojums";
    if(n.includes("uin")) return "UIN";
    return r.name.length>13?r.name.slice(0,13)+"…":r.name;
  };
  return <><div className="card">
    <Section title="Atskaites" text="Visi klienti vienā kompaktā pārskatā." icon={<ClipboardCheck/>}/>
    <div className="toolbar"><input type="month" value={p.month} onChange={e=>p.setMonth(e.target.value)}/></div>
    <div className="report-table-wrap"><table className="report-table"><thead><tr><th>Klients</th>{p.data.reports.map(r=><th key={r.id}>{shortName(r)}<small>Līdz {r.dueDay}.</small></th>)}</tr></thead><tbody>{p.data.clients.map(c=><tr key={c.id}><td><strong>{c.name}</strong></td>{p.data.reports.map(r=>{
      const s=p.status(r.id,c.id);
      const state=s.submitted&&s.taxSent?"both":s.submitted?"submitted":s.taxSent?"tax":"none";
      return <td key={r.id}><div className="report-checks">
        <button type="button" className={`report-status ${state}`} onClick={()=>{
          if(state==="none") p.toggleStatus(r.id,c.id,"submitted");
          else if(state==="submitted") p.toggleStatus(r.id,c.id,"taxSent");
          else if(state==="both") p.toggleStatus(r.id,c.id,"submitted");
          else p.toggleStatus(r.id,c.id,"submitted");
        }} title={state==="both"?"Iesniegta + nodokļu informācija":state==="submitted"?"Iesniegta":state==="tax"?"Nodokļu informācija":"Nav iesniegta"}>
          {state==="both"?<span>✓✓</span>:state==="submitted"?<span>✓</span>:state==="tax"?<span>✓</span>:<span>○</span>}
        </button>
      </div></td>;
    })}</tr>)}</tbody></table></div>
    <div className="report-legend"><span><b className="legend-dot green">✓</b> Iesniegta</span><span><b className="legend-dot green both">✓✓</b> Iesniegta + nodokļi paziņoti</span><span><b className="legend-dot yellow">○</b> Nav iesniegta</span></div>
  </div>
  <div className="card"><Section title="Pievienot atskaiti" text="Vari pievienot jebkuru savu atskaiti." icon={<Plus/>}/><div className="form-grid"><input placeholder="Atskaites nosaukums" value={p.reportName} onChange={e=>p.setReportName(e.target.value)}/><select value={p.reportPeriod} onChange={e=>p.setReportPeriod(e.target.value)}><option value="monthly">Katru mēnesi</option><option value="quarterly">Reizi ceturksnī</option><option value="yearly">Reizi gadā</option></select></div><div className="form-grid mt"><input inputMode="numeric" placeholder="Termiņa diena" value={p.reportDue} onChange={e=>p.setReportDue(e.target.value)}/><button className="primary-btn" onClick={p.addReport}><Plus/>Pievienot atskaiti</button></div></div>
  <div className="card"><Section title="Manas atskaites" text="Pašas pievienotās atskaites vari arī dzēst." icon={<Settings2/>}/><div className="manage-list">{p.data.reports.map(r=><div className="manage-row" key={r.id}><div><strong>{shortName(r)}</strong><span>{r.name} · līdz {r.dueDay}. datumam</span></div>{!r.builtIn&&<button className="danger-btn" onClick={()=>p.deleteReport(r.id)}><Trash2/>Dzēst</button>}</div>)}</div></div></>;
}

function Section({title,text,icon}){return <div className="section-title"><div><h2>{title}</h2>{text&&<p>{text}</p>}</div>{icon}</div>}

createRoot(document.getElementById("root")).render(<App/>);
