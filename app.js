import { PublicClientApplication } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@5/+esm";
import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@5.4.54/build/pdf.min.mjs";
import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";
import Chart from "https://cdn.jsdelivr.net/npm/chart.js@4.5.0/+esm";
pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdn.jsdelivr.net/npm/pdfjs-dist@5.4.54/build/pdf.worker.min.mjs";

const FOLDER_PATH = "Finance Dashboard/Statements";
const SCOPES = ["Files.Read"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const nav = [
 ["overview","Overview","Monthly financial snapshot"],
 ["analytics","Analytics","Charts and interactive drill-downs"],
 ["transactions","Transactions","Normalized transactions"],
 ["shared","Shared Household","Shared living expenses and settlement"],
 ["review","Review","Items requiring confirmation"],
 ["onedrive","OneDrive","Statement files and synchronization"],
 ["settings","Settings","Microsoft connection and local data"]
];

let msal = null, account = null;
let tx = JSON.parse(localStorage.getItem("finance.tx.v2") || "[]");
let charts = {};
let drillPredicate = null;
let drillLabel = "";
function getClientId(){
 return localStorage.getItem("finance.clientId") || localStorage.getItem("finance.clientId.backup") || "";
}
function setClientId(id){
 setClientId(id);
 localStorage.setItem("finance.clientId.backup",id);
}


const $ = id => document.getElementById(id);
function safeEl(id){const el=$(id); if(!el) console.warn("Missing element:",id); return el;}
const money = n => (n < 0 ? "−" : "") + "$" + Math.abs(n).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
const esc = s => String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]));

function buildNav(){
 const hosts=[$("sideNav"),$("mobileNav")].filter(Boolean);
 for(const host of hosts){
  host.innerHTML="";
  nav.forEach(([id,label])=>{
   const b=document.createElement("button");
   b.textContent=label;b.dataset.tab=id;
   if(id==="overview")b.classList.add("active");
   b.onclick=()=>show(id);
   host.appendChild(b);
  });
 }
}
function show(id){
 document.querySelectorAll(".tab").forEach(x=>x.classList.toggle("active",x.id===id));
 document.querySelectorAll("[data-tab]").forEach(x=>x.classList.toggle("active",x.dataset.tab===id));
 const t=nav.find(x=>x[0]===id);$("title").textContent=t[1];$("subtitle").textContent=t[2];
 window.scrollTo({top:0,behavior:"smooth"});
}
function parseNum(v){return parseFloat(String(v||"").replace(/[$,\s]/g,""))||0}
function parseCSV(text){
 const lines=text.replace(/\r/g,"").split("\n").filter(x=>x.trim());
 const split=s=>{let a=[],c="",q=false;for(let i=0;i<s.length;i++){let ch=s[i];if(ch==='"'){if(q&&s[i+1]==='"'){c+='"';i++}else q=!q}else if(ch===","&&!q){a.push(c);c=""}else c+=ch}a.push(c);return a};
 if(!lines.length)return[];const h=split(lines[0]).map(x=>x.trim());
 return lines.slice(1).map(line=>{const v=split(line),o={};h.forEach((k,i)=>o[k]=v[i]??"");return o});
}
function classify(desc, amount, credit){
 const d=desc.toUpperCase();
 if(credit>0 && /PAYROLL|DIRECT DEP|SALARY/.test(d)) return ["Income","Income",false];
 if(credit>0 && !/TRANSFER FROM/.test(d)) return ["Other Income","Income",false];
 if(/TRANSFER (FROM|TO)|CHASE CREDIT|CAPITAL ONE|CREDIT CARD|VENMO|ZELLE|PAYPAL/.test(d)) return ["Transfer","Transfer",false];
 if(/COMCAST|XFINITY|INTERNET/.test(d)) return ["Internet","Expense",true];
 if(/PUD|ELECTRIC|POWER|UTILITY/.test(d)) return ["Power","Expense",true];
 if(/FRED MEYER|SAFEWAY|QFC|COSTCO|GROCERY|INSTACART/.test(d)) return ["Groceries","Expense",true];
 if(/RENTER|ALLSTATE/.test(d)) return ["Renters Insurance","Expense",true];
 if(/RENT|PROPERTY MGMT|APARTMENT/.test(d)) return ["Rent","Expense",true];
 if(/CHEVRON|SHELL|ARCO|FUEL|GAS STATION/.test(d)) return ["Transportation","Expense",false];
 if(/DAIRY QUEEN|MCDONALD|STARBUCKS|CAFE|RESTAURANT/.test(d)) return ["Dining","Expense",false];
 if(/GIFTHEALTH|LILLYDIRECT|PHARM|MEDICAL|CLINIC/.test(d)) return ["Healthcare","Expense",false];
 if(/PAYPAL|AMAZON/.test(d)) return ["Needs Review","Expense",false];
 return ["Other","Expense",false];
}
function importBECU(text, source){
 const rows=parseCSV(text);let count=0;
 for(const r of rows){
  if(!r.Date||!r.Description)continue;
  const debit=parseNum(r.Debit), credit=parseNum(r.Credit), amount=credit>0?credit:debit;
  const [category,type,shared]=classify(r.Description,amount,credit);
  const key=[r.Date,r.Description,amount,source].join("|");
  if(tx.some(t=>t.key===key))continue;
  tx.push({key,date:normalizeDate(r.Date),description:cleanDesc(r.Description),category,type,shared,amount,source,confidence:category==="Needs Review"?"review":"auto"});
  count++;
 }
 save();return count;
}
function normalizeDate(s){
 const d=new Date(s);return Number.isNaN(d.getTime())?s:d.toISOString().slice(0,10);
}
function cleanDesc(s){return s.replace(/^POS Withdrawal - |^External Withdrawal - |^External Deposit - |^Deposit - /i,"").replace(/\s+- Card Ending In \d+$/i,"").trim()}
function save(){localStorage.setItem("finance.tx.v2",JSON.stringify(tx));render()}
function selectedMonthKey(){
 const v=$("monthSelect")?.value||"";
 const d=new Date(v+" 1");
 return Number.isNaN(d.getTime())?null:`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
}
function txForSelectedMonth(){
 const key=selectedMonthKey();
 return key?tx.filter(t=>String(t.date||"").startsWith(key)):tx;
}
function refreshMonthOptions(){
 const keys=[...new Set(tx.map(t=>String(t.date||"").slice(0,7)).filter(x=>/^\d{4}-\d{2}$/.test(x)))].sort().reverse();
 if(!keys.length)return;
 const current=$("monthSelect").value;
 $("monthSelect").innerHTML=keys.map(k=>{
   const [y,m]=k.split("-").map(Number);
   return `<option value="${new Date(y,m-1,1).toLocaleString(undefined,{month:"long",year:"numeric"})}">${new Date(y,m-1,1).toLocaleString(undefined,{month:"long",year:"numeric"})}</option>`;
 }).join("");
 if([...$("monthSelect").options].some(o=>o.value===current))$("monthSelect").value=current;
}
function render(){
 refreshMonthOptions();
 const monthTx=txForSelectedMonth();
 const expenses=monthTx.filter(t=>t.type==="Expense"), income=monthTx.filter(t=>t.type==="Income");
 const spend=expenses.reduce((a,t)=>a+Math.abs(t.amount),0), inc=income.reduce((a,t)=>a+Math.abs(t.amount),0), net=inc-spend;
 $("incomeKpi").textContent=money(inc);$("spendKpi").textContent=money(spend);$("netKpi").textContent=money(net);$("netKpi").className="big "+(net>=0?"good":"bad");$("rateKpi").textContent=(inc?net/inc*100:0).toFixed(1)+"%";

 const cats={};expenses.forEach(t=>cats[t.category]=(cats[t.category]||0)+Math.abs(t.amount));
 renderOverviewCategoryChart(cats);
 renderAnalyticsCharts();

 const q=($("search")?.value||"").toLowerCase(), f=$("filter")?.value||"", cf=$("categoryFilter")?.value||"";
 const rows=monthTx.filter(t=>(!q||(t.description+" "+t.category+" "+t.source).toLowerCase().includes(q))&&(!f||t.type===f)&&(!cf||t.category===cf)).sort((a,b)=>b.date.localeCompare(a.date));
 $("txBody").innerHTML=rows.map(t=>`<tr><td>${esc(t.date)}</td><td>${esc(t.description)}</td><td>${esc(t.category)}</td><td><span class="badge ${t.type.toLowerCase()}">${t.type}</span></td><td>${t.shared?'<span class="badge shared">Shared</span>':""}</td><td>${esc(t.source)}</td><td class="amount ${t.amount>0?"good":""}">${money(t.amount)}</td></tr>`).join("");

 const categories=[...new Set(monthTx.map(t=>t.category).filter(Boolean))].sort();
 const sel=$("categoryFilter"); if(sel){const old=sel.value;sel.innerHTML='<option value="">All categories</option>'+categories.map(c=>`<option>${esc(c)}</option>`).join("");if(categories.includes(old))sel.value=old}

 renderShared(expenses);renderReview();
 if(drillPredicate)renderDrilldown();
}
function renderBars(cats){
 const arr=Object.entries(cats).sort((a,b)=>b[1]-a[1]);const max=Math.max(1,...arr.map(x=>x[1]));
 $("categoryBars").innerHTML=arr.length?arr.map(([n,v],i)=>`<div class="barrow"><span>${esc(n)}</span><div class="track"><div class="bar ${i===1?"gold":""}" style="width:${v/max*100}%"></div></div><b>${money(v)}</b></div>`).join(""):`<div class="notice">Import statements to populate this view.</div>`;
}

function chartDestroy(name){if(charts[name]){charts[name].destroy();delete charts[name]}}
function chartColors(n){
 const palette=["#5a3b73","#7a5a92","#9a7caf","#b69dc5","#d0bdd9","#68456e","#8b6b86","#ac8fa4","#c8aebc","#e0ccd5"];
 return Array.from({length:n},(_,i)=>palette[i%palette.length]);
}
function renderOverviewCategoryChart(cats){
 const el=$("overviewCategoryChart");if(!el)return;chartDestroy("overviewCategory");
 const labels=Object.keys(cats), data=Object.values(cats);
 if(!labels.length)return;
 charts.overviewCategory=new Chart(el,{type:"doughnut",data:{labels,datasets:[{data,backgroundColor:chartColors(labels.length),borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"}},onClick:(e,els)=>{if(els.length){const cat=labels[els[0].index];setDrill(`Category: ${cat}`,t=>t.type==="Expense"&&t.category===cat);show("analytics")}}}});
}
function monthlySeries(){
 const buckets={};
 tx.filter(t=>t.type==="Expense"||t.type==="Income").forEach(t=>{
   const k=String(t.date||"").slice(0,7);if(!/^\d{4}-\d{2}$/.test(k))return;
   buckets[k]??={income:0,spending:0};
   if(t.type==="Income")buckets[k].income+=Math.abs(t.amount);else buckets[k].spending+=Math.abs(t.amount);
 });
 return Object.entries(buckets).sort((a,b)=>a[0].localeCompare(b[0])).map(([k,v])=>{const[y,m]=k.split("-").map(Number);return{key:k,label:new Date(y,m-1,1).toLocaleString(undefined,{month:"short",year:"2-digit"}),...v}});
}
function renderAnalyticsCharts(){
 const monthTx=txForSelectedMonth(), exp=monthTx.filter(t=>t.type==="Expense");

 const series=monthlySeries(); chartDestroy("monthly");
 if($("monthlyChart")&&series.length)charts.monthly=new Chart($("monthlyChart"),{type:"line",data:{labels:series.map(x=>x.label),datasets:[{label:"Income",data:series.map(x=>x.income),borderColor:"#5a3b73",backgroundColor:"#5a3b73",tension:.25},{label:"Spending",data:series.map(x=>x.spending),borderColor:"#b18a52",backgroundColor:"#b18a52",tension:.25}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"}},onClick:(e,els)=>{if(els.length){const key=series[els[0].index].key;setDrill(`Month: ${series[els[0].index].label}`,t=>String(t.date||"").startsWith(key)&&(t.type==="Expense"||t.type==="Income"))}}}});

 const cats={};exp.forEach(t=>cats[t.category]=(cats[t.category]||0)+Math.abs(t.amount));const catEntries=Object.entries(cats).sort((a,b)=>b[1]-a[1]);
 chartDestroy("category");
 if($("categoryChart")&&catEntries.length)charts.category=new Chart($("categoryChart"),{type:"doughnut",data:{labels:catEntries.map(x=>x[0]),datasets:[{data:catEntries.map(x=>x[1]),backgroundColor:chartColors(catEntries.length)}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"}},onClick:(e,els)=>{if(els.length){const cat=catEntries[els[0].index][0];setDrill(`Category: ${cat}`,t=>t.type==="Expense"&&t.category===cat&&txForSelectedMonth().includes(t))}}}});

 const merchants={};exp.forEach(t=>merchants[t.description]=(merchants[t.description]||0)+Math.abs(t.amount));const merch=Object.entries(merchants).sort((a,b)=>b[1]-a[1]).slice(0,10).reverse();
 chartDestroy("merchant");
 if($("merchantChart")&&merch.length)charts.merchant=new Chart($("merchantChart"),{type:"bar",data:{labels:merch.map(x=>x[0]),datasets:[{label:"Spending",data:merch.map(x=>x[1]),backgroundColor:"#7a5a92"}]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},onClick:(e,els)=>{if(els.length){const m=merch[els[0].index][0];setDrill(`Merchant: ${m}`,t=>t.type==="Expense"&&t.description===m&&txForSelectedMonth().includes(t))}}}});

 const sources={};exp.forEach(t=>sources[t.source]=(sources[t.source]||0)+Math.abs(t.amount));const src=Object.entries(sources).sort((a,b)=>b[1]-a[1]);
 chartDestroy("source");
 if($("sourceChart")&&src.length)charts.source=new Chart($("sourceChart"),{type:"bar",data:{labels:src.map(x=>x[0]),datasets:[{label:"Spending",data:src.map(x=>x[1]),backgroundColor:chartColors(src.length)}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},onClick:(e,els)=>{if(els.length){const s=src[els[0].index][0];setDrill(`Source: ${s}`,t=>t.type==="Expense"&&t.source===s&&txForSelectedMonth().includes(t))}}}});
}
function setDrill(label,predicate){drillLabel=label;drillPredicate=predicate;renderDrilldown()}
function renderDrilldown(){
 const rows=tx.filter(drillPredicate||(()=>false)).sort((a,b)=>b.date.localeCompare(a.date));
 $("drillTitle").textContent=drillLabel||"Drill-down";$("drillSub").textContent=rows.length?"Underlying transactions for this selection.":"No matching transactions.";
 const total=rows.reduce((a,t)=>a+Math.abs(t.amount),0);$("drillCount").textContent=rows.length;$("drillTotal").textContent=money(total);$("drillAvg").textContent=money(rows.length?total/rows.length:0);
 const cats=[...new Set(rows.map(t=>t.category).filter(Boolean))].sort();
 $("drillChips").innerHTML=cats.map(c=>`<button class="chip" data-cat="${esc(c)}">${esc(c)}</button>`).join("");
 $("drillChips").querySelectorAll("button").forEach(b=>b.onclick=()=>{const cat=b.dataset.cat;const base=drillPredicate;setDrill(`${drillLabel.split(" • ")[0]} • ${cat}`,t=>base(t)&&t.category===cat)});
 $("drillBody").innerHTML=rows.map(t=>`<tr><td>${esc(t.date)}</td><td>${esc(t.description)}</td><td>${esc(t.category)}</td><td>${esc(t.source)}</td><td class="amount">${money(t.amount)}</td></tr>`).join("");
}

function renderShared(expenses){
 const shared=expenses.filter(t=>t.shared),cats={};shared.forEach(t=>cats[t.category]=(cats[t.category]||0)+Math.abs(t.amount));
 const total=shared.reduce((a,t)=>a+Math.abs(t.amount),0); // V2 assumes imported statements are user's paid transactions.
 const partnerOwes=total/2;
 const rows=Object.entries(cats).sort((a,b)=>b[1]-a[1]);
 const html=rows.length?rows.map(([n,v])=>`<div class="row"><span>${esc(n)}</span><b>${money(v)}</b></div>`).join(""):`<div class="notice">No shared expenses loaded.</div>`;
 $("sharedRows").innerHTML=html;$("sharedDetail").innerHTML=html;$("settlement").textContent=total?`Partner share: ${money(partnerOwes)}`:"No shared expenses loaded";
}
function renderReview(){
 const items=tx.filter(t=>t.confidence==="review"||t.category==="Needs Review");
 $("reviewSummary").innerHTML=items.length?`<b>${items.length}</b> transaction${items.length===1?"":"s"} need categorization or matching.`:"No review items.";
 $("reviewList").innerHTML=items.length?items.map(t=>`<div class="row"><div><b>${esc(t.description)}</b><div class="meta">${esc(t.date)} • ${esc(t.source)}</div></div><div><span class="badge review">Needs Review</span> &nbsp; <b>${money(t.amount)}</b></div></div>`).join(""):`<div class="notice">Nothing waiting for review.</div>`;
}
async function initMsal(){
 const clientId=getClientId();
 $("clientId").value=clientId;
 $("redirectUri").textContent=location.href.split("#")[0].split("?")[0];
 if(!clientId)return;

 msal=new PublicClientApplication({
   auth:{
     clientId,
     authority:"https://login.microsoftonline.com/consumers",
     redirectUri:location.href.split("#")[0].split("?")[0]
   },
   cache:{
     cacheLocation:"localStorage",
     storeAuthStateInCookie:true
   }
 });

 await msal.initialize();

 let result=null;
 try{
   result=await msal.handleRedirectPromise();
 }catch(e){
   console.error("Redirect handling failed",e);
 }

 if(result?.account){
   account=result.account;
   msal.setActiveAccount(account);
 }else{
   account=msal.getActiveAccount()||msal.getAllAccounts()[0]||null;
   if(account)msal.setActiveAccount(account);
 }

 updateAuthUI();
}
function updateAuthUI(){
 $("signBtn").textContent=account?"Disconnect":"Connect Microsoft";$("syncBtn").disabled=!account;$("browseBtn").disabled=!account;if($("syncAllBtn"))$("syncAllBtn").disabled=!account;
 $("connectionStatus").className="status "+(account?"ok":"warn");$("connectionStatus").innerHTML=account?`Connected as <b>${esc(account.username||account.name||"Microsoft account")}</b>.`:"Not connected to OneDrive yet.";
}
async function sign(){
 if(account){
   try{
     await msal.logoutPopup({account,mainWindowRedirectUri:location.href.split("#")[0].split("?")[0]});
   }catch{
     await msal.logoutRedirect({account});
   }
   return;
 }

 if(!msal){
   show("settings");
   $("connectionStatus").textContent="Add your Microsoft Application (client) ID in Settings first.";
   return;
 }

 try{
   const result=await msal.loginPopup({
     scopes:SCOPES,
     prompt:"select_account"
   });
   if(result?.account){
     account=result.account;
     msal.setActiveAccount(account);
     updateAuthUI();
   }
 }catch(e){
   console.warn("Popup sign-in unavailable; falling back to redirect.",e);
   await msal.loginRedirect({scopes:SCOPES});
 }
}
async function token(){
 if(!account){
   account=msal?.getActiveAccount()||msal?.getAllAccounts()[0]||null;
   if(account)msal.setActiveAccount(account);
 }

 if(!account)throw new Error("Microsoft sign-in is not active.");

 try{
   return (await msal.acquireTokenSilent({account,scopes:SCOPES})).accessToken;
 }catch(e){
   console.warn("Silent token failed; requesting interactively.",e);
   const result=await msal.acquireTokenPopup({account,scopes:SCOPES});
   if(result?.account){
     account=result.account;
     msal.setActiveAccount(account);
     updateAuthUI();
   }
   return result.accessToken;
 }
}
function encPath(path){return path.split("/").map(encodeURIComponent).join("/")}
async function graph(url, opts={}){
 const access=await token();const res=await fetch(url,{...opts,headers:{Authorization:`Bearer ${access}`,...(opts.headers||{})}});
 if(!res.ok)throw new Error(`${res.status} ${await res.text()}`);return res;
}
async function listDrive(){
 $("driveFiles").innerHTML='<div class="status">Reading OneDrive folder…</div>';
 try{
  const url=`${GRAPH}/me/drive/root:/${encPath(FOLDER_PATH)}:/children?$select=id,name,size,file,folder,lastModifiedDateTime`;
  const data=await (await graph(url)).json();
  const files=data.value.filter(x=>x.file);window.__driveFiles=files;renderDriveFiles(files);$("syncAllBtn").disabled=false;$("recentFiles").innerHTML=files.slice(0,5).map(fileCard).join("");
  $("connectionStatus").className="status ok";$("connectionStatus").innerHTML=`Found <b>${files.length}</b> file${files.length===1?"":"s"} in ${FOLDER_PATH}.`;
 }catch(e){
  $("driveFiles").innerHTML=`<div class="status bad"><b>Could not read the statements folder.</b><br>${esc(e.message)}<br><br>Confirm that the folder exists exactly as <b>${FOLDER_PATH}</b>.</div>`;
 }
}
function fileCard(f){return `<div class="file"><div><b>${esc(f.name)}</b><br><small>${Math.round((f.size||0)/1024).toLocaleString()} KB</small></div><span class="badge">${esc((f.name.split(".").pop()||"file").toUpperCase())}</span></div>`}
function renderDriveFiles(files){
 $("driveFiles").innerHTML=files.length?files.map(f=>`<div class="file"><div><b>${esc(f.name)}</b><br><small>${Math.round((f.size||0)/1024).toLocaleString()} KB • ${esc(f.lastModifiedDateTime||"")}</small></div><button class="btn" data-id="${esc(f.id)}" data-name="${esc(f.name)}">Import</button></div>`).join(""):'<div class="status warn">Folder exists, but no files were found.</div>';
 $("driveFiles").querySelectorAll("button[data-id]").forEach(b=>b.onclick=()=>importDriveFile(b.dataset.id,b.dataset.name,b));
}
async function importDriveFile(id,name,button){
 button.disabled=true;button.textContent="Importing…";
 try{
  const res=await graph(`${GRAPH}/me/drive/items/${encodeURIComponent(id)}/content`);
  const blob=await res.blob();const rows=await processFile(name,blob);
  button.textContent=`Imported ${rows}`;
 }catch(e){button.textContent="Failed";alert(`Import failed: ${e.message}`)}
}

async function pdfToLines(blob){
 const data=new Uint8Array(await blob.arrayBuffer());
 const pdf=await pdfjsLib.getDocument({data}).promise;
 const lines=[];
 for(let p=1;p<=pdf.numPages;p++){
  const page=await pdf.getPage(p);
  const content=await page.getTextContent();
  const grouped=new Map();
  for(const item of content.items){
   const y=Math.round(item.transform?.[5]||0);
   if(!grouped.has(y))grouped.set(y,[]);
   grouped.get(y).push({x:item.transform?.[4]||0,s:item.str});
  }
  [...grouped.entries()].sort((a,b)=>b[0]-a[0]).forEach(([,items])=>{
   const line=items.sort((a,b)=>a.x-b.x).map(x=>x.s).join(" ").replace(/\s+/g," ").trim();
   if(line)lines.push(line);
  });
 }
 return lines;
}
function addTx({date,description,amount,source,category=null,type=null,shared=null,confidence="auto"}){
 if(!date||!description||!Number.isFinite(amount))return false;
 const inferred=classify(description,amount,amount>0?amount:0);
 category=category||inferred[0];type=type||inferred[1];shared=shared??inferred[2];
 const key=[date,description.toUpperCase().replace(/\s+/g," "),amount.toFixed(2),source].join("|");
 if(tx.some(t=>t.key===key))return false;
 tx.push({key,date,description:cleanDesc(description),category,type,shared,amount,source,confidence});
 return true;
}
function inferYear(lines){
 const joined=lines.slice(0,120).join(" ");
 const m=joined.match(/(?:Statement Date:?\s*|Statement Period.*?)(?:\w+\s+\d{1,2},\s*)?(20\d{2})/i) || joined.match(/\b(20\d{2})\b/);
 return m?Number(m[1]):new Date().getFullYear();
}
function importCreditCardLines(lines,source){
 const year=inferYear(lines);let count=0;
 const isChase=/CHASE/i.test(source)||lines.some(x=>/Manage your account online.*chase/i.test(x));
 const isCapital=/CAPITAL\s*ONE/i.test(source)||lines.some(x=>/Capital One/i.test(x));
 for(const line of lines){
  let m=line.match(/^(\d{2})\/(\d{2})\s+(.+?)\s+(-?\$?[\d,]+\.\d{2})(?:\s+[\d,]+)?$/);
  if(!m)continue;
  const date=`${year}-${m[1]}-${m[2]}`;
  let desc=m[3].trim(), raw=parseNum(m[4]), amount=-Math.abs(raw);
  if(/PAYMENT|CREDIT|REFUND|REDEMPTION/i.test(desc)) amount=Math.abs(raw);
  let type="Expense", category=null, shared=null;
  if(/PAYMENT|AUTOMATIC PAYMENT/i.test(desc)){type="Transfer";category="Transfer";shared=false}
  else if(/CREDIT|REFUND|REDEMPTION/i.test(desc)){type="Transfer";category="Credit / Refund";shared=false}
  else if(/INTEREST CHARGE/i.test(desc)){type="Expense";category="Interest";shared=false}
  if(addTx({date,description:desc,amount,source,category,type,shared,confidence:"statement"}))count++;
 }
 return count;
}
function importPayPalLines(lines,source){
 let count=0,current=null;
 const flush=()=>{
  if(!current)return;
  const desc=current.parts.join(" ").replace(/\s+/g," ").replace(/\bUSD\b.*$/,"").trim();
  const d=desc||"PayPal transaction";
  let type=current.amount<0?"Expense":"Income", category=null, shared=null;
  if(/General Credit Card Deposit|Bank Deposit|Transfer|Add Money|Withdrawal/i.test(d)){type="Transfer";category="Transfer";shared=false}
  else if(/Mobile Payment/i.test(d)&&current.amount>0){type="Income";category="Reimbursement / Payment";shared=false}
  if(addTx({date:current.date,description:d,amount:current.amount,source,category,type,shared,confidence:"statement"}))count++;
  current=null;
 };
 for(const line of lines){
  const m=line.match(/^(\d{2})\/(\d{2})\/(20\d{2})\s+(.+?)\s+USD\s+(-?[\d,]+\.\d{2})\s+[-\d,.]+\s+(-?[\d,]+\.\d{2})/);
  if(m){
   flush();
   current={date:`${m[3]}-${m[1]}-${m[2]}`,amount:parseNum(m[6]),parts:[m[4]]};
  }else if(current && !/^(ID:|Ref ID:|USD$|Page \d+|PAYPAL ACCOUNT|ACCOUNT ACTIVITY|DATE )/i.test(line)){
   if(!/Boeing Empl CU|PayPal Balance/i.test(line))current.parts.push(line.trim());
  }
 }
 flush();return count;
}
function importAmazonCSV(text,source){
 const rows=parseCSV(text);let count=0;
 if(!rows.length)return 0;
 const headers=Object.keys(rows[0]);
 const pick=(r,names)=>{for(const n of names){const k=headers.find(h=>h.toLowerCase()===n.toLowerCase()||h.toLowerCase().includes(n.toLowerCase()));if(k&&r[k])return r[k]}return""};
 for(const r of rows){
  const dateRaw=pick(r,["Order Date","Purchase Date","Date"]);
  const title=pick(r,["Product Name","Title","Item","Description"]);
  const totalRaw=pick(r,["Total Owed","Item Total","Order Total","Total","Amount"]);
  if(!dateRaw||!title||!totalRaw)continue;
  const d=new Date(dateRaw), date=Number.isNaN(d.getTime())?dateRaw:d.toISOString().slice(0,10);
  const amount=-Math.abs(parseNum(totalRaw));
  // Amazon export is detail/enrichment. Keep out of spending totals to avoid double-counting card statement charges.
  if(addTx({date,description:`Amazon detail: ${title}`,amount,source,category:"Amazon Purchase Detail",type:"Detail",shared:false,confidence:"detail"}))count++;
 }
 return count;
}
async function importPDF(blob,name){
 const lines=await pdfToLines(blob);
 if(/paypal/i.test(name)||lines.some(x=>/PAYPAL ACCOUNT/i.test(x)))return importPayPalLines(lines,name);
 return importCreditCardLines(lines,name);
}
async function importZIP(blob,name){
 const zip=await JSZip.loadAsync(await blob.arrayBuffer());let count=0;
 for(const [entryName,entry] of Object.entries(zip.files)){
  if(entry.dir)continue;
  if(entryName.toLowerCase().endsWith(".pdf")){
   const bytes=await entry.async("uint8array");
   count+=await importPDF(new Blob([bytes],{type:"application/pdf"}),`${name} / ${entryName}`);
  }else if(entryName.toLowerCase().endsWith(".csv")){
   const text=await entry.async("string");
   count+=/amazon/i.test(entryName)?importAmazonCSV(text,`${name} / ${entryName}`):importBECU(text,`${name} / ${entryName}`);
  }
 }
 save();return count;
}

async function processFile(name,blob){
 const lower=name.toLowerCase();let rows=0;
 if(lower.endsWith(".csv")){
  const text=await blob.text();
  rows=/amazon/i.test(name)?importAmazonCSV(text,name):importBECU(text,name);
 }else if(lower.endsWith(".pdf")){
  rows=await importPDF(blob,name);save();
 }else if(lower.endsWith(".zip")){
  rows=await importZIP(blob,name);
 }else{
  addReviewPlaceholder(name,"Unsupported statement format.");return 0;
 }
 return rows;
}
function addReviewPlaceholder(name,note){
 const key="placeholder|"+name;if(!tx.some(t=>t.key===key)){tx.push({key,date:new Date().toISOString().slice(0,10),description:name,category:"Needs Review",type:"Transfer",shared:false,amount:0,source:"Importer",confidence:"review",note});save()}
 alert(note);
}

async function syncAll(){
 const files=window.__driveFiles||[];
 if(!files.length){await listDrive();return}
 $("syncAllBtn").disabled=true;$("syncBtn").disabled=true;
 let imported=0,done=0,failed=0;
 $("syncStatus").className="status";$("syncStatus").textContent=`Syncing ${files.length} statement files…`;
 for(const f of files){
  try{
   if((f.size||0)===0){done++;continue}
   const res=await graph(`${GRAPH}/me/drive/items/${encodeURIComponent(f.id)}/content`);
   imported+=await processFile(f.name,await res.blob());done++;
   $("syncStatus").textContent=`Processed ${done}/${files.length} files • ${imported} new rows`;
  }catch(e){failed++;done++;console.error(f.name,e)}
 }
 save();
 $("syncStatus").className="status "+(failed?"warn":"ok");
 $("syncStatus").innerHTML=`<b>Sync complete.</b> ${done} files checked, ${imported} new rows imported${failed?`, ${failed} file(s) need review`:""}. Existing rows were skipped.`;
 $("syncAllBtn").disabled=false;$("syncBtn").disabled=false;
}

$("search")?.addEventListener("input",render);$("filter")?.addEventListener("change",render);$("categoryFilter")?.addEventListener("change",render);$("monthSelect")?.addEventListener("change",render);
if($("clearDrill"))$("clearDrill").onclick=()=>{drillPredicate=null;drillLabel="";$("drillTitle").textContent="Drill-down";$("drillSub").textContent="Choose a chart segment, month, KPI, or category.";$("drillBody").innerHTML="";$("drillCount").textContent="0";$("drillTotal").textContent="$0.00";$("drillAvg").textContent="$0.00";$("drillChips").innerHTML=""};
if($("incomeCard"))$("incomeCard").onclick=()=>{const key=selectedMonthKey();setDrill("Income",t=>t.type==="Income"&&(!key||String(t.date||"").startsWith(key)));show("analytics")};
if($("spendCard"))$("spendCard").onclick=()=>{const key=selectedMonthKey();setDrill("Spending",t=>t.type==="Expense"&&(!key||String(t.date||"").startsWith(key)));show("analytics")};
if($("netCard"))$("netCard").onclick=()=>{const key=selectedMonthKey();setDrill("Cash flow activity",t=>(t.type==="Income"||t.type==="Expense")&&(!key||String(t.date||"").startsWith(key)));show("analytics")};
if($("signBtn"))$("signBtn").onclick=sign;if($("syncBtn"))$("syncBtn").onclick=async()=>{await listDrive();await syncAll()};if($("browseBtn"))$("browseBtn").onclick=listDrive;if($("syncAllBtn"))$("syncAllBtn").onclick=syncAll;
if($("localBtn"))$("localBtn").onclick=()=>$("localPicker").click();if($("localPicker"))$("localPicker").onchange=async e=>{for(const f of e.target.files)await processFile(f.name,f)};
if($("saveConfig"))$("saveConfig").onclick=()=>{const id=$("clientId").value.trim();setClientId(id);alert("Saved. Reloading the app so Microsoft authentication can initialize.");location.reload()};
if($("clearData"))$("clearData").onclick=()=>{if(confirm("Clear locally cached imported transactions?")){tx=[];save()}};


async function startApp(){
 try{
   buildNav();
   render();
   await initMsal();
 }catch(e){
   console.error("Finance Dashboard startup failed",e);
   const box=$("fatalError");
   if(box){
     box.style.display="block";
     box.innerHTML=`<b>Dashboard startup error.</b><br>${esc(e?.message||String(e))}`;
   }
   // Ensure navigation still exists even if auth/chart initialization fails.
   try{buildNav()}catch{}
 }
}
startApp();
