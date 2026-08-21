import { PublicClientApplication } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@5/+esm";
import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@5.4.54/build/pdf.min.mjs";
import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";
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
let drillPredicate=null, drillLabel="";

const $ = id => document.getElementById(id);
const money = n => (n < 0 ? "−" : "") + "$" + Math.abs(n).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
const esc = s => String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]));

function buildNav(){
 for(const host of [$("sideNav"),$("mobileNav")]){
  nav.forEach(([id,label])=>{const b=document.createElement("button");b.textContent=label;b.dataset.tab=id;if(id==="overview")b.classList.add("active");b.onclick=()=>show(id);host.appendChild(b)})
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
function refreshMonthOptions(){
 const keys=[...new Set(tx.map(t=>String(t.date||"").slice(0,7)).filter(k=>/^\d{4}-\d{2}$/.test(k)))].sort().reverse();
 if(!keys.length)return;
 const current=$("monthSelect").value;
 $("monthSelect").innerHTML=keys.map(k=>{const[y,m]=k.split("-").map(Number);const label=new Date(y,m-1,1).toLocaleString(undefined,{month:"long",year:"numeric"});return `<option>${label}</option>`}).join("");
 if([...$("monthSelect").options].some(o=>o.value===current))$("monthSelect").value=current;
}
function monthTransactions(){
 const key=selectedMonthKey();return key?tx.filter(t=>String(t.date||"").startsWith(key)):tx;
}
function render(){
 refreshMonthOptions();
 const mt=monthTransactions();
 const expenses=mt.filter(t=>t.type==="Expense"), income=mt.filter(t=>t.type==="Income");
 const spend=expenses.reduce((a,t)=>a+Math.abs(t.amount),0), inc=income.reduce((a,t)=>a+Math.abs(t.amount),0), net=inc-spend;
 $("incomeKpi").textContent=money(inc);$("spendKpi").textContent=money(spend);$("netKpi").textContent=money(net);$("netKpi").className="big "+(net>=0?"good":"bad");$("rateKpi").textContent=(inc?net/inc*100:0).toFixed(1)+"%";
 renderCharts(mt);

 const q=($("search")?.value||"").toLowerCase(), f=$("filter")?.value||"", cf=$("categoryFilter")?.value||"";
 const rows=mt.filter(t=>(!q||(t.description+" "+t.category+" "+t.source).toLowerCase().includes(q))&&(!f||t.type===f)&&(!cf||t.category===cf)).sort((a,b)=>b.date.localeCompare(a.date));
 $("txBody").innerHTML=rows.map(t=>`<tr><td>${esc(t.date)}</td><td>${esc(t.description)}</td><td>${esc(t.category)}</td><td><span class="badge ${t.type.toLowerCase()}">${t.type}</span></td><td>${t.shared?'<span class="badge shared">Shared</span>':""}</td><td>${esc(t.source)}</td><td class="amount ${t.amount>0?"good":""}">${money(t.amount)}</td></tr>`).join("");
 const cats=[...new Set(mt.map(t=>t.category).filter(Boolean))].sort();
 if($("categoryFilter")){const old=$("categoryFilter").value;$("categoryFilter").innerHTML='<option value="">All categories</option>'+cats.map(c=>`<option>${esc(c)}</option>`).join("");if(cats.includes(old))$("categoryFilter").value=old}
 renderShared(expenses);renderReview();if(drillPredicate)renderDrill();
}
function renderBars(cats){
 const arr=Object.entries(cats).sort((a,b)=>b[1]-a[1]);const max=Math.max(1,...arr.map(x=>x[1]));
 $("categoryBars").innerHTML=arr.length?arr.map(([n,v],i)=>`<div class="barrow"><span>${esc(n)}</span><div class="track"><div class="bar ${i===1?"gold":""}" style="width:${v/max*100}%"></div></div><b>${money(v)}</b></div>`).join(""):`<div class="notice">Import statements to populate this view.</div>`;
}

const purple=["#3b1d4a","#573168","#704781","#896099","#a27ab0","#bb95c6","#d2b2d9","#6b4c78","#8d6d99","#ad91b6"];
function sumBy(rows,key){const o={};rows.forEach(t=>o[t[key]]=(o[t[key]]||0)+Math.abs(t.amount));return o}
function donut(elId,legendId,entries,labelPrefix){
 const el=$(elId), leg=$(legendId);if(!el)return;
 if(!entries.length){el.innerHTML='<div class="notice">No spending data for this period.</div>';if(leg)leg.innerHTML="";return}
 const total=entries.reduce((a,x)=>a+x[1],0),cx=150,cy=130,r=82,sw=46,C=2*Math.PI*r;let offset=0;
 const circles=entries.map(([name,val],i)=>{const frac=val/total,dash=frac*C, gap=C-dash;const html=`<circle data-name="${esc(name)}" cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${purple[i%purple.length]}" stroke-width="${sw}" stroke-dasharray="${dash} ${gap}" stroke-dashoffset="${-offset}" transform="rotate(-90 ${cx} ${cy})" style="cursor:pointer"/>`;offset+=dash;return html}).join("");
 el.innerHTML=`<svg viewBox="0 0 300 260" role="img">${circles}<text x="${cx}" y="${cy-4}" text-anchor="middle" font-size="12" fill="#6b7280">Total</text><text x="${cx}" y="${cy+20}" text-anchor="middle" font-size="21" font-weight="800" fill="#18222d">${money(total)}</text></svg>`;
 el.querySelectorAll("circle[data-name]").forEach(c=>c.onclick=()=>{const n=c.dataset.name;setDrill(`${labelPrefix}: ${n}`,t=>t.type==="Expense"&&t.category===n&&monthTransactions().includes(t));show("analytics")});
 if(leg)leg.innerHTML=entries.map(([n,v],i)=>`<button class="legendchip" data-name="${esc(n)}"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${purple[i%purple.length]};margin-right:5px"></span>${esc(n)} ${money(v)}</button>`).join("");
 if(leg)leg.querySelectorAll("button").forEach(b=>b.onclick=()=>{const n=b.dataset.name;setDrill(`${labelPrefix}: ${n}`,t=>t.type==="Expense"&&t.category===n&&monthTransactions().includes(t));show("analytics")});
}
function bars(elId,entries,kind){
 const el=$(elId);if(!el)return;if(!entries.length){el.innerHTML='<div class="notice">No data for this period.</div>';return}
 const max=Math.max(...entries.map(x=>x[1]),1),w=560,rowH=28,left=165,right=75,h=entries.length*rowH+20;
 el.innerHTML=`<svg viewBox="0 0 ${w} ${h}">${entries.map(([name,val],i)=>{const bw=(w-left-right)*(val/max);return `<text x="${left-8}" y="${i*rowH+18}" text-anchor="end" font-size="11" fill="#4b5563">${esc(name.length>24?name.slice(0,22)+"…":name)}</text><rect data-name="${esc(name)}" x="${left}" y="${i*rowH+6}" width="${Math.max(2,bw)}" height="15" rx="5" fill="${purple[i%purple.length]}" style="cursor:pointer"/><text x="${left+bw+7}" y="${i*rowH+18}" font-size="10" fill="#6b7280">${money(val)}</text>`}).join("")}</svg>`;
 el.querySelectorAll("rect[data-name]").forEach(r=>r.onclick=()=>{const n=r.dataset.name;if(kind==="merchant")setDrill(`Merchant: ${n}`,t=>t.type==="Expense"&&t.description===n&&monthTransactions().includes(t));else setDrill(`Source: ${n}`,t=>t.type==="Expense"&&t.source===n&&monthTransactions().includes(t))});
}
function monthlyLine(elId){
 const el=$(elId);if(!el)return;const buckets={};tx.filter(t=>t.type==="Expense"||t.type==="Income").forEach(t=>{const k=String(t.date||"").slice(0,7);if(!/^\d{4}-\d{2}$/.test(k))return;buckets[k]??={income:0,spend:0};if(t.type==="Income")buckets[k].income+=Math.abs(t.amount);else buckets[k].spend+=Math.abs(t.amount)});
 const rows=Object.entries(buckets).sort((a,b)=>a[0].localeCompare(b[0]));if(!rows.length){el.innerHTML='<div class="notice">Import multiple months to see the trend.</div>';return}
 const W=600,H=270,L=48,R=18,T=20,B=46,max=Math.max(1,...rows.flatMap(x=>[x[1].income,x[1].spend]));
 const x=i=>L+(rows.length===1?(W-L-R)/2:i*(W-L-R)/(rows.length-1)), y=v=>T+(H-T-B)*(1-v/max);
 const line=key=>rows.map((r,i)=>`${i?"L":"M"} ${x(i)} ${y(r[1][key])}`).join(" ");
 el.innerHTML=`<svg viewBox="0 0 ${W} ${H}"><line x1="${L}" y1="${H-B}" x2="${W-R}" y2="${H-B}" stroke="#d1d5db"/><path d="${line("income")}" fill="none" stroke="#573168" stroke-width="3"/><path d="${line("spend")}" fill="none" stroke="#b18a52" stroke-width="3"/>${rows.map((r,i)=>{const[yv,mv]=r[0].split("-").map(Number),lab=new Date(yv,mv-1,1).toLocaleString(undefined,{month:"short"});return `<g data-key="${r[0]}" style="cursor:pointer"><circle cx="${x(i)}" cy="${y(r[1].income)}" r="5" fill="#573168"/><circle cx="${x(i)}" cy="${y(r[1].spend)}" r="5" fill="#b18a52"/><text x="${x(i)}" y="${H-18}" text-anchor="middle" font-size="10" fill="#6b7280">${lab}</text></g>`}).join("")}<text x="${L}" y="12" font-size="10" fill="#573168">Income</text><text x="${L+48}" y="12" font-size="10" fill="#b18a52">Spending</text></svg>`;
 el.querySelectorAll("g[data-key]").forEach(g=>g.onclick=()=>{const k=g.dataset.key;const[yv,mv]=k.split("-").map(Number);setDrill(`Month: ${new Date(yv,mv-1,1).toLocaleString(undefined,{month:"long",year:"numeric"})}`,t=>String(t.date||"").startsWith(k)&&(t.type==="Expense"||t.type==="Income"))});
}
function renderCharts(mt){
 const exp=mt.filter(t=>t.type==="Expense"),cats=Object.entries(sumBy(exp,"category")).sort((a,b)=>b[1]-a[1]);
 donut("overviewCategoryChart","overviewLegend",cats,"Category");donut("categoryChart","categoryLegend",cats,"Category");monthlyLine("monthlyChart");
 const merch=Object.entries(sumBy(exp,"description")).sort((a,b)=>b[1]-a[1]).slice(0,9);bars("merchantChart",merch,"merchant");
 const src=Object.entries(sumBy(exp,"source")).sort((a,b)=>b[1]-a[1]).slice(0,9);bars("sourceChart",src,"source");
}
function setDrill(label,pred){drillLabel=label;drillPredicate=pred;renderDrill()}
function renderDrill(){
 const rows=tx.filter(drillPredicate||(()=>false)).sort((a,b)=>b.date.localeCompare(a.date));
 $("drillTitle").textContent=drillLabel||"Drill-down";$("drillSub").textContent=rows.length?"Underlying transactions for this selection.":"No matching transactions.";
 const total=rows.reduce((a,t)=>a+Math.abs(t.amount),0);$("drillCount").textContent=rows.length;$("drillTotal").textContent=money(total);$("drillAvg").textContent=money(rows.length?total/rows.length:0);
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
 const clientId=localStorage.getItem("finance.clientId")||"";
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

$("search").addEventListener("input",render);$("filter").addEventListener("change",render);$("categoryFilter").addEventListener("change",render);$("monthSelect").addEventListener("change",render);
$("clearDrill").onclick=()=>{drillPredicate=null;drillLabel="";$("drillTitle").textContent="Drill-down";$("drillSub").textContent="Tap a chart, KPI, merchant, source, or category.";$("drillCount").textContent="0";$("drillTotal").textContent="$0.00";$("drillAvg").textContent="$0.00";$("drillBody").innerHTML=""};
$("incomeCard").onclick=()=>{const k=selectedMonthKey();setDrill("Income",t=>t.type==="Income"&&(!k||String(t.date||"").startsWith(k)));show("analytics")};
$("spendCard").onclick=()=>{const k=selectedMonthKey();setDrill("Spending",t=>t.type==="Expense"&&(!k||String(t.date||"").startsWith(k)));show("analytics")};
$("netCard").onclick=()=>{const k=selectedMonthKey();setDrill("Cash flow",t=>(t.type==="Income"||t.type==="Expense")&&(!k||String(t.date||"").startsWith(k)));show("analytics")};
$("signBtn").onclick=sign;$("syncBtn").onclick=async()=>{await listDrive();await syncAll()};$("browseBtn").onclick=listDrive;$("syncAllBtn").onclick=syncAll;
$("localBtn").onclick=()=>$("localPicker").click();$("localPicker").onchange=async e=>{for(const f of e.target.files)await processFile(f.name,f)};
$("saveConfig").onclick=()=>{const id=$("clientId").value.trim();localStorage.setItem("finance.clientId",id);alert("Saved. Reloading the app so Microsoft authentication can initialize.");location.reload()};
$("clearData").onclick=()=>{if(confirm("Clear locally cached imported transactions?")){tx=[];save()}};
buildNav();render();await initMsal();
