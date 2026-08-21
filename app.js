import { PublicClientApplication } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@5/+esm";

const FOLDER_PATH = "Finance Dashboard/Statements";
const SCOPES = ["Files.Read"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const nav = [
 ["overview","Overview","Monthly financial snapshot"],
 ["transactions","Transactions","Normalized transactions"],
 ["shared","Shared Household","Shared living expenses and settlement"],
 ["review","Review","Items requiring confirmation"],
 ["onedrive","OneDrive","Statement files and synchronization"],
 ["settings","Settings","Microsoft connection and local data"]
];

let msal = null, account = null;
let tx = JSON.parse(localStorage.getItem("finance.tx.v2") || "[]");

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
 if(/TRANSFER (FROM|TO)|CHASE CREDIT|CAPITAL ONE|CREDIT CARD|VENMO|ZELLE/.test(d)) return ["Transfer","Transfer",false];
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
function render(){
 const expenses=tx.filter(t=>t.type==="Expense"), income=tx.filter(t=>t.type==="Income");
 const spend=expenses.reduce((a,t)=>a+Math.abs(t.amount),0), inc=income.reduce((a,t)=>a+Math.abs(t.amount),0), net=inc-spend;
 $("incomeKpi").textContent=money(inc);$("spendKpi").textContent=money(spend);$("netKpi").textContent=money(net);$("netKpi").className="big "+(net>=0?"good":"bad");$("rateKpi").textContent=(inc?net/inc*100:0).toFixed(1)+"%";
 const cats={};expenses.forEach(t=>cats[t.category]=(cats[t.category]||0)+Math.abs(t.amount));renderBars(cats);
 const q=($("search")?.value||"").toLowerCase(), f=$("filter")?.value||"";
 const rows=tx.filter(t=>(!q||(t.description+" "+t.category+" "+t.source).toLowerCase().includes(q))&&(!f||t.type===f)).sort((a,b)=>b.date.localeCompare(a.date));
 $("txBody").innerHTML=rows.map(t=>`<tr><td>${esc(t.date)}</td><td>${esc(t.description)}</td><td>${esc(t.category)}</td><td><span class="badge ${t.type.toLowerCase()}">${t.type}</span></td><td>${t.shared?'<span class="badge shared">Shared</span>':""}</td><td>${esc(t.source)}</td><td class="amount ${t.amount>0?"good":""}">${money(t.amount)}</td></tr>`).join("");
 renderShared(expenses);renderReview();
}
function renderBars(cats){
 const arr=Object.entries(cats).sort((a,b)=>b[1]-a[1]);const max=Math.max(1,...arr.map(x=>x[1]));
 $("categoryBars").innerHTML=arr.length?arr.map(([n,v],i)=>`<div class="barrow"><span>${esc(n)}</span><div class="track"><div class="bar ${i===1?"gold":""}" style="width:${v/max*100}%"></div></div><b>${money(v)}</b></div>`).join(""):`<div class="notice">Import statements to populate this view.</div>`;
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
 $("signBtn").textContent=account?"Disconnect":"Connect Microsoft";$("syncBtn").disabled=!account;$("browseBtn").disabled=!account;
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
  const files=data.value.filter(x=>x.file);renderDriveFiles(files);$("recentFiles").innerHTML=files.slice(0,5).map(fileCard).join("");
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
  const blob=await res.blob();await processFile(name,blob);
  button.textContent="Imported";
 }catch(e){button.textContent="Failed";alert(`Import failed: ${e.message}`)}
}
async function processFile(name,blob){
 const lower=name.toLowerCase();
 if(lower.endsWith(".csv")){
  const text=await blob.text();const rows=importBECU(text,name);alert(`Imported ${rows} new transaction rows from ${name}.`);
 } else if(lower.endsWith(".pdf")){
  addReviewPlaceholder(name,"PDF statement recognized. Chase/PayPal PDF transaction extraction is the next parser step.");
 } else if(lower.endsWith(".zip")){
  addReviewPlaceholder(name,"ZIP statement archive recognized. PayPal ZIP extraction is the next parser step.");
 } else addReviewPlaceholder(name,"Unsupported statement format.");
}
function addReviewPlaceholder(name,note){
 const key="placeholder|"+name;if(!tx.some(t=>t.key===key)){tx.push({key,date:new Date().toISOString().slice(0,10),description:name,category:"Needs Review",type:"Transfer",shared:false,amount:0,source:"Importer",confidence:"review",note});save()}
 alert(note);
}
$("search").addEventListener("input",render);$("filter").addEventListener("change",render);
$("signBtn").onclick=sign;$("syncBtn").onclick=listDrive;$("browseBtn").onclick=listDrive;
$("localBtn").onclick=()=>$("localPicker").click();$("localPicker").onchange=async e=>{for(const f of e.target.files)await processFile(f.name,f)};
$("saveConfig").onclick=()=>{const id=$("clientId").value.trim();localStorage.setItem("finance.clientId",id);alert("Saved. Reloading the app so Microsoft authentication can initialize.");location.reload()};
$("clearData").onclick=()=>{if(confirm("Clear locally cached imported transactions?")){tx=[];save()}};
buildNav();render();await initMsal();
