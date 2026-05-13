const SUPABASE_URL = 'https://qzipmakglwwtmecrtpyq.supabase.co';

const SUPABASE_KEY = 'sb_publishable_DtZYH_Ntg3JCZ-qus7Hk6A_pE1kauB_';

const supabase = window.supabase.createClient(
  SUPABASE_URL,
  SUPABASE_KEY
);

// ══ STATE ══
let tags=[];let labels=[];let activeFilter='all';let selectedColor=0;let editingId=null;
let sheetsWebAppUrl='';let sortField='date';let sortDir='desc';let deferredInstall=null;
let convPanelOpen=false;let addAvatarData='';let editAvatarData='';
let addEstado='';let editEstado='';let viewAllDays=false;let entrevistas=[];
let currentDocTagId=null;let currentDocFiles=[];let currentFichaTag=null;

const ESTADOS_UF=['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'];
const COLORS=[
  {hex:'#16a34a',bg:'rgba(22,163,74,.14)'},
  {hex:'#2563eb',bg:'rgba(37,99,235,.14)'},
  {hex:'#d97706',bg:'rgba(217,119,6,.14)'},
  {hex:'#dc2626',bg:'rgba(220,38,38,.14)'},
  {hex:'#7c3aed',bg:'rgba(124,58,237,.14)'},
  {hex:'#db2777',bg:'rgba(219,39,119,.14)'},
  {hex:'#ea580c',bg:'rgba(234,88,12,.14)'},
  {hex:'#0891b2',bg:'rgba(8,145,178,.14)'},
];

const APPS_SCRIPT=`// ── Contatos WA · Apps Script · Aba Banco ──
const SHEET_NAME = 'Banco';
const HEADERS = [
  'ID','Nome','Telefone','Data','Etiqueta','Cor da Etiqueta',
  'Estado','CPF','Placa','Modelo/Cor','Valor Aprovado (R$)',
  'Chave Pix','Ref. Nome','Ref. Telefone','Qtd. Documentos',
  'Criado Em','Atualizado Em'
];
function getOrCreateSheet(ss){let sheet=ss.getSheetByName(SHEET_NAME);if(!sheet){sheet=ss.insertSheet(SHEET_NAME);const hdr=sheet.getRange(1,1,1,HEADERS.length);hdr.setValues([HEADERS]);hdr.setBackground('#16a34a');hdr.setFontColor('#ffffff');hdr.setFontWeight('bold');sheet.setFrozenRows(1);sheet.setColumnWidths(1,HEADERS.length,150);}return sheet;}
function getLabelMap(labels){const map={};(labels||[]).forEach(l=>{map[l.id]=l;});return map;}
function tagToRow(t,labelMap){const lbl=t.labelId?(labelMap[t.labelId]||{}):{};const now=new Date().toLocaleString('pt-BR');return[String(t.id||''),t.name||'',t.phone||'',t.date||'',lbl.name||'',lbl.colorHex||'',t.estado||'',t.cpf||'',t.placa||'',t.modelo||'',t.valor||'',t.pix||'',t.refNome||'',t.refTel||'',String((t.docFiles||[]).length),now,now];}
function doGet(e){try{const action=(e.parameter.action||'').toLowerCase();const ss=SpreadsheetApp.getActiveSpreadsheet();const sheet=getOrCreateSheet(ss);if(action==='read'){const data=sheet.getDataRange().getValues();if(data.length<=1)return respond({status:'ok',tags:[],labels:[]});const rows=data.slice(1);let filtered=rows.filter(r=>r[0]!=='');const tags=filtered.map(r=>({id:Number(r[0])||0,name:r[1],phone:r[2],date:r[3],labelName:r[4],labelColorHex:r[5],estado:r[6],cpf:r[7],placa:r[8],modelo:r[9],valor:r[10],pix:r[11],refNome:r[12],refTel:r[13]}));let savedLabels=[];try{const lblSheet=ss.getSheetByName('__labels__');if(lblSheet){const raw=lblSheet.getRange('A1').getValue();if(raw)savedLabels=JSON.parse(raw);}}catch(e){}return respond({status:'ok',tags,labels:savedLabels});}return respond({status:'err',message:'Acao invalida.'});}catch(err){return respond({status:'err',message:err.toString()});}}
function doPost(e){try{const payload=JSON.parse(e.postData.contents);const action=(payload.action||'').toLowerCase();const ss=SpreadsheetApp.getActiveSpreadsheet();if(action==='write'){const sheet=getOrCreateSheet(ss);const incomingTags=payload.tags||[];const labelMap=getLabelMap(payload.labels||[]);const now=new Date().toLocaleString('pt-BR');const existingData=sheet.getDataRange().getValues();const existingIds={};existingData.slice(1).forEach((r,i)=>{if(r[0])existingIds[String(r[0])]=i+2;});const toInsert=[];incomingTags.forEach(t=>{const row=tagToRow(t,labelMap);const sid=String(t.id||'');if(existingIds[sid]){row[16]=now;sheet.getRange(existingIds[sid],1,1,row.length).setValues([row]);}else{toInsert.push(row);}});if(toInsert.length>0){sheet.getRange(sheet.getLastRow()+1,1,toInsert.length,HEADERS.length).setValues(toInsert);}if(payload.labels&&payload.labels.length){let lblSheet=ss.getSheetByName('__labels__');if(!lblSheet)lblSheet=ss.insertSheet('__labels__');lblSheet.getRange('A1').setValue(JSON.stringify(payload.labels));}return respond({status:'ok',saved:incomingTags.length});}return respond({status:'err',message:'Acao desconhecida.'});}catch(err){return respond({status:'err',message:err.toString()});}}
function respond(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}`;

// boot
(function init(){
  const today=new Date().toISOString().split('T')[0];
  document.getElementById('activeDate').value=today;
  document.getElementById('inputDate').value=today;
  document.getElementById('activeDate').addEventListener('change',()=>{updateDateBadge();loadTagsFromStorage();setFilter('all',document.querySelector('[data-filter="all"]'));});
  ['inputPhone','editPhone'].forEach(id=>{const el=document.getElementById(id);if(el)el.addEventListener('input',function(){let v=this.value.replace(/\D/g,'').slice(0,11);if(v.length>10)v=v.replace(/^(\d{2})(\d{5})(\d{4})$/,'($1) $2-$3');else if(v.length>6)v=v.replace(/^(\d{2})(\d{4})(\d*)$/,'($1) $2-$3');else if(v.length>2)v=v.replace(/^(\d{2})(\d*)$/,'($1) $2');this.value=v;});});
  document.getElementById('inputPhone').addEventListener('input',checkDuplicate);
  ['inputName','inputPhone','inputDate'].forEach(id=>{const el=document.getElementById(id);if(el)el.addEventListener('keydown',e=>{if(e.key==='Enter')addTag();});});
  const li=document.getElementById('inputLabelName');if(li)li.addEventListener('keydown',e=>{if(e.key==='Enter')createLabel();});
  window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();deferredInstall=e;if(!localStorage.getItem('pwaDismissed'))document.getElementById('pwaBanner').classList.add('show');document.getElementById('btnPwa').style.display='flex';});
  window.addEventListener('appinstalled',()=>{document.getElementById('pwaBanner').classList.remove('show');document.getElementById('btnPwa').style.display='none';showToast('✅ App instalado!');});
  buildColorPicker();buildEstadoGrid('estadoGrid','add');buildEstadoGrid('estadoGridEdit','edit');
  loadTheme();loadSheetsCredentials();loadLabelsFromStorage();loadMeta();updateDateBadge();loadTagsFromStorage();injectPWAManifest();
  ['editModal','sheetsModal','fichaModal'].forEach(id=>{const el=document.getElementById(id);if(el)el.addEventListener('click',e=>{if(e.target===e.currentTarget)el.classList.remove('open');});});
  document.addEventListener('keydown',e=>{if(e.key==='Escape'){closeEditModal();closeSheetsModal();closeFichaModal();}});
  buildEstadoFilter();loadEntrevistas();
  const einput=document.getElementById('entrevistaDateInput');if(einput)einput.value=today;
  const ta=document.getElementById('scriptDisplay');if(ta)ta.value=APPS_SCRIPT;
})();

//navegação
function goPage(pageId,navEl){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  document.getElementById('page-'+pageId).classList.add('active');
  if(navEl)navEl.classList.add('active');
  const titles={contacts:'Contatos',add:'Novo Contato',labels:'Etiquetas',agenda:'Agenda'};
  document.getElementById('topbarTitle').textContent=titles[pageId]||'';
  if(window.innerWidth<=768)toggleSidebar();
}
function toggleSidebar(){
  const sb=document.getElementById('sidebar');
  const ov=document.getElementById('sidebarOverlay');
  sb.classList.toggle('open');
  ov.style.display=sb.classList.contains('open')?'block':'none';
}

// pwa 
function injectPWAManifest(){const manifest={name:'Contatos WA',short_name:'Contatos WA',description:'Gestão de leads WhatsApp',start_url:window.location.href,display:'standalone',background_color:'#fafafa',theme_color:'#16a34a',icons:[{src:"data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect width='100' height='100' rx='20' fill='%2316a34a'/><text y='.9em' font-size='75' x='12'>💬</text></svg>",sizes:'192x192',type:'image/svg+xml'}]};const blob=new Blob([JSON.stringify(manifest)],{type:'application/json'});document.getElementById('pwaManifest').href=URL.createObjectURL(blob);}
async function installPwa(){if(!deferredInstall){showToast('⚠️ Instalação não disponível');return;}deferredInstall.prompt();const{outcome}=await deferredInstall.userChoice;if(outcome==='accepted')deferredInstall=null;}
function dismissPwa(){document.getElementById('pwaBanner').classList.remove('show');localStorage.setItem('pwaDismissed','1');}

// modo noturno/diurno
function loadTheme(){applyTheme(localStorage.getItem('waTheme')||'light');}
function toggleTheme(){applyTheme(document.documentElement.getAttribute('data-theme')==='dark'?'light':'dark');}
function applyTheme(t){document.documentElement.setAttribute('data-theme',t);localStorage.setItem('waTheme',t);const b=document.getElementById('btnTheme');if(b){const icon=b.querySelector('svg');if(icon)icon.style.display=t==='dark'?'none':'block';b.childNodes[b.childNodes.length-1].textContent=t==='dark'?' Tema':' Tema';}}

function goToToday(){const t=new Date().toISOString().split('T')[0];document.getElementById('activeDate').value=t;updateDateBadge();loadTagsFromStorage();setFilter('all',document.querySelector('[data-filter="all"]'));showToast('📍 Voltou para hoje!');}

// meta diraria 
function loadMeta(){const s=localStorage.getItem('waMeta');if(s)document.getElementById('metaInput').value=s;updateMeta();}
function saveMeta(){const v=document.getElementById('metaInput').value;if(v)localStorage.setItem('waMeta',v);updateMeta();}
function updateMeta(){
  const metaVal=parseInt(document.getElementById('metaInput').value)||0;
  const current=tags.length;
  const bar=document.getElementById('metaBar');
  const pct=document.getElementById('metaPct');
  const left=document.getElementById('metaLeft');
  const emoji=document.getElementById('metaEmoji');
  if(!metaVal){bar.style.width='0%';pct.textContent='—';left.textContent='Defina uma meta';emoji.textContent='';return;}
  const perc=Math.min(100,Math.round((current/metaVal)*100));
  bar.style.width=perc+'%';
  pct.textContent=perc+'%';
  if(perc>=100){bar.classList.add('done');left.textContent=`🎉 Meta batida! (${current}/${metaVal})`;left.className='ok';emoji.textContent='🏆';}
  else{bar.classList.remove('done');const rem=metaVal-current;left.textContent=`${current}/${metaVal} — faltam ${rem}`;left.className='';emoji.textContent=perc>=75?'🔥':perc>=50?'💪':perc>=25?'📈':'';}
}

function toggleConvPanel(){convPanelOpen=!convPanelOpen;document.getElementById('convPanel').classList.toggle('open',convPanelOpen);if(convPanelOpen)renderConvPanel();}
function renderConvPanel(){
  const panel=document.getElementById('convGrid');
  document.getElementById('convDate').textContent=fmtDate(getActiveDate());
  if(!labels.length||!tags.length){panel.innerHTML='<span style="font-size:12px;color:var(--text-3)">Adicione etiquetas e contatos.</span>';return;}
  const counts={};labels.forEach(l=>{counts[l.id]=0;});
  tags.forEach(t=>{if(t.labelId&&counts[t.labelId]!==undefined)counts[t.labelId]++;});
  const total=tags.length;let html='';
  labels.forEach((l,i)=>{
    const n=counts[l.id]||0;const pct=total?Math.round((n/total)*100):0;
    if(i>0)html+='<span class="conv-arrow-txt">→</span>';
    html+=`<div class="conv-box"><div class="conv-box-lbl" style="color:${l.colorHex}">${esc(l.name)}</div><div class="conv-box-val" style="color:${l.colorHex}">${n}</div><div style="font-size:10px;color:var(--text-3);margin-top:2px">${pct}%</div></div>`;
  });
  if(labels.length>=2){const first=counts[labels[0].id]||0;const last=counts[labels[labels.length-1].id]||0;const rate=first?Math.round((last/first)*100):0;html+=`<span class="conv-arrow-txt">=</span><div class="conv-rate">${rate}%<div style="font-size:10px;font-weight:400;margin-top:2px">${labels[0].name}→${labels[labels.length-1].name}</div></div>`;document.getElementById('scConv').textContent=rate+'%';}else{document.getElementById('scConv').textContent='—';}
  panel.innerHTML=html;
}

// corzinha
function buildColorPicker(){const c=document.getElementById('colorPicker');if(!c)return;c.innerHTML=COLORS.map((col,i)=>`<div class="color-dot ${i===0?'selected':''}" style="background:${col.hex}" onclick="selectColor(${i},this)"></div>`).join('');}
function selectColor(i,el){selectedColor=i;document.querySelectorAll('.color-dot').forEach(d=>d.classList.remove('selected'));el.classList.add('selected');}

// fotinha
function onAvatarChange(event,ctx){const file=event.target.files[0];if(!file)return;const reader=new FileReader();reader.onload=e=>{const data=e.target.result;if(ctx==='add'){addAvatarData=data;document.getElementById('avatarPreview').innerHTML=`<img src="${data}" style="width:100%;height:100%;object-fit:cover;border-radius:50%">`;}else{editAvatarData=data;document.getElementById('editAvatarPreview').innerHTML=`<img src="${data}" style="width:100%;height:100%;object-fit:cover;border-radius:50%">`;}};reader.readAsDataURL(file);}
function onNameInput(){const name=document.getElementById('inputName').value;const btn=document.getElementById('avatarPreview');if(!addAvatarData&&name){btn.textContent=initials(name);}else if(!addAvatarData){btn.textContent='📷';}}
function initials(name){if(!name)return'?';const p=name.trim().split(' ');return p.length>=2?(p[0][0]+p[p.length-1][0]).toUpperCase():name[0].toUpperCase();}

// ══ ESTADO ══
function buildEstadoGrid(gridId,ctx){const grid=document.getElementById(gridId);if(!grid)return;grid.innerHTML=ESTADOS_UF.map(uf=>`<div class="estado-opt" data-uf="${uf}" onclick="selEstado('${uf}','${ctx}',this)">${uf}</div>`).join('');}
function selEstado(uf,ctx,el){const gridId=ctx==='add'?'estadoGrid':'estadoGridEdit';document.querySelectorAll(`#${gridId} .estado-opt`).forEach(x=>x.classList.remove('sel'));if(ctx==='add'){if(addEstado===uf){addEstado='';return;}addEstado=uf;}else{if(editEstado===uf){editEstado='';return;}editEstado=uf;}el.classList.add('sel');}

function onLabelChange(){const val=document.getElementById('inputLabel').value;const lbl=val?getLabelById(val):null;document.getElementById('fichaFinSection').classList.toggle('visible',!!(lbl&&lbl.financeiro));}
function onEditLabelChange(){const val=document.getElementById('editLabel').value;const lbl=val?getLabelById(val):null;document.getElementById('fichaFinSectionEdit').classList.toggle('visible',!!(lbl&&lbl.financeiro));}

function tryParseText(){
  const txt=document.getElementById('parserInput').value;
  const fb=document.getElementById('parserFeedback');
  if(!txt.trim()){fb.className='parser-fb';return;}
  const get=(label)=>{const re=new RegExp('[*_]*'+label+'[*_]*\\s*:?\\s*[*_]*([^\\n*]+)','i');const m=txt.match(re);return m?m[1].replace(/\*/g,'').trim():'';};
  const nome=get('Nome completo'),cpf=get('CPF'),telRaw=get('Telefone'),refRaw=get('Nome e Telefone de refer[eê]ncia'),placa=get('Placa do carro'),modelo=get('Modelo e cor'),valor=get('Valor aprovado'),pix=get('Chave Pix');
  const tel=(telRaw.match(/[\d\s\(\)\-\+]{8,}/)||[''])[0].replace(/\D/g,'').slice(0,11);
  const telFmt=formatPhone(tel);
  const refTelM=refRaw.match(/[\+\d][\d\s\-\(\)]{7,}/);
  const refTel=refTelM?refTelM[0].replace(/\D/g,''):'';
  const refNome=refRaw.replace(/[\(\[]?[\+\d][\d\s\-\(\)]{7,}[\)\]]?/g,'').replace(/[()]/g,'').trim().split('/')[0].trim();
  let filled=0;
  if(nome){document.getElementById('inputName').value=nome;onNameInput();filled++;}
  if(telFmt){document.getElementById('inputPhone').value=telFmt;filled++;}
  if(cpf){document.getElementById('inputCpf').value=cpf;filled++;}
  if(refNome){document.getElementById('inputRefNome').value=refNome;filled++;}
  if(refTel){document.getElementById('inputRefTel').value=refTel;filled++;}
  if(placa){document.getElementById('inputPlaca').value=placa.toUpperCase();filled++;}
  if(modelo){document.getElementById('inputModelo').value=modelo;filled++;}
  if(valor){document.getElementById('inputValor').value=valor;filled++;}
  if(pix){document.getElementById('inputPix').value=pix;filled++;}
  if(filled>0){
    fb.className='parser-fb ok';fb.textContent=`✅ ${filled} campo${filled!==1?'s':''} preenchido${filled!==1?'s':''}!`;
    const finLbl=labels.find(l=>l.financeiro);
    if(finLbl){document.getElementById('inputLabel').value=finLbl.id;onLabelChange();}
    else{document.getElementById('fichaFinSection').classList.add('visible');}
    showToast('✨ Ficha preenchida!');
  }else{fb.className='parser-fb warn';fb.textContent='⚠️ Nenhum campo reconhecido.';}
}
function formatPhone(digits){if(!digits)return'';const v=digits.slice(0,11);if(v.length>10)return v.replace(/^(\d{2})(\d{5})(\d{4})$/,'($1) $2-$3');if(v.length>6)return v.replace(/^(\d{2})(\d{4})(\d*)$/,'($1) $2-$3');if(v.length>2)return v.replace(/^(\d{2})(\d*)$/,'($1) $2');return v;}

function findTagAnywhere(id){
  const numId=Number(id),strId=String(id);
  let t=tags.find(x=>x.id===numId||x.id===strId||String(x.id)===strId);
  if(t)return{tag:t,key:tagsKey(),arr:tags,inCurrent:true};
  const allKeys=Object.keys(localStorage).filter(k=>k.startsWith('waTags_')&&k!==tagsKey());
  for(const key of allKeys){try{const arr=JSON.parse(localStorage.getItem(key)||'[]');const found=arr.find(x=>x.id===numId||x.id===strId||String(x.id)===strId);if(found)return{tag:found,key,arr,inCurrent:false};}catch(e){}}
  return null;
}
function saveTagAnywhere(r){if(r.inCurrent)saveTags();else localStorage.setItem(r.key,JSON.stringify(r.arr));}

function toggleDocAnalisada(id,cb){const r=findTagAnywhere(id);if(!r)return;r.tag.docAnalisada=cb.checked;saveTagAnywhere(r);if(cb.checked)openFichaModal(id);}
function openFichaModal(id){
  const r=findTagAnywhere(id);if(!r)return;const t=r.tag;
  currentFichaTag=t;currentDocFiles=t.docFiles||[];
  const lbl=t.labelId?getLabelById(t.labelId):null;
  const avatarHTML=t.avatar?`<img src="${t.avatar}" style="width:48px;height:48px;border-radius:50%;object-fit:cover;border:2px solid var(--border)">`:`<div style="width:48px;height:48px;border-radius:50%;background:${lbl?lbl.colorBg:'var(--bg2)'};border:2px solid ${lbl?lbl.colorHex+'55':'var(--border)'};display:grid;place-items:center;font-weight:700;font-size:17px;color:${lbl?lbl.colorHex:'var(--text-3)'}">${initials(t.name)}</div>`;
  document.getElementById('fichaModalTitle').innerHTML=`<div style="display:flex;align-items:center;gap:12px">${avatarHTML}<div><div style="font-size:16px;font-weight:700">${esc(t.name)}</div>${lbl?`<span class="badge" style="color:${lbl.colorHex};background:${lbl.colorBg};border-color:${lbl.colorHex}44;font-size:11px">${esc(lbl.name)}</span>`:''}${t.estado?`<span class="estado-pill-row" style="margin-left:4px">${t.estado}</span>`:''}</div></div>`;
  let rows=`<div class="fv-row"><span class="fv-key">📱 Telefone</span><span class="fv-val">${esc(t.phone)||'—'}</span></div><div class="fv-row"><span class="fv-key">📅 Data</span><span class="fv-val">${fmtDate(t.date)}</span></div><div class="fv-row"><span class="fv-key">📍 Estado</span><span class="fv-val">${t.estado||'—'}</span></div>`;
  if(lbl&&lbl.financeiro){
    rows+=`<div style="font-size:11px;font-weight:600;color:var(--warning);text-transform:uppercase;letter-spacing:.6px;margin:14px 0 8px">💰 Ficha Financeira</div>`;
    rows+=`<div class="fv-row"><span class="fv-key">CPF</span><span class="fv-val">${esc(t.cpf||'—')}</span></div>`;
    const refWaNum=(t.refTel||'').replace(/\D/g,'');
    const refWaBtn=refWaNum?`<a href="https://wa.me/55${refWaNum}" target="_blank" rel="noopener" style="display:inline-flex;align-items:center;gap:3px;margin-left:8px;padding:2px 8px;border-radius:5px;font-size:10px;background:var(--accent-subtle);border:1px solid rgba(22,163,74,.2);color:var(--accent);text-decoration:none">💬 ${esc(t.refTel)}</a>`:`<span>${esc(t.refTel||'—')}</span>`;
    rows+=`<div class="fv-row"><span class="fv-key">Ref. Nome</span><span class="fv-val">${esc(t.refNome||'—')}</span></div>`;
    rows+=`<div class="fv-row"><span class="fv-key">Ref. Tel</span><span class="fv-val" style="display:flex;align-items:center;justify-content:flex-end;flex-wrap:wrap;gap:4px">${refWaBtn}</span></div>`;
    rows+=`<div class="fv-row"><span class="fv-key">Placa</span><span class="fv-val">${esc(t.placa||'—')}</span></div>`;
    rows+=`<div class="fv-row"><span class="fv-key">Veículo</span><span class="fv-val">${esc(t.modelo||'—')}</span></div>`;
    rows+=`<div class="fv-row"><span class="fv-key" style="color:var(--warning)">💵 Valor</span><span class="fv-val" style="color:var(--accent);font-weight:600">R$ ${esc(t.valor||'—')}</span></div>`;
    rows+=`<div class="fv-row"><span class="fv-key">Pix</span><span class="fv-val">${esc(t.pix||'—')}</span></div>`;
  }
  const waNum=(t.phone||'').replace(/\D/g,'');
  rows+=`<div style="margin-top:12px;display:flex;gap:8px">${waNum?`<a href="https://wa.me/55${waNum}" target="_blank" class="btn-primary" style="text-decoration:none">💬 WhatsApp</a>`:''}<button onclick="closeFichaModal();openEditModal(${t.id})" class="btn-secondary">✏️ Editar</button></div>`;
  document.getElementById('fichaModalContent').innerHTML=rows;
  const hasDocs=currentDocFiles&&currentDocFiles.length>0;
  const docAnalisada=t.docAnalisada||hasDocs;
  document.getElementById('docCheckModal').checked=docAnalisada;
  document.getElementById('docSavedBadge').style.display=hasDocs?'inline-flex':'none';
  document.getElementById('docUploadAreaModal').classList.toggle('open',docAnalisada);
  renderDocFileListModal();
  document.getElementById('fichaModal').classList.add('open');
}
function closeFichaModal(){document.getElementById('fichaModal').classList.remove('open');currentFichaTag=null;currentDocFiles=[];}
function onDocCheckModal(cb){document.getElementById('docUploadAreaModal').classList.toggle('open',cb.checked);if(!cb.checked)currentDocFiles=[];renderDocFileListModal();}
function onDocFilePickModal(event){readAndStoreDocFilesModal(Array.from(event.target.files));}
function onDocDropModal(event){event.preventDefault();document.getElementById('docDropZoneModal').classList.remove('drag-over');readAndStoreDocFilesModal(Array.from(event.dataTransfer.files));}
function readAndStoreDocFilesModal(files){let pending=files.length;files.forEach(file=>{const r=new FileReader();r.onload=e=>{currentDocFiles.push({name:file.name,data:e.target.result,type:file.type});pending--;if(pending===0)renderDocFileListModal();};r.readAsDataURL(file);});}
function renderDocFileListModal(){
  const container=document.getElementById('docFileListModal');if(!container)return;
  if(!currentDocFiles.length){container.innerHTML='';return;}
  container.innerHTML=currentDocFiles.map((f,i)=>`<div class="doc-chip">${f.type&&f.type.startsWith('image/')?'🖼️':'📄'} ${esc(f.name)}<button class="doc-chip-rm" onclick="removeDocFileModal(${i})">✕</button></div>`).join('');
  document.getElementById('docSavedBadge').style.display=currentDocFiles.length?'inline-flex':'none';
}
function removeDocFileModal(idx){currentDocFiles.splice(idx,1);renderDocFileListModal();if(!currentDocFiles.length)document.getElementById('docCheckModal').checked=false;}
async function saveClientFolder(){
  if(!currentFichaTag){showToast('⚠️ Nenhum contato!');return;}
  const folderName=`${currentFichaTag.name.replace(/[^a-zA-Z0-9À-ú ]/g,'_')}_${currentFichaTag.date||''}`;
  const zip=new JSZip();
  zip.file('ficha.txt',buildFichaText(currentFichaTag));
  for(const file of currentDocFiles){zip.file(file.name,file.data.split(',')[1],{base64:true});}
  const content=await zip.generateAsync({type:'blob'});
  const url=URL.createObjectURL(content);
  const a=document.createElement('a');a.href=url;a.download=`${folderName}.zip`;a.click();
  URL.revokeObjectURL(url);
  currentFichaTag.docFiles=currentDocFiles;
  const r=findTagAnywhere(currentFichaTag.id);if(r){r.tag.docFiles=currentDocFiles;saveTagAnywhere(r);}else saveTags();
  showToast(`✅ "${folderName}.zip" baixado!`);
  document.getElementById('docSavedBadge').style.display='inline-flex';
}
function buildFichaText(t){const lbl=t.labelId?getLabelById(t.labelId):null;let txt=`FICHA DO CLIENTE\n${'═'.repeat(40)}\n`;txt+=`Nome:     ${t.name}\nTelefone: ${t.phone}\nData:     ${fmtDate(t.date)}\nEstado:   ${t.estado||'—'}\nEtiqueta: ${lbl?lbl.name:'—'}\n`;if(lbl&&lbl.financeiro){txt+=`\n--- DADOS FINANCEIROS ---\n`;if(t.cpf)txt+=`CPF:       ${t.cpf}\n`;if(t.placa)txt+=`Placa:     ${t.placa}\n`;if(t.modelo)txt+=`Veículo:   ${t.modelo}\n`;if(t.valor)txt+=`Valor:     R$ ${t.valor}\n`;if(t.pix)txt+=`Pix:       ${t.pix}\n`;if(t.refNome)txt+=`Referência:${t.refNome}\n`;if(t.refTel)txt+=`Tel.Ref:   ${t.refTel}\n`;}if(t.docFiles&&t.docFiles.length)txt+=`\n--- DOCUMENTOS: ${t.docFiles.length} arquivo(s) ---\n`;return txt;}

// etiquetas
function saveLabels(){localStorage.setItem('waLabels',JSON.stringify(labels));}
function loadLabelsFromStorage(){const d=localStorage.getItem('waLabels');labels=d?JSON.parse(d):[];renderExistingLabels();renderLabelSelect();renderFilterChips();}
function createLabel(){
  const name=document.getElementById('inputLabelName').value.trim();
  if(!name){showToast('⚠️ Digite o nome!');return;}
  if(labels.find(l=>l.name.toLowerCase()===name.toLowerCase())){showToast('⚠️ Etiqueta já existe!');return;}
  const c=COLORS[selectedColor];
  const financeiro=document.getElementById('inputLabelFin').checked;
  labels.push({id:Date.now(),name,colorHex:c.hex,colorBg:c.bg,financeiro});
  saveLabels();renderExistingLabels();renderLabelSelect();renderFilterChips();
  document.getElementById('inputLabelName').value='';document.getElementById('inputLabelFin').checked=false;
  showToast(`🏷️ "${name}" criada!`);updateSummary();
  document.getElementById('navBadgeAgenda').textContent=entrevistas.length;
}
function deleteLabel(id){
  if(!confirm('Remover essa etiqueta?\nOs contatos vinculados ficarão sem etiqueta.'))return;
  labels=labels.filter(l=>l.id!==id);
  tags.forEach(t=>{if(t.labelId===id)t.labelId=null;});
  saveLabels();saveTags();renderExistingLabels();renderLabelSelect();
  if(activeFilter===String(id))setFilter('all',document.querySelector('[data-filter="all"]'));
  else{renderFilterChips();renderList();}
  updateSummary();
}
function renderExistingLabels(){
  const el=document.getElementById('existingLabels');if(!el)return;
  if(!labels.length){el.innerHTML='<span style="font-size:12px;color:var(--text-3)">Nenhuma etiqueta criada ainda.</span>';return;}
  el.innerHTML=labels.map(l=>`<span class="label-chip" style="color:${l.colorHex};background:${l.colorBg};border-color:${l.colorHex}44">${esc(l.name)}${l.financeiro?'<span class="fin-label-badge" style="font-size:9px;padding:1px 5px">💰</span>':''}<button class="label-chip-del" onclick="deleteLabel(${l.id})" title="Remover">✕</button></span>`).join('');
}
function renderLabelSelect(){
  ['inputLabel','editLabel'].forEach(id=>{const s=document.getElementById(id);if(!s)return;const cur=s.value;s.innerHTML='<option value="">— Sem etiqueta —</option>'+labels.map(l=>`<option value="${l.id}">${esc(l.name)}</option>`).join('');if(cur)s.value=cur;});
}
function renderFilterChips(){
  const bar=document.getElementById('filterBar');if(!bar)return;
  const extras=bar.querySelectorAll('.filter-pill:not([data-filter="all"]):not([data-filter="none"])');
  extras.forEach(e=>e.remove());
  labels.forEach(l=>{
    const btn=document.createElement('button');
    btn.className='filter-pill'+(activeFilter===String(l.id)?' active':'');
    btn.dataset.filter=l.id;
    btn.onclick=function(){setFilter(String(l.id),this);};
    btn.textContent=l.name;
    if(activeFilter===String(l.id))btn.style.cssText=`background:${l.colorHex};color:#fff;border-color:transparent`;
    bar.appendChild(btn);
  });
}
function getLabelById(id){return labels.find(l=>l.id===Number(id)||l.id===id);}
function getLabelByName(name){return labels.find(l=>l.name.toLowerCase()===name.toLowerCase());}

function setFilter(f,el){
  activeFilter=f;
  document.querySelectorAll('.filter-pill').forEach(c=>{c.classList.remove('active');c.style.background='';c.style.color='';c.style.borderColor='';});
  el.classList.add('active');
  if(f==='all'){el.style.cssText='background:var(--accent);color:#fff;border-color:transparent';}
  else if(f!=='none'){const lbl=getLabelById(f);if(lbl)el.style.cssText=`background:${lbl.colorHex};color:#fff;border-color:transparent`;}
  renderList();
}
function setSort(field,btn){
  if(sortField===field)sortDir=sortDir==='asc'?'desc':'asc';else{sortField=field;sortDir='desc';}
  document.querySelectorAll('.sort-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const dirs={date:'sortDirDate',name:'sortDirName',label:'sortDirLabel'};
  Object.values(dirs).forEach(id=>{const el=document.getElementById(id);if(el)el.textContent='↑';});
  const d=document.getElementById(dirs[field]);if(d)d.textContent=sortDir==='asc'?'↑':'↓';
  renderList();
}
function sortedTags(list){return[...list].sort((a,b)=>{let va,vb;if(sortField==='name'){va=a.name.toLowerCase();vb=b.name.toLowerCase();}else if(sortField==='date'){va=a.id;vb=b.id;}else{const la=getLabelById(a.labelId);const lb=getLabelById(b.labelId);va=(la?la.name:'zzz').toLowerCase();vb=(lb?lb.name:'zzz').toLowerCase();}if(va<vb)return sortDir==='asc'?-1:1;if(va>vb)return sortDir==='asc'?1:-1;return 0;});}
function toggleViewAll(){viewAllDays=!viewAllDays;const btn=document.getElementById('btnViewAll');btn.classList.toggle('active',viewAllDays);btn.querySelector('svg').style.color=viewAllDays?'var(--accent)':'';renderList();}

// check dulicada
function checkDuplicate(){
  const phone=document.getElementById('inputPhone').value.replace(/\D/g,'');
  const alert=document.getElementById('dupAlert');const msg=document.getElementById('dupMsg');
  if(phone.length<8){alert.classList.remove('show');return;}
  const allKeys=Object.keys(localStorage).filter(k=>k.startsWith('waTags_'));
  let found=null;
  for(const key of allKeys){try{const dt=JSON.parse(localStorage.getItem(key)||'[]');const m=dt.find(t=>t.phone.replace(/\D/g,'')===phone);if(m){const date=key.replace('waTags_','');found={name:m.name,date};break;}}catch(e){}}
  if(found){alert.classList.add('show');msg.innerHTML=`<strong>${esc(found.name)}</strong> cadastrado em <strong>${fmtDate(found.date)}</strong>`;}
  else{alert.classList.remove('show');}
}


function getActiveDate(){return document.getElementById('activeDate').value;}
function tagsKey(){return'waTags_'+getActiveDate();}
function saveTags(){localStorage.setItem(tagsKey(),JSON.stringify(tags));}
function loadTagsFromStorage(){const d=localStorage.getItem(tagsKey());tags=d?JSON.parse(d):[];renderList();updateSummary();updateMeta();if(convPanelOpen)renderConvPanel();document.getElementById('navBadgeContacts').textContent=tags.length;}

// adicionar tag
function addTag(){
  const name=document.getElementById('inputName').value.trim();
  const phone=document.getElementById('inputPhone').value.trim();
  const date=document.getElementById('inputDate').value;
  const labelId=document.getElementById('inputLabel').value;
  if(!name||!phone||!date){showToast('⚠️ Preencha nome, telefone e data!');return;}
  const phoneDigits=phone.replace(/\D/g,'');
  const allKeys=Object.keys(localStorage).filter(k=>k.startsWith('waTags_'));
  for(const key of allKeys){try{const dt=JSON.parse(localStorage.getItem(key)||'[]');if(dt.some(t=>t.phone.replace(/\D/g,'')===phoneDigits)){showToast('⚠️ Número já cadastrado!');return;}}catch(e){}}
  const lbl=labelId?getLabelById(labelId):null;
  const tag={id:Date.now(),name,phone,date,labelId:labelId?Number(labelId):null,avatar:addAvatarData||'',estado:addEstado||'',docFiles:[]};
  if(lbl&&lbl.financeiro){tag.cpf=document.getElementById('inputCpf').value.trim();tag.refNome=document.getElementById('inputRefNome').value.trim();tag.refTel=document.getElementById('inputRefTel').value.trim();tag.placa=document.getElementById('inputPlaca').value.trim();tag.modelo=document.getElementById('inputModelo').value.trim();tag.valor=document.getElementById('inputValor').value.trim();tag.pix=document.getElementById('inputPix').value.trim();}
  tags.push(tag);saveTags();renderList();updateSummary();updateMeta();if(convPanelOpen)renderConvPanel();
  ['inputName','inputPhone','inputCpf','inputRefNome','inputRefTel','inputPlaca','inputModelo','inputValor','inputPix'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});
  document.getElementById('inputDate').value=getActiveDate();
  document.getElementById('inputLabel').value='';
  document.getElementById('fichaFinSection').classList.remove('visible');
  document.getElementById('parserInput').value='';
  document.getElementById('parserFeedback').className='parser-fb';
  document.getElementById('dupAlert').classList.remove('show');
  const avatarBtn=document.getElementById('avatarPreview');avatarBtn.innerHTML='📷';
  addAvatarData='';addEstado='';buildEstadoGrid('estadoGrid','add');
  document.getElementById('inputName').focus();
  showToast('✅ Contato adicionado!');
  document.getElementById('navBadgeContacts').textContent=tags.length;
  // Navigate back to contacts
  goPage('contacts',document.getElementById('nav-contacts'));
}
function deleteTag(id){if(!confirm('Remover este contato?'))return;
tags=tags.filter(t=>t.id!==id);saveTags();
const allKeys=Object.keys(localStorage).filter(k=>k.startsWith('waTags_')&&k!==tagsKey());allKeys.forEach(key=>{try{const dt=JSON.parse(localStorage.getItem(key)||'[]');const filtered=dt.filter(t=>t.id!==id);if(filtered.length!==dt.length)localStorage.setItem(key,JSON.stringify(filtered));}catch(e){}});
renderList();updateSummary();updateMeta();if(convPanelOpen)renderConvPanel();document.getElementById('navBadgeContacts').textContent=tags.length;}

// edit
function openEditModal(id){
  const r=findTagAnywhere(id);if(!r)return;const t=r.tag;
  editingId=id;
  document.getElementById('editName').value=t.name;
  document.getElementById('editPhone').value=t.phone;
  document.getElementById('editDate').value=t.date;
  renderLabelSelect();document.getElementById('editLabel').value=t.labelId||'';
  editAvatarData=t.avatar||'';
  const avatarBtn=document.getElementById('editAvatarPreview');
  if(t.avatar)avatarBtn.innerHTML=`<img src="${t.avatar}" style="width:100%;height:100%;object-fit:cover;border-radius:50%">`;else avatarBtn.textContent=initials(t.name);
  editEstado=t.estado||'';buildEstadoGrid('estadoGridEdit','edit');
  if(t.estado)setTimeout(()=>{const opt=document.querySelector(`#estadoGridEdit [data-uf="${t.estado}"]`);if(opt)opt.classList.add('sel');},0);
  const lbl=t.labelId?getLabelById(t.labelId):null;
  const showFin=!!(lbl&&lbl.financeiro);
  document.getElementById('fichaFinSectionEdit').classList.toggle('visible',showFin);
  if(showFin){document.getElementById('editCpf').value=t.cpf||'';document.getElementById('editRefNome').value=t.refNome||'';document.getElementById('editRefTel').value=t.refTel||'';document.getElementById('editPlaca').value=t.placa||'';document.getElementById('editModelo').value=t.modelo||'';document.getElementById('editValor').value=t.valor||'';document.getElementById('editPix').value=t.pix||'';}
  document.getElementById('editModal').classList.add('open');
  setTimeout(()=>document.getElementById('editName').focus(),100);
}
function closeEditModal(){document.getElementById('editModal').classList.remove('open');editingId=null;}
function saveEdit(){
  if(!editingId)return;
  const r=findTagAnywhere(editingId);if(!r)return;const t=r.tag;
  const name=document.getElementById('editName').value.trim();
  const phone=document.getElementById('editPhone').value.trim();
  const date=document.getElementById('editDate').value;
  const lblId=document.getElementById('editLabel').value;
  if(!name||!phone||!date){showToast('⚠️ Preencha todos os campos!');return;}
  t.name=name;t.phone=phone;t.date=date;t.labelId=lblId?Number(lblId):null;t.avatar=editAvatarData||t.avatar||'';t.estado=editEstado||'';
  const lbl=lblId?getLabelById(lblId):null;
  if(lbl&&lbl.financeiro){t.cpf=document.getElementById('editCpf').value.trim();t.refNome=document.getElementById('editRefNome').value.trim();t.refTel=document.getElementById('editRefTel').value.trim();t.placa=document.getElementById('editPlaca').value.trim();t.modelo=document.getElementById('editModelo').value.trim();t.valor=document.getElementById('editValor').value.trim();t.pix=document.getElementById('editPix').value.trim();}
  saveTagAnywhere(r);renderList();updateSummary();closeEditModal();if(convPanelOpen)renderConvPanel();showToast('✅ Contato atualizado!');
}

// sumario
function updateSummary(){
  document.getElementById('scTotal').textContent=tags.length;
  document.getElementById('scWa').textContent=tags.filter(t=>t.phone).length;
  const used=new Set(tags.filter(t=>t.labelId).map(t=>t.labelId));
  document.getElementById('scLbl').textContent=used.size;
  document.getElementById('scLblTotal').textContent=labels.length;
  document.getElementById('scNone').textContent=tags.filter(t=>!t.labelId).length;
  if(labels.length>=2){const first=labels[0],last=labels[labels.length-1];const fn=tags.filter(t=>t.labelId===first.id).length;const ln=tags.filter(t=>t.labelId===last.id).length;const rate=fn?Math.round((ln/fn)*100):0;document.getElementById('scConv').textContent=rate+'%';}
  else{document.getElementById('scConv').textContent='—';}
  document.getElementById('navBadgeContacts').textContent=tags.length;
}

function searchQuery(){return document.getElementById('searchInput').value.toLowerCase().trim();}
function filtered(){
  let list=tags;
  const estadoFilter=document.getElementById('estadoFilter').value;
  if(viewAllDays){list=getAllDaysTags();}
  if(estadoFilter)list=list.filter(t=>t.estado===estadoFilter);
  if(activeFilter==='none')list=list.filter(t=>!t.labelId);
  else if(activeFilter!=='all')list=list.filter(t=>t.labelId===Number(activeFilter));
  const q=searchQuery();
  if(q)list=list.filter(t=>t.name.toLowerCase().includes(q)||t.phone.includes(q));
  return sortedTags(list);
}
function getAllDaysTags(){const allKeys=Object.keys(localStorage).filter(k=>k.startsWith('waTags_'));let all=[];allKeys.forEach(key=>{try{const dt=JSON.parse(localStorage.getItem(key)||'[]');dt.forEach(t=>{const copy={...t};copy._dayKey=key.replace('waTags_','');all.push(copy);});}catch(e){}});return all;}
function renderList(){
  const container=document.getElementById('mainList');
  const pill=document.getElementById('countPill');
  const visible=filtered();
  const listTitle=document.getElementById('listTitle');
  if(listTitle){if(viewAllDays)listTitle.textContent='Todos os contatos';else{const ef=document.getElementById('estadoFilter').value;listTitle.textContent=ef?`Contatos - UF ${ef}`:'Contatos do dia';}}
  pill.textContent=visible.length+(visible.length===1?' registro':' registros');
  if(!visible.length){container.innerHTML=`<div class="empty-state"><div class="empty-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/></svg></div><div class="empty-title">${viewAllDays?'Nenhum contato encontrado.':'Nenhum contato hoje.'}</div><div class="empty-sub">Adicione um novo contato para começar.</div></div>`;return;}
  if(activeFilter!=='all'){container.innerHTML=`<div class="contact-table">${visible.map(t=>rowHTML(t)).join('')}</div>`;return;}
  const order=[],groups={};
  visible.forEach(t=>{const k=t.labelId?String(t.labelId):'__none__';if(!groups[k]){groups[k]=[];order.push(k);}groups[k].push(t);});
  container.innerHTML=order.map(key=>{
    const items=groups[key];
    let dot='var(--text-3)',title='Sem etiqueta';
    if(key!=='__none__'){const lbl=getLabelById(key);if(lbl){dot=lbl.colorHex;title=lbl.name;}}
    return `<div class="group-header"><div class="group-dot" style="background:${dot}"></div><span class="group-name-text" style="color:${dot}">${esc(title)}</span><span class="group-count-pill">${items.length}</span></div><div class="contact-table">${items.map(t=>rowHTML(t)).join('')}</div>`;
  }).join('');
}
function rowHTML(t){
  const lbl=(t.labelId?getLabelById(t.labelId):null)||(t._labelName?getLabelByName(t._labelName):null);
  const waNum=t.phone.replace(/\D/g,'');
  const avatarHTML=t.avatar?`<div class="av" style="border-color:${lbl?lbl.colorHex+'55':'var(--border)'}"><img src="${t.avatar}"></div>`:`<div class="av" style="background:${lbl?lbl.colorBg:'var(--bg2)'};color:${lbl?lbl.colorHex:'var(--text-3)'};border-color:${lbl?lbl.colorHex+'55':'var(--border)'}">${initials(t.name)}</div>`;
  const badge=lbl?`<span class="badge" style="color:${lbl.colorHex};background:${lbl.colorBg};border-color:${lbl.colorHex}44"><span class="badge-dot" style="background:${lbl.colorHex}"></span>${esc(lbl.name)}</span>`:`<span style="font-size:11px;color:var(--text-3)">—</span>`;
  const estadoPill=t.estado?`<span class="estado-pill-row">${t.estado}</span>`:'';
  const fichaIcon=(lbl&&lbl.financeiro&&t.cpf)?'<span title="Ficha" style="font-size:10px;color:var(--warning);margin-left:3px">💰</span>':'';
  const docIcon=(t.docFiles&&t.docFiles.length)?'<span title="Docs" style="font-size:10px;color:var(--info);margin-left:3px">📎</span>':'';
  const dateDisplay=t._dayKey?fmtDate(t._dayKey):fmtDate(t.date);
  const docChecked=t.docAnalisada?'checked':'';
  return `<div class="contact-row">
    <div class="cr-avatar">${avatarHTML}</div>
    <div class="cr-name">${esc(t.name)}${fichaIcon}${docIcon}</div>
    <div class="cr-phone">${esc(t.phone)}</div>
    <div class="cr-badge" style="display:flex;align-items:center;gap:4px">${badge}${estadoPill}</div>
    <div class="cr-date">${dateDisplay}</div>
    <div class="cr-check"><input type="checkbox" class="doc-cb" ${docChecked} onchange="toggleDocAnalisada(${t.id},this)" title="Doc. analisada?"></div>
    <div class="cr-action"><a class="act-btn wa" href="https://wa.me/55${waNum}" target="_blank" rel="noopener" title="WhatsApp"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 11.5a8.38 8.38 0 01-.9 3.8 8.5 8.5 0 01-7.6 4.7 8.38 8.38 0 01-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 01-.9-3.8 8.5 8.5 0 014.7-7.6 8.38 8.38 0 013.8-.9h.5a8.48 8.48 0 018 8v.5z"/></svg></a></div>
    <div class="cr-action"><button class="act-btn info" onclick="openFichaModal(${t.id})" title="Ver ficha"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg></button></div>
    <div class="cr-action"><button class="act-btn" onclick="openEditModal(${t.id})" title="Editar"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></button></div>
    <div class="cr-action"><button class="act-btn danger" onclick="deleteTag(${t.id})" title="Remover"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a1 1 0 011-1h4a1 1 0 011 1v2"/></svg></button></div>
  </div>`;
}
function updateDateBadge(){const d=getActiveDate();if(!d)return;const[y,m,day]=d.split('-');document.getElementById('dateBadge').textContent=`${day}/${m}/${y}`;}
function fmtDate(iso){if(!iso)return'—';if(iso.includes('/'))return iso;const clean=iso.includes('T')?iso.split('T')[0]:iso;const[y,m,d]=clean.split('-');return`${d}/${m}/${y}`;}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function csvEsc(s){return String(s).replace(/"/g,'""');}
function showToast(msg){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),2800);}
function setSyncBar(msg,cls){const b=document.getElementById('syncBar');b.textContent=msg;b.className='sync-indicator'+(cls?' '+cls:'');}

// filtro por estado 
function buildEstadoFilter(){const sel=document.getElementById('estadoFilter');if(!sel)return;while(sel.options.length>1)sel.remove(1);ESTADOS_UF.forEach(uf=>{const opt=document.createElement('option');opt.value=uf;opt.textContent=uf;sel.appendChild(opt);});}

// csv
function copyReport(){
  if(!tags.length){showToast('⚠️ Nenhum contato!');return;}
  const d=fmtDate(getActiveDate());let txt=`📋 RELATÓRIO — ${d}\n${'─'.repeat(42)}\n`;
  const order=[],groups={};
  tags.forEach(t=>{const k=t.labelId?String(t.labelId):'__none__';if(!groups[k]){groups[k]=[];order.push(k);}groups[k].push(t);});
  order.forEach(k=>{const items=groups[k];const gname=k==='__none__'?'Sem etiqueta':(getLabelById(k)?.name||'?');txt+=`\n🏷️ ${gname.toUpperCase()} (${items.length})\n`;items.forEach((t,i)=>{txt+=`  ${String(i+1).padStart(2,'0')}. ${t.name}\n      📱 ${t.phone}   📅 ${fmtDate(t.date)}\n`;});});
  txt+=`\n${'─'.repeat(42)}\nTotal: ${tags.length} contato${tags.length!==1?'s':''}\n`;
  navigator.clipboard.writeText(txt).then(()=>showToast('✅ Relatório copiado!')).catch(()=>{const ta=document.createElement('textarea');ta.value=txt;document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);showToast('✅ Relatório copiado!');});
}
function exportCSV(){
  if(!tags.length){showToast('⚠️ Nenhum contato!');return;}
  const header=`Nome,Telefone,Data de Contato,Etiqueta,__labelId__`;
  const meta=`"__LABELS_META__","${csvEsc(JSON.stringify(labels))}","","",""`;
  const rows=tags.map(t=>{const lbl=t.labelId?getLabelById(t.labelId):null;return`"${csvEsc(t.name)}","${csvEsc(t.phone)}","${fmtDate(t.date)}","${csvEsc(lbl?lbl.name:'')}","${t.labelId||''}"`});
  const blob=new Blob(['\uFEFF'+[header,meta,...rows].join('\n')],{type:'text/csv;charset=utf-8;'});
  const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=`contatos_${getActiveDate()}.csv`;a.click();URL.revokeObjectURL(url);showToast('💾 CSV exportado!');
}
function importCSV(e){
  const file=e.target.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=function(ev){
    const lines=ev.target.result.replace(/\r/g,'').split('\n').filter(l=>l.trim());
    if(lines.length<2){showToast('⚠️ CSV inválido!');return;}
    let imported=0;
    for(let i=1;i<lines.length;i++){
      const cols=parseCSVLine(lines[i]);if(!cols.length)continue;
      if(cols[0]==='__LABELS_META__'){try{const imp=JSON.parse(cols[1]);imp.forEach(il=>{if(!labels.find(l=>l.name.toLowerCase()===il.name.toLowerCase()))labels.push(il);});saveLabels();renderExistingLabels();renderLabelSelect();renderFilterChips();}catch(err){}continue;}
      if(cols.length<3)continue;
      const name=cols[0].trim(),phone=cols[1].trim(),dateRaw=cols[2].trim(),lblName=cols[3]?cols[3].trim():'',lblIdRaw=cols[4]?cols[4].trim():'';
      if(!name||!phone)continue;
      let dateISO=dateRaw;if(/^\d{2}\/\d{2}\/\d{4}$/.test(dateRaw)){const[d,m,y]=dateRaw.split('/');dateISO=`${y}-${m}-${d}`;}
      let labelId=null;if(lblIdRaw){const l=getLabelById(lblIdRaw);if(l)labelId=l.id;}else if(lblName){const l=labels.find(l=>l.name.toLowerCase()===lblName.toLowerCase());if(l)labelId=l.id;}
      if(!tags.find(t=>t.phone===phone)){tags.push({id:Date.now()+i,name,phone,date:dateISO,labelId,docFiles:[]});imported++;}
    }
    saveTags();renderList();updateSummary();updateMeta();
    showToast(`📥 ${imported} contato${imported!==1?'s':''} importado${imported!==1?'s':''}!`);
  };
  reader.readAsText(file,'utf-8');e.target.value='';
}
function parseCSVLine(line){const cols=[];let cur='',inQ=false;for(let i=0;i<line.length;i++){const c=line[i];if(c==='"'){if(inQ&&line[i+1]==='"'){cur+='"';i++;}else inQ=!inQ;continue;}if(c===','&&!inQ){cols.push(cur);cur='';continue;}cur+=c;}cols.push(cur);return cols;}

// excluir todos os contatos
function deleteAllContacts(){
  if(!tags.length){showToast('⚠️ Nenhum contato para excluir!');return;}
  const d = fmtDate(getActiveDate());
  const confirmed = confirm(`🗑️ EXCLUIR TODOS OS CONTATOS\n\nDia: ${d}\nQuantidade: ${tags.length} contato(s)\n\nEssa ação não pode ser desfeita.\nTem certeza?`);
  if(!confirmed) return;
  tags = [];
  saveTags();
  renderList();
  updateSummary();
  updateMeta();
  showToast(`🗑️ ${tags.length === 0 ? 'Todos os contatos foram excluídos!' : 'Excluído!'}`);
}

// backup 
async function exportBackup(){
  showToast('⏳ Gerando backup...');
  try {
    const zip = new JSZip();
    // localStorage
    const allKeys = Object.keys(localStorage);
    const backupData = {};
    allKeys.forEach(key => {
      // Inclui tudo
      if(
        key.startsWith('waTags_') ||
        key.startsWith('waLabels') ||
        key.startsWith('waEntrevistas') ||
        key.startsWith('waMeta') ||
        key.startsWith('waSheetsUrl') ||
        key.startsWith('waTheme') ||
        key.startsWith('pwaDismissed')
      ){
        backupData[key] = localStorage.getItem(key);
      }
    });

    const tagKeys = Object.keys(backupData).filter(k => k.startsWith('waTags_'));
    let totalContatos = 0;
    tagKeys.forEach(k => { try { totalContatos += JSON.parse(backupData[k]).length; } catch(e){} });

    const manifest = {
      version: '1.0',
      app: 'Trupe',
      exportedAt: new Date().toISOString(),
      totalContatos,
      totalDias: tagKeys.length,
      keys: Object.keys(backupData)
    };

    zip.file('manifest.json', JSON.stringify(manifest, null, 2));
    zip.file('dados.json', JSON.stringify(backupData));

    const content = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    const today = new Date().toISOString().split('T')[0];
    const url = URL.createObjectURL(content);
    const a = document.createElement('a');
    a.href = url;
    a.download = `trupe_backup_${today}.trupe`;
    a.click();
    URL.revokeObjectURL(url);
    showToast(`✅ Backup exportado! (${totalContatos} contatos, ${tagKeys.length} dias)`);
  } catch(err) {
    showToast('❌ Erro ao gerar backup: ' + err.message);
    console.error(err);
  }
}

async function importBackup(e) {
  const file = e.target.files[0];
  if(!file) return;
  e.target.value = '';

  const ext = file.name.split('.').pop().toLowerCase();
  if(ext !== 'trupe' && ext !== 'zip') {
    showToast('⚠️ Arquivo inválido! Use um arquivo .trupe');
    return;
  }

  const confirmMsg = `⚠️ ATENÇÃO\n\nImportar este backup vai MESCLAR os dados com os atuais.\n\nContatos do mesmo dia serão atualizados (sem duplicar).\nEtiquetas e agenda serão combinadas.\n\nContinuar?`;
  if(!confirm(confirmMsg)) return;

  showToast('⏳ Restaurando backup...');

  try {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const manifestFile = zip.file('manifest.json');
    if(!manifestFile) { showToast('❌ Arquivo corrompido (sem manifest)'); return; }
    const manifest = JSON.parse(await manifestFile.async('string'));
    if(manifest.app !== 'Trupe') { showToast('❌ Backup não é do Trupe!'); return; }

    // carregar 
    const dadosFile = zip.file('dados.json');
    if(!dadosFile) { showToast('❌ Arquivo corrompido (sem dados)'); return; }
    const backupData = JSON.parse(await dadosFile.async('string'));

    let contatosAdded = 0, contatosUpdated = 0, diasMerged = 0;

    Object.entries(backupData).forEach(([key, value]) => {
      if(key.startsWith('waTags_')) {
        //  n duplicar contatos pelo id ou telefone
        let existing = [];
        try { existing = JSON.parse(localStorage.getItem(key) || '[]'); } catch(e){}
        let incoming = [];
        try { incoming = JSON.parse(value); } catch(e){}
        incoming.forEach(t => {
          const idx = existing.findIndex(x => String(x.id) === String(t.id) || (x.phone && x.phone === t.phone));
          if(idx >= 0) { existing[idx] = { ...existing[idx], ...t }; contatosUpdated++; }
          else { existing.push(t); contatosAdded++; }
        });
        if(incoming.length) diasMerged++;
        localStorage.setItem(key, JSON.stringify(existing));
      } else if(key === 'waLabels') {
        // Merge de etiquetas: adiciona as que não existem
        let existingLabels = [];
        try { existingLabels = JSON.parse(localStorage.getItem('waLabels') || '[]'); } catch(e){}
        let incomingLabels = [];
        try { incomingLabels = JSON.parse(value); } catch(e){}
        incomingLabels.forEach(il => {
          if(!existingLabels.find(l => l.id === il.id || l.name.toLowerCase() === il.name.toLowerCase())) {
            existingLabels.push(il);
          }
        });
        localStorage.setItem('waLabels', JSON.stringify(existingLabels));
      } else if(key === 'waEntrevistas') {
        // agenda m
        let existing = [];
        try { existing = JSON.parse(localStorage.getItem('waEntrevistas') || '[]'); } catch(e){}
        let incoming = [];
        try { incoming = JSON.parse(value); } catch(e){}
        incoming.forEach(ev => {
          if(!existing.find(x => x.id === ev.id)) existing.push(ev);
        });
        localStorage.setItem('waEntrevistas', JSON.stringify(existing));
      } else if(key === 'waMeta') {
        // Meta: só aplica se não houver meta local definida
        if(!localStorage.getItem('waMeta')) localStorage.setItem('waMeta', value);
      } else if(key === 'waTheme' || key === 'waSheetsUrl') {
        // Configurações: aplica apenas se não existir localmente
        if(!localStorage.getItem(key)) localStorage.setItem(key, value);
      }
    });

    // recarregar
    loadLabelsFromStorage();
    loadTagsFromStorage();
    loadEntrevistas();
    loadMeta();

    showToast(`✅ Backup restaurado! +${contatosAdded} novos, ${contatosUpdated} atualizados em ${diasMerged} dias`);
  } catch(err) {
    showToast('❌ Erro ao restaurar: ' + err.message);
    console.error(err);
  }
}

//google sheets
function loadSheetsCredentials(){sheetsWebAppUrl=localStorage.getItem('waSheetsUrl')||'';const inp=document.getElementById('sheetsUrlInput');if(inp&&sheetsWebAppUrl)inp.value=sheetsWebAppUrl;updateSheetsStatus();}
function updateSheetsStatus(){const badge=document.getElementById('navSheetsStatus');if(badge)badge.style.display=sheetsWebAppUrl?'inline':'none';}
function openSheetsModal(){loadSheetsCredentials();const ta=document.getElementById('scriptDisplay');if(ta)ta.value=APPS_SCRIPT;document.getElementById('sheetsModal').classList.add('open');}
function closeSheetsModal(){document.getElementById('sheetsModal').classList.remove('open');}
function saveSheetsCreds(){const url=(document.getElementById('sheetsUrlInput').value||'').trim();if(!url){showSheetsStatus('Cole a URL do Web App primeiro.','err');return;}sheetsWebAppUrl=url;localStorage.setItem('waSheetsUrl',url);updateSheetsStatus();showSheetsStatus('✅ URL salva!','ok');}
function showSheetsStatus(msg,cls){const el=document.getElementById('sheetsStatus');if(!el)return;el.textContent=msg;el.className='sheets-status '+cls;}
async function syncToSheets(){
  if(!sheetsWebAppUrl){showSheetsStatus('Configure a URL primeiro.','err');return;}
  const allKeys=Object.keys(localStorage).filter(k=>k.startsWith('waTags_'));
  let allTags=[];allKeys.forEach(key=>{try{const dt=JSON.parse(localStorage.getItem(key)||'[]');allTags=allTags.concat(dt);}catch(e){}});
  if(!allTags.length){showSheetsStatus('⚠️ Nenhum contato para enviar.','err');return;}
  setSyncBar('⏳ Enviando para o Banco…','syncing');
  try{
    const resp=await fetch(sheetsWebAppUrl,{method:'POST',headers:{'Content-Type':'text/plain'},body:JSON.stringify({action:'write',tags:allTags,labels})});
    let result={status:'ok'};try{result=await resp.json();}catch(_){}
    if(result.status==='ok'){setSyncBar(`✅ ${allTags.length} contatos enviados!`,'ok');showSheetsStatus(`✅ ${allTags.length} contato(s) enviados!`,'ok');showToast('✅ Banco atualizado!');}
    else{setSyncBar('❌ Erro no envio','err');showSheetsStatus('❌ '+(result.message||'Erro desconhecido'),'err');}
  }catch(err){setSyncBar('❌ Erro de conexão','err');showSheetsStatus('❌ '+err.message,'err');}
}
async function syncFromSheets(){
  if(!sheetsWebAppUrl){showSheetsStatus('Configure a URL primeiro.','err');return;}
  const url=`${sheetsWebAppUrl}?action=read`;
  setSyncBar('⏳ Puxando do Banco…','syncing');showSheetsStatus('⏳ Buscando dados…','ok');
  try{
    const resp=await fetch(url);if(!resp.ok)throw new Error(`HTTP ${resp.status}`);
    const data=await resp.json();if(data.status!=='ok')throw new Error(data.message||'Erro');
    const incoming=Array.isArray(data.tags)?data.tags:[];
    if(!incoming.length){setSyncBar('ℹ️ Banco vazio — nenhum contato encontrado','ok');showSheetsStatus('ℹ️ O Banco está vazio. Envie os contatos primeiro usando "Enviar para o Banco".','ok');return;}
    // Merge labels vindas do Sheets antes de processar os contatos
    if(Array.isArray(data.labels)&&data.labels.length){
      data.labels.forEach(il=>{if(!labels.find(l=>l.id===il.id))labels.push(il);});
      saveLabels();renderExistingLabels();renderLabelSelect();renderFilterChips();
    }
    let added=0,updated=0;
    incoming.forEach(t=>{
      // Normalizar data: remover parte de tempo se vier como ISO completo
      if(t.date&&t.date.includes('T'))t.date=t.date.split('T')[0];
      // Resolver labelId a partir do labelName caso não venha no objeto
      if(!t.labelId&&t.labelName){
        const found=labels.find(l=>l.name.toLowerCase()===t.labelName.toLowerCase());
        if(found)t.labelId=found.id;
      }
      const date=t.date||getActiveDate();const key='waTags_'+date;let dayTags=[];
      try{dayTags=JSON.parse(localStorage.getItem(key)||'[]');}catch(e){}
      const idx=dayTags.findIndex(x=>x.id===t.id||x.phone===t.phone);
      if(idx>=0){dayTags[idx]={...dayTags[idx],...t};updated++;}
      else{dayTags.push({...t,docFiles:t.docFiles||[]});added++;}
      localStorage.setItem(key,JSON.stringify(dayTags));
    });
    loadTagsFromStorage();setSyncBar(`✅ ${added} novos, ${updated} atualizados`,'ok');showSheetsStatus(`✅ ${incoming.length} registro(s): ${added} novos, ${updated} atualizados.`,'ok');showToast(`📥 ${incoming.length} contato(s) sincronizados!`);
  }catch(err){setSyncBar('❌ Erro','err');showSheetsStatus('❌ '+err.message,'err');}
}
function copyAppsScript(){const ta=document.getElementById('scriptDisplay');navigator.clipboard.writeText(ta?ta.value:APPS_SCRIPT).then(()=>showSheetsStatus('✅ Script copiado!','ok')).catch(()=>showSheetsStatus('Copie manualmente.','err'));}

function loadEntrevistas(){const d=localStorage.getItem('waEntrevistas');entrevistas=d?JSON.parse(d):[];renderEntrevistas();document.getElementById('navBadgeAgenda').textContent=entrevistas.length;}
function saveEntrevistas(){localStorage.setItem('waEntrevistas',JSON.stringify(entrevistas));}
function addEntrevista(){
  const dateVal=document.getElementById('entrevistaDateInput').value;
  const timeVal=document.getElementById('entrevistaTimeInput').value||'';
  const descRaw=document.getElementById('entrevistaDescInput').value.trim();
  if(!dateVal){showToast('⚠️ Selecione a data!');return;}
  const descricao=descRaw||'Entrevista agendada';
  entrevistas.push({id:Date.now(),date:dateVal,time:timeVal,descricao,done:false});
  saveEntrevistas();renderEntrevistas();
  document.getElementById('navBadgeAgenda').textContent=entrevistas.length;
  document.getElementById('entrevistaDescInput').value='';
  document.getElementById('entrevistaTimeInput').value='';

  const labelName=`Entrevista ${fmtDate(dateVal)}`;
  const purpleHex='#7c3aed';const purpleBg='rgba(124,58,237,.14)';
  if(!labels.find(l=>l.name.toLowerCase()===labelName.toLowerCase())){
    labels.push({id:Date.now()+1,name:labelName,colorHex:purpleHex,colorBg:purpleBg,financeiro:false,isEntrevista:true});
    saveLabels();renderExistingLabels();renderLabelSelect();renderFilterChips();updateSummary();
    showToast(`📅 Agendado! Etiqueta "${labelName}" criada 🟣`);
  } else {
    showToast(`📅 Agendado para ${fmtDate(dateVal)}!`);
  }
}
function removeEntrevista(id){if(!confirm('Remover este agendamento?'))return;entrevistas=entrevistas.filter(e=>e.id!==id);saveEntrevistas();renderEntrevistas();showToast('✅ Removido!');document.getElementById('navBadgeAgenda').textContent=entrevistas.length;}
function toggleRealizado(id){const e=entrevistas.find(x=>x.id===id);if(!e)return;e.done=!e.done;saveEntrevistas();renderEntrevistas();showToast(e.done?'✅ Marcado como realizado!':'↩️ Desmarcado!');}
function startEditDesc(id){
  const e=entrevistas.find(x=>x.id===id);if(!e)return;
  const span=document.getElementById('desc-'+id);if(!span)return;
  const old=e.descricao;
  span.innerHTML=`<input class="agenda-edit-input" id="einput-${id}" value="${esc(old)}" onblur="saveDesc(${id})" onkeydown="if(event.key==='Enter')this.blur();if(event.key==='Escape'){this.value='${esc(old)}';this.blur();}">`;
  const inp=document.getElementById('einput-'+id);if(inp){inp.focus();inp.select();}
}
function saveDesc(id){
  const inp=document.getElementById('einput-'+id);if(!inp)return;
  const val=inp.value.trim()||'Entrevista agendada';
  const e=entrevistas.find(x=>x.id===id);if(e){e.descricao=val;saveEntrevistas();}
  renderEntrevistas();
}
function renderEntrevistas(){
  const container=document.getElementById('entrevistaFichas');if(!container)return;
  if(!entrevistas.length){container.innerHTML='<div class="agenda-empty">📭 Nenhum agendamento.</div>';return;}
  const today=new Date().toISOString().split('T')[0];
  const sorted=[...entrevistas].sort((a,b)=>{const dc=a.date.localeCompare(b.date);if(dc!==0)return dc;return(a.time||'').localeCompare(b.time||'');});
  container.innerHTML=sorted.map(e=>{
    const isToday=e.date===today,isPast=e.date<today&&!e.done,isDone=!!e.done;
    let cls='';let statusHtml='';
    if(isDone){cls='done';statusHtml=`<span class="agenda-item-status" style="color:var(--accent)">✓ Realizado</span>`;}
    else if(isToday){cls='today';statusHtml=`<span class="agenda-item-status" style="color:#7c3aed">● Hoje</span>`;}
    else if(isPast){cls='past';statusHtml=`<span class="agenda-item-status" style="color:var(--danger)">⏪ Não realizado</span>`;}
    else{statusHtml=`<span class="agenda-item-status" style="color:var(--text-3)">Agendado</span>`;}
    const timeStr=e.time?`<div class="agenda-item-time">🕐 ${e.time}</div>`:'';
    const doneIcon=isDone?'✓':'○';
    return `<div class="agenda-item ${cls}">
      <button class="agenda-done-btn" onclick="toggleRealizado(${e.id})" title="${isDone?'Desmarcar':'Marcar como realizado'}">${doneIcon}</button>
      <div class="agenda-item-left">
        <div class="agenda-item-date">${fmtDate(e.date)}</div>
        ${timeStr}
      </div>
      <div class="agenda-item-center">
        ${statusHtml}
        <div class="agenda-item-desc" id="desc-${e.id}" ondblclick="startEditDesc(${e.id})" title="Duplo clique para editar">${esc(e.descricao)}</div>
      </div>
      ${isToday&&!isDone?'<span class="agenda-today-tag">HOJE</span>':''}
      <button onclick="removeEntrevista(${e.id})" style="background:none;border:none;color:var(--text-3);cursor:pointer;font-size:13px;padding:0;width:24px;height:24px;border-radius:4px;transition:color .14s;flex-shrink:0;" onmouseover="this.style.color='var(--danger)'" onmouseout="this.style.color='var(--text-3)'">✕</button>
    </div>`;
  }).join('');
}

async function testeSupabase() {
  const { data, error } = await supabase
    .from('contatos')
    .select('*');

  console.log('DATA:', data);
  console.log('ERROR:', error);
}

testeSupabase();