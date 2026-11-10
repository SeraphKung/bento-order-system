const MENU = ['原味蒜香舒肥雞','韓式辣味舒肥雞胸','日式照燒舒肥雞胸','檸檬椒鹽舒肥雞胸','油蔥梅花豬'];
const HEADERS = ['id','createdAt','source','name','items','total','status','receipt','requestId','updatedAt'];
function setup() {
  const p=PropertiesService.getScriptProperties();
  if (!p.getProperty('SHEET_ID')) {
    const ss=SpreadsheetApp.create('便當智慧點餐系統｜正式訂單');
    const s=ss.getSheets()[0]; s.setName('Orders'); s.appendRow(HEADERS); s.setFrozenRows(1);
    s.getRange('A1:J1').setBackground('#264d3b').setFontColor('#ffffff').setFontWeight('bold');
    s.setColumnWidths(1,10,160); s.setColumnWidth(5,420);
    p.setProperty('SHEET_ID',ss.getId());
    const m=ss.insertSheet('Menu'); m.appendRow(['品項','價格']); MENU.forEach(n=>m.appendRow([n,125])); m.setFrozenRows(1);
  }
  if(!p.getProperty('STAFF_KEY')) p.setProperty('STAFF_KEY',Utilities.getUuid()+Utilities.getUuid());
  console.log('試算表：https://docs.google.com/spreadsheets/d/'+p.getProperty('SHEET_ID'));
  console.log('管理金鑰（僅提供店員及廚房）：'+p.getProperty('STAFF_KEY'));
}
function sheet_(){const id=PropertiesService.getScriptProperties().getProperty('SHEET_ID');if(!id)throw Error('系統尚未初始化');return SpreadsheetApp.openById(id).getSheetByName('Orders');}
function staff_(key){const correct=PropertiesService.getScriptProperties().getProperty('STAFF_KEY');if(!correct||key!==correct)throw Error('管理金鑰不正確');}
function rows_(){const s=sheet_();return s.getLastRow()<2?[]:s.getRange(2,1,s.getLastRow()-1,10).getValues();}
function view_(r){return {id:r[0],createdAt:r[1],source:r[2],name:r[3],items:JSON.parse(r[4]),total:r[5],status:r[6],updatedAt:r[9]};}
function clean_(v,max){return String(v||'').trim().slice(0,max).replace(/^[=+@-]/,'\u200b$&');}
function api(action,data){
  data=data||{};
  if(action==='menu')return {menu:MENU,price:125};
  if(action==='list'){staff_(data.key);return rows_().slice(-300).reverse().map(view_);}
  if(action==='track'){const r=rows_().find(r=>r[0]===data.id&&r[7]===data.receipt);if(!r)throw Error('找不到訂單或憑證不正確');return view_(r);}
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try {
    if(action==='create'){
      if(!/^[a-zA-Z0-9-]{16,80}$/.test(data.requestId||''))throw Error('請求編號無效');
      const name=clean_(data.name,40);if(!name)throw Error('請輸入取餐稱呼');
      if(data.source==='staff')staff_(data.key);
      const existing=rows_().find(r=>r[8]===data.requestId);
      if(existing)return Object.assign(view_(existing),{receipt:existing[7]});
      if(!Array.isArray(data.quantities)||data.quantities.length!==5)throw Error('品項格式錯誤');
      const quantities=data.quantities;if(quantities.some(q=>!Number.isInteger(q)||q<0||q>20))throw Error('每款數量須為 0–20');
      const count=quantities.reduce((a,b)=>a+b,0);if(count<1||count>50)throw Error('每張訂單須為 1–50 份');
      const items=MENU.map((name,i)=>({name,quantity:quantities[i],price:125})).filter(i=>i.quantity);
      const now=new Date().toISOString(),id='B'+Utilities.formatDate(new Date(),'Asia/Taipei','yyyyMMdd')+'-'+Utilities.getUuid().slice(0,8).toUpperCase(),receipt=Utilities.getUuid()+Utilities.getUuid();
      const row=[id,now,data.source==='staff'?'店員':'顧客',name,JSON.stringify(items),count*125,'待接單',receipt,data.requestId,now];
      sheet_().appendRow(row);return Object.assign(view_(row),{receipt});
    }
    if(action==='status'){
      staff_(data.key);const rows=rows_(),index=rows.findIndex(r=>r[0]===data.id);if(index<0)throw Error('訂單不存在');
      const r=rows[index],next={'待接單':['製作中','已取消'],'製作中':['待取餐','已取消'],'待取餐':['已完成']}[r[6]]||[];
      if(!next.includes(data.status))throw Error('狀態已變更，請重新整理');
      sheet_().getRange(index+2,7).setValue(data.status);sheet_().getRange(index+2,10).setValue(new Date().toISOString());r[6]=data.status;r[9]=new Date().toISOString();return view_(r);
    }
    throw Error('不支援的操作');
  }finally{lock.releaseLock();}
}
function doGet(e){return HtmlService.createHtmlOutput(APP_HTML).setTitle('便當智慧點餐').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}
const APP_HTML = `<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><style>
:root{font-family:system-ui,sans-serif;color:#17372c;background:#f5f3eb}*{box-sizing:border-box}body{margin:0}header{background:#214b3a;color:white;padding:24px}header>div,main{max-width:1040px;margin:auto}h1{margin:8px 0;font-size:30px}h2{font-size:23px}small{opacity:.8}nav{display:flex;gap:8px;margin-top:20px;flex-wrap:wrap}button,input,select{font:inherit;border-radius:10px;padding:12px;border:1px solid #cbd3c9}button{cursor:pointer;background:#214b3a;color:white;border:0}button:disabled{opacity:.45;cursor:wait}nav button{background:#ffffff25}nav button.active{background:#efce75;color:#17372c}main{padding:20px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px}.card{background:white;border:1px solid #e0e4da;border-radius:16px;padding:20px;margin-bottom:14px}.dish{display:flex;justify-content:space-between;gap:12px;align-items:center}.dish input{width:75px}label{display:block;margin:12px 0}input{max-width:100%}.bar{display:flex;gap:10px;align-items:center;flex-wrap:wrap}.total{font-size:26px;font-weight:700}.error{color:#a12323}#message{min-height:28px;margin:10px 0}#message:empty{display:none}.tag{background:#f6e9bb;padding:6px 12px;border-radius:20px;display:inline-block}p{line-height:1.6}.muted{color:#647268}a{color:#214b3a} .wide{width:100%} .receipt{overflow-wrap:anywhere}footer{text-align:center;padding:20px;color:#647268}
</style></head><body><header><div><small>每日好食 · 現點現做</small><h1>便當智慧點餐</h1><div>五款招牌便當，每款 NT$125</div><nav><button data-mode="customer" class="active">顧客點餐</button><button data-mode="staff">店員點餐</button><button data-mode="kitchen">廚房接單</button><button data-mode="track">訂單進度</button></nav></div></header><main><div id="message" role="status" aria-live="polite"></div><div id="content"></div></main><footer>現場付款・請保留訂單憑證以查詢進度</footer><script>
const names=${JSON.stringify(MENU)};
let mode='customer',key='',timer,busy=false,pending=null;
const $=id=>document.getElementById(id),el=(tag,text)=>{const e=document.createElement(tag);e.textContent=text;return e;};
function call(action,data){return new Promise((resolve,reject)=>google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).api(action,data));}
function message(text,error=false){$('message').textContent=text;$('message').className=error?'error':'';}
function saved(){try{return JSON.parse(localStorage.getItem('bento-receipt')||'null')}catch(e){return null}}
function save(r){try{localStorage.setItem('bento-receipt',JSON.stringify({id:r.id,receipt:r.receipt}));}catch(e){}}
function switchMode(next){clearInterval(timer);mode=next;pending=null;message('');document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('active',b.dataset.mode===mode));render();}
document.querySelectorAll('nav button').forEach(b=>b.onclick=()=>switchMode(b.dataset.mode));
function login(after){$('content').innerHTML='<section class="card"><h2>工作人員驗證</h2><p>請輸入店家提供的管理金鑰。金鑰僅保存在此頁面記憶體。</p><form id="login"><label>管理金鑰 <input id="key" type="password" autocomplete="off" required></label><button>登入</button></form></section>';$('login').onsubmit=async e=>{e.preventDefault();const candidate=$('key').value;try{await call('list',{key:candidate});key=candidate;message('驗證成功');after();}catch(e){message(e.message,true);}};}
function render(){if((mode==='staff'||mode==='kitchen')&&!key){login(render);return}if(mode==='kitchen'){kitchen();timer=setInterval(()=>refreshKitchen(),5000);return}if(mode==='track'){trackForm();return}orderForm();}
function orderForm(){const root=$('content');root.innerHTML='<h2>'+(mode==='staff'?'店員代客點餐':'今天想吃哪一款？')+'</h2><p class="muted">每款限 20 份，每單最多 50 份。送出後由廚房確認。</p><form id="order"><div class="grid" id="menu"></div><section class="card"><label>取餐稱呼 <input id="name" maxlength="40" required placeholder="例如：王先生"></label><div class="bar"><span class="total" id="total">NT$0</span><button id="submit">送出訂單</button></div></section></form>';
names.forEach((n,i)=>{const d=el('div','');d.className='card dish';const box=el('div','');box.append(el('h3',n),el('p','NT$125'));const q=document.createElement('input');q.type='number';q.min=0;q.max=20;q.step=1;q.value=0;q.id='q'+i;q.setAttribute('aria-label',n+' 數量');q.oninput=()=>{pending=null;$('total').textContent='NT$'+names.reduce((s,_,j)=>s+Number($('q'+j).value||0)*125,0)};d.append(box,q);$('menu').append(d)});
$('name').oninput=()=>pending=null;$('order').onsubmit=async e=>{e.preventDefault();if(busy)return;const quantities=names.map((_,i)=>Number($('q'+i).value));const count=quantities.reduce((a,b)=>a+b,0);if(count<1||count>50||quantities.some(q=>!Number.isInteger(q)||q<0||q>20)){message('請選擇 1–50 份，每款最多 20 份。',true);return}pending=pending||crypto.randomUUID();busy=true;$('submit').disabled=true;message('訂單傳送中，請稍候…');try{const r=await call('create',{name:$('name').value,quantities,source:mode==='staff'?'staff':'customer',key,requestId:pending});save(r);pending=null;message('訂單已確認寫入，請保留下方憑證');showReceipt(r);}catch(e){message(e.message+'。重試會沿用同一請求編號，避免重複下單。',true);if($('submit'))$('submit').disabled=false;}finally{busy=false;}};}
function showReceipt(r){$('content').innerHTML='<section class="card"><h2>訂單已送出</h2><div id="receipt"></div><p class="receipt" id="credential"></p><button id="progress">查看進度</button></section>';$('receipt').append(orderCard(r));$('credential').textContent='查詢憑證：'+r.receipt;$('progress').onclick=()=>switchMode('track');}
function orderCard(r){const card=el('article','');card.className='card';card.append(el('h3',r.id),el('span',r.status),el('p',r.name+' · '+r.source+' · '+new Date(r.createdAt).toLocaleString('zh-TW')));r.items.forEach(i=>card.append(el('p',i.name+' × '+i.quantity)));card.append(el('strong','合計 NT$'+r.total));return card;}
function trackForm(){const s=saved();$('content').innerHTML='<section class="card"><h2>查詢訂單進度</h2><form id="track"><label>訂單編號 <input class="wide" id="id" required></label><label>查詢憑證 <input class="wide" id="receipt" required></label><button>查詢</button></form></section><div id="result"></div>';$('id').value=s?s.id:'';$('receipt').value=s?s.receipt:'';$('track').onsubmit=e=>{e.preventDefault();refreshTrack();clearInterval(timer);timer=setInterval(refreshTrack,5000)};if(s){refreshTrack();timer=setInterval(refreshTrack,5000)}}
let tracking=false;async function refreshTrack(){if(mode!=='track'||tracking)return;tracking=true;try{const r=await call('track',{id:$('id').value,receipt:$('receipt').value});if(mode!=='track')return;$('result').replaceChildren(orderCard(r));message('已同步 · '+new Date().toLocaleTimeString('zh-TW'));}catch(e){message(e.message,true)}finally{tracking=false}}
function kitchen(){$('content').innerHTML='<div class="bar"><h2>廚房訂單看板</h2><button id="refresh">立即同步</button><button id="logout">登出</button></div><p>每 5 秒更新。請依實際製作進度操作。</p><div id="orders"></div>';$('refresh').onclick=refreshKitchen;$('logout').onclick=()=>{key='';switchMode('kitchen')};refreshKitchen();}
let refreshing=false;async function refreshKitchen(){if(mode!=='kitchen'||refreshing)return;refreshing=true;try{const orders=await call('list',{key});if(mode!=='kitchen')return;$('orders').replaceChildren();const active=orders.filter(r=>!['已完成','已取消'].includes(r.status));if(!active.length)$('orders').append(el('p','目前沒有待處理訂單'));active.forEach(r=>{const card=orderCard(r);const targets={'待接單':['製作中','已取消'],'製作中':['待取餐','已取消'],'待取餐':['已完成']}[r.status]||[];targets.forEach(status=>{const b=el('button',status==='製作中'?'接單並開始製作':status);b.style.margin='12px 8px 0 0';b.onclick=async()=>{if(status==='已取消'&&!confirm('確定取消此訂單？'))return;b.disabled=true;try{await call('status',{key,id:r.id,status});await refreshKitchenAfter();}catch(e){message(e.message,true);b.disabled=false;}};card.append(b)});$('orders').append(card)});message('已同步 · '+new Date().toLocaleTimeString('zh-TW'));}catch(e){message('同步失敗：'+e.message,true)}finally{refreshing=false}}
async function refreshKitchenAfter(){refreshing=false;await refreshKitchen()}
render();
</script></body></html>`;
