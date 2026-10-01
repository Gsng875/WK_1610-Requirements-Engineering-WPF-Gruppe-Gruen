const $ = id => document.getElementById(id);
let me, csrf, current, decision, noticeTimer;
const labels = {offen:'Prüfung offen',freigegeben:'Freigegeben',abgelehnt:'Abgelehnt'};
const date = value => value ? new Date(value).toLocaleDateString('de-DE') : '–';
function notify(message, error=false) {
 $('notice').textContent=message; $('notice').className=error?'error-notice':''; $('notice').hidden=false;
 clearTimeout(noticeTimer); noticeTimer=setTimeout(()=>$('notice').hidden=true,6000);
}
async function api(url, data) {
 // Änderungen senden das Sicherheitstoken; der Server prüft die Berechtigung.
 const response=await fetch(url,{method:data?'POST':'GET',headers:data?{'Content-Type':'application/json','X-CSRF-Token':csrf||''}:{},body:data?JSON.stringify(data):undefined});
 const value=await response.json();
 if (!response.ok) {
  if(response.status===401 && me){me=null; $('confirmation').close(); showLogin();}
  throw Error(value.error || 'Anfrage fehlgeschlagen.');
 }
 return value;
}
function showLogin(){ $('workspace').hidden=true; $('login').hidden=false; $('password').value=''; $('username').focus(); }
function showPage(name){for(const page of ['overview','detail','rules']) $(page).hidden=page!==name; $('overview-nav').classList.toggle('active',name!=='rules'); $('rules-nav').classList.toggle('active',name==='rules');}
function textElement(tag,text,className){const e=document.createElement(tag);e.textContent=text;if(className)e.className=className;return e;}
async function overview(){
 // Themen kommen ausschließlich aus der serverseitig gefilterten Liste.
 const topics=await api('/api/topics');
 $('total').textContent=topics.length; $('pending').textContent=topics.filter(t=>t.status==='offen').length; $('done').textContent=topics.filter(t=>t.status!=='offen').length;
 $('topic-list').replaceChildren();
 for(const t of topics){
  const row=document.createElement('article');row.className='topic-row';
  const summary=document.createElement('div');summary.append(textElement('small',`TF-${String(t.id).padStart(3,'0')} · ${t.degree}`),textElement('h3',t.title),textElement('p',`${t.student} · geplanter Start ${date(t.plannedStart)}`));
  const meta=document.createElement('div');meta.className='row-meta';meta.append(textElement('span',labels[t.status],`badge ${t.status}`));
  const button=textElement('button',t.status==='offen'?'Thema prüfen →':'Entscheidung ansehen →','secondary');button.onclick=()=>loadTopic(t.id).catch(e=>notify(e.message,true));meta.append(button);row.append(summary,meta);$('topic-list').append(row);
 }
 showPage('overview');
}
async function loadTopic(id){
 // Textinhalt verhindert, dass Themen oder Kommentare als HTML ausgeführt werden.
 current=await api(`/api/topics/${id}`);
 for(const key of ['title','student','matriculation','degree','context','question','goal','method','deliverables','secondExaminer']) $(key).textContent=current[key];
 $('scope-text').textContent=current.scope; $('plannedStart').textContent=date(current.plannedStart); $('examiner').textContent=me.name;
 $('topic-id').textContent=`VORGANG TF-${String(id).padStart(3,'0')} · THEMENFREIGABE`;
 $('status').textContent=labels[current.status];$('status').className=`badge ${current.status}`;
 const ready=current.authorized&&current.complete;
 $('prerequisite').textContent=ready?'Voraussetzungen hinterlegt: Angaben vollständig, Zulassung bestätigt.':'Freigabe gesperrt: Die vorherige Zulassung ist noch nicht bestätigt.';
 $('prerequisite').className=ready?'info':'info warning';
 $('rating').value=current.review.rating||''; $('comment').value=current.review.comment||''; $('academic').checked=!!current.review.academic; $('scope').checked=!!current.review.scope;
 $('review-fields').disabled=current.status!=='offen';$('approve').disabled=!ready;
 $('decision-result').hidden=current.status==='offen';$('decision-result').textContent=`${labels[current.status]} am ${date(current.decidedAt)}. Die Rückmeldung ist gespeichert.`;
 $('history').replaceChildren();
 for(const h of current.history){const li=textElement('li',h.action);li.append(textElement('small',`${h.actor} · ${new Date(h.at).toLocaleString('de-DE')}`));$('history').append(li);}
 if(!current.history.length)$('history').append(textElement('li','Noch keine Einschätzung oder Entscheidung gespeichert.'));
 showPage('detail');window.scrollTo({top:0,behavior:'smooth'});
}
function review(){return {rating:$('rating').value,comment:$('comment').value,academic:$('academic').checked,scope:$('scope').checked};}
function askDecision(value){
 // Der Dialog fragt eine zweite, ausdrückliche Bestätigung ab.
 if(!$('review-form').reportValidity())return;
 const r=review();
 if(value==='freigegeben'&&(r.rating!=='geeignet'||!r.academic||!r.scope)){notify('Für die Freigabe bitte „Geeignet“ auswählen und beide Prüfkriterien bestätigen.',true);return;}
 decision=value;$('confirm-check').checked=false;
 $('confirm-title').textContent=value==='freigegeben'?'Thema verbindlich freigeben?':'Thema verbindlich ablehnen?';
 $('confirm-text').textContent=`${current.title} · ${current.student}`;$('confirmation').showModal();
}
async function enter(account){me=account.user;csrf=account.csrf;$('user-name').textContent=me.name;document.querySelector('.avatar').textContent=me.id==='weber'?'AW':'JK';$('login').hidden=true;$('workspace').hidden=false;await overview();}
$('login-form').onsubmit=async e=>{e.preventDefault();$('login-error').textContent='';const button=e.submitter;button.disabled=true;try{await enter(await api('/api/login',{username:$('username').value,password:$('password').value}));$('password').value='';}catch(e){$('login-error').textContent=e.message;}finally{button.disabled=false;}};
$('review-form').onsubmit=async e=>{e.preventDefault();const button=e.submitter;button.disabled=true;try{await api(`/api/topics/${current.id}/review`,review());await loadTopic(current.id);notify('Einschätzung gespeichert. Das Thema ist noch nicht freigegeben.');}catch(e){notify(e.message,true);}finally{button.disabled=false;}};
$('approve').onclick=()=>askDecision('freigegeben');$('reject').onclick=()=>askDecision('abgelehnt');$('cancel').onclick=()=>$('confirmation').close();
$('confirm-form').onsubmit=async e=>{e.preventDefault();$('confirm-submit').disabled=true;try{await api(`/api/topics/${current.id}/decision`,{...review(),decision,confirmed:$('confirm-check').checked});$('confirmation').close();await loadTopic(current.id);notify(decision==='freigegeben'?'Thema bestätigt und freigegeben.':'Thema abgelehnt. Rückmeldung gespeichert.');}catch(e){notify(e.message,true);}finally{$('confirm-submit').disabled=false;}};
for(const id of ['back','overview-nav']) $(id).onclick=()=>overview().catch(e=>notify(e.message,true));
$('rules-nav').onclick=()=>showPage('rules');
$('logout').onclick=async()=>{try{await api('/api/logout',{});me=null;csrf=null;showLogin();notify('Sie wurden abgemeldet.');}catch(e){notify(e.message,true);}};
api('/api/me').then(enter).catch(()=>showLogin());
