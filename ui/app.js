const $ = s => document.querySelector(s);
let rows = [], editing = null, connected = false, busy = false;
function message(text, error=false) { $('#message').hidden = !text; $('#message').textContent = text; $('#message').className = error ? 'error' : ''; }
function controls() { document.querySelectorAll('button').forEach(b => b.disabled = busy || (b.hasAttribute('data-connected') && !connected)); document.body.classList.toggle('busy',busy); }
async function call(method, ...args) { const result = await window.pywebview.api[method](...args); if (!result.ok) throw Error(result.error); return result.data; }
async function task(fn) { if(busy) return; busy=true; controls(); try { await fn(); } catch(e) {message(e.message,true);} finally {busy=false;controls();} }
function render() {
 const query=$('#search').value.toLocaleLowerCase('pt-BR');
 const filtered=rows.filter(r=>[r.IDCLIENTE,r.NOME,r.DATA].some(v=>v.toLocaleLowerCase('pt-BR').includes(query)));
 $('#rows').replaceChildren(); $('#count').textContent=`${filtered.length} de ${rows.length} registro(s)`;
 for(const row of filtered) { const tr=document.createElement('tr');
  for(const key of ['IDCLIENTE','NOME','DATA']) {const td=document.createElement('td');td.textContent=key==='DATA' && row[key] ? row[key].split('-').reverse().join('/') : row[key];tr.append(td);}
  const td=document.createElement('td');
  const edit=document.createElement('button');edit.textContent='Editar';edit.onclick=()=>openEditor(row);
  const del=document.createElement('button');del.textContent='Excluir';del.className='danger';
  del.onclick=()=>{if(busy || !confirm(`Excluir o cliente ${row.IDCLIENTE} — ${row.NOME}?`))return;task(async()=>{await call('delete',row.id,row.etag);message('Registro excluído.');await load();});};
  td.append(edit,del);tr.append(td);$('#rows').append(tr);
 }
 if(!filtered.length){const tr=document.createElement('tr'),td=document.createElement('td');td.colSpan=4;td.className='empty';td.textContent=rows.length?'Nenhum registro corresponde à busca.':'Nenhum registro na lista. Crie o primeiro.';tr.append(td);$('#rows').append(tr);}
 controls();
}
async function load(){try{rows=await call('rows');render();}catch(e){rows=[];render();$('#count').textContent='Não foi possível atualizar a lista';throw e;}}
function openEditor(row=null){if(busy)return;editing=row;$('#record').reset();$('#editor-error').textContent='';$('#editor-title').textContent=row?'Editar registro':'Novo registro';if(row)for(const k of ['IDCLIENTE','NOME','DATA'])$('#record').elements[k].value=row[k];$('#editor').showModal();}
$('#connection').onsubmit=e=>{e.preventDefault();task(async()=>{connected=false;rows=[];render();$('#account').textContent='Conectando…';$('#disconnect').hidden=true;const cfg=Object.fromEntries(new FormData(e.target));try{const user=await call('connect',cfg);connected=true;$('#account').textContent=user.username||'Conta conectada';$('#disconnect').hidden=false;$('#settings').open=false;await load();message('Conectado ao SharePoint.');}catch(err){$('#account').textContent=connected?'Conectado • falha ao carregar':'Desconectado';throw err;}});};
$('#record').onsubmit=e=>{e.preventDefault();task(async()=>{try{await call('save',Object.fromEntries(new FormData(e.target)),editing?.id||null,editing?.etag||null);}catch(err){$('#editor-error').textContent=err.message;return;}$('#editor').close();message('Registro salvo.');await load();});};
$('#new').onclick=()=>openEditor();$('#refresh').onclick=()=>task(async()=>{await load();message('Tabela atualizada.');});
$('#search').oninput=render;$('#close').onclick=$('#cancel').onclick=()=>{if(!busy)$('#editor').close();};
$('#editor').addEventListener('cancel',e=>{if(busy)e.preventDefault();});
$('#disconnect').onclick=()=>task(async()=>{await call('disconnect');connected=false;rows=[];render();$('#rows').innerHTML='<tr><td colspan="4" class="empty">Conecte sua conta para carregar os registros.</td></tr>';$('#count').textContent='Conecte para visualizar a lista';$('#account').textContent='Desconectado';$('#disconnect').hidden=true;$('#settings').open=true;message('Sessão encerrada neste aplicativo. Sua conta do Windows continua conectada.');});
window.addEventListener('pywebviewready',()=>task(async()=>{const cfg=await call('settings');for(const [k,v] of Object.entries(cfg)){if($('#connection').elements[k])$('#connection').elements[k].value=v;}}));
