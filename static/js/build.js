 const list=document.getElementById('list');
 const SLIDE_TITLES=window.SLIDE_TITLES, RESUME=window.RESUME, SERVER_CTX=window.SERVER_CTX, BUILD_ID=window.BUILD_ID;
 function currentOrder(){return [...list.children].map(li=>li.dataset.id);}
 function persist(){var d=loadDeck()||{}; d.active=true; d.order=currentOrder();
   if(!RESUME){d.ctx=SERVER_CTX;d.buildId=BUILD_ID;} saveDeck(d);}
 function setH(id,v){var e=document.getElementById(id);if(e)e.value=v||'';}
 function fillHiddenCtx(c){c=c||{};setH('h-client',c.client_name);setH('h-industry',c.industry);
   setH('h-transcript',c.transcript);setH('h-phase',c.phase);setH('h-recipient',c.recipient);
   var box=document.getElementById('h-multi');if(box){box.innerHTML='';
     (c.functions||[]).forEach(f=>box.insertAdjacentHTML('beforeend','<input type="hidden" name="functions" value="'+f+'">'));
     (c.work_types||[]).forEach(w=>box.insertAdjacentHTML('beforeend','<input type="hidden" name="work_types" value="'+w+'">'));}}
 function rebuild(order){list.innerHTML='';(order||[]).forEach(id=>list.appendChild(makeRow(id,SLIDE_TITLES[id]||'','in your deck')));}
 function renumber(){[...list.children].forEach((li,i)=>li.querySelector('.pos').textContent=(i+1));
   const n=list.children.length;document.getElementById('sum-count').textContent=n;
   document.getElementById('count-tag').textContent=n+' slides';persist();}
 function move(b,d){const li=b.closest('li');
   if(d<0&&li.previousElementSibling)list.insertBefore(li,li.previousElementSibling);
   if(d>0&&li.nextElementSibling)list.insertBefore(li.nextElementSibling,li);renumber();}
 function removeRow(b){b.closest('li').remove();renumber();}
 function addSlide(id,btn){id=id||document.getElementById('addsel').value;if(!id)return;
   if([...list.children].some(li=>li.dataset.id===id)){if(!btn)alert(id+' is already in the deck');return;}
   list.appendChild(makeRow(id,SLIDE_TITLES[id]||'','added manually'));renumber();
   if(btn){btn.disabled=true;btn.innerHTML='<i class="ti ti-check"></i> Added';}}
 function prefillAI(name,desc,domain,useCase){var t=document.getElementById('ca-topic');if(t)t.value=name||'';
   var p=document.getElementById('ca-problem');if(p){var full=desc||'';
     if(domain) full+=(full?' ':'')+'Domain: '+domain+'.';
     if(useCase) full+=(full?' ':'')+'Use case: '+useCase+'.';
     p.value=full;}
   var card=document.getElementById('ai-create');if(card)card.scrollIntoView({behavior:'smooth',block:'center'});
   if(t)t.focus();}
 function filterAdd(q){q=(q||'').trim().toLowerCase();const sel=document.getElementById('addsel');
   let first=null;
   sel.querySelectorAll('optgroup').forEach(g=>{let vis=false;
     g.querySelectorAll('option').forEach(o=>{
       const m=!q||o.textContent.toLowerCase().includes(q)||o.value.toLowerCase().includes(q);
       o.hidden=!m; if(m){vis=true; if(!first)first=o;}});
     g.hidden=!vis;});
   if(first)sel.value=first.value;
   const none=document.getElementById('addsearch-none'); if(none)none.style.display=first?'none':'block';}
 function makeRow(id,title,reason){const li=document.createElement('li');
   li.className='slide-item';li.draggable=true;li.dataset.id=id;
   var tg=id; if(id.indexOf('NEW:')===0)tg='AI'; else if(id==='SK:industry')tg='IND'; else if(id==='SK:skills')tg='SKL'; else if(id.indexOf('SK:')===0)tg='CAP'; else if(id.indexOf('FP:')===0)tg='FOOT';
   li.innerHTML='<span class="grip"><i class="ti ti-grip-vertical"></i></span><span class="pos"></span>'+
    '<span class="sid">'+tg+'</span><div class="s-main"><div class="s-title"></div>'+
    '<div class="s-reason">'+reason+'</div></div><div class="s-actions">'+
    '<button type="button" class="mini" onclick="move(this,-1)"><i class="ti ti-chevron-up"></i></button>'+
    '<button type="button" class="mini" onclick="move(this,1)"><i class="ti ti-chevron-down"></i></button>'+
    '<button type="button" class="mini rm" onclick="removeRow(this)"><i class="ti ti-x"></i></button></div>';
   li.querySelector('.s-title').textContent=title;return li;}
 let dragEl=null;
 list.addEventListener('dragstart',e=>{dragEl=e.target.closest('li');dragEl.classList.add('drag');});
 list.addEventListener('dragend',e=>{if(dragEl)dragEl.classList.remove('drag');renumber();});
 list.addEventListener('dragover',e=>{e.preventDefault();const li=e.target.closest('li');if(!li||li===dragEl)return;
   const r=li.getBoundingClientRect();const after=(e.clientY-r.top)/r.height>0.5;
   list.insertBefore(dragEl,after?li.nextElementSibling:li);});
 function syncOrder(){document.getElementById('order').value=[...list.children].map(li=>li.dataset.id).join(',');}
 function caEsc(s){var e=document.createElement('div');e.textContent=(s==null?'':s);return e.innerHTML;}
 // The generation core, shared by Draft with AI (live) and Create with AI (disabled).
 // /create_ai reloads this build's deep research + stakeholder profile + FULL transcript
 // by build_id, so the draft is grounded in the real account context, not just the brief.
 function caGenerate(brief, ind, btn, gapIdx){
   window._caBrief=brief; window._caInd=ind; window._caGap=(gapIdx==null?null:gapIdx);
   btn.disabled=true; var old=btn.innerHTML; btn.innerHTML='<i class="ti ti-loader"></i> Writing…';
   var loader=document.getElementById('ca-loader');
   var bar=document.getElementById('ca-bar');
   var idle=document.getElementById('ca-idle'); if(idle) idle.style.display='none';
   if(loader) loader.style.display='block';
   var card=document.getElementById('ai-create'); if(card) card.scrollIntoView({behavior:'smooth',block:'center'});
   var pct=0; var ticker=setInterval(function(){pct=Math.min(pct+Math.random()*7+3,88);if(bar)bar.style.width=pct+'%';},500);
   var fd=new FormData(); fd.append('brief',brief);
   fd.append('build_id',BUILD_ID||'');   // server reloads full research/profile/transcript by this
   fd.append('industry',ind);
   fd.append('client_name',(SERVER_CTX&&SERVER_CTX.client_name)||'');
   fd.append('recipient',(SERVER_CTX&&SERVER_CTX.recipient)||'');
   fd.append('functions',(((SERVER_CTX&&SERVER_CTX.functions)||[]).join(', ')));
   fd.append('context',(((SERVER_CTX&&SERVER_CTX.transcript)||'')).slice(0,2000));
   fetch('/create_ai',{method:'POST',body:fd}).then(r=>r.json()).then(d=>{
     clearInterval(ticker); if(bar)bar.style.width='100%';
     setTimeout(function(){
       btn.disabled=false; btn.innerHTML=old;
       if(loader)loader.style.display='none'; if(bar)bar.style.width='0%';
       if(!d.ok){alert(d.error||'Could not generate');return;}
       window._ca=d; caRender(d);
     },350);
   }).catch(function(){clearInterval(ticker);btn.disabled=false;btn.innerHTML=old;if(loader)loader.style.display='none';alert('Could not generate');});
 }

 // "Draft with AI" — on a gap card. The gap already says what's missing, so nothing is
 // typed in by hand: the need's name/description/domain/use-case ARE the brief.
 // One preview at a time, so the single-slot preview machinery below stays honest.
 //
 // Bound from the gap card's data-* attributes (see bindDraftButtons), NOT an inline
 // onclick: the gap text is arbitrary prose, and passing it through an HTML attribute as
 // JS source meant one double quote in a description broke the whole handler.
 function draftAI(name,desc,domain,useCase,gapIdx,btn){
   if(window._ca && !confirm('You have a drafted slide that has not been added to the deck yet. Drafting this one will replace it. Continue?')) return;
   var brief=(name||'')+'. Problem: '+(desc||'');
   if(domain) brief+=' Domain: '+domain+'.';
   if(useCase) brief+=' Use case: '+useCase+'.';
   caGenerate(brief, (SERVER_CTX&&SERVER_CTX.industry)||'', btn, gapIdx);
 }

 function bindDraftButtons(){
   document.querySelectorAll('.draft-btn').forEach(function(b){
     b.addEventListener('click', function(){
       draftAI(b.dataset.name, b.dataset.desc, b.dataset.domain, b.dataset.usecase,
               parseInt(b.dataset.idx, 10), b);
     });
   });
 }

 // DISABLED with the manual form it reads (see the {# … #} block in build.html).
 // Kept intact so re-enabling is a one-line uncomment, not a rewrite.
 function createAI(btn){
   var topic=(document.getElementById('ca-topic')||{}).value||'';
   var problem=(document.getElementById('ca-problem')||{}).value||'';
   topic=topic.trim(); problem=problem.trim();
   if(!topic||!problem){alert('Please fill in Topic / use case and Problem — both are required.');return;}
   var solution=((document.getElementById('ca-solution')||{}).value||'').trim();
   var results=((document.getElementById('ca-results')||{}).value||'').trim();
   var ind=((document.getElementById('ca-industry')||{}).value||'').trim()||(SERVER_CTX&&SERVER_CTX.industry)||'';
   var brief=topic+'. Problem: '+problem;
   if(solution) brief+=' Solution: '+solution;
   if(results) brief+=' Results: '+results;
   caGenerate(brief, ind, btn, null);
 }
 // Re-draft the SAME brief (the manual form may not be on the page any more).
 function caRegen(btn){
   if(!window._caBrief) return;
   caGenerate(window._caBrief, window._caInd||'', btn||document.getElementById('ca-genbtn')||{disabled:false,innerHTML:''}, window._caGap);
 }
 function caRender(d){
   var caps=(d.capabilities||[]).map(c=>'<li>'+caEsc(c)+'</li>').join('');
   var res=(d.results||[]).map(c=>'<li>'+caEsc(c)+'</li>').join('');
   var rv=d.review||{}; var strong=((rv.quality||'')+'').toLowerCase().indexOf('strong')>-1;
   document.getElementById('ca-preview').innerHTML=
    '<div class="card" style="background:#f3f8f7;border:1px solid #bcdfd8">'+
    '<div style="font-weight:700;font-size:16px">'+caEsc(d.title)+'</div>'+
    '<div class="hint" style="font-size:12px;margin-bottom:8px">'+caEsc(d.subhead)+'</div>'+
    '<div style="font-size:13px"><b>Challenge:</b> '+caEsc(d.challenge)+'</div>'+
    '<div style="font-size:13px;margin-top:4px"><b>Solution:</b> '+caEsc(d.solution)+'</div>'+
    '<div style="display:flex;gap:18px;margin-top:8px;flex-wrap:wrap">'+
      '<div style="flex:1;min-width:220px"><b style="font-size:12px">Capabilities</b><ul style="margin:4px 0 0;padding-left:18px;font-size:12px">'+caps+'</ul></div>'+
      '<div style="flex:1;min-width:220px"><b style="font-size:12px">Results</b><ul style="margin:4px 0 0;padding-left:18px;font-size:12px">'+res+'</ul></div></div>'+
    '<div style="margin-top:10px;font-size:12px;padding:8px;border-radius:6px;background:'+(strong?'#e8f3f1':'#fdecea')+';color:'+(strong?'#1f5a52':'#8a2a1e')+'">'+
      '<b>Self-review &mdash; '+caEsc(rv.quality||'')+'.</b> Weakest: '+caEsc(rv.weakest||'None')+'. Fix: '+caEsc(rv.fix||'None')+'</div>'+
    '<div style="display:flex;gap:8px;margin-top:10px">'+
      '<button type="button" class="btn btn-primary" onclick="addCreated()"><i class="ti ti-plus"></i> Add to deck</button>'+
      '<button type="button" class="btn" onclick="caRegen(this)"><i class="ti ti-refresh"></i> Regenerate</button>'+
      '<button type="button" class="btn" onclick="caDiscard()"><i class="ti ti-x"></i> Discard</button></div></div>';
   document.getElementById('ca-preview').style.display='block';
 }
 function addCreated(){
   var d=window._ca; if(!d) return;
   if([...list.children].some(li=>li.dataset.id===d.id)){alert('Already added');return;}
   list.appendChild(makeRow(d.id, d.title, 'drafted with AI')); renumber();
   // tick the gap this draft answered, so a long gap list stays readable
   if(window._caGap!=null){
     var card=document.getElementById('gap-'+window._caGap);
     if(card){ var done=card.querySelector('.gap-done'); if(done) done.style.display='inline';
       var b=card.querySelector('button'); if(b){ b.disabled=true; b.innerHTML='<i class="ti ti-check"></i> Drafted'; } }
   }
   caDiscard(); ['ca-topic','ca-problem','ca-solution','ca-results'].forEach(function(id){var el=document.getElementById(id);if(el)el.value='';});
 }
 function caDiscard(){
   var p=document.getElementById('ca-preview');if(p)p.style.display='none';
   var idle=document.getElementById('ca-idle');if(idle)idle.style.display='';
   window._ca=null; window._caBrief=null; window._caGap=null;
 }
 (function init(){
   var d=loadDeck();
   if(RESUME){
     if(d&&d.order&&d.order.length){rebuild(d.order);fillHiddenCtx(d.ctx);
       var dc=document.getElementById('deck-client');
       if(dc&&d.ctx&&d.ctx.client_name) dc.textContent=d.ctx.client_name;}
     else{var emp=document.getElementById('resume-empty');if(emp)emp.style.display='';}
   } else if(d&&d.buildId===BUILD_ID&&d.order&&d.order.length){
     // Same build re-opened (e.g. browser Back) — restore it as left, keeping any library adds.
     rebuild(d.order);
   } // else: brand-new build — keep the server-rendered picks (persist() below seeds the deck).
   renumber();
   bindDraftButtons();
   var caInd=document.getElementById('ca-industry');
   if(caInd&&SERVER_CTX&&SERVER_CTX.industry) caInd.value=SERVER_CTX.industry;
 })();
 window.addEventListener('pageshow',function(){var l=document.getElementById('loader');if(l)l.style.display='none';});
