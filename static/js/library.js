 function updatePill(){var d=loadDeck();var n=(d&&d.order)?d.order.length:0;
   document.getElementById('deck-pill-n').textContent=n;
   document.getElementById('deck-pill').style.display=n?'inline-flex':'none';}
 function libAdd(id,btn){var d=loadDeck()||{active:true,ctx:{client_name:'',industry:'',transcript:'',phase:'',recipient:'',functions:[],work_types:[]},order:[]};
   d.active=true;if(!d.order)d.order=[];
   if(d.order.indexOf(id)===-1){d.order.push(id);saveDeck(d);}
   if(btn){btn.innerHTML='<i class="ti ti-check"></i>';btn.style.color='#2C6E66';}
   updatePill();}
 (function(){var d=loadDeck();var inDeck=(d&&d.order)?d.order:[];
   document.querySelectorAll('.deck-add').forEach(b=>{if(inDeck.indexOf(b.dataset.id)>-1){b.innerHTML='<i class="ti ti-check"></i>';b.style.color='#2C6E66';}});
   updatePill();})();
 const cards=[...document.querySelectorAll('.lib-card')];
 const f={wt:'',kind:'',ind:'',q:''};
 function apply(){let n=0;
   cards.forEach(c=>{const ok=(!f.wt||c.dataset.wt===f.wt)&&(!f.kind||c.dataset.kind===f.kind)
     &&(!f.ind||c.dataset.ind===f.ind)&&(!f.q||c.dataset.text.indexOf(f.q)>-1);
     c.style.display=ok?'':'none';if(ok)n++;});
   document.getElementById('lib-count').textContent=n;
   document.getElementById('lib-empty').style.display=n?'none':'';}
 document.getElementById('lib-search').addEventListener('input',e=>{f.q=e.target.value.toLowerCase();apply();});
 document.getElementById('ind-sel').addEventListener('change',e=>{f.ind=e.target.value;apply();});
 function wire(attr,key){document.querySelectorAll('[data-'+attr+']').forEach(b=>b.addEventListener('click',()=>{
   f[key]=b.getAttribute('data-'+attr);
   b.parentNode.querySelectorAll('.fchip').forEach(x=>x.classList.remove('active'));
   b.classList.add('active');apply();}));}
 wire('wtf','wt');wire('kindf','kind');
 // ── bulk download: tick slides, download all selected as one .zip ──
 (function(){
   const selected=new Set();
   const bar=document.getElementById('sel-bar');
   const nEl=document.getElementById('sel-n');
   function refresh(){nEl.textContent=selected.size;bar.style.display=selected.size?'flex':'none';}
   document.querySelectorAll('.lib-sel').forEach(cb=>cb.addEventListener('change',()=>{
     if(cb.checked)selected.add(cb.value);else selected.delete(cb.value);
     const card=cb.closest('.lib-card');if(card)card.classList.toggle('sel',cb.checked);
     refresh();}));
   document.getElementById('sel-dl').addEventListener('click',()=>{
     if(!selected.size)return;
     window.location='/slides/download?ids='+[...selected].map(encodeURIComponent).join(',');});
   document.getElementById('sel-clear').addEventListener('click',()=>{
     selected.clear();
     document.querySelectorAll('.lib-sel').forEach(cb=>{cb.checked=false;const c=cb.closest('.lib-card');if(c)c.classList.remove('sel');});
     refresh();});
 })();
