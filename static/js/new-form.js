// Industry "Other…" — reveal a free-text box; on submit, turn the typed text
// into a real <option> on the SAME select (name="industry") so the rest of the
// form (cfSubmit's industry lookup, the backend) needs no special-casing.
function indChanged(){
  var sel = document.getElementById('ind-select');
  var other = document.getElementById('ind-other');
  if(!sel || !other) return;
  var isOther = sel.value === '__OTHER__';
  other.style.display = isOther ? 'block' : 'none';
  if(isOther) other.focus();
}

(function(){
  var f = document.getElementById('deckForm');
  if(!f) return;
  f.addEventListener('submit', function(e){
    var sel = document.getElementById('ind-select');
    var other = document.getElementById('ind-other');
    if(sel && sel.value === '__OTHER__'){
      var val = other ? other.value.trim() : '';
      if(!val){
        e.preventDefault();
        if(other){ other.style.borderColor = '#c0392b'; other.focus(); }
        return;
      }
      var opt = document.createElement('option');
      opt.value = val; opt.text = val; opt.selected = true;
      sel.appendChild(opt);
    }
    var checked = f.querySelectorAll('input[name="work_types"]:checked').length;
    var warn = document.getElementById('wt-warn');
    if(checked === 0){
      e.preventDefault();
      if(warn) warn.style.display = 'block';
    } else {
      if(warn) warn.style.display = 'none';
      var ld = document.getElementById('loader');
      if(ld) ld.style.display = 'flex';
    }
  });
})();
// "Any function" = no specific function (optional). Highlighting it clears the
// individual function chips; ticking any chip clears the highlight.
function setAnyActive(on){
  var b=document.getElementById('fn-any'); if(!b) return;
  if(on){b.style.background='#2C6E66';b.style.color='#fff';b.style.borderColor='#2C6E66';b.dataset.on='1';}
  else{b.style.background='';b.style.color='';b.style.borderColor='';b.dataset.on='';}
}
function toggleAnyFn(){
  var on=document.getElementById('fn-any').dataset.on==='1';
  if(!on){document.querySelectorAll('#fn-chips input[name="functions"]').forEach(function(c){c.checked=false;});setAnyActive(true);}
  else{setAnyActive(false);}
}
function fnChanged(){
  var any=[...document.querySelectorAll('#fn-chips input[name="functions"]')].some(function(c){return c.checked;});
  if(any) setAnyActive(false);
}
// F1 content shortcut: build ONE slide at a time via /from_content (now JSON,
// not a page nav) and queue it -- lets the salesperson repeat this for several
// slides (e.g. handed one-by-one by the MS team) before either reviewing just
// the queue or continuing into the full deck form, which folds the queue in
// at the right spot server-side (see /build's content_slide_ids handling).
var CF_QUEUE_KEY = 'j2w_content_queue';
function cfLoadQueue(){ try{ return JSON.parse(localStorage.getItem(CF_QUEUE_KEY)) || []; }catch(e){ return []; } }
function cfSaveQueue(q){ localStorage.setItem(CF_QUEUE_KEY, JSON.stringify(q)); }
function cfRenderQueue(){
  var wrap=document.getElementById('cf-queue-wrap');
  var list=document.getElementById('cf-queue');
  if(!wrap||!list) return;
  var q=cfLoadQueue();
  wrap.style.display = q.length ? 'block' : 'none';
  list.innerHTML='';
  q.forEach(function(item){
    var li=document.createElement('li');
    li.className='slide-item';
    li.innerHTML='<span class="sid">'+ (item.template_label||'Slide') +'</span>'+
      '<div class="s-main"><div class="s-title"></div><div class="s-reason">queued — will be placed automatically when you build the deck</div></div>'+
      '<div class="s-actions"><button type="button" class="mini rm" aria-label="remove"><i class="ti ti-x"></i></button></div>';
    li.querySelector('.s-title').textContent = item.title || 'Untitled';
    li.querySelector('.rm').addEventListener('click', function(){ cfRemoveQueued(item.id); });
    list.appendChild(li);
  });
}
function cfRemoveQueued(id){
  cfSaveQueue(cfLoadQueue().filter(function(item){ return item.id!==id; }));
  cfRenderQueue();
}
function cfSubmit(e){
  if(e) e.preventDefault();
  var cn=document.querySelector('#deckForm [name="client_name"]');
  var ind=document.querySelector('#deckForm [name="industry"]');
  var cf=document.getElementById('cf-client'); if(cf&&cn) cf.value=cn.value;
  var ci=document.getElementById('cf-industry'); if(ci&&ind) ci.value=ind.value;
  var content=document.getElementById('cf-content');
  var file=document.getElementById('cf-file');
  var hasContent=(content&&content.value.trim()) || (file&&file.files&&file.files.length);
  var err=document.getElementById('cf-error');
  if(!hasContent){
    if(err){ err.textContent='Paste the case-study content or attach a document first.'; err.style.display='block'; }
    return false;
  }
  if(err) err.style.display='none';
  var btn=document.getElementById('cf-submit-btn');
  var old=btn?btn.innerHTML:'';
  if(btn){ btn.disabled=true; btn.innerHTML='<i class="ti ti-loader"></i> Building…'; }
  var form=document.getElementById('contentForm');
  var fd=new FormData(form);
  fetch('/from_content',{method:'POST',body:fd}).then(function(r){return r.json();}).then(function(d){
    if(btn){ btn.disabled=false; btn.innerHTML=old; }
    if(!d.ok){
      if(err){ err.textContent=d.error||'Could not build that slide.'; err.style.display='block'; }
      return;
    }
    var q=cfLoadQueue();
    q.push({id:d.id, title:d.title, content_type:d.content_type, template_label:d.template_label});
    cfSaveQueue(q);
    cfRenderQueue();
    if(content) content.value='';
    if(file) file.value='';
    var fname=document.getElementById('cf-fname');
    if(fname) fname.textContent='Attach a document (PDF, Word, text)';
  }).catch(function(){
    if(btn){ btn.disabled=false; btn.innerHTML=old; }
    if(err){ err.textContent='Could not reach the server.'; err.style.display='block'; }
  });
  return false;
}
(function(){
  cfRenderQueue();
  var deckForm=document.getElementById('deckForm');
  if(deckForm){
    deckForm.addEventListener('submit', function(e){
      if(e.defaultPrevented) return;   // an earlier validation listener cancelled the submit
      var ids=cfLoadQueue().map(function(item){ return item.id; });
      var f=document.getElementById('content-slide-ids');
      if(f) f.value=ids.join(',');
      cfSaveQueue([]);
    });
  }
  var reviewForm=document.getElementById('cf-review-form');
  if(reviewForm){
    reviewForm.addEventListener('submit', function(e){
      var q=cfLoadQueue();
      if(!q.length){ e.preventDefault(); return; }
      var order=q.map(function(item){ return 'NEW:'+item.id; }).join(',');
      document.getElementById('cf-review-order').value=order;
      var cn=document.querySelector('#deckForm [name="client_name"]');
      document.getElementById('cf-review-client').value=cn?cn.value.trim():'Client';
      cfSaveQueue([]);
    });
  }
})();
window.addEventListener('pageshow',function(){var l=document.getElementById('loader');if(l)l.style.display='none';});

// "Research this account" — a live executive-grade strategic brief, additive
// to (never replacing) the file uploads above. Populates one editable
// textarea for review before the form is ever submitted; nothing is applied
// silently. See /research_account (api.py) + deep_research.strategic_brief().
function researchAccount(){
  var cn=document.querySelector('#deckForm [name="client_name"]');
  var rc=document.querySelector('#deckForm [name="recipient"]');
  var ind=document.querySelector('#deckForm [name="industry"]');
  var btn=document.getElementById('research-btn');
  var err=document.getElementById('research-error');
  var results=document.getElementById('research-results');
  var clientName=cn?cn.value.trim():'';
  var recipient=rc?rc.value.trim():'';
  if(!clientName){
    if(err){err.textContent='Enter a client name first.';err.style.display='block';}
    return;
  }
  if(err) err.style.display='none';
  var old=btn.innerHTML;
  btn.disabled=true;
  btn.innerHTML='<i class="ti ti-loader"></i> Researching… (this can take a bit — it\'s a full strategic brief)';
  var fd=new FormData();
  fd.append('client_name',clientName);
  fd.append('recipient',recipient);
  fd.append('industry',ind?ind.value:'');
  var rf=document.querySelector('#deckForm [name="research_file"]');
  var pf=document.querySelector('#deckForm [name="profile_file"]');
  if(rf&&rf.files&&rf.files.length) fd.append('research_file',rf.files[0]);
  if(pf&&pf.files&&pf.files.length) fd.append('profile_file',pf.files[0]);
  fetch('/research_account',{method:'POST',body:fd}).then(function(r){return r.json();}).then(function(d){
    btn.disabled=false; btn.innerHTML=old;
    if(!d.ok){
      if(err){err.textContent=d.error||'Could not research this account.';err.style.display='block';}
      return;
    }
    document.getElementById('auto-company-text').value=d.brief||'';
    if(results) results.style.display='block';
  }).catch(function(){
    btn.disabled=false; btn.innerHTML=old;
    if(err){err.textContent='Could not reach the research service.';err.style.display='block';}
  });
}

// "Find logo" — best-effort auto-fetch by client name (+ optional domain).
// NOT reliable enough to apply silently (see client_logo.py for what was
// actually tested) -- shows a preview so you can see whether it actually
// found the right thing before it ever reaches a deck. A manual file upload
// always wins over this if both are present (see logoFileChosen below).
// Owner's spec, 2026-07-09: this should fire AUTOMATICALLY from the client
// name already entered earlier in the form, not require a separate manual
// click as the primary path -- see maybeAutoFindLogo below, wired to the
// client-name field's blur. The button/domain field remain as a manual
// retry (e.g. the auto-search picked the wrong company, or found nothing).
var _logoLastSearched = null;
function findLogo(){
  var cn=document.querySelector('#deckForm [name="client_name"]');
  var dom=document.getElementById('logo-domain');
  var btn=document.getElementById('logo-find-btn');
  var err=document.getElementById('logo-error');
  var box=document.getElementById('logo-preview-box');
  var clientName=cn?cn.value.trim():'';
  if(!clientName){
    if(err){err.textContent='Enter a client name first.';err.style.display='block';}
    return;
  }
  _logoLastSearched = clientName;
  if(err) err.style.display='none';
  var old=btn.innerHTML;
  btn.disabled=true;
  btn.innerHTML='<i class="ti ti-loader"></i> Searching…';
  var fd=new FormData();
  fd.append('client_name',clientName);
  fd.append('domain',dom?dom.value.trim():'');
  fetch('/logo_preview',{method:'POST',body:fd}).then(function(r){return r.json();}).then(function(d){
    btn.disabled=false; btn.innerHTML=old;
    if(!d.ok){
      if(box) box.style.display='none';
      document.getElementById('client-logo-data-uri').value='';
      if(err){err.textContent=d.error||'Could not find a logo automatically.';err.style.display='block';}
      return;
    }
    document.getElementById('logo-preview-img').src=d.data_uri;
    document.getElementById('client-logo-data-uri').value=d.data_uri;
    if(box) box.style.display='block';
  }).catch(function(){
    btn.disabled=false; btn.innerHTML=old;
    if(err){err.textContent='Could not reach the logo search service.';err.style.display='block';}
  });
}
// Fires on the client-name field losing focus -- auto-finds the logo from
// the name alone, no button click needed. Skipped if a file was already
// manually chosen (that always wins) or this exact name was already
// searched (avoids re-firing on every blur when nothing changed).
function maybeAutoFindLogo(){
  var cn=document.querySelector('#deckForm [name="client_name"]');
  var fileInput=document.querySelector('#deckForm [name="client_logo_file"]');
  if(!cn) return;
  var name=cn.value.trim();
  if(!name || name===_logoLastSearched) return;
  if(fileInput && fileInput.files && fileInput.files.length) return;
  findLogo();
}
// If a file is chosen after an auto-found logo was previewed, the file wins —
// clear the auto-found state so /build doesn't get a mixed signal, and stop
// the automatic search from later overwriting this deliberate choice.
function logoFileChosen(input){
  if(!input.files || !input.files.length) return;
  document.getElementById('client-logo-data-uri').value='';
  var box=document.getElementById('logo-preview-box');
  if(box) box.style.display='none';
  var cn=document.querySelector('#deckForm [name="client_name"]');
  _logoLastSearched = cn ? cn.value.trim() : null;
}
(function(){
  var cn=document.querySelector('#deckForm [name="client_name"]');
  if(cn) cn.addEventListener('blur', maybeAutoFindLogo);
})();
