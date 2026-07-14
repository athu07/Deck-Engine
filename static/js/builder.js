// builder.js -- the Custom Slide Builder (/builder).
//
// PASTE THE WHOLE DOCUMENT (owner's spec, 2026-07-10). The MS team hands over a deck's
// worth of content in one file, marked up "slide 1 - case study". Splitting that into
// boxes by hand is busywork the app should do.
//
// Three steps, and the middle one is the point:
//
//   1. paste            -> POST /builder/parse
//   2. REVIEW THE SPLIT -> nothing has been generated yet. Fix any shape, drop a
//                          duplicate, or go back and fix the text. Free.
//   3. build            -> POST /builder/slide, once per slide, four at a time, each
//                          card filling in as its render lands. Ten slides is ~40s;
//                          watching them arrive beats watching a spinner.
//
// The queue lives in the SAME localStorage key the new-deck form reads
// ('j2w_content_queue'), so "Use in a deck" needs no server-side session.

var CB_QUEUE_KEY = 'j2w_content_queue';
var CB_CONCURRENCY = 4;              // the server renders with LibreOffice; don't stampede it
var CB_PARSED = [];                  // the reviewed slides, in order

function cbLoadQueue(){ try { return JSON.parse(localStorage.getItem(CB_QUEUE_KEY)) || []; } catch(e){ return []; } }
function cbSaveQueue(q){ localStorage.setItem(CB_QUEUE_KEY, JSON.stringify(q)); }

function cbEsc(s){ var d=document.createElement('div'); d.textContent=(s===null||s===undefined)?'':String(s); return d.innerHTML; }
function cbEl(id){ return document.getElementById(id); }

function cbError(msg){
  var e=cbEl('cb-error');
  if(!e) return;
  if(!msg){ e.style.display='none'; return; }
  e.textContent=msg; e.style.display='block';
  e.scrollIntoView({behavior:'smooth', block:'center'});
}

function cbLoader(on, title, sub){
  var l=cbEl('loader'); if(!l) return;
  if(title) cbEl('cb-loader-title').textContent=title;
  if(sub) cbEl('cb-loader-sub').textContent=sub;
  l.style.display = on ? 'flex' : 'none';
}

function cbSettings(){
  return {work_type:(cbEl('cb-worktype')||{}).value||'',
          industry:(cbEl('cb-industry')||{}).value||'',
          client_name:(cbEl('cb-client')||{}).value||''};
}

// ── step 1: parse. Pure text + one embedding call. Nothing is built. ──────────
function cbParse(){
  cbError('');
  var s=cbSettings();
  if(!s.work_type){ cbError("Choose a work type — it's needed to save these to the library."); return; }
  var c=cbEl('cb-content'), f=cbEl('cb-file');
  if(!((c && c.value.trim()) || (f && f.files && f.files.length))){
    cbError('Paste the content, or attach the document.'); return;
  }
  cbLoader(true, 'Reading your slides…',
           "Splitting the content, working out what each slide is, and checking the library.");
  fetch('/builder/parse', {method:'POST', body:new FormData(cbEl('cbForm'))})
    .then(function(r){ return r.json(); })
    .then(function(d){
      cbLoader(false);
      if(!d.ok){ cbError(d.error || 'Could not read that content.'); return; }
      CB_PARSED = d.slides.map(function(s){
        return {number:s.number, label:s.label, heading:s.heading||'', template:s.template,
                matched:s.matched, inferred:!!s.inferred, content:s.content,
                preview:s.preview, matches:s.matches||[], reuse:null};
      });
      cbRenderReview(d.templates);
    })
    .catch(function(){ cbLoader(false); cbError('Could not reach the server.'); });
}

// ── step 2: the review screen. This is where a bad split gets caught, for free. ──
function cbRenderReview(templates){
  var n=CB_PARSED.length;
  cbEl('cb-review-head').textContent = 'We read ' + n + ' slide' + (n===1?'':'s');
  var opts=function(sel){
    return templates.map(function(t){
      return '<option value="'+cbEsc(t.key)+'"'+(t.key===sel?' selected':'')+'>'+cbEsc(t.label)+'</option>';
    }).join('');
  };

  var list=cbEl('cb-review-list');
  list.innerHTML='';
  CB_PARSED.forEach(function(s, i){
    var dup = s.matches.length ? s.matches[0] : null;
    var row=document.createElement('div');
    row.className='cb-rev-row';
    row.style.cssText='border:1px solid var(--line-2);border-radius:8px;padding:12px;margin-bottom:8px';
    row.innerHTML=
      '<div style="display:flex;gap:12px;align-items:flex-start">'+
        '<span class="sid" style="flex:0 0 auto">'+(i+1)+'</span>'+
        '<div style="flex:1;min-width:0">'+
          (s.heading ? '<div style="font-weight:600;font-size:14px;margin-bottom:6px">'+cbEsc(s.heading)+'</div>' : '')+
          '<div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">'+
            '<select class="input cb-tpl" style="max-width:230px;font-size:13px">'+opts(s.template)+'</select>'+
            (s.matched
              ? '<span class="hint" style="font-size:12px"><i class="ti ti-check"></i> you named this shape</span>'
              : '<span class="hint" style="font-size:12px;color:#2C6E66"><i class="ti ti-sparkles"></i> we read this from the content — change it if it\'s wrong</span>') +
          '</div>'+
          '<div class="hint" style="font-size:12.5px;margin-top:6px">'+cbEsc(s.preview)+'</div>'+
        '</div>'+
        '<button type="button" class="mini rm" title="Drop this slide"><i class="ti ti-x"></i></button>'+
      '</div>'+
      (dup ?
        '<div class="cb-dup" style="margin-top:10px;background:#fff8f0;border:1px solid #f0d9bf;border-radius:8px;padding:10px 12px">'+
          '<div style="font-size:13px"><i class="ti ti-copy" style="color:#c47d27"></i> '+
            '<b>'+cbEsc(dup.id)+'</b> — '+cbEsc(dup.title)+' is <b>'+dup.percent+'% similar</b>. '+
            'Reuse it and we won\'t build this one at all.</div>'+
          '<div style="display:flex;gap:6px;margin-top:8px">'+
            '<button type="button" class="btn cb-dup-preview"><i class="ti ti-eye"></i> Preview '+cbEsc(dup.id)+'</button>'+
            '<button type="button" class="btn cb-dup-use"><i class="ti ti-exchange"></i> Use '+cbEsc(dup.id)+' instead</button>'+
          '</div></div>'
        : '');

    row.querySelector('.cb-tpl').addEventListener('change', function(){ s.template=this.value; });
    row.querySelector('.rm').addEventListener('click', function(){
      CB_PARSED.splice(CB_PARSED.indexOf(s), 1);
      if(!CB_PARSED.length){ cbEl('cb-review').style.display='none'; return; }
      cbRenderReview(templates);
    });
    if(dup){
      row.querySelector('.cb-dup-preview').addEventListener('click', function(){ cbPreviewCase(dup.id); });
      row.querySelector('.cb-dup-use').addEventListener('click', function(){
        s.reuse = {id:dup.id, title:dup.title};
        cbRenderReview(templates);
      });
    }
    if(s.reuse){
      row.style.background='#f3f8f7'; row.style.borderColor='#bcdfd8';
      row.innerHTML='<div style="display:flex;justify-content:space-between;align-items:center;gap:10px">'+
        '<div><b>'+cbEsc(s.reuse.id)+' · '+cbEsc(s.reuse.title)+'</b>'+
          '<div class="hint" style="font-size:12.5px"><i class="ti ti-check" style="color:#2C6E66"></i> '+
          'Reusing the library slide — nothing will be generated for slide '+(i+1)+'.</div></div>'+
        '<button type="button" class="btn cb-undo"><i class="ti ti-arrow-back-up"></i> Build mine instead</button></div>';
      row.querySelector('.cb-undo').addEventListener('click', function(){
        s.reuse=null; cbRenderReview(templates);
      });
    }
    list.appendChild(row);
  });

  var toBuild = CB_PARSED.filter(function(s){ return !s.reuse; }).length;
  var reused = CB_PARSED.length - toBuild;
  cbEl('cb-build-label').textContent = toBuild
    ? 'Build ' + (toBuild===CB_PARSED.length ? ('these ' + toBuild + ' slide' + (toBuild===1?'':'s'))
                                             : (toBuild + ' slide' + (toBuild===1?'':'s')))
    : 'Add the reused slides';
  cbEl('cb-build-note').textContent = reused
    ? reused + ' slide' + (reused===1?'':'s') + ' reused from the library, ' + toBuild + ' to build.'
    : (toBuild > 3 ? 'This takes about ' + Math.ceil(toBuild * 4) + ' seconds. They appear as they finish.' : '');
  cbEl('cb-review').style.display='block';
  cbEl('cb-review').scrollIntoView({behavior:'smooth', block:'start'});
}

function cbEditText(){
  cbEl('cb-review').style.display='none';
  cbEl('cb-content').scrollIntoView({behavior:'smooth', block:'center'});
  cbEl('cb-content').focus();
}

// ── step 3: build. One request per slide, CB_CONCURRENCY at a time, each card
//    filling in the moment its render lands rather than all at the end. ────────
function cbBuildParsed(){
  cbError('');
  if(!CB_PARSED.length) return;
  var s=cbSettings();
  cbEl('cb-build-btn').disabled=true;

  // reused slides need no build -- they go straight into the queue
  var q=cbLoadQueue();
  CB_PARSED.forEach(function(sl){
    if(sl.reuse && !q.some(function(it){ return it.id===sl.reuse.id; })){
      q.push({id:sl.reuse.id, title:sl.reuse.title, content_type:'case_study',
              template_label:'Case study', reused:true});
    }
  });
  cbSaveQueue(q); cbRenderQueue();

  var jobs = CB_PARSED.filter(function(sl){ return !sl.reuse; });
  if(!jobs.length){ cbEl('cb-review').style.display='none'; cbEl('cb-build-btn').disabled=false; return; }

  cbPrepareResults(jobs);
  // Results land in THIS array by their original paste position, not by finish order --
  // pushed to the queue only once every job is done (owner-reported, 2026-07-14: building
  // 4-at-a-time concurrently used to push each slide into the queue the moment IT finished,
  // so a fast slide could land ahead of a slower one -- a 3-slide paste came out 3rd-1st-2nd
  // instead of 1st-2nd-3rd). The visible cards still fill in as each one lands (cbFillResult
  // below), so progress still feels live -- only the final queue order is now deferred.
  var results = new Array(jobs.length);
  var done=0, next=0;
  var finish=function(){
    done++;
    cbEl('cb-progress').textContent = done + ' of ' + jobs.length + ' built';
    if(done===jobs.length){
      var q2 = cbLoadQueue();
      results.forEach(function(d){
        if(d && d.ok) q2.push({id:d.id, title:d.title, content_type:d.content_type,
                               template_label:d.template_label});
      });
      cbSaveQueue(q2); cbRenderQueue();
      cbEl('cb-build-btn').disabled=false;
      cbEl('cb-review').style.display='none';
    }
  };
  var pump=function(){
    if(next>=jobs.length) return;
    var i=next++, job=jobs[i];
    var fd=new FormData();
    fd.append('content', job.content);
    fd.append('template_hint', job.template);
    fd.append('work_type', s.work_type);
    fd.append('industry', s.industry);
    fd.append('client_name', s.client_name);
    fetch('/builder/slide', {method:'POST', body:fd})
      .then(function(r){ return r.json(); })
      .then(function(d){ results[i]=d; cbFillResult(i, d); })
      .catch(function(){ var d={ok:false, error:'Could not reach the server.'}; results[i]=d; cbFillResult(i, d); })
      .then(function(){ finish(); pump(); });
  };
  for(var k=0; k<Math.min(CB_CONCURRENCY, jobs.length); k++) pump();
}

// placeholder cards, so the slides appear in ORDER as they land, not in finish order
function cbPrepareResults(jobs){
  var list=cbEl('cb-result-list');
  list.innerHTML='';
  jobs.forEach(function(job, i){
    var card=document.createElement('div');
    card.id='cb-res-'+i;
    card.className='rev-slide-card';
    card.innerHTML='<div class="rev-slide-index">'+(i+1)+'</div>'+
      '<div class="hint" style="padding:22px 24px;font-size:13px"><i class="ti ti-loader"></i> building…</div>';
    list.appendChild(card);
  });
  cbEl('cb-progress').textContent='0 of '+jobs.length+' built';
  cbEl('cb-results').style.display='block';
  cbEl('cb-results').scrollIntoView({behavior:'smooth', block:'start'});
}

// Each built slide becomes the SAME editable slide card the review page shows: you click
// straight onto its title, paragraph or capability. There is no separate "Edit" step and
// no rendered image -- the card IS the slide, and it is editable the moment it lands.
function cbFillResult(i, d){
  var card=cbEl('cb-res-'+i);
  if(!card) return;
  if(!d.ok){
    card.style.borderColor='#f2d2d2'; card.style.background='#fdf2f2';
    card.innerHTML='<div class="rev-slide-index">'+(i+1)+'</div>'+
      '<div style="padding:20px 24px"><b>Couldn\'t build this slide</b>'+
      '<div class="hint" style="font-size:13px">'+cbEsc(d.error||'')+'</div></div>';
    return;
  }
  card.className='rev-slide-card '+d.content_type;
  card.dataset.slideId=d.id;
  card.dataset.contentType=d.content_type;
  card.innerHTML='<div class="rev-slide-index">'+(i+1)+'</div>'+
    '<button type="button" class="mini rm cb-drop" title="Remove this slide" '+
      'style="position:absolute;top:10px;left:12px;z-index:3"><i class="ti ti-x"></i></button>'+
    d.html;
  card.querySelector('.cb-drop').addEventListener('click', function(){
    cbSaveQueue(cbLoadQueue().filter(function(it){ return it.id!==d.id; }));
    cbRenderQueue();
    card.remove();
  });
  // auto-grow the editable paragraphs, exactly as the review page does
  card.querySelectorAll('textarea.se').forEach(function(t){
    function grow(){ t.style.height='auto'; t.style.height=(t.scrollHeight+2)+'px'; }
    t.addEventListener('input', grow); grow();
  });
  cbMarkDirty(false);
  card.addEventListener('input', function(){ cbMarkDirty(true); });
}

// ── reading the edits back ───────────────────────────────────────────────────
// Every field carries a data-path naming its place in the record ("title", "results.0",
// "columns.0.items.2"), so ONE serialiser reads any of the eight shapes.
function cbSetPath(obj, path, value){
  var parts=path.split('.'), cur=obj;
  for(var i=0;i<parts.length-1;i++){
    var key=parts[i], nextIsIndex=/^\d+$/.test(parts[i+1]);
    if(/^\d+$/.test(key)) key=parseInt(key,10);
    if(cur[key]===undefined) cur[key] = nextIsIndex ? [] : {};
    cur=cur[key];
  }
  var last=parts[parts.length-1];
  cur[/^\d+$/.test(last)?parseInt(last,10):last] = value;
}

function cbSerializeCard(card){
  var out={};
  card.querySelectorAll('[data-path]').forEach(function(el){
    cbSetPath(out, el.dataset.path, (el.value||'').trim());
  });
  out = cbClean(out);
  // capabilities are stored as "Name: what it delivers" strings, but edited as two fields
  // (as the review page has always shown them) -- re-join before saving
  if(Array.isArray(out.capabilities) && out.capabilities.length && typeof out.capabilities[0]==='object'){
    out.capabilities = out.capabilities.map(function(c){
      var t=(c&&c.title||'').trim(), b=(c&&c.body||'').trim();
      if(!t && !b) return '';
      return b ? (t + ': ' + b) : t;
    }).filter(Boolean);
  }
  return out;
}

function cbClean(v){
  if(Array.isArray(v)){
    return v.filter(function(x){ return x!==undefined && x!==null; })
            .map(cbClean)
            .filter(function(x){
              if(typeof x==='string') return x.length>0;
              if(Array.isArray(x)) return x.length>0;
              if(x && typeof x==='object') return Object.keys(x).some(function(k){
                var y=x[k]; return y && y.length; });
              return true;
            });
  }
  if(v && typeof v==='object'){
    var o={};
    Object.keys(v).forEach(function(k){ o[k]=cbClean(v[k]); });
    return o;
  }
  return v;
}

var CB_DIRTY=false;
function cbMarkDirty(on){
  CB_DIRTY=on;
  var s=cbEl('cb-dirty');
  if(s) s.style.display = on ? 'inline' : 'none';
}

// Save every card's edits before the slides leave this page. The staged record is the
// single source of truth, so once these land the edit follows the slide into the .pptx,
// into a generated deck, and into the shared library.
function cbSaveAll(){
  var cards=[].slice.call(document.querySelectorAll('#cb-result-list [data-slide-id]'));
  if(!cards.length) return Promise.resolve();
  var failed=[];
  return Promise.all(cards.map(function(card){
    var sid=card.dataset.slideId;
    return fetch('/builder/slide/'+encodeURIComponent(sid), {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify(cbSerializeCard(card))
    }).then(function(r){
      if(!r.ok){ failed.push(sid); return null; }
      return r.json();
    }).then(function(r){
      if(r && r.ok){
        var q=cbLoadQueue();
        q.forEach(function(it){ if(it.id===r.id) it.title=r.title; });
        cbSaveQueue(q);
      }
    }).catch(function(){ failed.push(sid); });
  })).then(function(){
    cbRenderQueue();
    // A failed save must NOT fall through to the download -- that would hand back a
    // .pptx built from the UN-edited slides, silently. Surface it and stop.
    if(failed.length) throw new Error('Could not save ' + failed.length + ' slide(s). Nothing was downloaded.');
    cbMarkDirty(false);
  });
}

// preview an EXISTING library slide, inline under its review row
function cbPreviewCase(id){
  cbLoader(true, 'Rendering that slide…', 'Fetching it from the library.');
  fetch('/builder/preview/'+encodeURIComponent(id))
    .then(function(r){ return r.json(); })
    .then(function(d){
      cbLoader(false);
      if(!d.ok || !d.png){ cbError('Could not render that slide.'); return; }
      var w=window.open('', '_blank');
      if(w) w.document.write('<title>'+cbEsc(d.id)+'</title><img src="'+cbEsc(d.png)+'" style="width:100%">');
      else cbError('Allow pop-ups to preview the library slide.');
    })
    .catch(function(){ cbLoader(false); cbError('Could not render that slide.'); });
}

// ── the queue (drives the sidebar counts and the download) ────────────────────
// It SURVIVES a reload -- that's what makes "Use in a deck" work. So say plainly how
// many slides are in it: a queue left over from an earlier visit would otherwise ride
// along silently in the next download.
function cbRenderQueue(){
  var q=cbLoadQueue();
  cbEl('cb-sum-count').textContent = q.length;
  cbEl('cb-sum-cases').textContent = q.filter(function(it){ return it.content_type==='case_study'; }).length;
  cbEl('cb-download').disabled = !q.length;
  cbEl('cb-clear').style.display = q.length ? 'block' : 'none';
}

function cbClear(){
  if(!confirm('Remove all queued slides? They stay in the AI history, but leave your deck.')) return;
  cbSaveQueue([]);
  cbRenderQueue();
  cbEl('cb-results').style.display='none';
  cbEl('cb-review').style.display='none';
}

// ── download: the queue as one .pptx, and a commit point for the library ──────
function cbDownload(){
  var q=cbLoadQueue();
  if(!q.length) return;
  cbLoader(true, 'Saving your edits…', 'Then building the .pptx.');
  cbSaveAll().then(function(){
    cbLoader(false);
    var s=cbSettings();
    var params=new URLSearchParams({
      ids: q.map(function(it){ return it.id; }).join(','),
      industry: s.industry, client_name: s.client_name
    });
    window.location = '/builder/download?' + params.toString();
  }).catch(function(err){ cbLoader(false); cbError(err && err.message || 'Could not save your edits.'); });
}

function cbUseInDeck(e){
  if(e) e.preventDefault();
  cbLoader(true, 'Saving your edits…', 'Then taking these slides to the deck builder.');
  cbSaveAll().then(function(){ window.location = '/new'; })
             .catch(function(err){ cbLoader(false); cbError(err && err.message || 'Could not save your edits.'); });
  return false;
}

(function(){
  cbRenderQueue();
  var u=cbEl('cb-usedeck');
  if(u) u.addEventListener('click', cbUseInDeck);
  window.addEventListener('pageshow', function(){ cbLoader(false); });
  window.addEventListener('beforeunload', function(e){
    if(CB_DIRTY){ e.preventDefault(); e.returnValue=''; }   // don't lose typed edits
  });
})();
