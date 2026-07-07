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
// F1 content shortcut: reuse the client/industry typed in the main form, validate, show loader.
function cfSubmit(){
  var cn=document.querySelector('#deckForm [name="client_name"]');
  var ind=document.querySelector('#deckForm [name="industry"]');
  var cf=document.getElementById('cf-client'); if(cf&&cn) cf.value=cn.value;
  var ci=document.getElementById('cf-industry'); if(ci&&ind) ci.value=ind.value;
  var content=document.querySelector('#contentForm [name="content"]');
  var file=document.querySelector('#contentForm [name="content_file"]');
  var hasContent=(content&&content.value.trim()) || (file&&file.files&&file.files.length);
  if(!hasContent){ alert('Paste the case-study content or attach a document first.'); return false; }
  var loader=document.getElementById('loader');
  if(loader){ var t=loader.querySelector('.loader-title'); if(t) t.textContent='Building your slide…'; loader.style.display='flex'; }
  return true;
}
window.addEventListener('pageshow',function(){var l=document.getElementById('loader');if(l)l.style.display='none';});
