// toast.js -- the app's one notification primitive: a small card, top right.
//
// Loaded on every page from _shell.html, so any script can call j2wToast(...).
// Built for the New-deck form's "you missed a field" feedback: the browser's own
// validation bubble shows ONE field at a time, can't be styled to match the app, and
// has nothing to attach to for the work-type checkboxes (no `required` attribute is
// possible on "at least one of these"). See static/js/new-form.js.
(function (w) {
  'use strict';

  var TTL = 4000;          // ms on screen before it fades itself out
  var MAX = 4;             // never stack more than this; oldest goes first

  function wrap() {
    var el = document.getElementById('j2w-toasts');
    if (!el) {
      el = document.createElement('div');
      el.id = 'j2w-toasts';
      el.className = 'toast-wrap';
      // polite, not assertive: a missed form field is not an emergency, and
      // assertive would interrupt a screen reader mid-sentence for each one.
      el.setAttribute('role', 'status');
      el.setAttribute('aria-live', 'polite');
      document.body.appendChild(el);
    }
    return el;
  }

  function dismiss(el) {
    if (!el || el.classList.contains('is-out')) return;
    el.classList.add('is-out');
    // remove after the fade; the animationend may never fire under
    // prefers-reduced-motion (the animation is disabled), so time it out too.
    var done = function () { if (el.parentNode) el.parentNode.removeChild(el); };
    el.addEventListener('animationend', done);
    setTimeout(done, 400);
  }

  /**
   * j2wToast("Please fill up the Client name field", "error")
   * kind: "error" (red rule) or anything else (teal rule). Returns the element.
   */
  function toast(message, kind) {
    if (!message) return null;
    var box = wrap();
    while (box.children.length >= MAX) dismiss(box.firstChild);

    var el = document.createElement('div');
    el.className = 'toast' + (kind === 'error' ? ' is-error' : '');

    var ic = document.createElement('i');
    ic.className = 'ti ' + (kind === 'error' ? 'ti-alert-circle' : 'ti-circle-check');
    ic.setAttribute('aria-hidden', 'true');

    var msg = document.createElement('div');
    msg.className = 'toast-msg';
    msg.textContent = message;          // textContent, never innerHTML -- the caller
                                        // may be passing a field label from the DOM
    el.appendChild(ic);
    el.appendChild(msg);
    el.addEventListener('click', function () { dismiss(el); });
    box.appendChild(el);
    setTimeout(function () { dismiss(el); }, TTL);
    return el;
  }

  w.j2wToast = toast;
})(window);
