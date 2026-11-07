/*
 * Horizonte Travel – multi-step access control.
 *
 * The site unlocks content in levels. Each form the visitor completes raises
 * their level by one:
 *
 *   Level 0  Visitor          -> Home only (company info, why travel, instructions)
 *   Level 1  Account created  -> register.html      (Step 1, includes CAPTCHA)
 *   Level 2  Email verified   -> verify.html        (Step 2)
 *   Level 3  Profile complete -> preferences.html   (Step 3)  => unlocks destinations.html
 *   Level 4  Full access      -> application.html   (Step 4, includes CAPTCHA)
 *                                                    => unlocks resources.html + prices
 *
 * Progress is stored in the visitor's browser (localStorage). This is a static
 * site (GitHub Pages), so this is a "soft" gate: a technical user could skip it.
 *
 * To RECEIVE the answers, set a form service endpoint here, e.g.:
 *   Formspree:  'https://formspree.io/f/YOUR_ID'
 *   FormSubmit: 'https://formsubmit.co/ajax/you@example.com'
 * If left empty, answers are only stored in the visitor's browser.
 */
var FORM_ENDPOINT = '';

var Access = (function () {
  var KEY = 'horizonte-travel-progress';

  // Each step: the page that completes it and a short label.
  var STEPS = [
    null,
    { page: 'register.html',    title: 'Create your account' },
    { page: 'verify.html',      title: 'Verify your email' },
    { page: 'preferences.html', title: 'Build your travel profile' },
    { page: 'application.html', title: 'Submit your traveler application' }
  ];

  function read() {
    try { return JSON.parse(localStorage.getItem(KEY) || 'null') || { level: 0 }; } catch (e) {}
    try { return JSON.parse(sessionStorage.getItem(KEY) || 'null') || { level: 0 }; } catch (e) {}
    return { level: 0 };
  }

  function write(data) {
    var text = JSON.stringify(data);
    try { localStorage.setItem(KEY, text); return; } catch (e) {}
    try { sessionStorage.setItem(KEY, text); } catch (e) {}
  }

  // Merge new data into the stored progress.
  function save(patch) {
    var data = read();
    Object.keys(patch).forEach(function (k) { data[k] = patch[k]; });
    write(data);
    return data;
  }

  function level() { return read().level || 0; }

  // Relative path to the site root from the current page (root or subfolder).
  function base() {
    var s = document.querySelector('script[src$="access.js"]');
    return s ? s.getAttribute('src').replace(/access\.js$/, '') : '';
  }

  // Page the visitor must complete next to reach `needed`.
  function nextStepPage() {
    var l = Math.min(level() + 1, 4);
    return base() + STEPS[l].page;
  }

  /*
   * Page guard. Call in <head>: Access.require(3) sends anyone below level 3
   * to the step they are missing.
   */
  function require(needed) {
    if (level() >= needed) return true;
    location.replace(nextStepPage() + '?locked=' + needed);
    return false;
  }

  /*
   * Fake "processing" screen with delays. Shows a full-screen overlay with a
   * spinner, a progress bar and rotating messages, then resolves.
   *   Access.processing(['Checking…', 'Saving…'], 4000).then(...)
   */
  function processing(messages, totalMs) {
    return new Promise(function (resolve) {
      var overlay = document.createElement('div');
      overlay.className = 'overlay';
      overlay.setAttribute('role', 'status');
      overlay.setAttribute('aria-live', 'polite');
      overlay.innerHTML =
        '<div class="overlay-card"><div class="spinner"></div>' +
        '<p class="overlay-msg"></p><div class="barra"><span></span></div></div>';
      document.body.appendChild(overlay);
      var msg = overlay.querySelector('.overlay-msg');
      var bar = overlay.querySelector('.barra span');
      var each = totalMs / messages.length;
      messages.forEach(function (m, i) {
        setTimeout(function () {
          msg.textContent = m;
          bar.style.width = Math.round(((i + 1) / messages.length) * 100) + '%';
        }, i * each);
      });
      setTimeout(function () { overlay.remove(); resolve(); }, totalMs + 400);
    });
  }

  function send(subject, data) {
    if (!FORM_ENDPOINT) return Promise.resolve(false);
    var body = new FormData();
    Object.keys(data).forEach(function (k) {
      var v = data[k];
      body.append(k, typeof v === 'object' ? JSON.stringify(v) : v);
    });
    body.append('_subject', subject + ' – Horizonte Travel');
    return fetch(FORM_ENDPOINT, { method: 'POST', body: body, headers: { Accept: 'application/json' } })
      .then(function (r) { return r.ok; }).catch(function () { return false; });
  }

  /*
   * Navigation and "locked" links. Any element with data-level="N" shows a
   * lock while the visitor is below level N, and its link points to the next
   * step they need instead of the locked page.
   */
  function decorate() {
    var l = level();
    document.querySelectorAll('[data-level]').forEach(function (a) {
      var need = +a.getAttribute('data-level');
      if (l >= need) {
        a.classList.remove('bloqueado-link');
      } else {
        a.classList.add('bloqueado-link');
        if (a.tagName === 'A') a.href = nextStepPage() + '?locked=' + need;
      }
    });
    // Show a banner on step pages when the visitor was redirected from a locked page.
    var locked = new URLSearchParams(location.search).get('locked');
    var banner = document.getElementById('aviso-bloqueo');
    if (locked && banner) {
      var names = { 3: 'Destinations & itineraries', 4: 'Prices & resources' };
      banner.textContent = '🔒 "' + (names[locked] || 'That page') +
        '" is locked. Complete this step to keep unlocking content.';
      banner.classList.remove('oculto');
    }
    var reset = document.getElementById('reiniciar');
    if (reset) reset.addEventListener('click', function (e) {
      e.preventDefault();
      if (confirm('Delete your progress on this browser and start over?')) {
        try { localStorage.removeItem(KEY); } catch (err) {}
        try { sessionStorage.removeItem(KEY); } catch (err) {}
        location.href = base() + 'index.html';
      }
    });
  }

  /*
   * Shared form validation: shows the message under each field.
   * `extra(field)` can return a custom error string.
   */
  function validator(form, extra) {
    function check(field) {
      var v = field.validity, m = '';
      if (field.type === 'radio' || field.type === 'checkbox') {
        // Groups are validated through the member that has the `required` attribute.
        var group = form.querySelectorAll('[name="' + field.name + '"]');
        var lead = form.querySelector('[name="' + field.name + '"][required]');
        if (!lead) return true;
        if (field !== lead) return check(lead);
        var anyChecked = Array.prototype.some.call(group, function (x) { return x.checked; });
        if (!anyChecked) m = group.length > 1 ? 'Please choose an option.' : 'Please check this box to continue.';
      } else if (v.valueMissing || (field.required && !field.value.trim())) m = 'This field is required.';
      else if (v.typeMismatch) m = 'Enter a valid email, e.g. name@email.com.';
      else if (v.tooShort) m = 'Please write at least ' + field.minLength + ' characters.';
      else if (v.patternMismatch) m = field.getAttribute('data-pattern-msg') || 'Invalid format.';
      else if (v.rangeUnderflow || v.rangeOverflow || v.badInput) m = 'Please enter a valid value.';
      if (!m && extra) m = extra(field) || '';
      var box = field.closest('.campo');
      if (box && box.querySelector('.error')) box.querySelector('.error').textContent = m;
      if (field.type !== 'radio' && field.type !== 'checkbox') field.classList.toggle('invalido', !!m);
      return !m;
    }
    form.querySelectorAll('input, textarea, select').forEach(function (f) {
      f.addEventListener('blur', function () { if (f.type !== 'radio' && f.type !== 'checkbox') check(f); });
      f.addEventListener(f.tagName === 'SELECT' || f.type === 'radio' || f.type === 'checkbox' ? 'change' : 'input', function () {
        var err = f.closest('.campo') && f.closest('.campo').querySelector('.error');
        if (err && err.textContent) check(f);
      });
    });
    // Validate every field inside `scope` (the whole form or one wizard page).
    return function validateAll(scope) {
      var first = null;
      (scope || form).querySelectorAll('input, textarea, select').forEach(function (f) {
        if (f.closest('.oculto')) return;
        if (!check(f) && !first) first = f;
      });
      if (first) first.focus();
      return !first;
    };
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', decorate);
  else decorate();

  return {
    STEPS: STEPS,
    data: read,
    save: save,
    level: level,
    require: require,
    nextStepPage: nextStepPage,
    processing: processing,
    send: send,
    validator: validator
  };
})();
