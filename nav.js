/* Armodilo Tools: shared top bar + single sign-in. Loaded by every page in this site.

   To add or rename a tool, edit the TOOLS list below (and the list in the home index.html).
   `key` is the name the tool uses to remember its session; leave it out for tools with no sign-in.

   How sign-in works: the home page signs in to every tool's Apps Script at once and saves each
   session. A tool page that has no session (or whose session ended) sends the visitor back to
   the home page instead of showing its own sign-in form. Each Apps Script still checks the
   session before returning data, so the data stays protected on the server side. */
(function () {
  var TOOLS = [
    { dir: 'sales-revenue',      name: 'Sales Revenue',      key: 'armo_token' },
    { dir: 'monthly-summary',    name: 'Monthly Summary',    key: 'ms_token' },
    { dir: 'product-allocation', name: 'Product Allocation', key: 'kt_token' },
    { dir: 'ci-builder',         name: 'Invoice Builder' }
  ];
  var script = document.currentScript;
  var root = script ? script.src.replace(/nav\.js(\?.*)?$/, '') : '../';
  var here = location.pathname;
  var current = null;
  TOOLS.forEach(function (t) { if (here.indexOf('/' + t.dir + '/') !== -1) current = t; });

  function ss(k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } }
  function ssSet(k, v) { try { sessionStorage.setItem(k, v); } catch (e) {} }
  function toHome(next) { location.replace(root + (next ? '?next=' + encodeURIComponent(next) : '')); }

  var css =
    '#armoNav{display:flex;align-items:center;gap:4px;height:40px;padding:0 14px;background:#1b1d1c;color:#e9e7e1;' +
    'font:500 13px/1 Inter,system-ui,-apple-system,"Segoe UI",Arial,sans-serif;overflow-x:auto;white-space:nowrap;scrollbar-width:none}' +
    '#armoNav::-webkit-scrollbar{display:none}' +
    '#armoNav a{color:#b9b6ad;text-decoration:none;padding:7px 10px;border-radius:6px;flex:none}' +
    '#armoNav a:hover{color:#fff;background:#2b2e2c}' +
    '#armoNav a:focus-visible{outline:2px solid #7cc3ab;outline-offset:1px}' +
    '#armoNav a.home{color:#fff;font-weight:600;letter-spacing:.01em;margin-right:8px;padding-left:4px}' +
    '#armoNav a[aria-current="page"]{color:#fff;background:#2f3331}' +
    '@media print{#armoNav{display:none!important}}';

  /* ---- single sign-in: tool pages never show their own sign-in form ---- */
  if (current && current.key) {
    if (!ss(current.key)) { toHome(current.dir); return; }
    css += '#login{display:none!important}';
    var bounced = false;
    var check = function () {
      var el = document.getElementById('login');
      if (!el || bounced) return;
      if (!el.hidden && !el.classList.contains('hidden')) {   // the tool wants a sign-in
        bounced = true;
        toHome(ss('armo_signout') ? '' : current.dir);
      }
    };
    document.addEventListener('click', function (e) {          // Sign out in a tool = sign out everywhere
      if (e.target && e.target.closest && e.target.closest('#logout, #logoutBtn')) ssSet('armo_signout', '1');
    }, true);
    document.addEventListener('DOMContentLoaded', function () {
      check();
      var el = document.getElementById('login');
      if (el) new MutationObserver(check).observe(el, { attributes: true, attributeFilter: ['hidden', 'class'] });
    });
  }

  var style = document.createElement('style');
  style.textContent = css;
  document.head.appendChild(style);

  function build() {
    var nav = document.createElement('nav');
    nav.id = 'armoNav';
    nav.setAttribute('aria-label', 'Armodilo tools');
    var home = document.createElement('a');
    home.className = 'home';
    home.href = root;
    home.textContent = 'Armodilo Tools';
    nav.appendChild(home);
    TOOLS.forEach(function (t) {
      var a = document.createElement('a');
      a.href = root + t.dir + '/';
      a.textContent = t.name;
      if (t === current) a.setAttribute('aria-current', 'page');
      nav.appendChild(a);
    });
    document.body.insertBefore(nav, document.body.firstChild);
  }
  if (document.body) build(); else document.addEventListener('DOMContentLoaded', build);
})();
